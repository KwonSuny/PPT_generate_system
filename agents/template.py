import json
from typing import Optional
from config.schema import PPTFinalBlueprint
from config.content import ContentConstraintError, bind_canvas, content_values
from config.sanitizers import normalize_blueprint_payload, parse_completion_payload, FONT_SCALE_STEPS
from config.state import PPTGenerationState
from config.llm_client import create_client, chat_completion_with_retry, resolve_model_name, model_request_options


def _strip_internal_fields(payload):
    if not isinstance(payload, dict):
        return payload
    cleaned = {key: value for key, value in payload.items() if key not in {"layout_feedback", "capacity_budget"}}
    cleaned["slides"] = [{k: v for k, v in slide.items() if k != "image_prompt"} for slide in payload.get("slides", [])]
    return cleaned


def select_layout_canvases(payload, slides, state=None):
    """页 ID 明确时按请求范围选取，不能把整套生成指令扩成额外页面。"""
    canvases = payload.get("canvases", [])
    expected_ids = [slide.get("id") for slide in slides]
    returned_ids = [canvas.get("slide_id") for canvas in canvases]
    if expected_ids and all(expected_ids) and len(set(expected_ids)) == len(expected_ids):
        if returned_ids and all(isinstance(value, str) and value for value in returned_ids):
            if len(set(returned_ids)) != len(returned_ids):
                raise ContentConstraintError("Layout 返回重复页 ID，无法安全对应请求页面")
            indexed = dict(zip(returned_ids, canvases))
            if all(value in indexed for value in expected_ids):
                if len(canvases) > len(slides) and state:
                    state.append_progress(f"Layout: 按页 ID 匹配本次 {len(slides)} 页，忽略 {len(canvases) - len(slides)} 张范围外画布")
                return {**payload, "canvases": [indexed[value] for value in expected_ids]}
            raise ContentConstraintError("Layout 未返回请求的页 ID：" + "、".join(value for value in expected_ids if value not in indexed))
        if any(returned_ids):
            raise ContentConstraintError("Layout 返回的页 ID 不完整，无法安全对应请求页面")
    if len(canvases) != len(slides):
        raise ContentConstraintError(f"Layout 本次输入 {len(slides)} 页、返回 {len(canvases)} 张画布，且缺少可唯一匹配的页 ID")
    return payload


def composition_guidance(slide):
    """根据已有语义选择构图候选，不预填文案、不绑定固定坐标。"""
    kind = slide.get("layout_type")
    points = slide.get("full_content") or []
    count = len(points)
    has_art = bool(slide.get("svg_code") or slide.get("image_path"))
    if kind == "title_page":
        candidates = ["标题与结论居中留白"]
        if has_art:
            candidates = ["左侧标题与结论、右侧独立主题图", "上方标题、下方横向主题图"]
    elif slide.get("table"):
        candidates = ["标题与结论之下放通栏表格", "表格为主、独立图解为辅的左右构图"] if has_art else ["标题与结论之下放通栏表格"]
    elif slide.get("diagram"):
        candidates = ["上方关系图、下方对应的解释列", "左侧关系图、右侧短说明"]
        if count > 3:
            candidates.insert(0, "上方关系图、下方二列解释")
    elif has_art:
        candidates = ["左侧分层文案、右侧独立图示", "上方宽幅图示、下方对应解释列"]
    elif count >= 3:
        candidates = ["二列分层要点", "三列短说明" if count == 3 else "二列二行分层要点"]
    else:
        candidates = ["一条清晰阅读轴上的分层说明", "留白充足的左右对照"]
    table = slide.get("table") or {}
    return {"candidates": candidates, "text_blocks": count,
            "table_rows": len(table.get("rows") or []), "table_columns": len(table.get("columns") or []),
            "art_is_separate": has_art,
            "selection_rule": "按内容长度、图示比例和用户风格选一种，必要时自由调整；候选不是必须填满的固定模板。"}


def get_layout_agent(api_key: str, model_name: Optional[str] = None):
    model_name = model_name or resolve_model_name("layout")
    client = create_client(api_key)

    def design_layout(state: PPTGenerationState, writer_output_json: dict) -> PPTFinalBlueprint:
        from config.icon_library import available_icons
        lean = _strip_internal_fields(writer_output_json)
        for slide in lean.get("slides", []):
            slide["content_catalog"] = content_values(slide)
            slide["composition_guidance"] = composition_guidance(slide)
        system_prompt = """
<role>
你是演示文稿艺术指导与信息排版师。你负责把内容转成清晰、有节奏、有主题辨识度的页面，而不是逐条往卡片里填文字。
</role>
<task>
你将语义内容排成 16:9 演示文稿，画布 13.333×7.5 英寸。只输出 JSON：schema_version=2、canvases。
每个画布有 slide_id、background_color、elements。输入一页时输出严格一个画布。
尊重用户风格，其次根据主题用途判断，未明确时采用简洁汇报。使用平面构图、清晰标题和信息层次，无须反复使用卡片或渐变。
存在 design_direction 时执行其构图方向与用户要求，只改变本次目标页，不改语义。layout_policy.allow_theme_change=true 时可改变背景和字体配色，否则沿用主题色族。用户明确要求优先于原版和邻页，主题统一不等于每页背景完全相同。
</task>
<creative_direction>
根据语义选择有主次的视觉构图：大字结论、侧边色面、分层流程、非等宽对照、宽幅时间轴或图文分区。背景、标题和重点文字共同服务叙事；可用主题深色底、浅色底、局部强调色面与克制的几何形状，不能把每页都做成白底蓝色标题加三列长正文。
选择一个主要视觉动作，周围内容保持安静。允许不对称、明显留白与不同字号层级；不要无意义装饰、全页纹理或把图示缩成角落小图。
深底使用浅色正文及清晰强调色，浅底使用深色正文；白底上的深色面板需要浅色面板文字，不能只看整页背景颜色。SVG内部颜色锁定时，可放到独立浅色承托区域。
标题常用28–40pt，结论18–24pt，正文18–24pt；可在提供的字号下限内缩放并重新测算高度。font_family可选择 Microsoft YaHei、DengXian、SimHei、SimSun、Segoe UI；同页最多两种字体，按主题选择而非随意混用。
同一页一个内容引用显示一次即可；不要用侧边大字再重复结论。必须展示 required_content_refs 中的原版独有文字引用。
</creative_direction>
<geometry_and_capacity>
所有 icon 必须指定本页文字 content_ref，且图标位于对应文本左侧，中心对齐、间距 0.2in，尺寸 0.3–0.5in。优先使用 icon_suggestions；不要用 circle 或编造名称代替图标。不要给同一段落重复配图标。
capacity_budget 在首次请求前按原文计算文字高度、图标留白与可行构图。先读取预算再分配空间；参考坐标可以重设计，但必须保留所需高度。列间有图标时至少预留图标宽度+0.2in文字间距+0.28in相邻板块间距，不能让下一列图标挤入上一列的文字。
先根据 composition_guidance 选择阅读顺序与主视觉，再安排容器。大标题、结论、说明有明确层级；整套保留统一标题基线、正文左沿与主题色族，但按语义切换构图。
参考原则：上图下说明时图中模块与下方列对应；左右图文时图示独占区域；对比信息成列且同级标题对齐。四条解释可采用二列二行；若预算提示这会压缩主图高度，使用四列短说明或切换分区，避免重复标题加长段落形成满页列表。
图标是信息锚点：只在能表达对应语义时使用，贴齐对应标题，禁止为每条要点堆放无意义圆圈。主题图优先保持完整比例和留白，不加全页装饰背景。
图元类型：textbox、shape、icon、image、svg、table、diagram。每个图元包含唯一 id 和 left/top/width/height（英寸）。
textbox.paragraphs 每段包含 content_ref、text、font_size、font_color、bold、alignment、space_after。
content_ref 必须逐字取自 content_catalog，例如 title、punchline、full_content/point-1；程序用引用回填原文，不能改写内容或省略数字。
正文优先 18/20pt，标题 32/40pt，注释 12/14pt。容器按实际中文字数和换行计算容量；允许缩放时须遵守字号下限并重新检查换行，不得缩成不可读的小字。
文本框上下内边距合计 0.16in，左右合计 0.22in；行高为字号×1.22/72in。单行 20pt 文本框至少 0.51in，18pt 至少 0.48in；换行及段间距另外计入。不要把要点框设为 0.4in 后紧贴正文框。
layout_feedback 是上一轮程序校验结果；按 required_text_heights 和字号下限重新安排整个本页的空间。先重排、再在 layout_policy 允许范围内缩放字号，正文默认至少18pt；不得新增页、删除引用或把内容改成图片。不能把首次猜测的过大字号当成永远不可降低的要求。
layout_feedback.violations 包含全部可定位错误、元素ID、坐标、缺额；逐项消除并检查关联组，不只修改第一个报错。repeated_geometry=true 时必须更换分区或构图，不能再次提交同一组坐标。程序会重新测量并再次验证；两轮后仍失败才使用稳定布局。
封面简洁，只放标题和副标题。其他页保留所有 full_content 和 punchline；表格/图示可与必要解释共同排版，不能遮挡。
原生表格只规划 type=table、content_ref=table 的容器和 table_style（字号、配色）。不要重述 table 数据、改成 SVG、拆散单元格或截断行列。列宽、行高由代码计算。表格来源由代码在 table 容器底部绘制，禁止再添加重复的 table/source 文本框。
table_style 的准确字段为 font_size、header_font_size、source_font_size、header_fill、header_color、font_color、body_fill、alternate_fill、border_color。颜色均为十六进制，不能使用 header_bg 等其它名称；背景/文字主题修改时同步更新表格，未指定时跟随当前页主题。表头字色对照 header_fill，单元格字色同时对照 body_fill 与 alternate_fill。
关系图只规划 type=diagram、content_ref=diagram 的容器和 diagram_style（font_size、label_font_size、fill_color、font_color、line_color）。节点和连线几何由代码计算；不要另画独立箭头和节点。
SVG 必须有一个 svg 容器，按 viewBox 保比例 contain，禁止拉伸。可以在图示下方放有引用的说明，但应留足边界与留白。
程序只裁剪透明外边距，SVG源码由上游锁定。参考预算的 paint_size 与字体单位，令图内最小标签尽量达到10pt；不能仅因容器很宽就认为图足够大，高度仍会限制 contain 比例。若标签先天密集，优先扩大图示或换位，不删图内文字。
图片使用输入 image_path，在独立插图区保持原始比例；不作为全文背景，不编造图注。
shape.shape_type 仅 rectangle、rounded_rectangle、ellipse、line、arrow。关系图和流程优先原生 diagram。
图标和文字按网格对齐，关联元素留 0.2in 间距。所有元素落在画布内，表格/图示与标题正文间至少 0.2in。
</geometry_and_capacity>
<output_check>
整套保持字体层级与色族联系，允许依据页面语义切换深浅背景和构图。必须保留全部证据、来源及已生成图片，不得凭空添加未引用的正文。
输出前核对：所有内容引用、文字所需高度、图标邻区留白、图示比例、实际文字衬底的对比度、装饰线端点。细线也须在画布内；不能以越界线段当作无关错误。
</output_check>
"""
        system_prompt += "\n字号刻度：" + str(FONT_SCALE_STEPS) + "\n图标白名单：" + ", ".join(available_icons())
        request = {"task": "只排版本次 slides；整套要求仅供理解风格，不能扩大本次页数或新增页面",
                   "expected_slide_ids": [slide.get("id") for slide in lean.get("slides", [])],
                   "expected_canvas_count": len(lean.get("slides", [])),
                   "style_context": state.requirements, "layout_feedback": writer_output_json.get("layout_feedback"), "content": lean}
        request['capacity_budget'] = writer_output_json.get('capacity_budget')
        request['layout_policy'] = writer_output_json.get('layout_policy',{})
        request['required_content_refs'] = writer_output_json.get('required_content_refs',[])
        if request['layout_feedback']:
            request['layout_feedback'] = {k:v for k,v in request['layout_feedback'].items() if k!='capacity_budget'}
        user = json.dumps(request, ensure_ascii=False)
        state.append_progress("Layout: 生成物理蓝图")
        response = chat_completion_with_retry(
            client, model=model_name,
            messages=[{"role": "system", "content": system_prompt}, {"role": "user", "content": user}],
            response_format={"type": "json_object"},
            **model_request_options('layout', model_name),
        )
        parsed = normalize_blueprint_payload(parse_completion_payload(response, state, "Layout"))
        slides = writer_output_json.get("slides", [])
        parsed = select_layout_canvases(parsed, slides, state)
        blueprint = PPTFinalBlueprint.model_validate(parsed)
        if len(blueprint.canvases) != len(slides):
            raise ContentConstraintError("Layout 返回的画布数量与输入页数不一致")
        blueprint.schema_version = 2
        blueprint.canvases = [bind_canvas(canvas, slide, preserve_component_style=writer_output_json.get('layout_policy',{}).get('allow_theme_change',False)) for canvas, slide in zip(blueprint.canvases, slides)]
        for canvas in blueprint.canvases:
            canvas.layout_resolved = False
        return blueprint

    return design_layout
