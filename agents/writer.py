import json
from typing import Optional
from pydantic import ValidationError
from config.content import ContentConstraintError
from config.schema import PPTOutline
from config.sanitizers import normalize_outline_payload, parse_completion_payload, model_to_dict
from config.state import PPTGenerationState
from config.llm_client import create_client, chat_completion_with_retry, resolve_model_name, model_request_options


def get_writer_agent(api_key: str, model_name: Optional[str] = None):
    model_name = model_name or resolve_model_name("writer")
    client = create_client(api_key)

    def write_ppt_content(state: PPTGenerationState, planner_output_json) -> PPTOutline:
        skeleton = model_to_dict(planner_output_json)
        system_prompt = """
<role>
你是演示文稿内容编辑与图解设计师，擅长把真实资料转化为有判断、有证据、可放进页面的文案与图示。
</role>
<task>
你负责补全文案和结构化证据，输出一个合法 JSON 对象：schema_version=2、topic、slides。
严格保留页数、顺序、页面 id、要点 bullet_ids，以及未授权的字段。文案面向观众，不写排版说明或泛泛口号。
新生成时复制 title、punchline、bullet_points、layout_type、theme；补充 full_content、icons 和所需 table/diagram/svg_code。
full_content 与要点对应，解释原因、过程、条件或证据，不重复标题凑字。普通页每条通常25–45字；有图示/表格且3条以上解释时优先15–28字，留出至少半页主视觉空间。以真实事实为先，不因字数建议删除已有数据或改变事实。结构化页面若无需外部解释可留 []，不得加入没有来源的数字或伪造引用。
预算来自同页内容总量：标题、结论、要点标题、正文、图示都会占空间。避免每页同时写长结论、三条长标题和三段长正文。
用户明确要求内联 SVG 表达时，生成与主题有关的 SVG 矢量插图或总览，写入适合页面的 svg_code，可与原生 table/diagram 共存；不得因为简单流程优先 diagram 就忽略 SVG 要求。仅允许 SVG 的指令不等于强制每页都使用 SVG。
用户明确风格优先，其次识别主题用途，未明确时采用汇报风格。无须图片时 image_prompt=null；不要为了装饰耗用生图额度。
图片用于独立插图区，不承载需要准确阅读的数据、文字或流程，image_path 由系统填充。
</task>
<native_components>
表格使用 table，格式为：
{"id":"table","columns":[{"id":"item","title":"项目"},{"id":"amount","title":"数值","format":{"style":"plain","decimals":2,"prefix":"","suffix":""}}],"rows":[{"id":"row-1","cells":[{"value":"原始项目"},{"value":12.5}]}],"source":"用户提供数据"}
列、行 id 必须唯一；每行 cells 与列数一致。value 保留原始字符串或数值，禁止截断行列、改写精度、符号和单位。
format.style 可为 plain、decimal、integer、percent；只有明确百分比时使用 percent（0.31 显示 31%）。数字字符串保持原样。
若输入已有 table 的数据，不得改写数值或删除行列；缺少来源时 source=null，不能编造出处。
简单关系图使用 diagram，格式为：
{"id":"diagram","direction":"LR","nodes":[{"id":"input","label":"输入","shape_type":"rectangle"},{"id":"output","label":"输出","shape_type":"rectangle"}],"edges":[{"id":"edge-1","source":"input","target":"output","source_port":"auto","target_port":"auto","label":""}]}
direction 为 LR 或 TB；节点与边 id 分别唯一。仅用于无环、连接相邻层的简单图；环路和复杂跨层拓扑使用 svg。
shape_type 仅支持 rectangle、rounded_rectangle；圆柱、菱形等复杂图形使用 SVG。采用 SVG 表达完整流程时 diagram=null，不要输出 nodes=[] 的占位图。
标签简洁；端口 auto 或与方向一致（LR 为 right/left，TB 为 bottom/top）。
</native_components>
<visual_design>
根据当前主题色族设计 SVG 的线条、节点、强调与留白。采用不同视觉语法表达不同信息：时间线强调顺序，对照图强调差异，流程图强调关系；不为所有主题重复画等宽方框。
SVG 用于复杂构图或用户明确要求的矢量表达：有效 xmlns="http://www.w3.org/2000/svg" 和正值 viewBox；宽高与构图一致，留足文字与连线间距。
SVG最终通常只占半页到三分之二页。不要把小图置于800×400的大量空白中；viewBox应贴合完整绘制范围并保留少量安全边距。以800单位宽的图为例，重要标签优先24–32单位而非10–14，节点也随字增加，避免标签密集到无法在正文至少18pt的同一页容纳。
SVG 不强制渐变/阴影/曲线；字体使用 Microsoft YaHei、SimHei；XML 特殊字符转义；不使用脚本、foreignObject 或外部资源。
</visual_design>
<output_check>
仅输出本次输入的 slides，整套历史只用于理解，不得把其它页一起返回。检查要点与正文一一对应、图示可读、数据精度、字段与页数完整。只输出 JSON，不输出代码围栏。
</output_check>
"""
        if state.operation_mode == "incremental_page":
            scope = state.modify_scope
            mode_hint = {
                "text": "仅修改文本和表格单元格值/列标题/来源、节点与边标签。保持要点与段落数量、行列/节点/边的数量和 id 顺序。保留所有视觉、格式、关系、配图和 SVG 字段。内容长度应接近原文，以适合已固定的容器。",
                "design": "仅修改 theme、layout_type、icons、image_prompt 等设计字段。本轮风格要求优先于旧版及邻页；要求更改背景/文字色时同步修改 theme.background/content/primary，不得只改说明而仍交回旧主题。允许同色族的深浅背景变体。所有 title、punchline、bullet_points、full_content、table 数据、diagram 关系/标签和 SVG 源码必须完全复制输入，不得改写任何数值。SVG 内部修改需使用页面修改。",
                "both": "按照用户要求修改目标页的内容和设计。保留页面 id；要点数量不变时保留 bullet_ids，数量改变时分配不重复的 id。",
            }.get(scope, "")
            lock_scope = scope if scope in {"text", "design"} else None
            mode_hint += "输入只有一页，输出严格一页。"
        else:
            lock_scope = "structure"
            mode_hint = "补全输入骨架，严格保留页数、标题、要点、版式、配色和稳定 id。"
        state.append_progress("Writer: 补全文案与结构化内容")
        prompt = mode_hint + "\n" + json.dumps(skeleton, ensure_ascii=False)
        for attempt in range(2):
            response = chat_completion_with_retry(
                client, model=model_name,
                messages=state.build_messages(system_prompt, prompt),
                response_format={"type": "json_object"},
                **model_request_options('writer', model_name, search=True),
            )
            data = parse_completion_payload(response, state, "Writer")
            expected = [slide.get('id') for slide in skeleton.get('slides', [])]
            returned = data.get('slides') if isinstance(data, dict) else None
            if returned and all(isinstance(s, dict) and s.get('id') for s in returned) and all(expected):
                ids = [s['id'] for s in returned]
                if len(ids)!=len(set(ids)):
                    raise ContentConstraintError('Writer 返回重复页面ID，页数与修改范围无法安全对应')
                indexed = dict(zip(ids, returned))
                if all(identity in indexed for identity in expected):
                    if len(returned)!=len(expected):
                        state.append_progress('Writer: 按稳定页面ID选取本次目标，忽略范围外页面')
                    data = {**data, 'slides':[indexed[identity] for identity in expected]}
            if not isinstance(data, dict) or not isinstance(data.get("slides"), list) or len(data["slides"]) != len(skeleton.get("slides", [])) or any(not isinstance(slide, dict) for slide in data["slides"]):
                raise ContentConstraintError("Writer 输出页数或页面结构不符合输入；已停止修改")
            try:
                outline = PPTOutline.model_validate(normalize_outline_payload(data, skeleton=skeleton, lock_scope=lock_scope))
            except (ValidationError, ContentConstraintError) as exc:
                if attempt or lock_scope in {"text", "design"}:
                    raise
                state.append_progress("Writer: 完整响应未通过结构校验，执行一次格式修复")
                prompt = (mode_hint + "\n原始骨架：" + json.dumps(skeleton, ensure_ascii=False) +
                          "\n待修复的完整内容：" + json.dumps(data, ensure_ascii=False) +
                          "\n校验错误：" + str(exc) +
                          "\n只修复字段格式，保留全部文字、数据、节点标签与连接关系，不删除页面。"
                          "若关系图含复杂图形或拓扑，将完整关系图转成 SVG，diagram=null；不得删减节点或连接以通过校验。"
                          "输出完整 JSON。")
                continue
            state.set("outline", outline.model_dump())
            return outline

    return write_ppt_content
