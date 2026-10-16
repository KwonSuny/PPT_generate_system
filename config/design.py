"""固定已提交内容的设计候选；生成与采用分别保存为持久任务。"""
from copy import deepcopy
import json
import os
import re
from pathlib import Path

from agents.template import composition_guidance, get_layout_agent
from config.content import ContentConstraintError, bind_canvas, content_values
from config.icon_library import available_icons, normalize_icon_name, load_lucide_svg
from config.llm_client import resolve_model_name
from config.preview import render_preview, tool_signature
from config.schema import PPTFinalBlueprint
from config.storage import content_hash
from config.visual import rule_issues
from config.layout_budget import design_request_policy, font_floor


DIRECTIONS = [
    {"id": "clear", "name": "清晰汇报", "description": "突出结论，按内容分组，语义图标作为阅读锚点。", "instruction": "选择适合内容的列式解释或图文分区；结论醒目，正文左对齐，使用少量对应语义的本地图标。"},
    {"id": "visual", "name": "图解重点", "description": "扩大主要图解或数据区，用不对称布局组织解释。", "instruction": "图解、表格或关键结论形成主视觉；解释置于侧栏或下方，与清晰汇报候选采用不同构图。避免重复等大卡片；不要仅更换颜色。"},
]


class DesignBudgetExceeded(RuntimeError):
    pass


def icon_suggestions(slide):
    hints = [(r"种植|幼苗|培育|生长", "sprout"), (r"采摘|成熟|叶片", "leaf"),
             (r"烘焙|加热", "flame"), (r"温度|摄氏", "thermometer"), (r"水洗|浸泡|萃取|冲煮", "droplet"),
             (r"运输|配送|物流", "truck"), (r"饮用|咖啡|饮料", "coffee")]
    names = set(available_icons())
    result = []
    for ref, text in content_values(slide).items():
        if ref.startswith("full_content/"):
            name = next((name for pattern, name in hints if name in names and re.search(pattern, text)), None)
            if name:
                result.append({"content_ref": ref, "icon_name": name})
    return result


def arrange_text_and_icons(canvas, minimum_sizes=None, flow_text=False, slide=None, style_floors=None, bottom_diagram=False):
    canvas = canvas.model_copy(deep=True)
    paragraphs = [element for element in canvas.elements if element.type == "textbox"]
    for element in canvas.elements:
        if element.type in (style_floors or {}):
            style=getattr(element,element.type+"_style")
            for field,minimum in style_floors[element.type].items():setattr(style,field,max(getattr(style,field),minimum))
    original_positions = [(e.left,e.top,e.width,e.height) for e in paragraphs]
    for element in paragraphs:
        for paragraph in element.paragraphs or []:
            role_floor = min(18,(minimum_sizes or {}).get(paragraph.content_ref,18)) if (paragraph.content_ref or "").startswith("full_content/") else 0
            paragraph.font_size = max(paragraph.font_size, role_floor, (minimum_sizes or {}).get(paragraph.content_ref, 0))
    for icon in (element for element in canvas.elements if element.type == "icon"):
        icon.icon_name=normalize_icon_name(icon.icon_name)
        if slide and not icon.icon_color:
            icon.icon_color=slide.get("theme",{}).get("primary")
        if icon.content_ref:
            continue
        nearby = [element for element in paragraphs if len(element.paragraphs or []) == 1
                  and 0 <= element.left - icon.left - icon.width <= .6
                  and abs(element.top - icon.top) <= .5]
        nearby.sort(key=lambda element: abs(element.top - icon.top) + abs(element.left - icon.left - icon.width))
        if not nearby or (len(nearby) > 1 and abs(nearby[0].top - nearby[1].top) < .03 and abs(nearby[0].left - nearby[1].left) < .03):
            raise ContentConstraintError("图标未绑定且无法唯一对应附近文字")
        icon.content_ref = nearby[0].paragraphs[0].content_ref
    # 新候选允许重排；沿用列宽、行序、字号及图示区域，旧版与文本锁不进入此分支。
    from renderer import PPTXRenderer
    if flow_text or bottom_diagram:
        from config.schema import CanvasElement
        indexed = {p.content_ref:p for e in paragraphs for p in e.paragraphs or []}
        groups = [[ref] for ref in ("title","punchline") if ref in indexed]
        point_ids = slide.get("bullet_ids") or [f"point-{i+1}" for i in range(len(slide.get("full_content") or []))]
        groups.extend([[ref for ref in (f"bullet_points/{point}",f"full_content/{point}") if ref in indexed] for point in point_ids])
        used = {ref for group in groups for ref in group}
        groups.extend([[ref] for ref in indexed if ref not in used])
        elements=[];top=.35
        body_groups=[]
        for index, group in enumerate(group for group in groups if group):
            if bottom_diagram and group[0] not in {"title","punchline"}:
                body_groups.append(group)
                continue
            meta = group[0] in {"title","punchline"}
            element=CanvasElement(id=f"flow-{index}",type="textbox",left=.8 if meta else 1.35,top=top,width=11.7 if meta else 11.15,height=1,
                paragraphs=[indexed[ref].model_copy(deep=True) for ref in group])
            element.height = PPTXRenderer.required_text_height(element)+.01
            elements.append(element);top += element.height+.14
        if bottom_diagram:
            for start in range(0,len(body_groups),2):
                row=[]
                for column,group in enumerate(body_groups[start:start+2]):
                    element=CanvasElement(id=f"diagram-note-{start+column}",type="textbox",left=1.4 if column==0 else 7.5,top=top,width=5.35,height=1,
                        paragraphs=[indexed[ref].model_copy(deep=True) for ref in group])
                    element.height=PPTXRenderer.required_text_height(element)+.01
                    row.append(element);elements.append(element)
                top+=max(e.height for e in row)+.14
            component=next(e for e in canvas.elements if e.type=="diagram")
            component.left=.8;component.width=11.7;component.top=top+.02;component.height=7.35-component.top
            if component.height<.8:
                raise ContentConstraintError("图解下置无法在保持原文与字号时容纳流程图")
        elif top > 7.5:
            raise ContentConstraintError("单列阅读轴无法容纳原文字与字号，请调整构图")
        canvas.elements = [e for e in canvas.elements if e.type not in {"textbox","shape"}] + elements
        paragraphs=elements
    else:
        PPTXRenderer._reflow_text_regions(canvas, 7.5, max_move=7.5)
        if original_positions != [(e.left,e.top,e.width,e.height) for e in paragraphs]:
            # 没有内容锚点的卡片背景不能停留在旧位置遮蔽重新排列的分组。
            canvas.elements = [e for e in canvas.elements if not (e.type=="shape" and not e.content_ref and e.shape_type in {"rectangle","rounded_rectangle","ellipse"} and e.width>1 and e.height>.6)]
    for icon in (element for element in canvas.elements if element.type == "icon"):
        partner = next((element for element in paragraphs if any(p.content_ref == icon.content_ref for p in element.paragraphs or [])), None)
        if partner:
            icon.left = partner.left - icon.width - .2
            icon.top = partner.top + (partner.height - icon.height) / 2
    for icon in (element for element in canvas.elements if element.type=="icon"):
        partner=next((element for element in paragraphs if any(p.content_ref==icon.content_ref for p in element.paragraphs or [])),None)
        if not partner:
            continue
        blockers=[e for e in paragraphs if e is not partner and e.left<partner.left and e.left+e.width>icon.left and min(e.top+e.height,icon.top+icon.height)>max(e.top,icon.top)]
        shift=max([0]+[e.left+e.width+.2-icon.left for e in blockers])
        if shift:
            column=[e for e in paragraphs if abs(e.left-partner.left)<.03 and abs(e.width-partner.width)<.03]
            if any(e.left+e.width+shift>13.333-.15 for e in column):
                raise ContentConstraintError("图标挤入相邻列，当前右侧留白不足以安全整理")
            refs={p.content_ref for e in column for p in e.paragraphs or []}
            for e in column:e.left+=shift
            for e in canvas.elements:
                if e.type=="icon" and e.content_ref in refs:e.left+=shift
    return canvas


def validate_icons(canvas, slide):
    values = content_values(slide)
    anchors = {}
    for element in canvas.elements:
        for paragraph in element.paragraphs or []:
            if paragraph.content_ref:
                anchors.setdefault(paragraph.content_ref, []).append(element)
    for element in canvas.elements:
        if element.type != "icon":
            continue
        name = normalize_icon_name(element.icon_name)
        if name == "circle" or name not in set(available_icons()) or not load_lucide_svg(name):
            raise ContentConstraintError("图标需要有效本地名称与具体语义，不能采用 circle 或占位圆点")
        if element.content_ref not in values or element.content_ref not in anchors:
            raise ContentConstraintError("设计图标必须绑定本页文字 content_ref")
        partners = anchors[element.content_ref]
        partner = min(partners, key=lambda item: abs(item.top - element.top))
        gap = partner.left - element.left - element.width
        distance = abs(element.top + element.height / 2 - partner.top - partner.height / 2)
        if gap < 0.15 or gap > 0.6 or distance > 0.12:
            raise ContentConstraintError("图标与对应文本的中心或间距未对齐")


def validate_page(canvas, slide):
    validate_icons(canvas, slide)
    issues = rule_issues(canvas)
    if issues:
        raise ContentConstraintError("候选几何检查未通过："+"；".join(item["message"]+"("+",".join(item["element_ids"])+")" for item in issues))


def composition_signature(blueprint, pages):
    return content_hash([sorted((element.type, element.content_ref or "", round(element.left, 1), round(element.top, 1),
                                round(element.width, 1), round(element.height, 1),
                                tuple(sorted(p.content_ref or "" for p in element.paragraphs or [])))
                               for element in blueprint.canvases[number - 1].elements if element.type != "shape") for number in pages])


def validate_candidate(blueprint, baseline, outline, pages, allow_theme_change=False):
    if len(blueprint.canvases) != len(baseline.canvases) or len(outline["slides"]) != len(baseline.canvases):
        raise ContentConstraintError("候选页数与原版本不一致")
    for number, canvas in enumerate(blueprint.canvases, 1):
        if number not in pages:
            if canvas.model_dump() != baseline.canvases[number - 1].model_dump():
                raise ContentConstraintError("候选改变了未选择的页面")
            continue
        previous_text = {p.text for element in baseline.canvases[number - 1].elements if element.type == "textbox" for p in element.paragraphs or []}
        candidate_text = {p.text for element in canvas.elements if element.type == "textbox" for p in element.paragraphs or []}
        candidate_text.update(element.table.source for element in canvas.elements if element.type=='table' and element.table and element.table.source)
        if previous_text - candidate_text:
            raise ContentConstraintError("候选遗漏了原版已显示的文字，不能只保留正文而删除标题或要点")
        source = deepcopy(outline["slides"][number - 1])
        if allow_theme_change:
            source['theme']={**source['theme'],'background':canvas.background_color}
        bound = bind_canvas(canvas, source, preserve_component_style=allow_theme_change)
        if bound.model_dump() != canvas.model_dump():
            raise ContentConstraintError("候选没有忠实保留文字、数据或 SVG")
        validate_icons(canvas, outline["slides"][number - 1])
        issues = rule_issues(canvas)
        if issues:
            raise ContentConstraintError("候选规则检查未通过：" + "；".join(dict.fromkeys(item["message"] for item in issues)))


def build_candidates(state, store, job, payload, run_dir, api_key):
    from config.jobs import atomic_json, StageOutcomeUnknown
    from main import _layout_page, _render_blueprint
    if not api_key:
        raise ValueError("请先连接模型服务，设计候选不能用离线演示替代")
    outline = deepcopy(state.get("outline"))
    baseline = PPTFinalBlueprint.model_validate(state.get("blueprint"))
    pages = payload["page_numbers"]
    if not pages or len(pages) > 3 or any(type(p) is not int or p < 1 or p > len(baseline.canvases) for p in pages):
        raise ValueError("设计候选需要选择 1–3 个有效页面")
    signature = tool_signature()
    max_calls = min(18, max(1, int(os.environ.get("PPT_DESIGN_MAX_LAYOUT_CALLS", "12"))))
    used = sum(len(json.loads(stage["usage_json"] or "[]")) or int(stage["status"] in {"running", "unknown", "succeeded"})
               for stage in store.stages(job["id"]) if stage["stage_key"].startswith("layout/candidate-"))
    saved_attempts = json.loads(job["snapshot_json"]).get("data", {}).get("design_call_attempts", 0)
    used = max(used, int(saved_attempts))
    designer = get_layout_agent(api_key=api_key, model_name=resolve_model_name("layout"))
    def limited_designer(current_state, request):
        nonlocal used
        if used >= max_calls:
            raise DesignBudgetExceeded("本次候选排版请求预算已用完")
        used += 1
        current_state.set("design_call_attempts", used)
        return designer(current_state, request)
    result = {"base_version_id": job["base_version_id"], "pages": pages, "outline_hash": content_hash(outline),
              "human_review": "not_reviewed", "max_layout_calls": max_calls, "candidates": []}
    original_file = state.get("last_file")
    policy=design_request_policy(payload.get('design_request',''))
    result['layout_policy']=policy
    baseline_preview = state.execute_stage("preview/design-baseline", {"filename": original_file, "runtime": signature},
        lambda: render_preview(original_file, Path(run_dir) / "generated/design/baseline", signature=signature), external=False)
    if len(baseline_preview["pages"]) != len(baseline.canvases):
        raise ContentConstraintError("原版对照截图页数不一致")
    result["baseline_assets"] = [{**store.describe_asset(path, "preview"), "page": number} for number, path in enumerate(baseline_preview["pages"], 1)]
    state.set("design_candidates", result)
    signatures = {composition_signature(baseline, pages)}
    for index, direction in enumerate(DIRECTIONS):
        candidate = {**direction, "status": "building", "assets": [], "reason": "", "human_review": "not_reviewed"}
        result["candidates"].append(candidate)
        state.set("design_candidates", result)
        folder = Path(run_dir) / "generated/design" / direction["id"]
        folder.mkdir(parents=True, exist_ok=True)
        try:
            blueprint = baseline.model_copy(deep=True)
            for number in pages:
                slide = deepcopy(outline["slides"][number - 1])
                choices = composition_guidance(slide)["candidates"]
                request = {"schema_version": 2, "topic": outline.get("topic", state.topic), "slides": [slide],
                           "design_direction": {**direction, "preferred_composition": choices[min(index, len(choices) - 1)],
                                                "icon_suggestions": icon_suggestions(slide), "user_request": payload.get("design_request", "")}}
                request['layout_policy']=policy
                request['required_content_refs']=sorted({p.content_ref for e in baseline.canvases[number-1].elements for p in e.paragraphs or [] if p.content_ref})
                floors = {p.content_ref:p.font_size for element in baseline.canvases[number-1].elements for p in element.paragraphs or [] if p.content_ref}
                if policy['allow_font_resize']:
                    floors={ref:font_floor(ref,payload.get('design_request','')) for ref in floors}
                request['layout_feedback']={'text_font_size_floors':floors}
                style_floors = {element.type:{field:value for field,value in getattr(element,element.type+"_style").model_dump().items() if "font_size" in field}
                                for element in baseline.canvases[number-1].elements if element.type in {"table","diagram"}}
                if policy['allow_font_resize']:
                    style_floors={kind:{field:min(value,12 if field=='source_font_size' else 14 if field=='label_font_size' else 18)
                        for field,value in values.items()} for kind,values in style_floors.items()}
                free_flow = not re.search(r"双栏|两栏|二列|三列|两列|三栏|分栏|(?:不要|禁止).*单列|(?:上方|下方|左侧|右侧).*(?:图|表)", state.requirements+payload.get("design_request", ""))
                pure_text = not any(slide.get(key) for key in ("svg_code","image_path","table","diagram")) and slide.get("layout_type")!="title_page"
                blueprint.canvases[number - 1] = _layout_page(state, limited_designer, request, number, str(folder),
                                                            stage_prefix=f"layout/candidate-{direction['id']}", canvas_validator=lambda canvas: validate_page(canvas, slide),
                                                            canvas_transform=lambda canvas: arrange_text_and_icons(canvas, floors, flow_text=index==1 and pure_text and free_flow, slide=slide, style_floors=style_floors,
                                                            bottom_diagram=index==1 and free_flow and bool(slide.get("diagram")) and not any(slide.get(key) for key in ("svg_code","image_path","table"))))
            validate_candidate(blueprint, baseline, outline, pages, policy['allow_theme_change'])
            fingerprint = composition_signature(blueprint, pages)
            if fingerprint in signatures:
                raise ContentConstraintError("候选与原版或已有候选的构图相同，请调整设计要求后再生成")
            output = folder / (direction["id"] + ".pptx")
            state.execute_stage(f"render/design-{direction['id']}", {"blueprint": blueprint.model_dump(), "filename": str(output)},
                                lambda: _render_blueprint(blueprint, str(output)), external=False)
            validate_candidate(blueprint, baseline, outline, pages, policy['allow_theme_change'])
            preview = state.execute_stage(f"preview/design-{direction['id']}", {"filename": str(output), "runtime": signature},
                                           lambda: render_preview(output, folder / "previews", signature=signature), external=False)
            if len(preview["pages"]) != len(baseline.canvases):
                raise ContentConstraintError("候选截图页数不一致")
            blue_path = folder / "blueprint.json"; atomic_json(blue_path, blueprint.model_dump())
            candidate["assets"] = [store.describe_asset(output, "pptx"), store.describe_asset(blue_path, "blueprint"),
                                   store.describe_asset(preview["pdf"], "pdf")]
            candidate["assets"].extend({**store.describe_asset(path, "preview"), "page": page} for page, path in enumerate(preview["pages"], 1))
            candidate.update(status="ready", composition_hash=fingerprint, rules="passed", visual_review="not_run",
                             icon_count=sum(element.type == "icon" for number in pages for element in blueprint.canvases[number - 1].elements))
            signatures.add(fingerprint)
        except Exception as exc:
            unknown = isinstance(exc, StageOutcomeUnknown) or any(stage["status"] == "unknown" for stage in store.stages(job["id"]))
            candidate.update(status="unknown" if unknown else "rejected", reason=str(exc) if isinstance(exc, (ContentConstraintError, DesignBudgetExceeded)) else type(exc).__name__)
            state.set("design_candidates", result)
            atomic_json(Path(run_dir) / "generated/design/candidates.json", result)
            if unknown:
                raise StageOutcomeUnknown("候选调用结果未知，已保存其它候选；确认后才可重试") from exc
        result["calls_used"] = used
        state.set("design_call_attempts", used)
        state.set("design_candidates", result)
        atomic_json(Path(run_dir) / "generated/design/candidates.json", result)
    if not any(item['status']=='ready' for item in result['candidates']):
        raise ContentConstraintError('未产出可用设计候选；已保存各方向的具体失败原因，可调整要求后重试')
    return result


def prepare_adoption(state, store, job, payload):
    source = store.job(payload["source_job_id"], job["owner_id"])
    if (not source or source["conversation_id"] != job["conversation_id"] or source["base_version_id"] != job["base_version_id"]
            or json.loads(source["request_json"]).get("kind") != "design_candidates"):
        raise ContentConstraintError("候选不属于当前文稿基准版本")
    result = store.rebase_snapshot(json.loads(source["snapshot_json"])).get("data", {}).get("design_candidates", {})
    candidate = next((item for item in result.get("candidates", []) if item["id"] == payload["candidate_id"]), None)
    if not candidate or candidate["status"] != "ready" or not store.verify_assets(candidate["assets"]):
        raise ContentConstraintError("候选尚未就绪或文件被修改，不能采用")
    outline = state.get("outline")
    if result.get("outline_hash") != content_hash(outline):
        raise ContentConstraintError("候选的原始语义内容与当前基准不一致")
    blue_asset = next(item for item in candidate["assets"] if item["role"] == "blueprint")
    if blue_asset["sha256"] != payload["candidate_hash"]:
        raise ContentConstraintError("候选已更新，不能采用未经重新预览的文件")
    blueprint = PPTFinalBlueprint.model_validate(json.loads(store.asset_path(blue_asset["path"]).read_text(encoding="utf8")))
    baseline = PPTFinalBlueprint.model_validate(state.get("blueprint"))
    theme_change=result.get('layout_policy',{}).get('allow_theme_change',False)
    validate_candidate(blueprint, baseline, outline, result["pages"], theme_change)
    if theme_change:
        outline=deepcopy(outline)
        for number in result['pages']:
            canvas=blueprint.canvases[number-1]
            theme=outline['slides'][number-1]['theme']
            theme['background']=canvas.background_color
            title=next((p for e in canvas.elements for p in e.paragraphs or [] if p.content_ref=='title'),None)
            body=next((p for e in canvas.elements for p in e.paragraphs or [] if (p.content_ref or '').startswith('full_content/')),None)
            if title:theme['primary']=title.font_color
            if body:theme['content']=body.font_color
            native=next((e for e in canvas.elements if e.type in {'table','diagram'}),None)
            if native:
                style=getattr(native,native.type+'_style')
                theme['content']=style.font_color
                theme['primary']=style.header_fill if native.type=='table' else style.line_color
        state.set('outline',outline)
    state.set_operation("continue", scope="design", pages=result["pages"])
    state.set("requested_mode", "modify_design")
    state.set("blueprint", blueprint.model_dump())
    state.set("layout_fallbacks", [item for item in state.get("layout_fallbacks", []) if item["page"] not in result["pages"]])
    state.set("preview_assets", {"pdf": str(store.asset_path(next(item for item in candidate["assets"] if item["role"] == "pdf")["path"])),
                                "pages": [str(store.asset_path(item["path"])) for item in sorted((item for item in candidate["assets"] if item["role"] == "preview"), key=lambda item: item["page"])]})
    state.set("visual_review", {"status": "not_run", "human_review": "not_reviewed", "pages": [], "reason": "用户选择设计候选；规则检查通过，尚未执行视觉模型审核"})
    state.set("visual_review_path", None); state.set("visual_candidates", [])
    state.set("design_selection", {"source_job_id": source["id"], "candidate_id": candidate["id"], "pages": result["pages"], "selected_by_user": True})
    state.set("design_candidates", None)
    output = str(store.asset_path(next(item for item in candidate["assets"] if item["role"] == "pptx")["path"]))
    state.set("last_file", output); state.set("result_kind", "agent_generated")
    return output
