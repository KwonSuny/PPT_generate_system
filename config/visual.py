"""持久预览、视觉审核与最多两轮几何修复。"""

import hashlib
import json
import os
from pathlib import Path

from agents.visual_review import VisualPageReview, review_page
from config.content import ContentConstraintError
from config.llm_client import resolve_model_name
from config.preview import PreviewError, render_preview, tool_signature
from config.schema import PPTFinalBlueprint, SlideCanvas
from renderer import PPTXRenderer


class VisionBudgetExceeded(RuntimeError):
    pass


def rule_issues(canvas, include_quality=False):
    from config.layout_budget import layout_issues
    return layout_issues(canvas, include_svg_readability=include_quality)


def _protected(canvas):
    payload = canvas.model_dump()
    payload.pop("layout_resolved", None)
    for element in payload["elements"]:
        for key in ("left", "top", "width", "height", "column_widths", "row_heights", "diagram_geometry"):
            element.pop(key, None)
        for edge in (element.get("diagram") or {}).get("edges", []):
            edge.pop("source_port", None)
            edge.pop("target_port", None)
    return payload


def apply_patches(canvas, patches):
    candidate = SlideCanvas.model_validate(canvas).model_copy(deep=True)
    protected = _protected(candidate)
    indexed = {element.id: element for element in candidate.elements}
    touched = set()
    for patch in patches:
        key = (patch.element_id, patch.edge_id)
        if key in touched or patch.element_id not in indexed:
            raise ContentConstraintError("视觉修复的图元 ID 重复或不存在")
        touched.add(key)
        element = indexed[patch.element_id]
        for name in ("left", "top", "width", "height"):
            value = getattr(patch, name)
            if value is not None:
                limit = 1.0 if name in ("left", "top") else 2.0
                if abs(value - getattr(element, name)) > limit:
                    raise ContentConstraintError("视觉修复超出单轮移动或尺寸限制")
                setattr(element, name, value)
        if patch.edge_id:
            edge = next((edge for edge in element.diagram.edges if edge.id == patch.edge_id), None) if element.diagram else None
            if not edge:
                raise ContentConstraintError("视觉修复引用了不存在的连接线")
            for name in ("source_port", "target_port"):
                value = getattr(patch, name)
                if value:
                    if getattr(edge, name) != "auto" and getattr(edge, name) != value:
                        raise ContentConstraintError("视觉修复不能改变已明确指定的连接端口")
                    setattr(edge, name, value)
    candidate.layout_resolved = False
    candidate = PPTXRenderer.resolve_canvas(candidate, preserve_svg_crop=True)
    if _protected(candidate) != protected:
        raise ContentConstraintError("视觉修复改变了文字、样式或内容，未采用候选")
    return candidate


def run_visual_pipeline(state, filename, run_dir, api_key, page_numbers=None, used_calls=0):
    from config.jobs import atomic_json
    from main import _render_blueprint
    original_outline = state.get("outline")
    blueprint_payload = state.get("blueprint")
    previous = state.get("visual_review") or {}
    report = {"status": "not_run", "human_review": "not_reviewed", "pages": [], "rounds": 0,
              "model": resolve_model_name("vision"), "reason": "", "calls_used": used_calls}
    state.set("visual_review", report)
    state.set("preview_assets", {})
    state.set("visual_candidates", [])
    report_dir = Path(run_dir) / "generated/visual"
    report_dir.mkdir(parents=True, exist_ok=True)
    selected_output = filename
    chosen_output, chosen_blueprint, chosen_preview = filename, blueprint_payload, {}
    baseline_pages, baseline_status = [], "not_run"
    try:
        if os.environ.get("PPT_PREVIEW_ENABLED", "1") == "0":
            report["reason"] = "成品预览已由配置关闭，未执行视觉审核"
            return filename
        signature = tool_signature()
        disabled = os.environ.get("PPT_VISUAL_REVIEW", "1") == "0"
        blueprint = PPTFinalBlueprint.model_validate(blueprint_payload) if blueprint_payload else None
        targets = {number - 1 for number in page_numbers} if page_numbers else set(range(len(blueprint.canvases))) if blueprint else set()
        if blueprint and any(index < 0 or index >= len(blueprint.canvases) for index in targets):
            raise ContentConstraintError("视觉审核目标页超出范围")
        max_calls = min(15, max(1, int(os.environ.get("PPT_VISION_MAX_CALLS", "8"))))
        frozen = blueprint.model_dump() if blueprint else None
        initial_issues = []
        for round_no in range(3):
            report["rounds"] = round_no
            preview = state.execute_stage(f"preview/round{round_no}", {"filename": selected_output, "runtime": signature},
                lambda: render_preview(selected_output, report_dir / "previews", signature=signature), external=False)
            state.set("preview_assets", preview)
            if round_no == 0:
                chosen_preview = preview
            report["pptx_sha256"] = preview["pptx_sha256"]
            report["pages"] = []
            plans = {}
            if disabled or not blueprint or not api_key or state.get("result_kind") == "offline_demo":
                report["reason"] = "视觉审核已由配置关闭" if disabled else "已生成成品预览；离线演示、缺少密钥或语义蓝图，未执行视觉模型审核"
                break
            if len(preview["pages"]) != len(blueprint.canvases):
                raise PreviewError("成品预览与蓝图页数不一致")
            for index, canvas in enumerate(blueprint.canvases):
                image = preview["pages"][index]
                page_hash = hashlib.sha256(Path(image).read_bytes()).hexdigest()
                result = {"page": index + 1, "slide_id": canvas.slide_id, "image_sha256": page_hash, "issues": [], "status": "not_run"}
                rules = rule_issues(canvas, include_quality=True)
                result["rule_issues"] = rules
                if index not in targets:
                    inherited = next((page for page in previous.get("pages", []) if page.get("slide_id") == canvas.slide_id and page.get("image_sha256") == page_hash), None)
                    if inherited:
                        result.update(status=inherited.get("status", "not_run"), issues=inherited.get("issues", []),
                                      reason=inherited.get("reason", ""), inherited=True)
                    report["pages"].append(result)
                    continue
                def operation(canvas=canvas, image=image):
                    if report["calls_used"] >= max_calls:
                        raise VisionBudgetExceeded("视觉请求预算已用完")
                    report["calls_used"] += 1
                    return review_page(state, canvas, image, api_key, report["model"])
                try:
                    reviewed = state.execute_stage(f"vision/page{index + 1}/round{round_no}",
                        {"image_path": image, "canvas": canvas.model_dump()}, operation,
                        model=report["model"], result_model=VisualPageReview)
                    result.update(issues=[issue.model_dump() for issue in reviewed.issues],
                                  status="needs_review" if reviewed.issues or rules else "passed")
                    if result["status"] == "needs_review" and reviewed.patches:
                        plans[index] = reviewed.patches
                except Exception as error:
                    result["reason"] = "视觉请求预算已用完" if isinstance(error, VisionBudgetExceeded) else "视觉调用失败或结果无效，未自动重发：" + type(error).__name__
                report["pages"].append(result)
            if not initial_issues:
                initial_issues = report["pages"]
            report["status"] = "passed" if all(page["status"] == "passed" for page in report["pages"]) else "needs_review"
            if round_no == 0:
                baseline_pages, baseline_status = report["pages"], report["status"]
            elif all(report["pages"][index]["status"] == "passed" for index in targets):
                chosen_output, chosen_blueprint, chosen_preview = selected_output, blueprint.model_dump(), preview
            report.setdefault("attempts", []).append({"round": round_no, "pages": report["pages"]})
            if report["status"] == "passed" or not plans or round_no == 2 or state.modify_scope == "text":
                if plans and state.modify_scope == "text":
                    report["reason"] = "仅修改文本时保持物理布局，视觉问题交由人工复核"
                break
            candidate = blueprint.model_copy(deep=True)
            try:
                for index, patches in plans.items():
                    candidate.canvases[index] = apply_patches(candidate.canvases[index], patches)
                for index in range(len(candidate.canvases)):
                    if index not in targets and candidate.canvases[index].model_dump() != frozen["canvases"][index]:
                        raise ContentConstraintError("视觉修复改变了非目标页")
                if state.get("outline") != original_outline:
                    raise ContentConstraintError("视觉修复改变了原始事实内容")
                if candidate.model_dump() == blueprint.model_dump():
                    report["reason"] = "建议未产生几何变化，停止重复修复"
                    break
                candidate_path = report_dir / f"round{round_no + 1}" / Path(filename).name
                state.execute_stage(f"render/visual{round_no + 1}", {"blueprint": candidate.model_dump()},
                    lambda: _render_blueprint(candidate, str(candidate_path)), external=False)
                blueprint = candidate
                selected_output = str(candidate_path)
                state.set("visual_candidates", [*(state.get("visual_candidates") or []), selected_output])
                state.set("blueprint", candidate.model_dump())
            except ContentConstraintError as error:
                report["reason"] = str(error)
                break
        report["initial_pages"] = initial_issues
        state.append_progress("Visual: 模型未发现页面问题，仍需人工验收" if report["status"] == "passed" else "Visual: 页面保留未审核或待复核状态")
    except (PreviewError, ContentConstraintError, OSError, ValueError) as error:
        report["reason"] = str(error) if isinstance(error, (PreviewError, ContentConstraintError)) else "成品预览或审核未完成：" + type(error).__name__
        report["status"] = "not_run"
        state.set("preview_assets", {})
        state.append_progress("Visual: " + report["reason"])
    finally:
        if chosen_output != selected_output:
            report["candidate_status"] = report["status"]
            report["pages"], report["status"] = baseline_pages, baseline_status
            report["reason"] = report["reason"] or "修复候选未完成有效的截图与审核，保留初始成品供人工复核"
        report["pptx_sha256"] = chosen_preview.get("pptx_sha256")
        state.set("blueprint", chosen_blueprint)
        state.set("preview_assets", chosen_preview)
        state.set("visual_review", report)
        destination = report_dir / "review.json"
        atomic_json(destination, report)
        state.set("visual_review_path", str(destination))
    return chosen_output
