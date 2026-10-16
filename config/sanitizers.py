import json
import re
from typing import Any, Optional


DEFAULT_THEME = {
    "primary": "#1F2937",
    "secondary": "#2563EB",
    "content": "#4B5563",
    "background": "#F8FAFC",
}

VALID_LAYOUTS = {"title_page", "standard_text", "column", "timeline", "svg", "table", "diagram"}
VALID_ELEMENT_TYPES = {"textbox", "shape", "svg", "image", "icon", "table", "diagram"}
VALID_SHAPE_TYPES = {"rectangle", "rounded_rectangle", "ellipse", "line", "arrow"}

# Design tokens：全局排版字号刻度（Pt，降序）。
# Layout Agent 只允许从该刻度表选字号；sanitizer 与 renderer 会把任意浮点字号
# 吸附到最近的刻度值，从数据层消灭 15.9 / 12.35 这类碎字号。
FONT_SCALE_STEPS = [40, 32, 28, 24, 20, 18, 16, 14, 12, 11, 10]


def snap_font_size(size: float) -> float:
    """把任意字号吸附到排版刻度表；等距时向下吸附（更保守，防溢出）。"""
    try:
        size = float(size)
    except (TypeError, ValueError):
        return 12.0
    if size <= 0:
        return 12.0
    if size >= FONT_SCALE_STEPS[0]:
        return float(FONT_SCALE_STEPS[0])
    for i in range(len(FONT_SCALE_STEPS) - 1):
        lower = FONT_SCALE_STEPS[i + 1]
        upper = FONT_SCALE_STEPS[i]
        if lower <= size <= upper:
            return float(lower if size - lower <= upper - size else upper)
    return float(FONT_SCALE_STEPS[-1])


def model_to_dict(value: Any) -> dict:
    if hasattr(value, "model_dump"):
        return value.model_dump()
    if isinstance(value, dict):
        return value
    return {}


def strip_json_fences(text: str) -> str:
    cleaned = (text or "").strip()
    if cleaned.startswith("```json"):
        return cleaned.split("```json", 1)[1].split("```", 1)[0].strip()
    if cleaned.startswith("```"):
        return cleaned.split("```", 1)[1].split("```", 1)[0].strip()
    return cleaned


def parse_json_object(text: str) -> Any:
    cleaned = strip_json_fences(text)
    data = json.loads(cleaned)
    if isinstance(data, dict) and "type" in data and "text" in data and isinstance(data["text"], str):
        inner = data["text"].strip()
        if inner.startswith("{") or inner.startswith("[") or inner.startswith("```"):
            return parse_json_object(inner)
    return data


def parse_completion_payload(response, state, role):
    """保存真实输出后检查结束状态，不能把截断内容修补成成功结果。"""
    from config.content import ContentConstraintError
    choices = getattr(response, "choices", None) or []
    if not choices:
        raise ContentConstraintError(f"{role} 未返回模型输出；已停止本轮任务")
    choice = choices[0]
    raw = (getattr(choice.message, "content", None) or "").strip()
    if raw:
        state.append_message("assistant", raw)
    finish = getattr(choice, "finish_reason", None)
    if finish == "length":
        raise ContentConstraintError(f"{role} 输出达到模型长度上限而被截断；已停止本轮任务，请减少单次内容或调整输出预算")
    if finish not in {None, "stop"} or not raw:
        raise ContentConstraintError(f"{role} 未返回完整内容；已停止本轮任务")
    try:
        return parse_json_object(raw)
    except (ValueError, TypeError) as exc:
        raise ContentConstraintError(f"{role} 返回的 JSON 不完整或格式无效；已停止本轮任务，原始响应已保存") from exc


def _as_list(value: Any) -> list:
    return value if isinstance(value, list) else []


def _as_content_text(value: Any) -> str:
    # 部分模型把段落包成单字段对象，甚至再次编码成 JSON 字符串。
    candidate = value
    if isinstance(value, str) and value.strip().startswith("{"):
        try:
            candidate = json.loads(value)
        except (ValueError, TypeError):
            pass
    if isinstance(candidate, dict) and set(candidate) == {"text"} and isinstance(candidate["text"], str):
        return candidate["text"]
    return str(value)


def _as_theme(value: Any) -> dict:
    if not isinstance(value, dict):
        return dict(DEFAULT_THEME)
    theme = dict(DEFAULT_THEME)
    for key in ("primary", "secondary", "content", "background"):
        raw = value.get(key)
        if isinstance(raw, str) and re.fullmatch(r"#?[0-9a-fA-F]{6}", raw.strip()):
            color = raw.strip()
            theme[key] = color if color.startswith("#") else f"#{color}"
    return theme


def _normalize_slide(raw_slide: Any, index: int, skeleton_slide: Optional[dict], lock_scope: Optional[str]) -> dict:
    raw = raw_slide if isinstance(raw_slide, dict) else {}
    base = skeleton_slide if isinstance(skeleton_slide, dict) else {}

    if base:
        if lock_scope == "structure":
            # 锁定骨架（标题/要点/版式/配色），仅允许扩写正文与视觉扩展字段
            slide = dict(base)
            for key in ("full_content", "icons", "svg_code", "image_prompt", "table", "diagram"):
                if key in {"table", "diagram"} and base.get(key):
                    continue
                if key in raw:
                    slide[key] = raw[key]
            slide["image_path"] = raw.get("image_path") if raw.get("image_path") else None
        elif lock_scope == "text":
            from config.content import lock_slide
            slide = lock_slide(base, raw, "text")
        elif lock_scope == "design":
            from config.content import lock_slide
            slide = lock_slide(base, raw, "design")
        else:
            slide = {**base, **raw}
    else:
        slide = {**base, **raw}

    if base:
        for key in ("id", "bullet_ids"):
            if base.get(key) and (key == "id" or lock_scope in {"text", "design", "structure"} or len(raw.get("bullet_points") or []) == len(base.get("bullet_points") or [])):
                slide[key] = base[key]

    slide["title"] = str(slide.get("title") or f"第 {index + 1} 页")
    slide["punchline"] = slide.get("punchline") if slide.get("punchline") is not None else None
    as_text = str if lock_scope in {"text", "design"} else _as_content_text
    slide["bullet_points"] = [as_text(item) for item in _as_list(slide.get("bullet_points"))]
    slide["full_content"] = [as_text(item) for item in _as_list(slide.get("full_content"))]
    slide["icons"] = [str(item) for item in _as_list(slide.get("icons"))]

    layout = slide.get("layout_type")
    if layout not in VALID_LAYOUTS:
        slide["layout_type"] = "title_page" if index == 0 else "standard_text"

    slide["svg_code"] = slide.get("svg_code") if slide.get("svg_code") else None
    slide["theme"] = _as_theme(slide.get("theme"))
    slide["image_prompt"] = slide.get("image_prompt") if slide.get("image_prompt") else None
    slide["image_path"] = slide.get("image_path") if slide.get("image_path") else None

    # 新页面可用完整 SVG 表达复杂流程；不能让空的原生占位对象阻断它。
    if lock_scope not in {"text", "design"} and not base.get("diagram") and slide["svg_code"]:
        diagram = slide.get("diagram")
        if diagram is None or (isinstance(diagram, dict) and diagram.get("nodes") == [] and not diagram.get("edges")):
            slide["diagram"] = None
            if slide["layout_type"] == "diagram":
                slide["layout_type"] = "svg"

    bullet_count = len(slide["bullet_points"])
    if bullet_count and lock_scope not in {"text", "design"}:
        structured_only = slide.get("layout_type") in {"table", "diagram", "svg"} and not slide["full_content"]
        if not structured_only and len(slide["full_content"]) < bullet_count:
            slide["full_content"].extend(slide["bullet_points"][len(slide["full_content"]):])
        elif len(slide["full_content"]) > bullet_count:
            slide["full_content"] = slide["full_content"][:bullet_count]

        if len(slide["icons"]) < bullet_count:
            slide["icons"].extend(["circle"] * (bullet_count - len(slide["icons"])))
        elif len(slide["icons"]) > bullet_count:
            slide["icons"] = slide["icons"][:bullet_count]

    return slide


def normalize_outline_payload(
    payload: Any,
    skeleton: Any = None,
    lock_skeleton: bool = False,
    lock_scope: Optional[str] = None,
) -> dict:
    # 兼容旧参数：lock_skeleton=True 等价于 lock_scope="structure"
    if lock_scope is None and lock_skeleton:
        lock_scope = "structure"

    base = model_to_dict(skeleton)
    data = payload if isinstance(payload, dict) else {"slides": payload if isinstance(payload, list) else []}

    base_slides = _as_list(base.get("slides"))
    raw_slides = _as_list(data.get("slides"))
    # 有锁策略时页数以骨架为准（不允许增删页）；否则取两者最大
    if lock_scope in ("structure", "text", "design") and base_slides:
        count = len(base_slides)
    else:
        count = max(len(raw_slides), len(base_slides))

    slides = []
    for i in range(count):
        raw_slide = raw_slides[i] if i < len(raw_slides) else {}
        base_slide = base_slides[i] if i < len(base_slides) else None
        slides.append(_normalize_slide(raw_slide, i, base_slide, lock_scope))

    return {
        "schema_version": data.get("schema_version", base.get("schema_version", 1)),
        "topic": str(data.get("topic") or base.get("topic") or "未命名主题"),
        "slides": slides,
    }


def _normalize_paragraph(raw: Any) -> Optional[dict]:
    if isinstance(raw, str):
        return {
            "text": raw,
            "font_size": 14.0,
            "font_color": "#333333",
            "bold": False,
            "alignment": "left",
            "space_after": 8.0,
        }
    if not isinstance(raw, dict):
        return None
    text = raw.get("text")
    if text is None:
        if not raw.get("content_ref"):
            return None
        text = ""
    alignment = raw.get("alignment") if raw.get("alignment") in {"left", "center", "right"} else "left"
    return {
        "content_ref": raw.get("content_ref"),
        "text": str(text),
        "font_size": snap_font_size(raw.get("font_size") or 14.0),
        "font_color": str(raw.get("font_color") or "#333333"),
        "bold": bool(raw.get("bold", False)),
        "alignment": alignment,
        "space_after": float(8.0 if raw.get("space_after") is None else raw["space_after"]),
    }


def _normalize_element(raw: Any) -> Optional[dict]:
    if not isinstance(raw, dict):
        return None
    item = dict(raw)

    if item.get("type") == "text":
        item["type"] = "textbox"
    if item.get("type") not in VALID_ELEMENT_TYPES:
        return None

    for key, default in (("left", 0.0), ("top", 0.0), ("width", 1.0), ("height", 0.4)):
        try:
            item[key] = float(item.get(key, default))
        except (TypeError, ValueError):
            item[key] = default

    if item["type"] == "textbox":
        paragraphs = item.get("paragraphs")
        if not isinstance(paragraphs, list):
            paragraphs = [item["text"]] if item.get("text") is not None else []
        normalized = [_normalize_paragraph(p) for p in paragraphs]
        item["paragraphs"] = [p for p in normalized if p is not None]
        if not item["paragraphs"]:
            return None

    if item["type"] == "shape":
        if item.get("shape_type") == "circle":
            item["shape_type"] = "ellipse"
        if item.get("shape_type") not in VALID_SHAPE_TYPES:
            item["shape_type"] = "rectangle"

    if item["type"] == "icon":
        # 白名单外 icon_name 模糊归一到库内最近似图标（如 mountain-peak -> mountain），
        # 无法可靠匹配时保留原值，由渲染层回退占位圆点。
        from config.icon_library import resolve_icon_name
        resolved = resolve_icon_name(item.get("icon_name"))
        if resolved:
            item["icon_name"] = resolved
        else:
            item["icon_name"] = str(item.get("icon_name") or "")

    return item


def normalize_blueprint_payload(payload: Any) -> dict:
    data = payload if isinstance(payload, dict) else {"canvases": payload if isinstance(payload, list) else []}
    canvases = []
    for canvas in _as_list(data.get("canvases")):
        if not isinstance(canvas, dict):
            continue
        elements = [_normalize_element(el) for el in _as_list(canvas.get("elements"))]
        canvases.append({
            "slide_id": canvas.get("slide_id"),
            "layout_resolved": bool(canvas.get("layout_resolved", False)),
            "background_color": str(canvas.get("background_color") or "#FFFFFF"),
            "elements": [el for el in elements if el is not None],
        })
    return {"schema_version": data.get("schema_version", 1), "canvases": canvases}
