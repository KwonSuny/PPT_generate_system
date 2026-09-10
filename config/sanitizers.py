import json
import re
from typing import Any, Optional


DEFAULT_THEME = {
    "primary": "#1F2937",
    "secondary": "#2563EB",
    "content": "#4B5563",
    "background": "#F8FAFC",
}

VALID_LAYOUTS = {"title_page", "standard_text", "column", "timeline",  "svg"}
VALID_ELEMENT_TYPES = {"textbox", "shape", "svg", "image", "icon"}
VALID_SHAPE_TYPES = {"rectangle", "rounded_rectangle", "ellipse", "line", "arrow"}


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


def _as_list(value: Any) -> list:
    return value if isinstance(value, list) else []


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
            for key in ("full_content", "icons", "svg_code", "image_prompt"):
                if key in raw:
                    slide[key] = raw[key]
            slide["image_path"] = raw.get("image_path") if raw.get("image_path") else None
        elif lock_scope == "text":
            # 锁定版式（layout_type/theme），允许改写文字字段
            slide = {**base, **raw}
            slide["layout_type"] = base.get("layout_type")
            slide["theme"] = base.get("theme")
        elif lock_scope == "design":
            # 锁定文字（标题/要点/正文），允许改写版式与配色
            slide = {**base, **raw}
            for key in ("title", "punchline", "bullet_points", "full_content"):
                if key in base:
                    slide[key] = base[key]
        else:
            slide = {**base, **raw}
    else:
        slide = {**base, **raw}

    slide["title"] = str(slide.get("title") or f"第 {index + 1} 页")
    slide["punchline"] = slide.get("punchline") if slide.get("punchline") is not None else None
    slide["bullet_points"] = [str(item) for item in _as_list(slide.get("bullet_points"))]
    slide["full_content"] = [str(item) for item in _as_list(slide.get("full_content"))]
    slide["icons"] = [str(item) for item in _as_list(slide.get("icons"))]

    layout = slide.get("layout_type")
    if layout not in VALID_LAYOUTS:
        slide["layout_type"] = "title_page" if index == 0 else "standard_text"

    slide["svg_code"] = slide.get("svg_code") if slide.get("svg_code") else None
    slide["theme"] = _as_theme(slide.get("theme"))
    slide["image_prompt"] = slide.get("image_prompt") if slide.get("image_prompt") else None
    slide["image_path"] = slide.get("image_path") if slide.get("image_path") else None

    bullet_count = len(slide["bullet_points"])
    if bullet_count:
        if len(slide["full_content"]) < bullet_count:
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
        return None
    alignment = raw.get("alignment") if raw.get("alignment") in {"left", "center", "right"} else "left"
    return {
        "text": str(text),
        "font_size": float(raw.get("font_size") or 14.0),
        "font_color": str(raw.get("font_color") or "#333333"),
        "bold": bool(raw.get("bold", False)),
        "alignment": alignment,
        "space_after": float(raw.get("space_after") or 8.0),
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


    return item


def normalize_blueprint_payload(payload: Any) -> dict:
    data = payload if isinstance(payload, dict) else {"canvases": payload if isinstance(payload, list) else []}
    canvases = []
    for canvas in _as_list(data.get("canvases")):
        if not isinstance(canvas, dict):
            continue
        elements = [_normalize_element(el) for el in _as_list(canvas.get("elements"))]
        canvases.append({
            "background_color": str(canvas.get("background_color") or "#FFFFFF"),
            "elements": [el for el in elements if el is not None],
        })
    return {"canvases": canvases}
