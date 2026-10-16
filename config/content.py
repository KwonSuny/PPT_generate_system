"""语义内容与物理蓝图之间的确定性契约。"""

import re
import unicodedata
from copy import deepcopy
from decimal import Decimal, ROUND_HALF_UP

from config.schema import PagePolicy, PPTOutline, SlideCanvas, SlideElement


class ContentConstraintError(ValueError):
    """可以直接反馈给用户的页数、映射或容量限制。"""


def _page_number(value):
    if value.isdigit():
        if len(value) > 3:
            raise ContentConstraintError("新生成支持 1–50 页，请缩小页数要求")
        return int(value)
    digits = dict(zip("一二三四五六七八九", range(1, 10))) | {"两": 2}
    if value in digits:
        return digits[value]
    if re.fullmatch(r"[一二两三四五六七八九]?十[一二三四五六七八九]?", value):
        left, right = value.split("十", 1)
        return (digits.get(left, 1) * 10) + digits.get(right, 0)
    raise ContentConstraintError("页数格式无效，请输入整数或范围，例如 3、2–5")


def parse_page_policy(raw="", text="", allow_default=True):
    """表单优先；没有表单值时识别文字中的明确页数要求。"""
    raw = unicodedata.normalize("NFKC", raw or "").strip()
    number = r"(?:[0-9]+|[一二两三四五六七八九十]+)"
    if raw:
        match = re.fullmatch(rf"\s*({number})\s*(?:页)?\s*(?:[-–—~～至到]\s*({number}))?\s*页?\s*", raw)
        if not match or len(raw) > 32:
            raise ContentConstraintError("页数格式无效，请输入整数或范围，例如 3、2–5")
        bounds = [(_page_number(match[1]), _page_number(match[2] or match[1]))]
    else:
        text = unicodedata.normalize("NFKC", text or "")
        pattern = rf"(?<![\d.])(第\s*)?({number})\s*(?:页\s*)?(?:[-–—~～至到]\s*({number})\s*)?页"
        bounds = [(_page_number(m[2]), _page_number(m[3] or m[2])) for m in re.finditer(pattern, text) if not m[1]]
        named = rf"(?:总?页数)\s*(?:[:：=为是])?\s*({number})\s*(?:[-–—~～至到]\s*({number}))?"
        bounds.extend((_page_number(m[1]), _page_number(m[2] or m[1])) for m in re.finditer(named, text))
        if not bounds:
            return PagePolicy() if allow_default else None
        if len(set(bounds)) != 1:
            raise ContentConstraintError("文字中有相互冲突的总页数要求，请在页数栏明确指定")
    low, high = bounds[0]
    if not 1 <= low <= high <= 50:
        raise ContentConstraintError("页数范围需满足 1 ≤ 下限 ≤ 上限 ≤ 50")
    return PagePolicy(min_slides=low, max_slides=high)


def check_page_count(count, policy):
    policy = PagePolicy.model_validate(policy or {})
    if not policy.min_slides <= count <= policy.max_slides:
        raise ContentConstraintError(
            f"实际生成 {count} 页，要求为 {policy.min_slides}–{policy.max_slides} 页；已停止导出"
        )


def identified_slide(payload, page_number=1):
    slide = deepcopy(payload)
    slide["id"] = slide.get("id") or f"slide-{page_number}"
    count = len(slide.get("bullet_points") or [])
    ids = slide.get("bullet_ids") or [f"point-{i + 1}" for i in range(count)]
    if len(ids) != count or len(set(ids)) != count or any(not re.fullmatch(r"[A-Za-z0-9_-]+", v) for v in ids):
        raise ContentConstraintError("要点 ID 必须唯一且与要点数量对应")
    slide["bullet_ids"] = ids
    return SlideElement.model_validate(slide).model_dump()


def identified_outline(payload):
    result = deepcopy(payload)
    result["schema_version"] = 2
    result["slides"] = [identified_slide(slide, i + 1) for i, slide in enumerate(result["slides"])]
    if len({slide["id"] for slide in result["slides"]}) != len(result["slides"]):
        raise ContentConstraintError("页面 ID 必须唯一")
    return PPTOutline.model_validate(result)


def cell_text(cell, column):
    value = cell.value
    fmt = cell.format or column.format
    if value is None:
        return "—"
    if isinstance(value, str) or fmt.style == "plain":
        text = str(value)
    else:
        number = Decimal(str(value))
        decimals = 0 if fmt.style == "integer" else fmt.decimals
        if fmt.style == "percent":
            number *= 100
        number = number.quantize(Decimal(1).scaleb(-decimals), rounding=ROUND_HALF_UP)
        text = f"{number:.{decimals}f}" + ("%" if fmt.style == "percent" else "")
    return fmt.prefix + text + fmt.suffix


def content_values(slide):
    slide = SlideElement.model_validate(slide)
    values = {"title": slide.title}
    if slide.punchline is not None:
        values["punchline"] = slide.punchline
    ids = slide.bullet_ids or [f"point-{i + 1}" for i in range(len(slide.bullet_points))]
    for name in ("bullet_points", "full_content"):
        for i, text in enumerate(getattr(slide, name)):
            point_id = ids[i] if i < len(ids) else f"point-{i + 1}"
            values[f"{name}/{point_id}"] = text
    if slide.table:
        for column in slide.table.columns:
            values[f"table/columns/{column.id}/title"] = column.title
        for row in slide.table.rows:
            for cell, column in zip(row.cells, slide.table.columns):
                values[f"table/rows/{row.id}/cells/{column.id}"] = cell_text(cell, column)
        values["table/source"] = slide.table.source or ""
    if slide.diagram:
        values.update({f"diagram/nodes/{node.id}/label": node.label for node in slide.diagram.nodes})
        values.update({f"diagram/edges/{edge.id}/label": edge.label for edge in slide.diagram.edges})
    return values


def lock_slide(old, proposed, scope):
    """在模型清洗之外再校验一次，避免 mock、回放或新 Agent 绕过锁定。"""
    old, proposed = deepcopy(old), deepcopy(proposed)
    if scope == "design":
        result = {**old, **proposed}
        for key in ("id", "bullet_ids", "title", "punchline", "bullet_points", "full_content", "table", "diagram", "svg_code"):
            if key in old:
                result[key] = old[key]
        return result
    if scope != "text":
        proposed["id"] = old.get("id") or proposed.get("id")
        if len(old.get("bullet_points") or []) == len(proposed.get("bullet_points") or []):
            proposed["bullet_ids"] = old.get("bullet_ids") or proposed.get("bullet_ids") or []
        return proposed
    result = deepcopy(old)
    for key in ("title", "punchline", "bullet_points", "full_content"):
        if key in proposed:
            if key in ("bullet_points", "full_content") and len(proposed[key] or []) != len(old.get(key) or []):
                raise ContentConstraintError("仅修改文本时不能增删要点或段落；请使用页面修改重新排版")
            result[key] = proposed[key]
    if old.get("table"):
        incoming = proposed.get("table")
        base = old["table"]
        if not incoming or [c["id"] for c in incoming["columns"]] != [c["id"] for c in base["columns"]] or [r["id"] for r in incoming["rows"]] != [r["id"] for r in base["rows"]]:
            raise ContentConstraintError("仅修改文本时不能增删、重排行列；请使用页面修改重新排版")
        for i, column in enumerate(result["table"]["columns"]):
            column["title"] = incoming["columns"][i]["title"]
        for i, row in enumerate(result["table"]["rows"]):
            if len(incoming["rows"][i]["cells"]) != len(row["cells"]):
                raise ContentConstraintError("仅修改文本时不能增删单元格")
            for j, cell in enumerate(row["cells"]):
                cell["value"] = incoming["rows"][i]["cells"][j]["value"]
        result["table"]["source"] = incoming.get("source")
    elif proposed.get("table"):
        raise ContentConstraintError("旧页面没有表格，添加表格需要重新排版")
    if old.get("diagram"):
        incoming, base = proposed.get("diagram"), old["diagram"]
        if not incoming or [n["id"] for n in incoming["nodes"]] != [n["id"] for n in base["nodes"]] or [e["id"] for e in incoming["edges"]] != [e["id"] for e in base["edges"]]:
            raise ContentConstraintError("仅修改文本时不能增删、重排节点或连接线")
        for i, node in enumerate(result["diagram"]["nodes"]):
            node["label"] = incoming["nodes"][i]["label"]
        for i, edge in enumerate(result["diagram"]["edges"]):
            edge["label"] = incoming["edges"][i]["label"]
    elif proposed.get("diagram"):
        raise ContentConstraintError("旧页面没有原生关系图，添加图形需要重新排版")
    return result


def _paragraph_reference(paragraph, values, strict):
    ref = paragraph.content_ref
    if ref is not None:
        if ref not in values:
            raise ContentConstraintError(f"未知内容引用 {ref}，无法可靠回填文本")
        return ref
    matches = [key for key, text in values.items() if text.strip() == paragraph.text.strip()]
    if len(matches) == 1:
        return matches[0]
    if strict:
        reason = "对应多个内容字段" if matches else "无法对应语义内容"
        raise ContentConstraintError(f"文字“{paragraph.text[:24]}”{reason}，需要明确 content_ref")
    return None


def bind_canvas(canvas, slide, strict=True, preserve_component_style=False):
    """Layout 只决定容器位置，文字、数值与 SVG 源码由代码回填。"""
    canvas = SlideCanvas.model_validate(canvas).model_copy(deep=True)
    slide = SlideElement.model_validate(slide)
    if strict and slide.layout_type in {"table", "diagram", "svg"} and not getattr(slide, "svg_code" if slide.layout_type == "svg" else slide.layout_type):
        raise ContentConstraintError(f"{slide.layout_type} 页面缺少对应的结构化内容")
    values = content_values(slide)
    canvas.slide_id = slide.id
    covered, components = set(), {"table": 0, "diagram": 0, "svg": 0, "image": 0}
    for i, element in enumerate(canvas.elements):
        element.id = element.id or f"{slide.id}-element-{i + 1}"
        for paragraph in (element.paragraphs or []) if element.type == "textbox" else []:
            ref = _paragraph_reference(paragraph, values, strict)
            if ref:
                paragraph.content_ref = ref
                if strict:
                    paragraph.text = values[ref]
                covered.add(ref)
        if element.type in ("table", "diagram"):
            spec = getattr(slide, element.type)
            if spec is None:
                raise ContentConstraintError(f"排版新增了语义层不存在的 {element.type}")
            element.content_ref = element.type
            setattr(element, element.type, spec.model_copy(deep=True))
            if strict:
                style = getattr(element, element.type + '_style')
                colors = ({'header_fill': slide.theme.primary, 'font_color': slide.theme.content,
                           'body_fill': slide.theme.background} if element.type == 'table' else
                          {'line_color': slide.theme.primary, 'font_color': slide.theme.content,
                           'fill_color': slide.theme.background})
                for name, color in colors.items():
                    if not preserve_component_style or name not in style.model_fields_set:
                        setattr(style, name, color)
            components[element.type] += 1
            covered.update(ref for ref in values if ref.startswith(element.type + "/"))
        elif element.type == "svg":
            if not slide.svg_code:
                raise ContentConstraintError("SVG 图元没有对应的 SVG 源码")
            element.svg_code = slide.svg_code
            components["svg"] += 1
        elif element.type == "image":
            if not slide.image_path:
                raise ContentConstraintError("图片图元没有对应的已生成图片")
            element.image_path = slide.image_path
            components["image"] += 1
    if len({el.id for el in canvas.elements}) != len(canvas.elements):
        raise ContentConstraintError("图元 ID 必须唯一")
    if strict:
        for kind in ("table", "diagram"):
            if getattr(slide, kind) and components[kind] != 1:
                raise ContentConstraintError(f"该页需要且只能有一个原生 {kind} 容器")
        if slide.svg_code and components["svg"] != 1:
            raise ContentConstraintError("该页需要一个保持比例的 SVG 容器")
        if slide.image_path and not components["image"]:
            raise ContentConstraintError("排版遗漏了已生成的图片")
        required = {"title"}
        if slide.punchline:
            required.add("punchline")
        if slide.layout_type != "title_page":
            required.update(ref for ref in values if ref.startswith("full_content/"))
        if required - covered:
            raise ContentConstraintError("排版遗漏内容引用：" + "、".join(sorted(required - covered)))
    return canvas


def lock_text_canvas(old_canvas, old_slide, new_slide):
    from renderer import PPTXRenderer
    old_values, new_values = content_values(old_slide), content_values(new_slide)
    changed = {ref for ref in old_values.keys() | new_values.keys() if old_values.get(ref) != new_values.get(ref)}
    canvas = bind_canvas(old_canvas, old_slide, strict=False)
    # 旧版本保存的是渲染前蓝图；先复现旧导出使用的几何与字体，再锁定。
    canvas = PPTXRenderer.resolve_canvas(canvas, legacy=True)
    covered = set()
    for element in canvas.elements:
        touched = False
        for paragraph in (element.paragraphs or []) if element.type == "textbox" else []:
            ref = paragraph.content_ref
            if ref in changed:
                if ref not in new_values:
                    raise ContentConstraintError(f"仅修改文本时不能移除 {ref} 对应的物理段落")
                paragraph.text = new_values[ref]
                touched = True
            if ref:
                covered.add(ref)
        if element.type in ("table", "diagram"):
            prefix = element.type + "/"
            setattr(element, element.type, getattr(SlideElement.model_validate(new_slide), element.type))
            covered.update(ref for ref in new_values if ref.startswith(prefix))
            touched = any(ref.startswith(prefix) for ref in changed)
        if touched:
            PPTXRenderer.validate_capacity(element)
    missing = changed - covered
    # 正文已占据物理段落时，要点骨架允许只作为内部内容保留。
    missing = {ref for ref in missing if not (ref.startswith("bullet_points/") and ref.replace("bullet_points/", "full_content/", 1) in covered)}
    if missing:
        raise ContentConstraintError("旧蓝图无法唯一映射修改内容（可能在 SVG 内或未显示）：" + "、".join(sorted(missing)))
    return canvas
