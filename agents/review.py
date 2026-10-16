from typing import Tuple, List
from config.content import check_page_count, ContentConstraintError
from config.schema import PPTOutline, PagePolicy


def review_ppt_outline(outline: PPTOutline, page_policy=None) -> Tuple[bool, List[str]]:
    errors = []
    try:
        check_page_count(len(outline.slides), page_policy or PagePolicy())
    except ContentConstraintError as exc:
        errors.append(str(exc))
    for i, slide in enumerate(outline.slides, 1):
        if slide.layout_type == "svg":
            if not slide.svg_code or len(slide.svg_code) < 50 or "```" in slide.svg_code:
                errors.append(f"第 {i} 页缺少有效的纯 SVG 源码")
        if slide.layout_type == "table" and not slide.table:
            errors.append(f"第 {i} 页缺少原生表格数据")
        if slide.table and not slide.table.source:
            errors.append(f"第 {i} 页的表格缺少数据来源说明，需要复核")
        if slide.layout_type == "diagram" and not slide.diagram:
            errors.append(f"第 {i} 页缺少原生关系图节点与连接")
        for bullet in slide.bullet_points:
            if len(bullet) > 50:
                errors.append(f"第 {i} 页要点超过 50 字，请精简")
    return not errors, errors
