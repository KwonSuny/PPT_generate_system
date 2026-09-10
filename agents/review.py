# agents/review.py
from config.schema import PPTOutline
from typing import List, Tuple

def review_ppt_outline(outline: PPTOutline) -> Tuple[bool, List[str]]:
    errors = []
    
    # 检查页面数量
    actual_slides_count = len(outline.slides)
    if not (1 <= actual_slides_count <= 5):
        errors.append(f"页面数量为 {actual_slides_count} 页，不符合 1-5 页的严格要求！")

    # 检查图表键名
    for i, slide in enumerate(outline.slides):
        if slide.layout_type == "chart" and slide.chart_data:
            for key in slide.chart_data.keys():
                if any(char in key for char in ["(", ")", "（", "）", "亿", "万", "用户", "%"]):
                    errors.append(f"第 {i+1} 页图表键名 '{key}' 违规！请将单位移到标题。")
                    
        # 检查 SVG 流程图合规性
        if slide.layout_type == "svg":
            if not slide.svg_code or len(slide.svg_code.strip()) < 50:
                errors.append(f"第 {i+1} 页被指定为 'svg' 布局，但 `svg_code` 为空或代码残缺！")
            elif "```" in slide.svg_code:
                errors.append(f"第 {i+1} 页的 `svg_code` 包含了 Markdown 标记（如 ```），请输出纯净的 XML 代码。")

        # 检查文字长度
        for bp in slide.bullet_points:
            if len(bp) > 50:
                errors.append(f"第 {i+1} 页文本过长（{len(bp)}字）: '{bp[:10]}...'。请精简到 50 字以内。")

    return len(errors) == 0, errors