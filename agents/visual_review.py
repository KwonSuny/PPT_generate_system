"""审核成品截图，输出问题与受限几何修复建议。"""

import base64
import json
from pathlib import Path
from typing import Literal

from pydantic import BaseModel, ConfigDict, Field, field_validator, model_validator

from config.content import ContentConstraintError
from config.llm_client import chat_completion_with_retry, create_client, resolve_model_name, model_request_options
from config.sanitizers import parse_completion_payload


class GeometryPatch(BaseModel):
    model_config = ConfigDict(extra="forbid")
    element_id: str
    left: float | None = Field(default=None, ge=0, allow_inf_nan=False)
    top: float | None = Field(default=None, ge=0, allow_inf_nan=False)
    width: float | None = Field(default=None, gt=0, allow_inf_nan=False)
    height: float | None = Field(default=None, gt=0, allow_inf_nan=False)
    edge_id: str | None = None
    source_port: Literal["left", "right", "top", "bottom"] | None = None
    target_port: Literal["left", "right", "top", "bottom"] | None = None

    @model_validator(mode="after")
    def meaningful(self):
        geometry = any(getattr(self, key) is not None for key in ("left", "top", "width", "height"))
        ports = self.source_port is not None or self.target_port is not None
        if not geometry and not ports or ports and not self.edge_id or self.edge_id and not ports:
            raise ValueError("修复建议缺少有效的几何或端口变更")
        return self


class VisualIssue(BaseModel):
    model_config = ConfigDict(extra="forbid")
    kind: Literal["overlap", "alignment", "overflow", "contrast", "arrow", "spacing", "other"]
    severity: Literal["minor", "major"]
    element_ids: list[str] = Field(default_factory=list, max_length=12)
    bbox: list[float] = Field(min_length=4, max_length=4)
    message: str = Field(min_length=1, max_length=300)
    suggestion: str = Field(default="", max_length=300)

    @field_validator("bbox", mode="before")
    @classmethod
    def named_coordinates(cls, value):
        if isinstance(value, dict) and set(value) == {"x", "y", "width", "height"}:
            return [value[key] for key in ("x", "y", "width", "height")]
        return value

    @model_validator(mode="after")
    def bounded(self):
        x, y, width, height = self.bbox
        if not all(0 <= value <= 1 for value in self.bbox) or width <= 0 or height <= 0 or x + width > 1.001 or y + height > 1.001:
            raise ValueError("问题位置必须在页面内，采用 0–1 比例坐标")
        return self


class VisualPageReview(BaseModel):
    model_config = ConfigDict(extra="forbid")
    slide_id: str
    issues: list[VisualIssue] = Field(default_factory=list, max_length=20)
    patches: list[GeometryPatch] = Field(default_factory=list, max_length=10)


def review_page(state, canvas, image_path, api_key, model=None):
    model = model or resolve_model_name("vision")
    image = Path(image_path)
    if image.stat().st_size > 10 * 1024 * 1024:
        raise ContentConstraintError("审核截图超出 10 MiB 限制")
    payload = canvas.model_dump()
    from config.layout_budget import layout_issues
    payload['rule_feedback'] = layout_issues(canvas)
    # 只传画布结构与文件名，截图使用内联数据，避免暴露本机目录。
    for element in payload["elements"]:
        if element.get("image_path"):
            element["image_path"] = Path(element["image_path"]).name
        element.pop("svg_code", None)
    system = """你审核真实 PPT 成品截图，检查文本裁切、遮挡、箭头连接、对齐、表格可读性与颜色对比。
不要评价主题优劣，不改写任何文字、事实或数值。输入蓝图用于定位元素，实际问题必须结合截图判断。
rule_feedback 是程序测得的坐标、间距、容量和图内字号估计；结合截图确认，修复时同时考虑关联元素，不能只移动报错框造成新的冲突。
只报告有具体位置的排版缺陷；直角或曲线、圆角或方角等设计偏好不能当作错误。连线不断开、不穿字且端点正确时，不因样式偏好报告箭头问题。
只输出 JSON：{slide_id,issues,patches}。没有发现问题时 issues=[]、patches=[]，不要虚构问题。
issues 每项包含 kind（overlap/alignment/overflow/contrast/arrow/spacing/other）、severity（minor/major）、
element_ids（必须取自输入图元 id）、bbox、message、suggestion。
bbox 必须使用带名称的对象，例如 {"x":0.3,"y":0.2,"width":0.2,"height":0.1}；均为页面 0–1 比例坐标。
width/height 是宽高，不是右下角位置，必须满足 x+width<=1、y+height<=1，禁止混用坐标格式。
可确定修复时才提供 patches：element_id，及 left/top/width/height（英寸，页面13.333×7.5）；
最多移动 1 英寸、改变尺寸 2 英寸。只修改已有容器，不新增/删除图元，不修改文字、字号、配色、SVG 或数据。
diagram 连线可建议 edge_id 与 source_port/target_port（left/right/top/bottom）。不得改变源/目标节点或关系文字。
不能用缩小字体解决文本溢出。不确定、需要内容调整或配色调整时，仅报告问题，不提供几何修复。
"""
    state.append_progress(f"Visual: 审核页面 {canvas.slide_id}")
    response = chat_completion_with_retry(
        create_client(api_key, timeout=120), model=model,
        messages=[{"role": "system", "content": system}, {"role": "user", "content": [
            {"type": "text", "text": json.dumps(payload, ensure_ascii=False)},
            {"type": "image_url", "image_url": {"url": "data:image/png;base64," + base64.b64encode(image.read_bytes()).decode("ascii")}},
        ]}], response_format={"type": "json_object"},
        **model_request_options('vision',model),
    )
    result = VisualPageReview.model_validate(parse_completion_payload(response, state, "Visual"))
    ids = {element.id for element in canvas.elements}
    if result.slide_id != canvas.slide_id or any(set(issue.element_ids) - ids for issue in result.issues) or any(patch.element_id not in ids for patch in result.patches):
        raise ContentConstraintError("视觉审核返回了范围外的页面或图元 ID")
    if result.patches and not result.issues:
        raise ContentConstraintError("视觉审核没有问题依据却建议修改，未应用修复")
    return result
