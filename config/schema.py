import re
import math
from typing import List, Optional, Literal
from pydantic import BaseModel, Field, field_validator, model_validator


class PagePolicy(BaseModel):
    min_slides: int = Field(default=2, ge=1, strict=True)
    max_slides: int = Field(default=5, ge=1, strict=True)

    @model_validator(mode="after")
    def ordered(self):
        if self.min_slides > self.max_slides:
            raise ValueError("页数范围的下限不能大于上限")
        return self


class NumberFormat(BaseModel):
    style: Literal["plain", "decimal", "integer", "percent"] = "plain"
    decimals: int = Field(default=2, ge=0, le=8)
    prefix: str = ""
    suffix: str = ""


class TableCell(BaseModel):
    value: str | int | float | None
    format: Optional[NumberFormat] = None

    @field_validator("value", mode="before")
    @classmethod
    def preserve_value(cls, value):
        if isinstance(value, bool) or not isinstance(value, (str, int, float, type(None))):
            raise ValueError("单元格只能使用文字、数值或 null")
        if isinstance(value, float) and not math.isfinite(value):
            raise ValueError("单元格数值必须有限")
        return value


class TableColumn(BaseModel):
    id: str = Field(pattern=r"^[A-Za-z0-9_-]+$")
    title: str
    width_weight: Optional[float] = Field(default=None, gt=0, allow_inf_nan=False)
    alignment: Literal["auto", "left", "center", "right"] = "auto"
    format: NumberFormat = Field(default_factory=NumberFormat)


class TableRow(BaseModel):
    id: str = Field(pattern=r"^[A-Za-z0-9_-]+$")
    cells: List[TableCell]


class TableSpec(BaseModel):
    id: str = "table"
    columns: List[TableColumn] = Field(min_length=1)
    rows: List[TableRow] = Field(min_length=1)
    source: Optional[str] = None

    @model_validator(mode="after")
    def rectangular(self):
        for group in (self.columns, self.rows):
            if len({item.id for item in group}) != len(group):
                raise ValueError("表格行和列的 ID 必须分别唯一")
        if any(len(row.cells) != len(self.columns) for row in self.rows):
            raise ValueError("表格每行的单元格数量必须与列数相同")
        return self


class TableStyle(BaseModel):
    font_size: float = Field(default=18, ge=14, le=32, allow_inf_nan=False)
    header_font_size: float = Field(default=18, ge=14, le=32, allow_inf_nan=False)
    header_fill: str = "#005088"
    header_color: str = "#FFFFFF"
    font_color: str = "#334155"
    body_fill: str = "#FFFFFF"
    alternate_fill: str = "#F8FAFC"
    border_color: str = "#CBD5E1"
    source_font_size: float = Field(default=12, ge=10, le=18, allow_inf_nan=False)


Port = Literal["auto", "top", "left", "bottom", "right"]


class DiagramNode(BaseModel):
    id: str = Field(pattern=r"^[A-Za-z0-9_-]+$")
    label: str
    shape_type: Literal["rectangle", "rounded_rectangle"] = "rectangle"


class DiagramEdge(BaseModel):
    id: str = Field(pattern=r"^[A-Za-z0-9_-]+$")
    source: str
    target: str
    source_port: Port = "auto"
    target_port: Port = "auto"
    label: str = ""


class DiagramSpec(BaseModel):
    id: str = "diagram"
    direction: Literal["LR", "TB"] = "LR"
    nodes: List[DiagramNode] = Field(min_length=1)
    edges: List[DiagramEdge] = Field(default_factory=list)

    @model_validator(mode="after")
    def valid_references(self):
        ids = {node.id for node in self.nodes}
        if len(ids) != len(self.nodes) or len({edge.id for edge in self.edges}) != len(self.edges):
            raise ValueError("关系图的节点和边 ID 必须分别唯一")
        if any(edge.source not in ids or edge.target not in ids or edge.source == edge.target for edge in self.edges):
            raise ValueError("关系图连接必须引用两个不同的已有节点")
        return self


class DiagramStyle(BaseModel):
    font_size: float = Field(default=18, ge=14, le=32, allow_inf_nan=False)
    label_font_size: float = Field(default=14, ge=12, le=24, allow_inf_nan=False)
    fill_color: str = "#F8FAFC"
    font_color: str = "#334155"
    line_color: str = "#005088"


class GeometryBox(BaseModel):
    left: float = Field(allow_inf_nan=False)
    top: float = Field(allow_inf_nan=False)
    width: float = Field(gt=0, allow_inf_nan=False)
    height: float = Field(gt=0, allow_inf_nan=False)


class NodeGeometry(GeometryBox):
    id: str


class EdgeGeometry(BaseModel):
    id: str
    source_port: Port
    target_port: Port
    start_x: float = Field(allow_inf_nan=False)
    start_y: float = Field(allow_inf_nan=False)
    end_x: float = Field(allow_inf_nan=False)
    end_y: float = Field(allow_inf_nan=False)
    label_box: Optional[GeometryBox] = None


class DiagramGeometry(BaseModel):
    nodes: List[NodeGeometry]
    edges: List[EdgeGeometry]


def re_split_legend(s: str):
    # split by common separators: comma, Chinese comma, semicolon, Chinese semicolon, slash, pipe
    return re.split(r"[,，;；/|、]", s)

class ColorPalette(BaseModel):
    """PPT 单页视觉调色盘"""
    primary: str = Field(description="主色调，用于当前页大标题、核心高亮（如 #005088）")
    secondary: str = Field(description="辅助色，用于副标题、图表对比色（如 #11CAA0）")
    content: str = Field(description="正文色，用于段落文本（如 #64748B）")
    background: str = Field(description="当前页面专属背景色（如 #F8FAFC）")

class SlideElement(BaseModel):
    """单页 PPT 中的内容与视觉模型"""
    id: Optional[str] = None
    bullet_ids: List[str] = Field(default_factory=list)
    title: str = Field(description="当前页面的小标题。")
    punchline: Optional[str] = Field(
        default=None, 
        description="【Planner生成】当前页面的核心结论、破局金句，字数控制在25字以内。"
    )
    bullet_points: List[str] = Field(
        default_factory=list,
        description="【Planner生成】当前页面的核心要点骨架（粗略大纲），通常为2-4条简短提示。"
    )
    full_content: List[str] = Field(
        default_factory=list,
        description="【Writer生成】由writer阶段填充的详细长文本扩写，每条内容不超过50字，须与bullet_points数量严格对应。"
    )
    layout_type: Literal['title_page', 'standard_text', 'column', 'timeline', 'svg', 'table', 'diagram'] = Field(
        description="表格选择 table，简单关系图选择 diagram，复杂构图保留 svg"
    )
    table: Optional[TableSpec] = None
    diagram: Optional[DiagramSpec] = None
    svg_code: Optional[str] = Field(
        default=None,
        description="【Writer生成】当 layout_type == 'svg' 时，在此处填入现代美学标准的 SVG 代码。"
    )
    theme: ColorPalette = Field(
        description="【Planner/Writer生成】该页PPT专属的个性化视觉色调盘，属于整体色族系列，禁止每页跨度突兀过大。"
    )
    image_prompt: Optional[str] = Field(
        default=None,
        description="【Writer生成】若此页需要图形点缀，提供精炼的中文核心意境描述词（无任何具体文本/图表数据信息）。反之留空。"
    )
    image_path: Optional[str] = Field(
        default=None,
        description="【系统自动填充】由图像模型生成后的本地临时图片绝对路径。"
    )
    icons: List[str] = Field(
        default_factory=list,
        description="【Writer生成】为每个要点匹配的 Lucide 图标库名称，数量必须与 bullet_points 严格一致。若无则留空。"
    )
    @field_validator('bullet_points', 'full_content', mode='before')
    @classmethod
    def convert_null_to_list(cls, v):
        if v is None:
            return []
        return v

class PPTOutline(BaseModel):
    """语义层：整份 PPT 的内容大纲与核心意境集合"""
    schema_version: Literal[1, 2] = 1
    topic: str = Field(description="PPT的全局核心主题")
    slides: List[SlideElement] = Field(description="幻灯片页面列表")

class TextParagraph(BaseModel):
    """物理层：单个段落的精细物理控制"""
    content_ref: Optional[str] = None
    text: str = Field(description="文本内容")
    font_size: float = Field(default=14.0, gt=0, allow_inf_nan=False, description="字号（Pt）")
    font_family: Optional[Literal['Microsoft YaHei', 'DengXian', 'SimHei', 'SimSun', 'Segoe UI']] = None
    font_color: str = Field(default="#333333", description="16进制颜色码")
    bold: bool = Field(default=False, description="是否加粗")
    alignment: Literal['left', 'center', 'right'] = Field(default='left', description="对齐方式")
    space_after: float = Field(default=12.0, ge=0, allow_inf_nan=False, description="段后距（Pt）")

class CanvasElement(BaseModel):
    """画布基础图元：全页通用的物理原子组件"""
    id: Optional[str] = None
    content_ref: Optional[str] = None
    type: Literal['textbox', 'shape', 'svg', 'image', 'icon', 'table', 'diagram'] = Field(description="图元类型")
    left: float = Field(allow_inf_nan=False)
    top: float = Field(allow_inf_nan=False)
    width: float = Field(allow_inf_nan=False)
    height: float = Field(allow_inf_nan=False)
    table: Optional[TableSpec] = None
    table_style: TableStyle = Field(default_factory=TableStyle)
    column_widths: List[float] = Field(default_factory=list)
    row_heights: List[float] = Field(default_factory=list)
    diagram: Optional[DiagramSpec] = None
    diagram_style: DiagramStyle = Field(default_factory=DiagramStyle)
    diagram_geometry: Optional[DiagramGeometry] = None
    
    # 文本框专属属性
    paragraphs: Optional[List[TextParagraph]] = Field(default=None, description="当 type=='textbox' 时必填")
    
    # 几何形状专属属性
    shape_type: Optional[Literal['rectangle', 'rounded_rectangle', 'ellipse', 'line', 'arrow']] = Field(
        default=None, 
        description="形状类型。支持矩形、圆角矩形、正圆、装饰分割线、流程引导箭头"
    )
    fill_color: Optional[str] = Field(default=None, description="填充色（16进制）")
    line_color: Optional[str] = Field(default=None, description="边框色（16进制），None表示无边框")
    native_arrow: bool = False
    
    # 矢量图专属属性 
    svg_code: Optional[str] = Field(default=None, description="当 type=='svg' 时由系统回填")
    svg_crop: Optional[List[float]] = Field(default=None, description="程序确认的透明外边距裁剪比例，旧版不裁剪")

    icon_name: Optional[str] = Field(default=None, description="当 type=='icon' 时必填，如 'user', 'cpu', 'globe'")
    icon_color: Optional[str] = Field(default=None, description="图标的渲染颜色，默认使用主色或辅助色")

   


    # 点缀图像专属属性
    image_path: Optional[str] = Field(default=None, description="当 type=='image' 时必填，本地图片绝对路径")

    @field_validator('svg_crop')
    @classmethod
    def valid_svg_crop(cls, value):
        if value is not None and (len(value) != 4 or not all(math.isfinite(v) for v in value)
                                  or not (0 <= value[0] < value[2] <= 1 and 0 <= value[1] < value[3] <= 1)):
            raise ValueError("SVG 裁剪比例无效")
        return value


class SlideCanvas(BaseModel):
    """物理层：单页幻灯片的画布元信息和图元集合"""
    slide_id: Optional[str] = None
    layout_resolved: bool = False
    background_color: str = Field(default="#FFFFFF", description="单页画布的物理背景颜色")
    elements: List[CanvasElement] = Field(default_factory=list, description="物理图元集合，由下至上层叠渲染")

class PPTFinalBlueprint(BaseModel):
    """物理层：最终交付给渲染器的图元硬编码蓝图"""
    schema_version: Literal[1, 2] = 1
    canvases: List[SlideCanvas] = Field(description="多页幻灯片画布集合")
