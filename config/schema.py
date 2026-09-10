# models/schema.py
import re
from pydantic import BaseModel, Field
from typing import List, Optional, Literal, Dict
from pydantic import BaseModel, Field, field_validator


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
    layout_type: Literal['title_page', 'standard_text', 'column', 'timeline', 'svg'] = Field(
        description="推荐的排版布局。若有逻辑流程、表格绘制、架构图，务必选择 'svg'"
    )
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
    topic: str = Field(description="PPT的全局核心主题")
    slides: List[SlideElement] = Field(description="幻灯片页面列表")

class TextParagraph(BaseModel):
    """物理层：单个段落的精细物理控制"""
    text: str = Field(description="文本内容")
    font_size: float = Field(default=14.0, description="字号（Pt）")
    font_color: str = Field(default="#333333", description="16进制颜色码")
    bold: bool = Field(default=False, description="是否加粗")
    alignment: Literal['left', 'center', 'right'] = Field(default='left', description="对齐方式")
    space_after: float = Field(default=12.0, description="段后距（Pt）")

class CanvasElement(BaseModel):
    """画布基础图元：全页通用的物理原子组件"""
    type: Literal['textbox', 'shape', 'svg', 'image', 'icon'] = Field(description="图元类型")
    left: float
    top: float
    width: float
    height: float
    
    # 文本框专属属性
    paragraphs: Optional[List[TextParagraph]] = Field(default=None, description="当 type=='textbox' 时必填")
    
    # 几何形状专属属性
    shape_type: Optional[Literal['rectangle', 'rounded_rectangle', 'ellipse', 'line', 'arrow']] = Field(
        default=None, 
        description="形状类型。支持矩形、圆角矩形、正圆、装饰分割线、流程引导箭头"
    )
    fill_color: Optional[str] = Field(default=None, description="填充色（16进制）")
    line_color: Optional[str] = Field(default=None, description="边框色（16进制），None表示无边框")
    
    # 矢量图专属属性 
    svg_code: Optional[str] = Field(default=None, description="当 type=='svg' 时由系统回填")
    icon_name: Optional[str] = Field(default=None, description="当 type=='icon' 时必填，如 'user', 'cpu', 'globe'")
    icon_color: Optional[str] = Field(default=None, description="图标的渲染颜色，默认使用主色或辅助色")

   


    # 点缀图像专属属性
    image_path: Optional[str] = Field(default=None, description="当 type=='image' 时必填，本地图片绝对路径")

class SlideCanvas(BaseModel):
    """物理层：单页幻灯片的画布元信息和图元集合"""
    background_color: str = Field(default="#FFFFFF", description="单页画布的物理背景颜色")
    elements: List[CanvasElement] = Field(default_factory=list, description="物理图元集合，由下至上层叠渲染")

class PPTFinalBlueprint(BaseModel):
    """物理层：最终交付给渲染器的图元硬编码蓝图"""
    canvases: List[SlideCanvas] = Field(description="多页幻灯片画布集合")