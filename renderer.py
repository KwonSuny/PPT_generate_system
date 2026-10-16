import io
import math
import os
import re
import unicodedata
from typing import Iterable, List, Optional, Tuple

from config_manager import configure_cairo_runtime, RenderDependencyError

try:
    configure_cairo_runtime()
    import cairosvg
except (ImportError, OSError) as exc:
    raise RenderDependencyError("SVG 渲染运行时不可用；请配置 PPT_CAIRO_BIN，并运行 python worker.py --check-runtime 检查") from exc
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Inches, Pt

from config.schema import CanvasElement, PPTFinalBlueprint, SlideCanvas, DiagramGeometry
from config.sanitizers import snap_font_size
from config.content import ContentConstraintError, cell_text


class PPTXRenderer:

    # 全局字族（latin + 东亚双轨，写入段落级与主题级）
    FONT_LATIN = "Segoe UI"
    FONT_EAST_ASIAN = "Microsoft YaHei"

    def __init__(self, blueprint: PPTFinalBlueprint):
        self.blueprint = blueprint
        self.prs = Presentation()
        self.slide_w = 13.333
        self.slide_h = 7.5
        self.prs.slide_width = Inches(self.slide_w)
        self.prs.slide_height = Inches(self.slide_h)
        self.lint_report: List[str] = []
        self._fixed_layout = False

    # ------------------------------------------------------------------
    # 基础工具
    # ------------------------------------------------------------------

    def _hex_to_rgb(self, hex_str: str) -> RGBColor:
        hex_str = (hex_str or "#333333").strip().lstrip("#")
        if len(hex_str) == 3:
            hex_str = "".join(ch * 2 for ch in hex_str)
        if len(hex_str) != 6:
            hex_str = "333333"
        return RGBColor(*(int(hex_str[i:i + 2], 16) for i in (0, 2, 4)))

    def _relative_luminance(self, hex_str: str) -> float:
        hex_str = (hex_str or "#FFFFFF").strip().lstrip("#")
        if len(hex_str) == 3:
            hex_str = "".join(ch * 2 for ch in hex_str)
        if len(hex_str) != 6:
            hex_str = "FFFFFF"
        values = [int(hex_str[i:i + 2], 16) / 255.0 for i in (0, 2, 4)]
        linear = [
            v / 12.92 if v <= 0.03928 else ((v + 0.055) / 1.055) ** 2.4
            for v in values
        ]
        return 0.2126 * linear[0] + 0.7152 * linear[1] + 0.0722 * linear[2]

    def _contrast_ratio(self, fg: str, bg: str) -> float:
        l1 = self._relative_luminance(fg)
        l2 = self._relative_luminance(bg)
        lighter, darker = max(l1, l2), min(l1, l2)
        return (lighter + 0.05) / (darker + 0.05)

    def _safe_text_color(self, font_color: str, background_color: str) -> str:
        """保留可读的原配色，否则调整明度以满足对比度。"""
        if self._contrast_ratio(font_color, background_color) >= 4.5:
            return font_color
        dark_ratio = self._contrast_ratio("#111827", background_color)
        light_ratio = self._contrast_ratio("#FFFFFF", background_color)
        target = (17,24,39) if dark_ratio >= light_ratio else (255,255,255)
        original = tuple(self._hex_to_rgb(font_color))
        for step in range(1,11):
            mix = step/10
            adjusted = '#'+''.join(f'{round(a*(1-mix)+b*mix):02X}' for a,b in zip(original,target))
            if self._contrast_ratio(adjusted,background_color)>=4.5:
                return adjusted
        return "#111827" if dark_ratio >= light_ratio else "#FFFFFF"

    @staticmethod
    def effective_background(element, canvas):
        color = canvas.background_color
        for plate in canvas.elements:
            if plate.type=='shape' and plate.fill_color and plate.shape_type not in {'line','arrow'}:
                if (plate.left <= element.left+.02 and plate.top <= element.top+.02
                    and plate.left+plate.width >= element.left+element.width-.02
                    and plate.top+plate.height >= element.top+element.height-.02):
                    color = plate.fill_color
        return color

    @staticmethod
    def background_samples(element, canvas):
        colors=set()
        for x in (.1,.5,.9):
            for y in (.15,.5,.85):
                px,py=element.left+element.width*x,element.top+element.height*y
                color=canvas.background_color
                for plate in canvas.elements:
                    if (plate.type=='shape' and plate.fill_color and plate.shape_type not in {'line','arrow'}
                        and plate.left<=px<=plate.left+plate.width and plate.top<=py<=plate.top+plate.height):
                        color=plate.fill_color
                colors.add(color)
        return colors

    @classmethod
    def fit_readable_text(cls, canvas, floors):
        from config.sanitizers import FONT_SCALE_STEPS
        from config.layout_budget import font_floor
        result=canvas.model_copy(deep=True)
        for element in result.elements:
            if element.type!='textbox':
                continue
            for _ in range(len(FONT_SCALE_STEPS)):
                if cls.required_text_height(element)<=element.height:
                    break
                changed=False
                for p in element.paragraphs or []:
                    minimum=floors.get(p.content_ref,font_floor(p.content_ref))
                    choices=[s for s in FONT_SCALE_STEPS if minimum<=s<p.font_size]
                    if choices:
                        p.font_size=max(choices);changed=True
                if not changed:
                    break
        return result

    def _set_font_name(self, font_obj, latin: str = None, ea: str = None):
        """同时设置 latin 与东亚字族（PowerPoint 对中英文走不同字族槽位）。"""
        font_obj.name = latin or self.FONT_LATIN
        rPr = font_obj._rPr
        for tag in ("a:ea", "a:cs"):
            existing = rPr.find(qn(tag))
            if existing is not None:
                rPr.remove(existing)
        ea_el = rPr.makeelement(qn("a:ea"), {"typeface": ea or self.FONT_EAST_ASIAN})
        rPr.append(ea_el)

    def _patch_theme_fonts(self):
        """补丁默认模板的主题 fontScheme（major/minor 的 latin+ea），保证成品
        二次编辑时新增文本也落在同一字族，而不是回退 Calibri。

        theme 是独立 OPC part（theme1.xml），不在 slideMaster 的元素树内，
        需要通过 master 的主题关系定位后直接改写其 XML blob。
        """
        try:
            from pptx.opc.constants import RELATIONSHIP_TYPE as RT
            from lxml import etree

            master = self.prs.slide_masters[0]
            theme_part = master.part.part_related_by(RT.THEME)
            root = etree.fromstring(theme_part.blob)
            ns_a = "http://schemas.openxmlformats.org/drawingml/2006/main"

            changed = False
            for group_tag in ("majorFont", "minorFont"):
                group = root.find(f".//{{{ns_a}}}{group_tag}")
                if group is None:
                    continue
                latin = group.find(f"{{{ns_a}}}latin")
                if latin is not None:
                    latin.set("typeface", self.FONT_LATIN)
                    changed = True
                ea = group.find(f"{{{ns_a}}}ea")
                if ea is not None:
                    ea.set("typeface", self.FONT_EAST_ASIAN)
                    changed = True

            if changed:
                theme_part._blob = etree.tostring(
                    root, xml_declaration=True, encoding="UTF-8", standalone=True
                )
        except Exception as e:
            print(f"[Renderer Warning] Theme font patch failed: {e}")

    def _is_large_background(self, el: CanvasElement) -> bool:
        area = max(el.width, 0) * max(el.height, 0)
        slide_area = self.slide_w * self.slide_h
        covers_center = (
            el.left <= self.slide_w * 0.25
            and el.top <= self.slide_h * 0.25
            and (el.left + el.width) >= self.slide_w * 0.75
            and (el.top + el.height) >= self.slide_h * 0.75
        )
        return area >= slide_area * 0.45 and covers_center

    def _layer_key(self, el: CanvasElement) -> int:
        # Large images are treated as backgrounds. Text and icons render last.
        if el.type == "image" and self._is_large_background(el):
            return 0
        if el.type in ("shape", "svg", "image"):
            return 1
        if el.type in ("textbox", "icon"):
            return 2
        return 1

    def _normalize_element(self, el: CanvasElement) -> Optional[CanvasElement]:
        item = el.model_copy()
        if item.type=='shape' and item.shape_type=='line' and min(item.width,item.height)>=0 and max(item.width,item.height)>0:
            return item
        item.left = max(0.0, min(float(item.left), self.slide_w))
        item.top = max(0.0, min(float(item.top), self.slide_h))
        item.width = max(0.01, float(item.width))
        item.height = max(0.01, float(item.height))
        if item.left + item.width > self.slide_w:
            item.width = max(0.01, self.slide_w - item.left)
        if item.top + item.height > self.slide_h:
            item.height = max(0.01, self.slide_h - item.top)
        if item.width <= 0.02 or item.height <= 0.02:
            return None
        return item

    # ------------------------------------------------------------------
    # 渲染前几何 lint：确定性检查 + 可自动修复项直接修复
    # ------------------------------------------------------------------

    def _lint_canvas(self, elements: List[CanvasElement], page_no: int, shrink_svg=True):
        """渲染前对单页画布做几何检查。图标-文本错位这类可确定性修复的
        直接修复；重叠等复杂问题只记录，不打断渲染。"""
        textboxes = [el for el in elements if el.type == "textbox" and el.paragraphs]
        icons = [el for el in elements if el.type == "icon"]

        # 1) 图标与相邻文本框垂直中心对齐（错位问题的根治点）
        for icon in icons:
            partner = self._find_icon_partner(icon, textboxes)
            if partner is None:
                continue
            # 新排版只在邻区留白足够时贴齐锚点，否则交给模型重排整列。
            if not shrink_svg:
                left = partner.left - icon.width - .2
                top = partner.top + (partner.height-icon.height)/2
                blockers = [e for e in elements if e is not icon and e is not partner and e.type!='shape'
                            and min(top+icon.height, e.top+e.height)>max(top,e.top)
                            and left < e.left+e.width+.28 and e.left < left+icon.width+.28]
                if left >= 0 and top >= 0 and top+icon.height <= self.slide_h and not blockers:
                    icon.left = left
            icon_cy = icon.top + icon.height / 2
            text_cy = partner.top + partner.height / 2
            delta = abs(icon_cy - text_cy)
            if delta > 0.06:
                new_top = max(0.0, partner.top + partner.height / 2 - icon.height / 2)
                self.lint_report.append(
                    f"p{page_no} icon '{icon.icon_name}': 中心偏差 {delta:.2f}in，已重排至文本框中心"
                )
                icon.top = new_top

        # 2) 图标与文本框水平间距不足（呼吸间距）
        for icon in icons:
            partner = self._find_icon_partner(icon, textboxes)
            if partner is None:
                continue
            gap = partner.left - (icon.left + icon.width)
            if 0 <= gap < 0.12:
                self.lint_report.append(
                    f"p{page_no} icon '{icon.icon_name}': 与文本框间距 {gap:.2f}in 偏窄"
                )

        # 3) 元素两两重叠（textbox 之间才算问题；装饰 shape/svg 允许衬底）
        for i in range(len(textboxes)):
            for j in range(i + 1, len(textboxes)):
                a, b = textboxes[i], textboxes[j]
                if self._boxes_overlap(a, b):
                    self.lint_report.append(
                        f"p{page_no} textbox 重叠: '{(a.paragraphs[0].text or '')[:10]}' vs '{(b.paragraphs[0].text or '')[:10]}'"
                    )

        # 4) SVG 与其他元素重叠：SVG 是页面视觉主体，确定性收缩其边界避让
        #    Coffee.pptx Slide5 实测：SVG 底 7.30 vs 卡片顶 5.72，重叠 1.42in，
        #    收缩 SVG 底边比挪一堆卡片更稳（SVG 高度由模型估算，初始猜错是常态）。
        svgs = [el for el in elements if shrink_svg and el.type == "svg" and el.svg_code]
        for svg_el in svgs:
            overlapping = [
                el for el in elements
                if el is not svg_el
                and el.type in ("textbox", "icon", "shape")
                and self._boxes_overlap(svg_el, el)
            ]
            if not overlapping:
                continue
            svg_cy = svg_el.top + svg_el.height / 2
            below = [el for el in overlapping if el.top + el.height / 2 >= svg_cy]
            if below:
                target_bottom = min(el.top for el in below) - 0.15
                new_h = target_bottom - svg_el.top
                if 0.5 <= new_h < svg_el.height - 0.05:
                    self.lint_report.append(
                        f"p{page_no} svg 与下方元素重叠，高度 {svg_el.height:.2f}in 收缩为 {new_h:.2f}in"
                    )
                    svg_el.height = new_h
                    continue
            above = [el for el in overlapping if el not in below]
            if above:
                target_top = max(el.top + el.height for el in above) + 0.15
                new_h = svg_el.top + svg_el.height - target_top
                if 0.5 <= new_h < svg_el.height - 0.05:
                    self.lint_report.append(
                        f"p{page_no} svg 与上方元素重叠，top {svg_el.top:.2f}in 下移为 {target_top:.2f}in"
                    )
                    svg_el.top = target_top
                    svg_el.height = new_h

        # 5) 迷你箭头视觉噪音：总宽 < 0.4in 的箭头等中心扩展到 0.5in
        for el in elements:
            if el.type == "shape" and el.shape_type == "arrow" and el.width < 0.4:
                center = el.left + el.width / 2
                new_left = max(0.0, center - 0.25)
                new_width = min(0.5, self.slide_w - new_left)
                if new_width > el.width + 0.02:
                    self.lint_report.append(
                        f"p{page_no} 迷你箭头宽 {el.width:.2f}in 放大为 {new_width:.2f}in"
                    )
                    el.left = new_left
                    el.width = new_width

    def _find_icon_partner(self, icon: CanvasElement, textboxes: List[CanvasElement]) -> Optional[CanvasElement]:
        """找图标右侧最近的文本框（水平并排组合）。"""
        if icon.content_ref:
            anchored = [tb for tb in textboxes if any(p.content_ref==icon.content_ref for p in tb.paragraphs or [])]
            return min(anchored, key=lambda tb: abs(tb.top-icon.top)) if anchored else None
        icon_right = icon.left + icon.width
        icon_cy = icon.top + icon.height / 2
        best, best_dist = None, None
        for tb in textboxes:
            if tb.left < icon_right - 0.05:
                continue  # 不在图标右侧
            gap = tb.left - icon_right
            if gap > 2.5:
                continue  # 太远，不算组合
            tb_cy = tb.top + tb.height / 2
            vertical_dist = abs(tb_cy - icon_cy)
            score = gap + vertical_dist
            if best_dist is None or score < best_dist:
                best, best_dist = tb, score
        return best

    @staticmethod
    def _boxes_overlap(a: CanvasElement, b: CanvasElement) -> bool:
        return (
            a.left < b.left + b.width - 0.05
            and b.left < a.left + a.width - 0.05
            and a.top < b.top + b.height - 0.05
            and b.top < a.top + a.height - 0.05
        )

    # ------------------------------------------------------------------
    # 主渲染入口
    # ------------------------------------------------------------------

    @staticmethod
    def _text_lines(text, font_size, width):
        if width <= 0:
            return math.inf
        capacity = width * 72 / (font_size * 1.04)
        lines = 0
        for line in text.split("\n"):
            units = sum(1 if unicodedata.east_asian_width(ch) in {"W", "F"} else 0.56 for ch in line)
            lines += max(1, math.ceil(units / max(capacity, 0.01)))
        return lines

    @classmethod
    def _text_height(cls, text, font_size, width):
        return cls._text_lines(text, font_size, width) * font_size * 1.22 / 72 + 0.16

    @classmethod
    def required_text_height(cls, element):
        paragraphs = element.paragraphs or []
        return sum(cls._text_lines(p.text, p.font_size, element.width - 0.22) * p.font_size * 1.22
                   + (p.space_after if i < len(paragraphs) - 1 else 0)
                   for i, p in enumerate(paragraphs)) / 72 + 0.16

    @classmethod
    def validate_capacity(cls, element):
        """固定字号与容器；容量不足时拒绝，不能缩字或改变几何。"""
        if element.type == "textbox":
            demand = cls.required_text_height(element)
            if demand > element.height + 0.005:
                raise ContentConstraintError(f"文本框 {element.id or ''} 内容超出容量：需要约 {demand:.2f}in，现有 {element.height:.2f}in；请精简文字或选择重新排版")
        elif element.type == "table":
            if not element.table or len(element.column_widths) != len(element.table.columns) or len(element.row_heights) != len(element.table.rows) + 1:
                raise ContentConstraintError("表格缺少可锁定的列宽或行高")
            if any(v <= 0 or not math.isfinite(v) for v in element.column_widths + element.row_heights):
                raise ContentConstraintError("表格列宽与行高必须为有限正数")
            if abs(sum(element.column_widths) - element.width) > 0.005 or sum(element.row_heights) > element.height + 0.005:
                raise ContentConstraintError("表格几何与容器尺寸不一致")
            style, spec = element.table_style, element.table
            rows = [[column.title for column in spec.columns]] + [
                [cell_text(cell, column) for cell, column in zip(row.cells, spec.columns)] for row in spec.rows
            ]
            for i, row in enumerate(rows):
                size = style.header_font_size if i == 0 else style.font_size
                for j, text in enumerate(row):
                    needed = cls._text_height(text, size, element.column_widths[j] - 0.22)
                    if needed > element.row_heights[i] + 0.005:
                        raise ContentConstraintError(f"表格第 {i + 1} 行第 {j + 1} 列文字超出容量；请精简内容或重新排版")
            if spec.source:
                needed = cls._text_height(spec.source, style.source_font_size, element.width - 0.22)
                if needed > element.height - sum(element.row_heights) + 0.005:
                    raise ContentConstraintError("表格来源说明超出预留区域；请精简说明或重新排版")
        elif element.type == "diagram":
            if not element.diagram or not element.diagram_geometry:
                raise ContentConstraintError("关系图缺少可锁定的节点与连接线几何")
            nodes = {node.id: node for node in element.diagram_geometry.nodes}
            edges = {edge.id: edge for edge in element.diagram_geometry.edges}
            if len(nodes) != len(element.diagram_geometry.nodes) or len(edges) != len(element.diagram_geometry.edges) or set(nodes) != {n.id for n in element.diagram.nodes} or set(edges) != {e.id for e in element.diagram.edges}:
                raise ContentConstraintError("关系图语义与物理 ID 不一致")
            for node in element.diagram.nodes:
                box = nodes[node.id]
                if box.left < 0 or box.top < 0 or box.left + box.width > element.width + 0.001 or box.top + box.height > element.height + 0.001:
                    raise ContentConstraintError("关系图节点超出容器范围")
                if cls._text_height(node.label, element.diagram_style.font_size, box.width - 0.22) > box.height + 0.005:
                    raise ContentConstraintError(f"节点 {node.id} 文字超出容量；请精简文字或重新排版")
            for edge in element.diagram.edges:
                route = edges[edge.id]
                if route.source_port == "auto" or route.target_port == "auto":
                    raise ContentConstraintError("关系图物理端口必须已解析")
                expected = cls._port(nodes[edge.source], route.source_port) + cls._port(nodes[edge.target], route.target_port)
                if any(abs(actual - target) > 0.001 for actual, target in zip((route.start_x, route.start_y, route.end_x, route.end_y), expected)):
                    raise ContentConstraintError("连接线端点与节点端口不对齐")
                box = route.label_box
                if edge.label and (box is None or cls._text_height(edge.label, element.diagram_style.label_font_size, box.width - 0.10) > box.height + 0.005):
                    raise ContentConstraintError(f"连接线 {edge.id} 标签超出预留区域；请精简标签或重新排版")

    @classmethod
    def _resolve_table(cls, element):
        if not element.table:
            raise ContentConstraintError("表格图元必须包含 TableSpec")
        spec, style = element.table, element.table_style
        weights = []
        for i, column in enumerate(spec.columns):
            texts = [column.title] + [cell_text(row.cells[i], column) for row in spec.rows]
            units = max(sum(1 if unicodedata.east_asian_width(ch) in {"W", "F"} else 0.56 for ch in text) for text in texts)
            weights.append(column.width_weight or max(1, math.sqrt(units)))
        minimum = 0.65
        remaining = element.width - minimum * len(weights)
        if remaining <= 0:
            raise ContentConstraintError("表格列数超过当前容器容量，请减少列数或扩大表格区域")
        element.column_widths = [minimum + remaining * weight / sum(weights) for weight in weights]
        rows = [[column.title for column in spec.columns]] + [
            [cell_text(cell, column) for cell, column in zip(row.cells, spec.columns)] for row in spec.rows
        ]
        heights = []
        for i, row in enumerate(rows):
            size = style.header_font_size if i == 0 else style.font_size
            heights.append(max(0.45, max(cls._text_height(text, size, width - 0.22) for text, width in zip(row, element.column_widths))))
        source_height = cls._text_height(spec.source, style.source_font_size, element.width - 0.22) if spec.source else 0
        remaining = element.height - source_height - sum(heights)
        if remaining < -0.005:
            raise ContentConstraintError(f"表格文字需要约 {sum(heights) + source_height:.2f}in 高度，当前只有 {element.height:.2f}in；请减少内容或重新排版")
        element.row_heights = [height + max(0, remaining) / len(heights) for height in heights]
        cls.validate_capacity(element)

    @staticmethod
    def _port(box, port):
        return {"top": (box.left + box.width / 2, box.top), "left": (box.left, box.top + box.height / 2),
                "bottom": (box.left + box.width / 2, box.top + box.height), "right": (box.left + box.width, box.top + box.height / 2)}[port]

    @classmethod
    def _resolve_diagram(cls, element):
        if not element.diagram:
            raise ContentConstraintError("关系图图元必须包含 DiagramSpec")
        spec = element.diagram
        incoming = {node.id: [] for node in spec.nodes}
        outgoing = {node.id: [] for node in spec.nodes}
        for edge in spec.edges:
            incoming[edge.target].append(edge.source)
            outgoing[edge.source].append(edge.target)
        pending = {key: len(value) for key, value in incoming.items()}
        ready = [node.id for node in spec.nodes if not pending[node.id]]
        ranks, order = {}, []
        while ready:
            node_id = ready.pop(0)
            ranks[node_id] = max((ranks[source] + 1 for source in incoming[node_id]), default=0)
            order.append(node_id)
            for target in outgoing[node_id]:
                pending[target] -= 1
                if pending[target] == 0:
                    ready.append(target)
        if len(order) != len(spec.nodes):
            raise ContentConstraintError("原生关系图暂不支持环路；复杂拓扑请使用 SVG")
        layers = [[node.id for node in spec.nodes if ranks[node.id] == rank] for rank in range(max(ranks.values()) + 1)]
        horizontal = spec.direction == "LR"
        major, minor = (element.width, element.height) if horizontal else (element.height, element.width)
        label_width = max((sum(1 if unicodedata.east_asian_width(ch) in {"W", "F"} else 0.56 for ch in edge.label) * element.diagram_style.label_font_size * 1.04 / 72 + 0.10 for edge in spec.edges if edge.label), default=0)
        gap = max(1.15, (max(0.55, label_width) + 0.08) * 2) if horizontal else max(0.70, (element.diagram_style.label_font_size * 1.22 / 72 + 0.24) * 2)
        node_major = (major - gap * (len(layers) - 1)) / len(layers)
        node_minor = min(1.15 if horizontal else 3.0, (minor - 0.30 * (max(map(len, layers)) - 1)) / max(map(len, layers)))
        if node_major < (1.1 if horizontal else 0.6) or node_minor < (0.55 if horizontal else 1.1):
            raise ContentConstraintError("关系图节点过密，请减少节点或使用 SVG")
        boxes = {}
        for rank, layer in enumerate(layers):
            spacing = minor / len(layer)
            for i, node_id in enumerate(layer):
                a = rank * (node_major + gap)
                b = spacing * (i + 0.5) - node_minor / 2
                boxes[node_id] = {"id": node_id, "left": a if horizontal else b, "top": b if horizontal else a,
                                  "width": node_major if horizontal else node_minor, "height": node_minor if horizontal else node_major}
        geometry = DiagramGeometry.model_validate({"nodes": list(boxes.values()), "edges": []})
        by_id = {box.id: box for box in geometry.nodes}
        routes = []
        for edge in spec.edges:
            if ranks[edge.target] != ranks[edge.source] + 1:
                raise ContentConstraintError("跨层连接可能遮挡节点；请拆分关系或使用 SVG")
            source_port, target_port = ("right", "left") if horizontal else ("bottom", "top")
            if edge.source_port not in {"auto", source_port} or edge.target_port not in {"auto", target_port}:
                raise ContentConstraintError("连接端口与图的方向冲突；请调整端口或使用 SVG")
            sx, sy = cls._port(by_id[edge.source], source_port)
            tx, ty = cls._port(by_id[edge.target], target_port)
            box = None
            if edge.label:
                size = element.diagram_style.label_font_size
                label_h = size * 1.22 / 72 + 0.16
                label_w = max(0.55, sum(1 if unicodedata.east_asian_width(ch) in {"W", "F"} else 0.56 for ch in edge.label) * size * 1.04 / 72 + 0.10)
                candidates = []
                if horizontal:
                    available = gap if abs(sy - ty) < 0.01 else gap / 2
                    if label_w > available - 0.06:
                        raise ContentConstraintError(f"连接线 {edge.id} 标签过长，请精简或扩大图的区域")
                    candidates.append({"left": sx + (available - label_w) / 2, "top": sy - label_h - 0.04, "width": label_w, "height": label_h})
                    if abs(sy - ty) >= 0.01:
                        candidates.append({"left": tx - available + (available - label_w) / 2, "top": ty - label_h - 0.04, "width": label_w, "height": label_h})
                else:
                    available = gap if abs(sx - tx) < 0.01 else gap / 2
                    candidates.append({"left": sx + 0.08, "top": sy + (available - label_h) / 2, "width": label_w, "height": label_h})
                    if abs(sx - tx) >= 0.01:
                        candidates.append({"left": tx + 0.08, "top": ty - available + (available - label_h) / 2, "width": label_w, "height": label_h})
                from config.schema import GeometryBox
                existing_labels = [GeometryBox.model_validate(route["label_box"]) for route in routes if route["label_box"]]
                for candidate in candidates:
                    trial = GeometryBox.model_validate(candidate)
                    inside = trial.left >= 0 and trial.top >= 0 and trial.left + trial.width <= element.width and trial.top + trial.height <= element.height
                    if inside and not any(cls._boxes_overlap(trial, other) for other in existing_labels + list(by_id.values())):
                        box = candidate
                        break
                if box is None:
                    raise ContentConstraintError(f"连接线 {edge.id} 标签区域冲突，请精简标签或使用 SVG")
            routes.append({"id": edge.id, "source_port": source_port, "target_port": target_port,
                           "start_x": sx, "start_y": sy, "end_x": tx, "end_y": ty, "label_box": box})
        element.diagram_geometry = DiagramGeometry.model_validate({"nodes": list(boxes.values()), "edges": routes})
        cls.validate_capacity(element)

    @classmethod
    def _reflow_text_regions(cls, canvas, slide_h, max_move=1.0):
        """最终候选只在原列和图示间的区域内收紧空白，保留阅读顺序。"""
        anchors = [element for element in canvas.elements if element.type in {"table", "diagram", "svg", "image"}]
        regions = {}
        for element in canvas.elements:
            if element.type != "textbox":
                continue
            related = [anchor for anchor in anchors if min(element.left + element.width, anchor.left + anchor.width) - max(element.left, anchor.left) > 0.05]
            if any(anchor.top < element.top < anchor.top + anchor.height for anchor in related):
                continue
            start = max([0.15] + [anchor.top + anchor.height + 0.15 for anchor in related if anchor.top + anchor.height <= element.top])
            end = min([slide_h - 0.15] + [anchor.top - 0.15 for anchor in related if anchor.top >= element.top])
            regions.setdefault((start, end), []).append(element)
        for (start, end), elements in regions.items():
            if not any(cls.required_text_height(element) > element.height + 0.005 or element.top + element.height > end + 0.005 for element in elements):
                continue
            rows = []
            for element in sorted(elements, key=lambda element: (element.top, element.left)):
                if not rows or abs(rows[-1][0].top - element.top) >= 0.03:
                    rows.append([])
                rows[-1].append(element)
            proposals = []
            valid = True
            for row in rows:
                top = start
                for element in row:
                    if any(other is not element and min(element.left + element.width, other.left + other.width) - max(element.left, other.left) > 0.05 for other in row):
                        valid = False
                        break
                    predecessors = [(other, previous_top, height) for other, previous_top, height in proposals
                                    if min(element.left + element.width, other.left + other.width) - max(element.left, other.left) > 0.05]
                    top = max([top] + [previous_top + height + min(0.16, max(0.05, element.top - other.top - other.height))
                                       for other, previous_top, height in predecessors])
                if not valid:
                    break
                # 同一行的各列先统一顶边，再按各自文字高度安排后继。
                for element in row:
                    height = cls.required_text_height(element) + 0.01
                    if top + height > end + 0.001 or abs(top - element.top) > max_move:
                        valid = False
                        break
                    proposals.append((element, top, height))
                if not valid:
                    break
            if valid:
                for element, top, height in proposals:
                    element.top, element.height = top, height

    @classmethod
    def resolve_canvas(cls, canvas, legacy=False, fit_text=False, reflow_text=False, preserve_svg_crop=False):
        canvas = SlideCanvas.model_validate(canvas).model_copy(deep=True)
        if canvas.layout_resolved:
            return canvas
        renderer = cls(PPTFinalBlueprint(canvases=[canvas]))
        if not legacy:
            for element in canvas.elements:
                line = element.type=='shape' and element.shape_type=='line' and min(element.width,element.height)>=0 and max(element.width,element.height)>0
                if element.left < 0 or element.top < 0 or ((element.width <= 0 or element.height <= 0) and not line) or element.left + element.width > renderer.slide_w + 0.001 or (element.top + element.height > renderer.slide_h + 0.001 and not (reflow_text and fit_text and element.type == "textbox")):
                    raise ContentConstraintError(f"图元 {element.id or element.type} 超出画布范围")
        canvas.elements = [item for item in (renderer._normalize_element(el) for el in canvas.elements) if item is not None]
        renderer._lint_canvas(canvas.elements, 1, shrink_svg=legacy)
        if not legacy and fit_text and reflow_text:
            for element in canvas.elements:
                for paragraph in (element.paragraphs or []) if element.type == "textbox" else []:
                    paragraph.font_size = snap_font_size(paragraph.font_size)
                    paragraph.space_after = min(paragraph.space_after, 8)
            cls._reflow_text_regions(canvas, renderer.slide_h)
            renderer._lint_canvas(canvas.elements, 1, shrink_svg=False)
        for element in canvas.elements:
            if element.type == "textbox":
                scale = renderer._fit_font_scale(element.paragraphs, element.width, element.height) if legacy else 1.0
                for paragraph in element.paragraphs or []:
                    paragraph.font_size = snap_font_size(paragraph.font_size * scale)
                    if scale < 1:
                        paragraph.space_after = min(paragraph.space_after, 8)
                    paragraph.font_color = renderer._safe_text_color(paragraph.font_color, renderer.effective_background(element,canvas))
                if not legacy:
                    if fit_text:
                        demand = cls.required_text_height(element)
                        blockers = [other.top - 0.05 for other in canvas.elements if other is not element
                                    and other.type in {"textbox", "table", "diagram", "svg", "image", "icon"}
                                    and other.top >= element.top
                                    and other.left < element.left + element.width
                                    and element.left < other.left + other.width]
                        available = min([renderer.slide_h] + blockers) - element.top
                        # 只在新排版的空白区域内扩高；已保存的文本修改不会进入此分支。
                        if element.height < demand <= available:
                            element.height = demand
                    cls.validate_capacity(element)
            elif element.type == "table":
                cls._resolve_table(element)
            elif element.type == "diagram":
                cls._resolve_diagram(element)
            elif element.type == "svg" and not legacy:
                if not renderer._svg_viewbox_ratio(element.svg_code or ""):
                    raise ContentConstraintError("SVG 必须声明有效的 viewBox 才能保证缩放比例")
                from config.layout_budget import svg_metrics
                if not preserve_svg_crop:
                    element.svg_crop = svg_metrics(element.svg_code or '').get('crop')
            elif element.type == "shape" and element.shape_type == "arrow" and not legacy:
                element.native_arrow = True
        renderer._lint_canvas(canvas.elements, 1, shrink_svg=legacy)
        for element in canvas.elements:
            if not legacy and (element.top < 0 or element.top + element.height > renderer.slide_h + 0.001):
                raise ContentConstraintError(f"图元 {element.id or element.type} 超出画布范围")
            if element.type in {"table", "diagram"} and any(renderer._boxes_overlap(element, other) for other in canvas.elements if other is not element and other.type in {"textbox", "table", "diagram", "svg", "image", "icon"}):
                raise ContentConstraintError("表格或关系图与其他内容重叠，请重新排版")
        canvas.layout_resolved = True
        return canvas

    def render_all(self, output_path: str = "output.pptx"):
        print("[Renderer Engine] Rendering structured blueprint to PPTX...")
        self._patch_theme_fonts()
        blank_layout = self.prs.slide_layouts[6]

        for page_no, canvas_data in enumerate(self.blueprint.canvases, 1):
            slide = self.prs.slides.add_slide(blank_layout)
            bg_color = canvas_data.background_color or "#FFFFFF"

            bg = slide.background
            fill = bg.fill
            fill.solid()
            fill.fore_color.rgb = self._hex_to_rgb(bg_color)

            self._fixed_layout = canvas_data.layout_resolved
            normalized = canvas_data.elements if self._fixed_layout else [
                item
                for item in (self._normalize_element(el) for el in canvas_data.elements)
                if item is not None
            ]
            # 渲染前几何 lint（在排序/绘制前执行，修复结果直接生效）
            if not self._fixed_layout:
                self._lint_canvas(normalized, page_no)

            for el in sorted(normalized, key=self._layer_key):
                if el.type == "textbox":
                    self._execute_draw_textbox(slide, el, self.effective_background(el,canvas_data))
                elif el.type == "shape":
                    self._execute_draw_shape(slide, el)
                elif el.type == "svg":
                    self._execute_draw_svg(slide, el)
                elif el.type == "image":
                    self._execute_draw_image(slide, el, bg_color)
                elif el.type == "icon":
                    self._execute_draw_icon(slide, el, bg_color)
                elif el.type == "table":
                    if not el.column_widths:
                        self._resolve_table(el)
                    self._execute_draw_table(slide, el, self.effective_background(el,canvas_data))
                elif el.type == "diagram":
                    if not el.diagram_geometry:
                        self._resolve_diagram(el)
                    self._execute_draw_diagram(slide, el)

        if self.lint_report:
            print(f"[Renderer Lint] {len(self.lint_report)} 项几何检查记录：")
            for line in self.lint_report[:10]:
                print(f"  - {line}")
            if len(self.lint_report) > 10:
                print(f"  ... 共 {len(self.lint_report)} 项")

        self.prs.save(output_path)
        print(f"[Renderer Engine] PPTX saved to: {output_path}")

    # ------------------------------------------------------------------
    # 文本框：字号量化 / 字体强制 / 锚定 / 单次缩放
    # ------------------------------------------------------------------

    def _estimated_line_count(self, text: str, font_size: float, width_in: float) -> int:
        if not text:
            return 1
        # 中文字符全宽、ASCII 半宽的混合宽度估计（此前 0.95 系数对纯中文低估）
        cn_chars = sum(1 for ch in text if ord(ch) > 0x2E7F)
        eff_len = cn_chars + (len(text) - cn_chars) * 0.55
        chars_per_line = max(int((width_in * 72) / max(font_size * 1.02, 1)), 1)
        return max(int((eff_len + chars_per_line - 1) // chars_per_line), 1)

    def _fit_font_scale(self, paragraphs: Iterable, width_in: float, height_in: float) -> float:
        """按逐段累加真实需求高度（此前用 max_font × 全部行数，混合字号严重高估）。"""
        paragraphs = list(paragraphs or [])
        if not paragraphs:
            return 1.0
        required_pt = 0.0
        for p in paragraphs:
            lines = self._estimated_line_count(p.text, p.font_size, width_in)
            required_pt += lines * p.font_size * 1.22
        required_pt += max(len(paragraphs) - 1, 0) * 4
        available_pt = max((height_in - 0.16) * 72, 8)
        if required_pt <= available_pt:
            return 1.0
        return max(0.62, min(1.0, available_pt / required_pt))

    def _execute_draw_textbox(self, slide, el: CanvasElement, background_color: str):
        tx_box = slide.shapes.add_textbox(
            Inches(el.left), Inches(el.top), Inches(el.width), Inches(el.height)
        )
        tf = tx_box.text_frame
        tf.word_wrap = True
        # 去掉 PowerPoint 二次缩放：缩放只由 _fit_font_scale 确定性计算一次
        tf.auto_size = MSO_AUTO_SIZE.NONE
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = tf.margin_right = Inches(0.11)
        tf.margin_top = tf.margin_bottom = Inches(0.08)

        if not el.paragraphs:
            return

        align_mapping = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}
        font_scale = 1.0 if self._fixed_layout else self._fit_font_scale(el.paragraphs, el.width, el.height)

        for idx, p_data in enumerate(el.paragraphs):
            p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
            p.text = p_data.text
            # 字号：先按内容自适应缩放，再吸附排版刻度，消除碎字号
            scaled = p_data.font_size * font_scale
            quantized = p_data.font_size if self._fixed_layout else snap_font_size(scaled)
            p.font.size = Pt(quantized)
            self._set_font_name(p.font, latin=p_data.font_family, ea=p_data.font_family)
            p.font.color.rgb = self._hex_to_rgb(
                p_data.font_color if self._fixed_layout else self._safe_text_color(p_data.font_color, background_color)
            )
            p.font.bold = p_data.bold
            p.alignment = align_mapping.get(p_data.alignment, PP_ALIGN.LEFT)
            p.space_after = Pt(max(0, min(p_data.space_after, 8)) if font_scale < 1 else p_data.space_after)

    def _native_text(self, text_frame, text, size, color, bold=False, alignment="left"):
        text_frame.clear()
        text_frame.word_wrap = True
        text_frame.auto_size = MSO_AUTO_SIZE.NONE
        text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        text_frame.margin_left = text_frame.margin_right = Inches(0.11)
        text_frame.margin_top = text_frame.margin_bottom = Inches(0.08)
        paragraph = text_frame.paragraphs[0]
        paragraph.text = text
        paragraph.font.size = Pt(size)
        paragraph.font.color.rgb = self._hex_to_rgb(color)
        paragraph.font.bold = bold
        paragraph.space_after = Pt(0)
        paragraph.alignment = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}[alignment]
        self._set_font_name(paragraph.font)

    def _execute_draw_table(self, slide, el, background_color='#FFFFFF'):
        self.validate_capacity(el)
        spec, style = el.table, el.table_style
        frame = slide.shapes.add_table(len(spec.rows) + 1, len(spec.columns), Inches(el.left), Inches(el.top), Inches(el.width), Inches(sum(el.row_heights)))
        frame.name = el.id or "table"
        table = frame.table
        table.first_row = table.first_col = table.horz_banding = table.vert_banding = False
        for i, width in enumerate(el.column_widths):
            table.columns[i].width = Inches(width)
        for i, height in enumerate(el.row_heights):
            table.rows[i].height = Inches(height)
        for i in range(len(spec.rows) + 1):
            for j, column in enumerate(spec.columns):
                cell = table.cell(i, j)
                raw = spec.rows[i - 1].cells[j] if i else None
                text = column.title if i == 0 else cell_text(raw, column)
                fill = style.header_fill if i == 0 else (style.body_fill if i % 2 else style.alternate_fill)
                color = style.header_color if i == 0 else style.font_color
                alignment = column.alignment
                if alignment == "auto":
                    alignment = "right" if raw and isinstance(raw.value, (int, float)) else "left"
                cell.fill.solid()
                cell.fill.fore_color.rgb = self._hex_to_rgb(fill)
                # 表格单元格的锚点属于 tcPr，不能只设置文本框的 bodyPr。
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                cell.margin_left = cell.margin_right = Inches(0.11)
                cell.margin_top = cell.margin_bottom = Inches(0.08)
                self._native_text(cell.text_frame, text, style.header_font_size if i == 0 else style.font_size,
                                  self._safe_text_color(color, fill), i == 0, alignment)
                properties = cell._tc.get_or_add_tcPr()
                for side in ("lnL", "lnR", "lnT", "lnB"):
                    existing = properties.find(qn("a:" + side))
                    if existing is not None:
                        properties.remove(existing)
                    border = OxmlElement("a:" + side)
                    border.set("w", str(int(Pt(0.7))))
                    solid = OxmlElement("a:solidFill")
                    rgb = OxmlElement("a:srgbClr")
                    rgb.set("val", style.border_color.lstrip("#"))
                    solid.append(rgb)
                    border.append(solid)
                    properties.append(border)
        if el.height - sum(el.row_heights) > 0.16:
            box = slide.shapes.add_textbox(Inches(el.left), Inches(el.top + sum(el.row_heights)), Inches(el.width), Inches(el.height - sum(el.row_heights)))
            box.name = (el.id or "table") + "-source"
            self._native_text(box.text_frame, spec.source or "", style.source_font_size, self._safe_text_color(style.font_color, background_color))

    @staticmethod
    def _arrow_head(connector):
        line = connector._element.spPr.get_or_add_ln()
        old = line.find(qn("a:tailEnd"))
        if old is not None:
            line.remove(old)
        ending = OxmlElement("a:tailEnd")
        ending.set("type", "triangle")
        ending.set("w", "sm")
        ending.set("len", "sm")
        line.append(ending)

    def _execute_draw_diagram(self, slide, el):
        self.validate_capacity(el)
        spec, style = el.diagram, el.diagram_style
        nodes, boxes = {}, {box.id: box for box in el.diagram_geometry.nodes}
        for node in spec.nodes:
            box = boxes[node.id]
            shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE if node.shape_type == "rectangle" else MSO_SHAPE.ROUNDED_RECTANGLE,
                                          Inches(el.left + box.left), Inches(el.top + box.top), Inches(box.width), Inches(box.height))
            shape.name = f"{el.id or 'diagram'}-node-{node.id}"
            shape.fill.solid()
            shape.fill.fore_color.rgb = self._hex_to_rgb(style.fill_color)
            shape.line.color.rgb = self._hex_to_rgb(style.line_color)
            shape.line.width = Pt(1.1)
            shape._element.spPr.append(OxmlElement("a:effectLst"))
            self._native_text(shape.text_frame, node.label, style.font_size, self._safe_text_color(style.font_color, style.fill_color), alignment="center")
            nodes[node.id] = shape
        routes = {route.id: route for route in el.diagram_geometry.edges}
        port_indices = {"top": 0, "left": 1, "bottom": 2, "right": 3}
        for edge in spec.edges:
            route = routes[edge.id]
            straight = abs(route.start_x - route.end_x) < 0.001 or abs(route.start_y - route.end_y) < 0.001
            connector = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT if straight else MSO_CONNECTOR.ELBOW,
                                                  Inches(el.left + route.start_x), Inches(el.top + route.start_y),
                                                  Inches(el.left + route.end_x), Inches(el.top + route.end_y))
            connector.name = f"{el.id or 'diagram'}-edge-{edge.id}"
            connector.begin_connect(nodes[edge.source], port_indices[route.source_port])
            connector.end_connect(nodes[edge.target], port_indices[route.target_port])
            connector.line.color.rgb = self._hex_to_rgb(style.line_color)
            connector.line.width = Pt(1.6)
            self._arrow_head(connector)
            # 连接线位于节点下方，端口绑定仍保留在 OOXML 中。
            first_node = next(iter(nodes.values()))._element
            first_node.addprevious(connector._element)
            if route.label_box:
                box = route.label_box
                label = slide.shapes.add_textbox(Inches(el.left + box.left), Inches(el.top + box.top), Inches(box.width), Inches(box.height))
                label.name = f"{el.id or 'diagram'}-label-{edge.id}"
                self._native_text(label.text_frame, edge.label, style.label_font_size, style.font_color, alignment="center")
                label.text_frame.margin_left = label.text_frame.margin_right = Inches(0.05)

    # ------------------------------------------------------------------
    # 形状
    # ------------------------------------------------------------------

    def _execute_draw_shape(self, slide, el: CanvasElement):
        if el.shape_type=='line':
            line=slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(el.left), Inches(el.top),
                Inches(el.left+el.width), Inches(el.top+el.height))
            line.line.color.rgb=self._hex_to_rgb(el.line_color or el.fill_color or '#CBD5E1')
            line.line.width=Pt(.8)
            return
        if el.shape_type in ("line", "arrow"):
            color = self._hex_to_rgb(el.line_color or el.fill_color or "#CBD5E1")
            line_width = max(min(el.height * 36, 1.8), 0.8)
            line_left = el.left + 0.08
            line_top = el.top + el.height / 2
            line_right = el.left + el.width - 0.08
            line = slide.shapes.add_connector(
                1,
                Inches(line_left),
                Inches(line_top),
                Inches(line_right),
                Inches(line_top),
            )
            line.line.color.rgb = color
            line.line.width = Pt(line_width)
            if el.shape_type == "arrow":
                if el.native_arrow:
                    self._arrow_head(line)
                    return
                head_size = max(min(el.height * 2.3, 0.2), 0.12)
                head_left = el.left + el.width - head_size - 0.04
                head_top = el.top + el.height / 2 - head_size / 2
                head = slide.shapes.add_shape(
                    MSO_SHAPE.ISOSCELES_TRIANGLE,
                    Inches(head_left),
                    Inches(head_top),
                    Inches(head_size),
                    Inches(head_size),
                )
                head.rotation = 90
                head.fill.solid()
                head.fill.fore_color.rgb = color
                head.line.fill.background()
            return

        shape_mapping = {
            "rectangle": MSO_SHAPE.RECTANGLE,
            "rounded_rectangle": MSO_SHAPE.ROUNDED_RECTANGLE,
            "ellipse": MSO_SHAPE.OVAL,
        }
        shape = slide.shapes.add_shape(
            shape_mapping.get(el.shape_type, MSO_SHAPE.RECTANGLE),
            Inches(el.left),
            Inches(el.top),
            Inches(el.width),
            Inches(el.height),
        )

        if el.fill_color:
            shape.fill.solid()
            shape.fill.fore_color.rgb = self._hex_to_rgb(el.fill_color)
        else:
            shape.fill.background()

        if el.line_color:
            shape.line.color.rgb = self._hex_to_rgb(el.line_color)
        else:
            shape.line.fill.background()

    def _add_picture_preserve_ratio(self, slide, image, left, top, width, height, fit: str):
        pic = slide.shapes.add_picture(image, Inches(left), Inches(top))
        img_w = pic.width
        img_h = pic.height
        box_w = Inches(width)
        box_h = Inches(height)
        scale = max(box_w / img_w, box_h / img_h) if fit == "cover" else min(box_w / img_w, box_h / img_h)
        pic.width = int(img_w * scale)
        pic.height = int(img_h * scale)
        pic.left = int(Inches(left) + (box_w - pic.width) / 2)
        pic.top = int(Inches(top) + (box_h - pic.height) / 2)
        return pic

    def _add_readability_overlay(self, slide, background_color: str):
        fill_color = "#000000" if self._relative_luminance(background_color) < 0.45 else "#FFFFFF"
        overlay = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(0),
            Inches(0),
            Inches(self.slide_w),
            Inches(self.slide_h),
        )
        overlay.fill.solid()
        overlay.fill.fore_color.rgb = self._hex_to_rgb(fill_color)
        overlay.fill.transparency = 35
        overlay.line.fill.background()

    # ------------------------------------------------------------------
    # SVG：viewBox 保比例 contain
    # ------------------------------------------------------------------

    @staticmethod
    def _svg_viewbox_ratio(svg_code: str) -> Optional[Tuple[float, float]]:
        m = re.search(
            r'viewBox\s*=\s*["\']\s*([-\d.]+)[ ,]+([-\d.]+)[ ,]+([\d.]+)[ ,]+([\d.]+)',
            svg_code,
        )
        if not m:
            return None
        w, h = float(m.group(3)), float(m.group(4))
        if w <= 0 or h <= 0:
            return None
        return w, h

    def _execute_draw_svg(self, slide, el: CanvasElement):
        if not el.svg_code:
            return
        try:
            svg_code = self._clean_svg_code(el.svg_code)

            png_bytes = cairosvg.svg2png(bytestring=svg_code.encode("utf-8"), scale=3.0, dpi=300)

            crop_ratio = (1, 1)
            if el.svg_crop:
                from PIL import Image
                with Image.open(io.BytesIO(png_bytes)) as image:
                    crop = (math.floor(el.svg_crop[0]*image.width), math.floor(el.svg_crop[1]*image.height),
                            math.ceil(el.svg_crop[2]*image.width), math.ceil(el.svg_crop[3]*image.height))
                    painted = image.convert('RGBA').getchannel('A').getbbox()
                    if painted and (crop[0]>painted[0] or crop[1]>painted[1] or crop[2]<painted[2] or crop[3]<painted[3]):
                        raise ContentConstraintError("SVG 透明边缘测量已变化，不能裁剪可见内容，请重新排版")
                    cropped = image.crop(crop)
                    buffer = io.BytesIO(); cropped.save(buffer, format='PNG')
                    png_bytes = buffer.getvalue()
                    crop_ratio = (cropped.width/image.width, cropped.height/image.height)

            pic = slide.shapes.add_picture(
                io.BytesIO(png_bytes),
                Inches(el.left),
                Inches(el.top)
            )
            # 保比例 contain：此前直接强制 width/height，宽高比不符时整图变形
            box = self._svg_viewbox_ratio(svg_code)
            if box:
                vb_w, vb_h = box[0]*crop_ratio[0], box[1]*crop_ratio[1]
                scale = min(el.width / vb_w, el.height / vb_h)
                draw_w, draw_h = vb_w * scale, vb_h * scale
                pic.left = int(Inches(el.left + (el.width - draw_w) / 2))
                pic.top = int(Inches(el.top + (el.height - draw_h) / 2))
                pic.width = int(Inches(draw_w))
                pic.height = int(Inches(draw_h))
            else:
                pic.width = Inches(el.width)
                pic.height = Inches(el.height)

        except Exception as e:
            if self._fixed_layout:
                raise ContentConstraintError("SVG 渲染失败，已停止导出，请检查 SVG 源码") from e
            print(f"[Renderer Warning] Failed to convert or place SVG: {e}")

    def _clean_svg_code(self, svg_code: str) -> str:
        cleaned = svg_code.strip()
        if cleaned.startswith("```"):
            cleaned = re.sub(r"^```(?:svg|xml)?\s*", "", cleaned, flags=re.IGNORECASE)
            cleaned = re.sub(r"\s*```$", "", cleaned)
        cleaned = re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F]", "", cleaned)
        cleaned = re.sub(r"&(?!amp;|lt;|gt;|quot;|apos;|#\d+;|#x[0-9a-fA-F]+;)", "&amp;", cleaned)
        return cleaned

    def _execute_draw_image(self, slide, el: CanvasElement, background_color: str):
        if not el.image_path or not os.path.exists(el.image_path):
            if self._fixed_layout:
                raise ContentConstraintError("图片资产缺失，已停止导出")
            print(f"[Renderer Warning] Missing local image: {el.image_path}; skipped.")
            return
        try:
            is_background = self._is_large_background(el)
            self._add_picture_preserve_ratio(
                slide,
                el.image_path,
                el.left,
                el.top,
                el.width,
                el.height,
                fit="cover" if is_background else "contain",
            )
            if is_background:
                self._add_readability_overlay(slide, background_color)
        except Exception as e:
            print(f"[Renderer Warning] Failed to draw image: {e}")

    # ------------------------------------------------------------------
    # 图标：Lucide 本地库矢量渲染，失败回退圆点
    # ------------------------------------------------------------------

    def _execute_draw_icon(self, slide, el: CanvasElement, background_color: str):
        color = self._safe_text_color(el.icon_color or "#334155", background_color)
        svg_markup = self._build_lucide_svg(el.icon_name, color)
        if svg_markup:
            try:
                png_bytes = cairosvg.svg2png(
                    bytestring=svg_markup.encode("utf-8"),
                    scale=3.0,
                    dpi=300,
                    background_color="transparent",
                )
                slide.shapes.add_picture(
                    io.BytesIO(png_bytes),
                    Inches(el.left),
                    Inches(el.top),
                    Inches(el.width),
                    Inches(el.height),
                )
                return
            except Exception as e:
                print(f"[Renderer Warning] Icon '{el.icon_name}' render failed: {e}")
        # 回退：半透明圆点（旧行为，保证未知 icon_name 也能出图）
        shape = slide.shapes.add_shape(
            MSO_SHAPE.OVAL,
            Inches(el.left),
            Inches(el.top),
            Inches(el.width),
            Inches(el.height),
        )
        shape.fill.solid()
        shape.fill.fore_color.rgb = self._hex_to_rgb(color)
        shape.fill.transparency = 18
        shape.line.color.rgb = self._hex_to_rgb(color)

    def _build_lucide_svg(self, icon_name: Optional[str], color: str) -> Optional[str]:
        """从 assets/icons 的本地 Lucide SVG 库构建上色后的独立 SVG 字符串。

        icon_library 返回 <svg> 的 innerXML（无外层标签），这里统一补壳、
        上色并强制 24x24 viewBox（Lucide 标准画布）。
        """
        if not icon_name:
            return None
        from config.icon_library import load_lucide_svg

        body = load_lucide_svg(icon_name)
        if not body:
            return None
        hex_color = (color or "#334155").strip()
        if not hex_color.startswith("#"):
            hex_color = "#" + hex_color
        stroke = f'stroke="{hex_color}"'
        body = re.sub(r'stroke\s*=\s*"[^"]*"', stroke, body)
        body = body.replace("currentColor", hex_color)
        return (
            '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" '
            f'fill="none" {stroke} stroke-width="2" stroke-linecap="round" '
            f'stroke-linejoin="round">{body}</svg>'
        )
