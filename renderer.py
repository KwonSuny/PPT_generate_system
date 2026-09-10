import io
import os
import re
from typing import Iterable

import cairosvg
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_AUTO_SIZE, PP_ALIGN
from pptx.util import Inches, Pt

from config.schema import CanvasElement, PPTFinalBlueprint


class PPTXRenderer:

    def __init__(self, blueprint: PPTFinalBlueprint):
        self.blueprint = blueprint
        self.prs = Presentation()
        self.slide_w = 13.333
        self.slide_h = 7.5
        self.prs.slide_width = Inches(self.slide_w)
        self.prs.slide_height = Inches(self.slide_h)

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
        """Preserve the model color when readable, otherwise force ink/white."""
        if self._contrast_ratio(font_color, background_color) >= 4.5:
            return font_color
        dark_ratio = self._contrast_ratio("#111827", background_color)
        light_ratio = self._contrast_ratio("#FFFFFF", background_color)
        return "#111827" if dark_ratio >= light_ratio else "#FFFFFF"

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

    def _normalize_element(self, el: CanvasElement) -> CanvasElement | None:
        item = el.model_copy()
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

    def render_all(self, output_path: str = "output.pptx"):
        print("[Renderer Engine] Rendering structured blueprint to PPTX...")
        blank_layout = self.prs.slide_layouts[6]

        for canvas_data in self.blueprint.canvases:
            slide = self.prs.slides.add_slide(blank_layout)
            bg_color = canvas_data.background_color or "#FFFFFF"

            bg = slide.background
            fill = bg.fill
            fill.solid()
            fill.fore_color.rgb = self._hex_to_rgb(bg_color)

            normalized = [
                item
                for item in (self._normalize_element(el) for el in canvas_data.elements)
                if item is not None
            ]
            for el in sorted(normalized, key=self._layer_key):
                if el.type == "textbox":
                    self._execute_draw_textbox(slide, el, bg_color)
                elif el.type == "shape":
                    self._execute_draw_shape(slide, el)
                elif el.type == "svg":
                    self._execute_draw_svg(slide, el)
                elif el.type == "image":
                    self._execute_draw_image(slide, el, bg_color)
                elif el.type == "icon":
                    self._execute_draw_icon(slide, el, bg_color)

        self.prs.save(output_path)
        print(f"[Renderer Engine] PPTX saved to: {output_path}")

    def _estimated_line_count(self, text: str, font_size: float, width_in: float) -> int:
        if not text:
            return 1
        chars_per_line = max(int((width_in * 72) / max(font_size * 0.95, 1)), 1)
        return max((len(text) + chars_per_line - 1) // chars_per_line, 1)

    def _fit_font_scale(self, paragraphs: Iterable, width_in: float, height_in: float) -> float:
        paragraphs = list(paragraphs or [])
        if not paragraphs:
            return 1.0
        total_lines = 0
        max_font = 1.0
        for p in paragraphs:
            max_font = max(max_font, p.font_size)
            total_lines += self._estimated_line_count(p.text, p.font_size, width_in)
        available_pt = max((height_in - 0.18) * 72, 8)
        required_pt = total_lines * max_font * 1.28 + max(len(paragraphs) - 1, 0) * 4
        if required_pt <= available_pt:
            return 1.0
        return max(0.68, available_pt / required_pt)

    def _execute_draw_textbox(self, slide, el: CanvasElement, background_color: str):
        tx_box = slide.shapes.add_textbox(
            Inches(el.left), Inches(el.top), Inches(el.width), Inches(el.height)
        )
        tf = tx_box.text_frame
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        tf.margin_left = tf.margin_right = Inches(0.11)
        tf.margin_top = tf.margin_bottom = Inches(0.08)

        if not el.paragraphs:
            return

        align_mapping = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}
        font_scale = self._fit_font_scale(el.paragraphs, el.width, el.height)

        for idx, p_data in enumerate(el.paragraphs):
            p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
            p.text = p_data.text
            p.font.size = Pt(max(9.0, p_data.font_size * font_scale))
            p.font.color.rgb = self._hex_to_rgb(
                self._safe_text_color(p_data.font_color, background_color)
            )
            p.font.bold = p_data.bold
            p.alignment = align_mapping.get(p_data.alignment, PP_ALIGN.LEFT)
            p.space_after = Pt(max(0, min(p_data.space_after, 8)) if font_scale < 1 else p_data.space_after)

    def _execute_draw_chart(self, slide, el: CanvasElement, background_color: str):
        data = el.chart_data or {}
        if not isinstance(data, dict) or not data:
            return

        panel_left = el.left + 0.08
        panel_top = el.top + 0.08
        panel_width = max(el.width - 0.16, 1.2)
        panel_height = max(el.height - 0.16, 1.2)

        panel = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(panel_left),
            Inches(panel_top),
            Inches(panel_width),
            Inches(panel_height),
        )
        panel.fill.solid()
        panel.fill.fore_color.rgb = self._hex_to_rgb(el.fill_color or "#F8FAFC")
        panel.line.color.rgb = self._hex_to_rgb(el.line_color or "#E2E8F0")

        title = el.chart_title or "数据概览"
        title_box = slide.shapes.add_textbox(
            Inches(el.left + 0.2),
            Inches(el.top + 0.16),
            Inches(max(el.width - 0.4, 2.0)),
            Inches(0.45),
        )
        tf = title_box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title
        p.font.size = Pt(16)
        p.font.bold = True
        p.font.color.rgb = self._hex_to_rgb(self._safe_text_color("#0F172A", background_color))

        items = list(data.items())
        max_value = max(abs(v) for _, v in items) if items else 1.0
        bar_area_left = el.left + 0.24
        bar_area_top = el.top + 1.15
        bar_area_width = max(el.width - 0.48, 2.0)
        bar_area_height = max(el.height - 1.7, 1.0)
        bar_gap = 0.18
        bar_width = max((bar_area_width - bar_gap * (len(items) - 1)) / len(items), 0.6)

        legend_items = el.legend_items or [k for k, _ in items]
        legend_y = el.top + el.height - 0.7
        for idx, (label, value) in enumerate(items):
            color = self._hex_to_rgb(["#2563EB", "#14B8A6", "#F59E0B", "#EC4899", "#8B5CF6"][idx % 5])
            bar_height = max(0.14, min(1.8, (abs(value) / max_value) * max(bar_area_height - 0.55, 0.3)))
            bar_left = bar_area_left + idx * (bar_width + bar_gap)
            bar = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                Inches(bar_left),
                Inches(bar_area_top + (bar_area_height - bar_height) / 2),
                Inches(bar_width),
                Inches(bar_height),
            )
            bar.fill.solid()
            bar.fill.fore_color.rgb = color
            bar.line.fill.background()

            label_box = slide.shapes.add_textbox(
                Inches(bar_left),
                Inches(bar_area_top + bar_area_height + 0.08),
                Inches(bar_width),
                Inches(0.25),
            )
            label_tf = label_box.text_frame
            label_tf.word_wrap = True
            label_p = label_tf.paragraphs[0]
            label_p.text = label
            label_p.font.size = Pt(8)
            label_p.font.color.rgb = self._hex_to_rgb(self._safe_text_color("#334155", background_color))

            value_box = slide.shapes.add_textbox(
                Inches(bar_left),
                Inches(bar_area_top + (bar_area_height - bar_height) / 2 - 0.28),
                Inches(bar_width),
                Inches(0.24),
            )
            value_tf = value_box.text_frame
            value_tf.word_wrap = True
            value_p = value_tf.paragraphs[0]
            value_p.text = str(int(value)) if float(value).is_integer() else f"{value:.1f}"
            value_p.font.size = Pt(8)
            value_p.font.bold = True
            value_p.font.color.rgb = self._hex_to_rgb(self._safe_text_color("#111827", background_color))

            if idx < len(legend_items):
                legend_box = slide.shapes.add_textbox(
                    Inches(el.left + 0.24 + idx * 1.2),
                    Inches(max(legend_y, el.top + 0.3)),
                    Inches(1.0),
                    Inches(0.24),
                )
                legend_tf = legend_box.text_frame
                legend_tf.word_wrap = True
                legend_p = legend_tf.paragraphs[0]
                legend_p.text = legend_items[idx]
                legend_p.font.size = Pt(8)
                legend_p.font.color.rgb = self._hex_to_rgb(self._safe_text_color("#334155", background_color))

        if items and max_value > 0:
            baseline = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                Inches(el.left + 0.14),
                Inches(el.top + el.height - 0.9),
                Inches(el.width - 0.28),
                Inches(0.02),
            )
            baseline.fill.solid()
            baseline.fill.fore_color.rgb = self._hex_to_rgb("#CBD5E1")
            baseline.line.fill.background()

    def _execute_draw_shape(self, slide, el: CanvasElement):
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

    def _execute_draw_svg(self, slide, el: CanvasElement):
        if not el.svg_code:
            return
        try:
            svg_code = self._clean_svg_code(el.svg_code)
            
            png_bytes = cairosvg.svg2png(bytestring=svg_code.encode("utf-8"), scale=3.0, dpi=300)
            
            pic = slide.shapes.add_picture(
                io.BytesIO(png_bytes), 
                Inches(el.left), 
                Inches(el.top)
            )
            pic.width = Inches(el.width)
            pic.height = Inches(el.height)
            
        except Exception as e:
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

    def _execute_draw_icon(self, slide, el: CanvasElement, background_color: str):
        color = self._safe_text_color(el.icon_color or "#334155", background_color)
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
