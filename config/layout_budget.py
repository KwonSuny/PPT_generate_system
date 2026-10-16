"""新排版的容量预算与可定位的几何反馈；不修改语义内容。"""
import io
import math
import re
from functools import lru_cache
from xml.etree import ElementTree as ET

from config.content import ContentConstraintError, content_values
from config.schema import CanvasElement, TextParagraph


ICON_GAP = .20
NEIGHBOR_GAP = .28


def font_floor(ref, requirements=''):
    role = '标题' if ref=='title' else '来源' if ref=='table/source' else '正文'
    match = re.search(role+r'[^\d]{0,8}(\d{1,2})(?:\s*[–—\-~～至]\s*\d{1,2})?\s*(?:pt|磅|号)', requirements, re.I)
    if match:
        return max(10, min(40, float(match[1])))
    return 28 if ref=='title' else 12 if ref=='table/source' else 18


def font_resize_locked(request):
    return bool(re.search(r'(?:保持|保留|不改|不变|禁止缩|不改变|不要改).{0,5}(?:字号|字体大小)|(?:字号|字体大小).{0,5}(?:不变|不改)|不(?:要)?缩(?:小)?字',request))


def design_request_policy(request):
    locked = font_resize_locked(request)
    return {'allow_font_resize': not locked and bool(re.search(r'缩放|缩小|缩字|字号|字体大小|适配|符合排版',request)),
            'allow_theme_change': bool(re.search(r'背景|颜色|色彩|主题|风格|配色|对比度', request))}


@lru_cache(maxsize=64)
def svg_metrics(code):
    """仅裁掉透明边缘；不解析外部资源，不用白色作为空白判据。"""
    from renderer import cairosvg
    from PIL import Image
    result = {"crop": None, "label_size_estimated": True}
    try:
        if len(code) > 250000 or re.search(r'(?:href\s*=\s*["\'](?!#)|url\(\s*["\']?(?!#)|@import|<!DOCTYPE|<!ENTITY)', code, re.I):
            return result
        root = ET.fromstring(code)
        view = [float(v) for v in re.split(r'[ ,]+', root.attrib['viewBox'].strip())]
        if len(view) != 4 or not all(math.isfinite(v) for v in view) or min(view[2:]) <= 0:
            return result
        result['viewbox'] = view
        result['paint_size'] = view[2:]
        if 'width' in root.attrib or 'height' in root.attrib:
            dimensions = [re.fullmatch(r'([\d.]+)(?:px)?', root.get(k, '')) for k in ('width', 'height')]
            if not all(dimensions):
                return result
            w, h = [float(item[1]) for item in dimensions]
            if min(w, h) <= 0 or abs(w/h-view[2]/view[3]) > .001:
                return result
        # 预算测量不能新增外部读取；不支持的资源交给原有导出校验。
        if not .2 <= view[2]/view[3] <= 20:
            return result
        raster_width = 1000
        png = cairosvg.svg2png(bytestring=code.encode('utf8'), output_width=raster_width,
                               output_height=max(1, round(raster_width*view[3]/view[2])), dpi=300)
        with Image.open(io.BytesIO(png)) as image:
            alpha = image.convert('RGBA').getchannel('A')
            bounds = alpha.getbbox()
            if bounds:
                pad = 8
                bounds = (max(0, bounds[0]-pad), max(0, bounds[1]-pad),
                          min(image.width, bounds[2]+pad), min(image.height, bounds[3]+pad))
                crop = [bounds[0]/image.width, bounds[1]/image.height, bounds[2]/image.width, bounds[3]/image.height]
                if (crop[2]-crop[0])*(crop[3]-crop[1]) < .90:
                    result['crop'] = crop
                    result['paint_size'] = [view[2]*(crop[2]-crop[0]), view[3]*(crop[3]-crop[1])]
        rules = {}
        for selector, declarations in re.findall(r'([^{}]+)\{([^{}]+)\}', code):
            for name in selector.strip().split(','):
                rules[name.strip()] = dict(re.findall(r'([\w-]+)\s*:\s*([^;]+)', declarations))
        sizes = []
        def walk(node, inherited=None, scale=1):
            attrs = dict(inherited or {})
            attrs.update(rules.get(node.tag.rsplit('}', 1)[-1], {}))
            for cls in node.get('class', '').split():
                attrs.update(rules.get('.'+cls, {}))
            attrs.update(node.attrib)
            attrs.update(dict(re.findall(r'([\w-]+)\s*:\s*([^;]+)', node.get('style', ''))))
            for match in re.finditer(r'scale\(([^)]+)\)', node.get('transform', '')):
                factors = [abs(float(v)) for v in re.split(r'[ ,]+', match[1].strip())]
                scale *= min(factors)
            if node.tag.rsplit('}', 1)[-1] in {'text', 'tspan'} and (node.text or '').strip():
                match = re.fullmatch(r'([\d.]+)(px|pt)?', str(attrs.get('font-size', '')))
                if match:
                    sizes.append(float(match[1]) * (300/72 if match[2]=='pt' else 1) * scale)
            for child in node:
                walk(child, attrs, scale)
        walk(root)
        if sizes:
            result['minimum_label_units'] = min(sizes)
    except Exception:
        # 预算估计失败不等于内容错误；原有 SVG 导出校验仍负责成品完整性。
        pass
    return result


def svg_display(element):
    metrics = svg_metrics(element.svg_code or '')
    size = metrics.get('viewbox', [0, 0, 1, 1])[2:]
    crop = element.svg_crop
    if crop:
        size = [size[0]*(crop[2]-crop[0]), size[1]*(crop[3]-crop[1])]
    scale = min(element.width/size[0], element.height/size[1])
    result = {'visible_width': round(size[0]*scale, 3), 'visible_height': round(size[1]*scale, 3)}
    if metrics.get('minimum_label_units'):
        result['estimated_min_label_pt'] = round(metrics['minimum_label_units']*scale*72, 2)
        result['target_min_label_pt'] = 10
    return result


def capacity_budget(slide, floors=None, component_floors=None):
    from renderer import PPTXRenderer
    from config.layout_fallback import stable_layout, StableLayoutUnavailable
    values = content_values(slide)
    text = []
    for ref, value in values.items():
        if not value or ref.startswith(('table/', 'diagram/')):
            continue
        size = max((floors or {}).get(ref, 0), 32 if ref=='title' else 18 if ref.startswith('full_content/') else 20)
        demand = []
        for width in (3.4, 5.5, 12.133):
            element = CanvasElement(type='textbox', left=0, top=0, width=width, height=1,
                paragraphs=[TextParagraph(content_ref=ref, text=value, font_size=size, space_after=0)])
            demand.append({'width': width, 'minimum_height': round(PPTXRenderer.required_text_height(element)+.04, 3)})
        text.append({'content_ref': ref, 'characters': len(value), 'font_size': size, 'width_choices': demand})
    result = {'canvas': [13.333, 7.5], 'safe_rect': [.6, .4, 12.133, 6.7],
              'text': text, 'icon': {'width': .4, 'text_gap': ICON_GAP, 'neighbor_gap': NEIGHBOR_GAP,
                                   'column_gutter_with_icon': .4+ICON_GAP+NEIGHBOR_GAP},
              'policy': '先预留图标与图示，再按精确原文安排正文；参考方案仅证明一种有限可行组合。不能删字；允许缩放时遵守字号下限并重新计算容量。'}
    if slide.get('svg_code'):
        result['svg'] = dict(svg_metrics(slide['svg_code']))
        result['svg']['policy'] = '外边距裁剪由程序测量；按 paint_size 估算比例。标签目标10pt，仅为可读性估计，不能以缩图掩盖冲突。'
    try:
        canvas, name = stable_layout(slide, floors, component_floors)
        result['stable_reference'] = {'composition': name, 'elements': [
            {'id': e.id, 'type': e.type, 'content_refs': [p.content_ref for p in e.paragraphs or []],
             'rect': rect(e), **({'display': svg_display(e)} if e.type=='svg' else {})} for e in canvas.elements]}
    except StableLayoutUnavailable as exc:
        result['stable_reference'] = {'available': False, 'reason': str(exc), 'capacity_conflict': exc.capacity_conflict}
    return result


def rect(element):
    return {key: round(getattr(element, key), 3) for key in ('left', 'top', 'width', 'height')}


def layout_issues(canvas, include_svg_readability=True):
    from renderer import PPTXRenderer
    issues = []
    visible = [e for e in canvas.elements if e.type!='shape']
    def add(kind, elements, message, **detail):
        issues.append({'kind': kind, 'severity': 'major', 'element_ids': [e.id for e in elements],
                       'message': message, 'rectangles': [rect(e) for e in elements], **detail})
    for e in canvas.elements:
        overflow = {key: round(max(0, v), 3) for key, v in
                    (('left', -e.left), ('top', -e.top), ('right', e.left+e.width-13.333), ('bottom', e.top+e.height-7.5))}
        line = e.type=='shape' and e.shape_type=='line' and min(e.width,e.height)>=0 and max(e.width,e.height)>0
        if max(overflow.values()) > .001 or (min(e.width, e.height) <= 0 and not line):
            add('bounds', [e], '图元超出画布', excess=overflow)
        if e.type=='textbox':
            renderer=PPTXRenderer.__new__(PPTXRenderer)
            background=renderer.effective_background(e,canvas)
            if any(renderer._contrast_ratio(p.font_color,color)<4.5 for color in renderer.background_samples(e,canvas) for p in e.paragraphs or []):
                add('contrast',[e],'文字与实际承托色面对比不足',background=background)
            needed = PPTXRenderer.required_text_height(e)
            if needed > e.height+.001:
                add('text_capacity', [e], '文字容量不足', minimum_height=round(needed+.01, 3),
                    height_deficit=round(needed-e.height, 3), content_refs=[p.content_ref for p in e.paragraphs or []])
        elif e.type in {'table', 'diagram'}:
            try:
                PPTXRenderer.validate_capacity(e)
            except ContentConstraintError as exc:
                add('component_capacity', [e], str(exc))
        elif e.type=='svg' and include_svg_readability:
            display = svg_display(e)
            if display.get('estimated_min_label_pt', 10) < 10:
                add('svg_readability', [e], 'SVG 标签显示过小；扩大主视觉区域或切换构图', **display)
    for i, a in enumerate(visible):
        for b in visible[i+1:]:
            ox = min(a.left+a.width, b.left+b.width)-max(a.left, b.left)
            oy = min(a.top+a.height, b.top+b.height)-max(a.top, b.top)
            if ox > .03 and oy > .03:
                add('overlap', [a,b], '内容容器发生重叠', overlap_width=round(ox,3), overlap_height=round(oy,3),
                    action='扩大或换位整个内容分区，勿只增高当前文本框')
    textboxes = [e for e in visible if e.type=='textbox']
    renderer = PPTXRenderer.__new__(PPTXRenderer)
    for icon in (e for e in visible if e.type=='icon'):
        partner = renderer._find_icon_partner(icon, textboxes)
        if partner is None:
            add('icon_anchor', [icon], '图标缺少可确定的文字锚点；填写 content_ref')
            continue
        gap = partner.left-icon.left-icon.width
        center = abs(partner.top+partner.height/2-icon.top-icon.height/2)
        if abs(gap-ICON_GAP) > .08 or center > .12:
            add('icon_alignment', [icon,partner], '图标与锚点文字的间距或中心未对齐', actual_gap=round(gap,3),
                target_gap=ICON_GAP, center_delta=round(center,3))
        for other in textboxes:
            if other is partner:
                continue
            oy = min(icon.top+icon.height, other.top+other.height)-max(icon.top,other.top)
            gap_x = max(other.left-icon.left-icon.width, icon.left-other.left-other.width, 0)
            if oy > .03 and gap_x < NEIGHBOR_GAP-.001:
                add('icon_neighbor', [icon,other], '图标接近相邻板块，列间距未包含图标宽度', actual_gap=round(gap_x,3),
                    minimum_gap=NEIGHBOR_GAP, deficit=round(NEIGHBOR_GAP-gap_x,3),
                    action='把图标与对应标题作为一组移动或调整列宽，为相邻正文保留独立间距')
    return issues


def layout_fingerprint(canvas):
    from config.storage import content_hash
    return content_hash(sorted((e.type, tuple(sorted(p.content_ref or '' for p in e.paragraphs or [])),
                                e.content_ref or '', *(round(getattr(e,k),2) for k in ('left','top','width','height')))
                               for e in canvas.elements if e.type!='shape'))
