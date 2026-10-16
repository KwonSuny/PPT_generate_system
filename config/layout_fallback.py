"""自由排版约束失败后，以有限构图搜索保留语义、字号和原生组件。"""
from config.content import ContentConstraintError, bind_canvas, content_values
from config.schema import CanvasElement, SlideCanvas, SlideElement, TextParagraph
from config.visual import rule_issues
from renderer import PPTXRenderer


class StableLayoutUnavailable(ContentConstraintError):
    def __init__(self, message, capacity_conflict):
        super().__init__(message)
        self.capacity_conflict = capacity_conflict


def stable_layout(slide, text_floors=None, component_floors=None, compact=False):
    slide = SlideElement.model_validate(slide)
    raw = slide.model_dump()
    identity = slide.id or 'slide'
    values = content_values(raw)
    floors, styles = text_floors or {}, component_floors or {}
    width, bottom, gap = 12.133, 7.10, .28

    def text(ref, left, top, container_width, default_size, bold=False):
        if compact:
            default_size = 28 if ref=='title' else 18
        size = max(default_size, floors.get(ref, 0))
        element = CanvasElement(id=identity+'-'+ref.replace('/', '-'), type='textbox', left=left, top=top,
            width=container_width, height=.5, paragraphs=[TextParagraph(content_ref=ref, text=values[ref], font_size=size,
            font_color=slide.theme.primary if bold else slide.theme.content, bold=bold, space_after=0)])
        element.height = PPTXRenderer.required_text_height(element) + .04
        return element

    header = [text('title', .6, .4, width, 32, True)]
    top = header[0].top + header[0].height + .15
    if values.get('punchline'):
        header.append(text('punchline', .6, top, width, 20))
        top = header[-1].top + header[-1].height + .15
    top += .10
    point_ids = list(dict.fromkeys(ref.split('/', 1)[1] for ref in values
                    if ref.startswith(('bullet_points/', 'full_content/'))))
    kinds = [kind for kind in ('table', 'diagram') if getattr(slide, kind)]
    if slide.svg_code:
        kinds.append('svg')
    if slide.image_path:
        kinds.append('image')

    def notes(left, start, note_width, columns):
        items, y = [], start
        column_width = (note_width - gap*(columns-1))/columns
        if column_width <= .5:
            raise ContentConstraintError('稳定布局没有足够的正文列宽')
        for offset in range(0, len(point_ids), columns):
            row = []
            for index, point in enumerate(point_ids[offset:offset+columns]):
                group, group_y = [], y
                for field, size, bold in (('bullet_points', 20, True), ('full_content', 18, False)):
                    ref = field+'/'+point
                    if ref not in values or not values[ref]:
                        continue
                    item = text(ref, left+index*(column_width+gap), group_y, column_width, size, bold)
                    group.append(item)
                    group_y += item.height + .12
                row.append((group, group_y-.12 if group else y))
            items.extend(item for group, _ in row for item in group)
            y = max(end for _, end in row) + .24
        return items, max(0, y-start-.24) if point_ids else 0

    def art(left, art_top, art_width, height):
        items = []
        available_width = (art_width-gap*(len(kinds)-1))/len(kinds)
        for index, kind in enumerate(kinds):
            element = CanvasElement(id=identity+'-'+kind, type=kind, content_ref=kind,
                left=left+index*(available_width+gap), top=art_top, width=available_width, height=height)
            if kind in {'table', 'diagram'}:
                style = getattr(element, kind+'_style')
                for name, minimum in styles.get(kind, {}).items():
                    setattr(style, name, max(getattr(style, name), minimum))
            items.append(element)
        return items

    candidates, failures = [], []
    def consider(name, elements):
        try:
            canvas = bind_canvas(SlideCanvas(slide_id=slide.id, background_color=slide.theme.background,
                                             elements=header+elements), raw)
            canvas = PPTXRenderer.resolve_canvas(canvas)
            if rule_issues(canvas):
                raise ContentConstraintError('稳定布局存在内容重叠或越界')
            shown = {paragraph.content_ref: paragraph.text for element in canvas.elements if element.type=='textbox'
                     for paragraph in element.paragraphs or []}
            if any(shown.get(ref) != value for ref, value in values.items()
                   if value and not ref.startswith(('table/', 'diagram/'))):
                raise ContentConstraintError('稳定布局未完整展示文字引用')
            score = 0
            for element in canvas.elements:
                if element.type == 'svg':
                    from config.layout_budget import svg_display
                    display = svg_display(element)
                    score += display['visible_width']
                elif element.type in {'image', 'table', 'diagram'}:
                    score += element.width*element.height
            if not kinds:
                score = -abs(min(2, len(point_ids) or 1)-int(name.split('-')[-1]))
            candidates.append((score, name, canvas))
        except ContentConstraintError as error:
            failures.append(str(error))

    if not kinds:
        for columns in range(1, min(3, len(point_ids) or 1)+1):
            items, height = notes(.6, top, width, columns)
            if top+height <= bottom:
                consider('text-columns-'+str(columns), items)
    else:
        for columns in range(1, min(4, len(point_ids) or 1)+1):
            items, height = notes(.6, 0, width, columns)
            note_top = bottom-height
            art_height = note_top-.25-top if point_ids else bottom-top
            if art_height < 1.25:
                continue
            for item in items:
                item.top += note_top
            consider('art-top-notes-'+str(columns), art(.6, top, width, art_height)+items)
        if point_ids:
            for art_width in (6.3, 7.0, 5.8):
                note_width = width-art_width-gap
                for side in ('left', 'right'):
                    note_left = .6+art_width+gap if side=='left' else .6
                    art_left = .6 if side=='left' else .6+note_width+gap
                    items, height = notes(note_left, top, note_width, 1)
                    if top+height <= bottom:
                        consider('art-'+side+'-notes', art(art_left, top, art_width, bottom-top)+items)
    if not candidates:
        detail = failures[-1] if failures else '正文与图示超过可用空间'
        capacity = not failures or all(any(term in message for term in ('容量', '文字需要', '太窄', '太矮', '空间过小'))
                                      for message in failures)
        action = '请允许扩页或明确可精简的内容' if capacity else '请检查原始图示或内容约束'
        raise StableLayoutUnavailable('稳定布局仍无法同时满足完整内容、字号和页数要求：'+detail+
                                      '；'+action+'，系统未删改原文', capacity)
    _, name, canvas = max(candidates, key=lambda candidate: candidate[0])
    return canvas, name
