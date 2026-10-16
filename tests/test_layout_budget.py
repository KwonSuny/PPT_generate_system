import copy
import tempfile
import unittest
from pathlib import Path

import main
from config.content import identified_slide
from config.layout_budget import capacity_budget, layout_issues, svg_metrics, svg_display
from config.schema import CanvasElement, PPTFinalBlueprint
from config.state import PPTGenerationState
from renderer import PPTXRenderer
from test_content import outline_fixture, canvas_for


SVG = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 800 400"><path d="M100 200 L700 200" stroke="blue" stroke-width="8"/><text x="350" y="180" font-size="24">LABEL</text></svg>'


class LayoutBudgetTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix='layout_budget_')
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.slide = identified_slide(copy.deepcopy(outline_fixture()['slides'][0]))

    def test_first_request_has_measured_budget_and_repeated_conflict_has_coordinates(self):
        bad = canvas_for(self.slide)
        bad.elements[0].height = .1
        bad.elements[0].width = .25
        calls = []
        def designer(state, request):
            calls.append(copy.deepcopy(request))
            return PPTFinalBlueprint(canvases=[bad])
        main._layout_page(PPTGenerationState(topic='验证'), designer, {'slides':[self.slide]}, 1, str(self.root))
        self.assertEqual(len(calls), 3)
        self.assertIn('stable_reference', calls[0]['capacity_budget'])
        self.assertTrue(calls[1]['layout_feedback']['violations'])
        issue = calls[1]['layout_feedback']['violations'][0]
        self.assertIn('rectangles', issue)
        self.assertGreater(issue['height_deficit'], 0)
        self.assertTrue(calls[2]['layout_feedback']['repeated_geometry'])
        self.assertIn('切换', calls[2]['layout_feedback']['repair_plan'])

    def test_explicit_icon_anchor_wins_and_neighbor_gap_enters_repair(self):
        canvas = canvas_for(self.slide)
        text = canvas.elements[1]
        text.left=5; text.top=3; text.width=3; text.height=1
        other = text.model_copy(deep=True); other.id='neighbor';other.left=1;other.width=3.7
        other.paragraphs[0].content_ref='other'
        icon = CanvasElement(id='shield', type='icon',content_ref=text.paragraphs[0].content_ref,
                             icon_name='shield',left=4.4,top=3.3,width=.4,height=.4)
        canvas.elements += [other,icon]
        self.assertIs(PPTXRenderer.__new__(PPTXRenderer)._find_icon_partner(icon, [other,text]), text)
        issues = layout_issues(canvas, include_svg_readability=False)
        self.assertTrue(any(i['kind']=='icon_neighbor' and 'neighbor' in i['element_ids'] for i in issues))

    def test_crop_only_transparent_margins_keeps_source_and_old_layout(self):
        metrics=svg_metrics(SVG)
        self.assertIsNotNone(metrics['crop'])
        self.assertLess(metrics['paint_size'][1], 100)
        slide={**self.slide, 'svg_code':SVG}
        canvas=canvas_for(self.slide)
        art=CanvasElement(id='art',type='svg',svg_code=SVG,left=1,top=3,width=10,height=1.4)
        canvas.elements.append(art)
        resolved=PPTXRenderer.resolve_canvas(canvas)
        revised=next(e for e in resolved.elements if e.type=='svg')
        self.assertEqual(revised.svg_code, SVG)
        self.assertGreater(svg_display(revised)['estimated_min_label_pt'], 20)
        main._render_blueprint(PPTFinalBlueprint(schema_version=2,canvases=[resolved]), str(self.root/'crop.pptx'))
        old=canvas.model_copy(deep=True);old.layout_resolved=True
        self.assertIsNone(PPTXRenderer.resolve_canvas(old).elements[-1].svg_crop)
        opaque=SVG.replace('<path', '<rect width="800" height="400" fill="white"/><path')
        self.assertIsNone(svg_metrics(opaque)['crop'])

    def test_svg_overlap_is_not_hidden_by_shrinking_and_budget_keeps_art(self):
        canvas=canvas_for(self.slide)
        self.slide['svg_code']=SVG
        art=CanvasElement(id='art',type='svg',svg_code=SVG,left=1,top=1.5,width=10,height=3)
        canvas.elements.append(art)
        result=PPTXRenderer.resolve_canvas(canvas)
        self.assertEqual(result.elements[-1].height, 3)
        self.assertTrue(any(i['kind']=='overlap' for i in layout_issues(result)))
        self.assertTrue(any(e['type']=='svg' for e in capacity_budget(self.slide)['stable_reference']['elements']))

    def test_measurement_does_not_read_external_resources_or_clip_visible_content(self):
        self.assertIsNone(svg_metrics(SVG.replace('<path', '<image href="file:///private.png"/><path'))['crop'])
        self.assertIsNone(svg_metrics(SVG.replace('<path', '<style>@import "file:///private.css";</style><path'))['crop'])
        self.assertIsNone(svg_metrics(SVG.replace('viewBox=', 'width="200" height="500" viewBox='))['crop'])
        canvas=canvas_for(self.slide);canvas.layout_resolved=True
        canvas.elements.append(CanvasElement(id='art',type='svg',svg_code=SVG,svg_crop=[.3,.3,.4,.4],
            left=1,top=3,width=5,height=2))
        with self.assertRaisesRegex(Exception, 'SVG 渲染失败'):
            main._render_blueprint(PPTFinalBlueprint(schema_version=2,canvases=[canvas]), str(self.root/'bad.pptx'))

    def test_visual_patches_keep_saved_crop_and_quality_does_not_block_stable_delivery(self):
        from agents.visual_review import GeometryPatch
        from config.visual import apply_patches, rule_issues
        canvas=canvas_for(self.slide);canvas.layout_resolved=True
        art=CanvasElement(id='art',type='svg',svg_code=SVG,left=1,top=3,width=5,height=.5)
        canvas.elements.append(art)
        revised=apply_patches(canvas, [GeometryPatch(element_id='art',left=1.1)])
        self.assertIsNone(revised.elements[-1].svg_crop)
        self.assertFalse(any(i['kind']=='svg_readability' for i in rule_issues(revised)))
        self.assertTrue(any(i['kind']=='svg_readability' for i in rule_issues(revised,include_quality=True)))

    def test_seeded_native_font_floor_is_not_replaced_by_first_proposal(self):
        from test_content import table_slide
        slide=table_slide();canvas=canvas_for(slide)
        requests=[]
        def designer(state, payload):
            requests.append(payload)
            proposal=canvas.model_copy(deep=True)
            if len(requests)>1:
                proposal.elements[-1].table_style.font_size=20
            return PPTFinalBlueprint(canvases=[proposal])
        result=main._layout_page(PPTGenerationState(topic='表格'), designer, {'slides':[slide],
            'layout_feedback':{'component_font_size_floors':{'table':{'font_size':20}}}},1,str(self.root))
        self.assertEqual(len(requests),2)
        self.assertEqual(result.elements[-1].table_style.font_size,20)
