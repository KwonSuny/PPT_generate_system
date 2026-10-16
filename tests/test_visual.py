import hashlib
import json
import os
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from PIL import Image
from pydantic import ValidationError

import main
from agents.visual_review import GeometryPatch, VisualIssue, VisualPageReview, review_page
from config.content import ContentConstraintError
from config.preview import PreviewError, render_preview
from config.schema import PPTFinalBlueprint
from config.state import PPTGenerationState
from config.storage import PPTStore, content_hash
from config.visual import apply_patches, rule_issues, run_visual_pipeline
from renderer import PPTXRenderer
from test_content import canvas_for, table_slide


class VisualPipelineTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory(prefix="ppt_visual_test_")
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name)
        self.slide = table_slide()
        self.canvas = PPTXRenderer.resolve_canvas(canvas_for(self.slide))
        self.blueprint = PPTFinalBlueprint(schema_version=2, canvases=[self.canvas]).model_dump()
        self.filename = self.root / "result.pptx"
        main._render_blueprint(PPTFinalBlueprint.model_validate(self.blueprint), str(self.filename))
        self.state = PPTGenerationState(topic="视觉回归")
        self.state.set("blueprint", self.blueprint)
        self.state.set("outline", {"schema_version": 2, "topic": "视觉回归", "slides": [self.slide]})
        environment = patch.dict(os.environ, {"PPT_PREVIEW_ENABLED": "1", "PPT_VISUAL_REVIEW": "1", "PPT_VISION_MAX_CALLS": "8"})
        environment.start()
        self.addCleanup(environment.stop)

    def fake_tools(self, arguments, timeout=120):
        if "--convert-to" in arguments:
            Path(arguments[arguments.index("--outdir") + 1], "deck.pdf").write_bytes(b"%PDF-test")
            self.assertTrue(any(argument.startswith("-env:UserInstallation=file:") for argument in arguments))
        elif arguments[0] == "pdfinfo":
            return "Pages: 1\n"
        elif arguments[0] == "pdftoppm":
            Image.new("RGB", (1600, 900), "white").save(Path(arguments[-1]).with_name("page-1.png"))
        return ""

    def preview(self, filename, output_dir, signature=None):
        directory = Path(output_dir)
        directory.mkdir(parents=True, exist_ok=True)
        image = directory / "page-1.png"
        pdf = directory / "deck.pdf"
        Image.new("RGB", (100, 56), "white").save(image)
        pdf.write_bytes(b"%PDF-test")
        return {"pdf": str(pdf), "pages": [str(image)], "pptx_sha256": hashlib.sha256(Path(filename).read_bytes()).hexdigest()}

    def review(self, state, canvas, *args):
        return VisualPageReview(slide_id=canvas.slide_id, issues=[{"kind": "alignment", "severity": "minor", "element_ids": [canvas.elements[0].id],
            "bbox": [0, 0, 0.2, 0.2], "message": "合成对齐问题"}],
            patches=[{"element_id": canvas.elements[0].id, "left": canvas.elements[0].left + 0.1}])

    def test_preview_checks_page_count_hash_and_cache_integrity(self):
        signature = {"versions": {"soffice": "test"}, "fonts": "test"}
        tools = {"soffice": "soffice", "pdfinfo": "pdfinfo", "pdftoppm": "pdftoppm"}
        with patch("config.preview._run", side_effect=self.fake_tools) as run:
            first = render_preview(self.filename, self.root / "previews", signature=signature, tools=tools)
            second = render_preview(self.filename, self.root / "previews", signature=signature, tools=tools)
        self.assertEqual(run.call_count, 3)
        self.assertEqual(first["pages"], second["pages"])
        self.assertEqual(first["page_count"], 1)
        Path(first["pages"][0]).write_bytes(b"changed")
        with self.assertRaises(PreviewError):
            render_preview(self.filename, self.root / "previews", signature=signature, tools=tools)

    def test_missing_runtime_keeps_pptx_unreviewed_and_does_not_call_model(self):
        original = self.filename.read_bytes()
        with patch("config.visual.tool_signature", side_effect=PreviewError("缺少工具")), patch("config.visual.review_page") as model:
            output = run_visual_pipeline(self.state, str(self.filename), self.root, "test-key")
        model.assert_not_called()
        self.assertEqual(output, str(self.filename))
        self.assertEqual(self.filename.read_bytes(), original)
        self.assertEqual(self.state.get("visual_review")["status"], "not_run")

    def test_pdf_page_mismatch_does_not_publish_a_preview(self):
        def broken(arguments, timeout=120):
            return "Pages: 2\n" if arguments[0] == "pdfinfo" else self.fake_tools(arguments, timeout)
        parent = self.root / "invalid-previews"
        with patch("config.preview._run", side_effect=broken), self.assertRaises(PreviewError):
            render_preview(self.filename, parent, signature={"test": True},
                           tools={"soffice": "soffice", "pdfinfo": "pdfinfo", "pdftoppm": "pdftoppm"})
        self.assertFalse(any(parent.iterdir()))

    def test_failed_candidate_preview_keeps_the_initial_pptx_and_images(self):
        calls = 0
        def interrupted(filename, output_dir, signature=None):
            nonlocal calls
            calls += 1
            if calls == 2:
                raise PreviewError("候选截图失败")
            return self.preview(filename, output_dir, signature)
        with patch("config.visual.tool_signature", return_value={}), patch("config.visual.render_preview", side_effect=interrupted), patch("config.visual.review_page", side_effect=self.review):
            output = run_visual_pipeline(self.state, str(self.filename), self.root, "test-key")
        self.assertEqual(output, str(self.filename))
        self.assertEqual(self.state.get("blueprint"), self.blueprint)
        self.assertTrue(self.state.get("preview_assets").get("pages"))
        self.assertEqual(self.state.get("visual_review")["status"], "needs_review")

    def test_geometry_patch_keeps_text_data_style_and_rejects_unbounded_changes(self):
        old = self.canvas.model_dump()
        result = apply_patches(self.canvas, [GeometryPatch(element_id=self.canvas.elements[0].id, left=0.8)])
        self.assertEqual(result.elements[0].left, 0.8)
        self.assertEqual(result.elements[0].paragraphs, self.canvas.elements[0].paragraphs)
        self.assertEqual(result.elements[1].table, self.canvas.elements[1].table)
        self.assertEqual(self.canvas.model_dump(), old)
        with self.assertRaises(ContentConstraintError):
            apply_patches(self.canvas, [GeometryPatch(element_id=self.canvas.elements[0].id, left=5)])
        with self.assertRaises(ValidationError):
            GeometryPatch(element_id="title", text="改写事实")
        self.assertFalse(rule_issues(self.canvas))

    def test_text_mode_reports_visual_issue_without_geometry_repair(self):
        self.state.modify_scope = "text"
        with patch("config.visual.tool_signature", return_value={}), patch("config.visual.render_preview", side_effect=self.preview), patch("config.visual.review_page", side_effect=self.review), patch("config.visual.apply_patches") as repair:
            output = run_visual_pipeline(self.state, str(self.filename), self.root, "test-key")
        repair.assert_not_called()
        self.assertEqual(output, str(self.filename))
        self.assertEqual(self.state.get("blueprint"), self.blueprint)
        self.assertEqual(self.state.get("visual_review")["status"], "needs_review")

    def test_two_repair_rounds_keep_unapproved_candidates_and_initial_version(self):
        with patch("config.visual.tool_signature", return_value={}), patch("config.visual.render_preview", side_effect=self.preview), patch("config.visual.review_page", side_effect=self.review) as model:
            output = run_visual_pipeline(self.state, str(self.filename), self.root, "test-key")
        self.assertEqual(model.call_count, 3)
        self.assertEqual(output, str(self.filename))
        self.assertEqual(self.state.get("blueprint"), self.blueprint)
        self.assertEqual(len(self.state.get("visual_candidates")), 2)
        self.assertEqual(self.state.get("visual_review")["status"], "needs_review")

    def test_only_reviewed_candidate_is_adopted_and_budget_stops_new_calls(self):
        clean = VisualPageReview(slide_id=self.canvas.slide_id)
        with patch("config.visual.tool_signature", return_value={}), patch("config.visual.render_preview", side_effect=self.preview), patch("config.visual.review_page", side_effect=[self.review(self.state, self.canvas), clean]):
            output = run_visual_pipeline(self.state, str(self.filename), self.root, "test-key")
        self.assertNotEqual(output, str(self.filename))
        self.assertEqual(Path(output).name, self.filename.name)
        self.assertEqual(self.state.get("visual_review")["status"], "passed")
        self.assertAlmostEqual(self.state.get("blueprint")["canvases"][0]["elements"][0]["left"], 0.8)
        self.state.set("blueprint", self.blueprint)
        with patch("config.visual.tool_signature", return_value={}), patch("config.visual.render_preview", side_effect=self.preview), patch("config.visual.review_page") as model:
            run_visual_pipeline(self.state, str(self.filename), self.root, "test-key", used_calls=8)
        model.assert_not_called()
        self.assertNotEqual(self.state.get("visual_review")["status"], "passed")

    def test_vision_rejects_unknown_ids_and_preserves_raw_response(self):
        image = self.root / "page.png"
        Image.new("RGB", (10, 10), "white").save(image)
        bad = {"slide_id": "wrong-slide", "issues": [], "patches": []}
        response = SimpleNamespace(choices=[SimpleNamespace(finish_reason="stop", message=SimpleNamespace(content=json.dumps(bad)))])
        with patch("agents.visual_review.create_client"), patch("agents.visual_review.chat_completion_with_retry", return_value=response):
            with self.assertRaises(ContentConstraintError):
                review_page(self.state, self.canvas, image, "test-key")
        self.assertEqual(json.loads(self.state.messages[-1]["content"]), bad)

    def test_named_boxes_are_normalized_but_ambiguous_or_out_of_bounds_boxes_fail(self):
        base = {"kind": "arrow", "severity": "minor", "message": "端点偏移"}
        issue = VisualIssue.model_validate({**base, "bbox": {"x": 0.3, "y": 0.2, "width": 0.2, "height": 0.1}})
        self.assertEqual(issue.bbox, [0.3, 0.2, 0.2, 0.1])
        for box in ([0.375, 0.2, 0.925, 0.6], [0, 0, 0, 0], {"x1": 0.3, "y1": 0.2, "x2": 0.5, "y2": 0.3}):
            with self.subTest(box=box), self.assertRaises(ValidationError):
                VisualIssue.model_validate({**base, "bbox": box})

    def test_worker_persists_preview_assets_visual_report_and_usage(self):
        from config.jobs import execute_job
        from config.llm_client import record_model_response
        store = PPTStore(self.root)
        run = self.root / "run"
        run.mkdir()
        request = {"user_input": "回归", "filename": "result.pptx", "requirements": "",
                   "page_numbers": None, "use_incremental": False, "mode": "new"}
        job, _ = store.create_job(request, self.state.snapshot(), run, "visual-job", content_hash(request))
        def generate(user_input, filename, state, **kwargs):
            state.set("result_kind", "generated")
            state.set("generation_status", "succeeded")
            main._render_blueprint(PPTFinalBlueprint.model_validate(self.blueprint), filename)
            return filename
        def reviewed(state, canvas, *args):
            record_model_response("qwen3.8-flash", {"total_tokens": 42}, "synthetic-vision-request")
            return VisualPageReview(slide_id=canvas.slide_id)
        with patch.dict(os.environ, {"DASHSCOPE_API_KEY": "test-key"}), patch("config_manager.load_base_url", return_value=""), patch("main.run_ppt_system", side_effect=generate), patch("config.visual.tool_signature", return_value={}), patch("config.visual.render_preview", side_effect=self.preview), patch("config.visual.review_page", side_effect=reviewed):
            execute_job(self.root, job["id"])
        finished = store.job(job["id"])
        self.assertEqual(finished["status"], "succeeded")
        artifacts = store.version_artifacts(finished["result_version_id"])
        self.assertTrue({"pptx", "pdf", "preview", "visual-review"}.issubset({item["role"] for item in artifacts}))
        self.assertTrue(store.verify_assets(artifacts))
        vision = next(stage for stage in store.stages(job["id"]) if stage["stage_key"].startswith("vision/"))
        self.assertEqual(json.loads(vision["usage_json"])[0]["usage"]["total_tokens"], 42)
        self.assertEqual(store.load_state(job["conversation_id"]).get("visual_review")["status"], "passed")


if __name__ == "__main__":
    unittest.main()
