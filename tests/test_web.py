import json
import os
import tempfile
import unittest
from pptx import Presentation
from pathlib import Path
from unittest.mock import patch

import app as web
import main
from config.archives import read_archive
from config.jobs import execute_job
from config.state import PPTGenerationState
from main import _render_offline_demo


class WebGenerationTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory(prefix="ppt_web_test_")
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name)
        for replacement in (
            patch.object(web, "OUTPUT_ROOT", self.root),
            patch.object(web, "dispatch_pending", side_effect=self.dispatch),
            patch("config_manager.load_api_key", return_value=""),
            patch("config_manager.load_base_url", return_value=""),
            patch.dict(os.environ, {"DASHSCOPE_API_KEY": "", "PPT_PREVIEW_ENABLED": "0"}),
        ):
            replacement.start()
            self.addCleanup(replacement.stop)
        self.client = web.app.test_client()

    def dispatch(self, root):
        for job in web._store().pending_dispatch():
            execute_job(root, job["id"])

    def seed_cache(self, canvases=3):
        run_dir = self.root / "previous"
        run_dir.mkdir()
        state = PPTGenerationState(topic="原主题")
        state.set("outline", {"topic": "原主题", "slides": [
            {"title": f"第 {number} 页"} for number in range(1, 4)
        ]})
        state.set("blueprint", {"canvases": [{"elements": []} for _ in range(canvases)]})
        original = run_dir / "original.pptx"
        _render_offline_demo("原主题", str(original))
        state.set("last_file", str(original))
        state.last_updated = "2026-10-11 10:00:00"
        (run_dir / "state.json").write_text(
            json.dumps(state.snapshot(), ensure_ascii=False), encoding="utf-8"
        )
        snapshot, manifest, digest = read_archive(web._store(), run_dir.name)
        conversation_id = web._store().import_legacy(run_dir.name, digest, snapshot, manifest)
        with self.client.session_transaction() as session:
            session[web.SESSION_ID_KEY] = conversation_id
        return run_dir

    def fake_generation(self, user_input, filename, state, **kwargs):
        state.set("result_kind", "generated")
        return _render_offline_demo(user_input, filename)

    def test_preview_is_bound_to_committed_version_and_inline_png_checks_hash(self):
        from PIL import Image
        from config.storage import content_hash
        self.seed_cache()
        store = web._store()
        with self.client.session_transaction() as session:
            conversation = session[web.SESSION_ID_KEY]
        base = store.load_state(conversation)
        run = self.root / "preview-version"
        run.mkdir()
        image = run / "page-1.png"
        Image.new("RGB", (16, 9), "white").save(image)
        request = {"user_input": "预览回归", "filename": "result.pptx", "requirements": "",
                   "page_numbers": None, "use_incremental": False, "mode": "new"}
        job, _ = store.create_job(request, base.snapshot(), run, "preview-job", content_hash(request),
                                 conversation, base.get("_version_id"))
        store.claim_job(job["id"], "preview-worker")
        version = store.finish_job(job["id"], "preview-worker", base.snapshot(),
                                   [store.describe_asset(base.get("last_file"), "pptx"), store.describe_asset(image, "preview")])
        data = self.client.get("/state").get_json()
        url = data["preview"][0]["image_url"]
        response = self.client.get(url)
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.headers["Content-Type"], "image/png")
        self.assertTrue(response.headers["Content-Disposition"].startswith("inline"))
        self.assertEqual(response.headers["X-Content-Type-Options"], "nosniff")
        draft = store.load_state(conversation)
        draft.set("outline", {"slides": [{"title": "失败草稿标题"}]})
        failed_run = self.root / "failed-preview"
        failed_run.mkdir()
        failed, _ = store.create_job(request, draft.snapshot(), failed_run, "failed-preview", content_hash(request), conversation, version)
        store.claim_job(failed["id"], "failed-worker")
        store.fail_job(failed["id"], "failed-worker", draft.snapshot(), "synthetic-failure")
        self.assertEqual(self.client.get("/state").get_json()["preview"][0]["title"], base.get("outline")["slides"][0]["title"])
        image.write_bytes(b"modified")
        self.assertEqual(self.client.get(url).status_code, 409)
        self.assertNotIn("image_url", self.client.get("/state").get_json()["preview"][0])

    def test_page_ranges_with_spaces(self):
        self.assertEqual(web._parse_page_numbers("1 - 3，5; 3"), [1, 2, 3, 5])
        self.assertEqual(web._parse_page_numbers("3-1"), [1, 2, 3])

    def test_oversized_range_is_rejected_before_expansion(self):
        self.assertIsNone(web._parse_page_numbers("1-999999999"))
        self.assertIsNone(web._parse_page_numbers("1-999999999", max_page=3))

    def test_invalid_page_input_never_runs_pipeline_or_changes_version(self):
        previous = self.seed_cache()
        for raw in ("invalid-page", "0", "4", "1,4", "-1", "1-10000"):
            with self.subTest(raw=raw), patch.object(main, "run_ppt_system", return_value=None) as full, patch.object(
                main, "run_incremental_modify_pages", return_value=None
            ) as incremental:
                response = self.client.post("/generate", data={
                    "topic": "改标题", "mode": "modify_text", "modify_page": raw,
                })
                self.assertEqual(response.status_code, 400)
                full.assert_not_called()
                incremental.assert_not_called()
                self.assertEqual(self.client.get("/state").json["last_file"], str(previous / "original.pptx"))
                self.assertEqual([p.name for p in self.root.iterdir() if p.is_dir()], ["previous"])

    def test_blank_pages_mean_global_edit(self):
        self.seed_cache()
        with patch.object(main, "run_incremental_modify_pages", side_effect=self.fake_generation) as incremental:
            response = self.client.post("/generate", data={
                "topic": "改标题", "mode": "modify_text", "modify_page": "",
            })
        self.assertEqual(response.status_code, 200)
        self.assertEqual(incremental.call_args.kwargs["page_numbers"], [1, 2, 3])

    def test_valid_target_pages_use_incremental_pipeline(self):
        self.seed_cache()
        with patch.object(main, "run_ppt_system") as full, patch.object(
            main, "run_incremental_modify_pages", side_effect=self.fake_generation
        ) as incremental:
            response = self.client.post("/generate", data={
                "topic": "调整设计", "mode": "modify_design", "modify_page": "1 - 2",
            })
        self.assertEqual(response.status_code, 200)
        full.assert_not_called()
        self.assertEqual(incremental.call_args.kwargs["page_numbers"], [1, 2])

    def test_modification_without_cache_is_rejected(self):
        with patch.object(main, "run_ppt_system") as full:
            response = self.client.post("/generate", data={"topic": "改标题", "mode": "modify_text"})
        self.assertEqual(response.status_code, 400)
        full.assert_not_called()
        self.assertIsNone(self.client.get("/state").json["conversation_id"])

    def test_mismatched_blueprint_is_rejected_before_task_creation(self):
        previous = self.seed_cache(canvases=2)
        with patch.object(main, "run_incremental_modify_pages", return_value=None) as incremental:
            response = self.client.post("/generate", data={"topic": "修改", "mode": "modify_design"})
        self.assertEqual(response.status_code, 400)
        incremental.assert_not_called()
        self.assertEqual(self.client.get("/state").json["last_file"], str(previous / "original.pptx"))

    def test_unknown_mode_is_rejected(self):
        with patch.object(main, "run_ppt_system") as full:
            response = self.client.post("/generate", data={"topic": "主题", "mode": "unexpected"})
        self.assertEqual(response.status_code, 400)
        full.assert_not_called()

    def test_run_directories_do_not_collide_with_same_timestamp(self):
        with patch.object(web, "datetime") as clock:
            clock.now.return_value.strftime.return_value = "20261011_100000"
            first, second = web._create_run_dir(), web._create_run_dir()
        self.assertNotEqual(first, second)
        self.assertTrue(first.is_dir() and second.is_dir())

    def test_filename_preserves_one_extension_and_chinese_name(self):
        for name, expected in (
            ("Coffee.pptx", "Coffee.pptx"), ("REPORT.PPTX", "REPORT.pptx"),
            ("季度汇报.pptx", "季度汇报.pptx"), ("../../report.pptx", "report.pptx"),
            ("CON.pptx", "presentation_CON.pptx"),
        ):
            with self.subTest(name=name), patch.object(main, "run_ppt_system", side_effect=self.fake_generation) as full:
                response = self.client.post("/generate", data={"topic": "主题", "filename": name})
                self.assertEqual(response.status_code, 200)
                self.assertEqual(Path(full.call_args.kwargs["filename"]).name, expected)

    def test_pipeline_failure_does_not_create_success_snapshot(self):
        with patch.object(main, "run_ppt_system", return_value=None):
            self.client.post("/generate", data={"topic": "主题"})
        with self.client.session_transaction() as session:
            self.assertEqual(list(session), [web.SESSION_ID_KEY])
            self.assertEqual(web._store().load_state(session[web.SESSION_ID_KEY]).history_snapshots, [])
        response = self.client.get("/state")
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.json["generation_status"], "failed")
        self.assertEqual(response.json["last_file"], "")

    def test_broken_export_is_reported_as_failure(self):
        def broken_export(user_input, filename, **kwargs):
            Path(filename).write_bytes(b"not a pptx")
            return filename
        with patch.object(main, "run_ppt_system", side_effect=broken_export):
            self.client.post("/generate", data={"topic": "主题"})
        response = self.client.get("/state")
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.json["generation_status"], "failed")
        self.assertEqual(response.json["last_file"], "")

    def test_failed_modification_keeps_last_successful_file(self):
        previous = self.seed_cache()
        original_file = previous / "original.pptx"
        original_bytes = original_file.read_bytes()
        with patch.object(main, "run_incremental_modify_pages", return_value=None):
            self.client.post("/generate", data={"topic": "修改", "mode": "modify_text", "modify_page": "2"})
        response = self.client.get("/state")
        self.assertEqual(response.json["generation_status"], "failed")
        self.assertEqual(response.json["last_file"], str(original_file))
        self.assertEqual(original_file.read_bytes(), original_bytes)

    def test_state_restores_timestamp_and_disables_caching(self):
        self.seed_cache()
        response = self.client.get("/state")
        self.assertEqual(response.json["last_updated"], "2026-10-11 10:00:00")
        self.assertIn("no-store", response.headers["Cache-Control"])

    def test_unreadable_previous_file_does_not_break_state_endpoint(self):
        previous = self.seed_cache()
        broken_path = previous / "original.pptx"
        broken_path.write_bytes(b"broken")
        self.assertEqual(self.client.get("/state").status_code, 200)

    def test_cookie_contains_only_id_and_other_browser_starts_blank(self):
        self.seed_cache()
        with self.client.session_transaction() as session:
            conversation_id = session[web.SESSION_ID_KEY]
            self.assertEqual(list(session), [web.SESSION_ID_KEY])
        (self.root / "current_session.txt").write_text("previous", encoding="utf-8")
        other = web.app.test_client()
        self.assertIsNone(other.get("/state").json["conversation_id"])
        self.assertEqual(other.post(f"/api/conversations/{conversation_id}/select").status_code, 200)
        self.assertEqual(other.get("/state").json["conversation_id"], conversation_id)
        other.post("/api/reset")
        self.assertIsNone(other.get("/state").json["conversation_id"])
        self.assertEqual(self.client.get("/state").json["conversation_id"], conversation_id)
        self.assertEqual((self.root / "current_session.txt").read_text(), "previous")

    def test_duplicate_submission_is_idempotent_and_different_input_conflicts(self):
        form = {"topic": "主题", "request_id": "test-request-123"}
        with patch.object(web, "dispatch_pending"):
            self.assertEqual(self.client.post("/generate", data=form).status_code, 200)
            job_id = self.client.get("/state").json["job_id"]
            self.assertEqual(self.client.post("/generate", data=form).status_code, 200)
            self.assertEqual(self.client.get("/state").json["job_id"], job_id)
            self.assertEqual(self.client.post("/generate", data=form | {"topic": "不同主题"}).status_code, 409)
        self.assertEqual(len(web._store().conversations()), 1)
        self.assertEqual(len([path for path in self.root.iterdir() if path.is_dir()]), 1)

    def test_pending_modification_and_stale_baseline_are_visible_conflicts(self):
        self.seed_cache()
        info = self.client.get("/state").json
        form = {"topic": "修改", "mode": "modify_text", "base_version_id": info["version_id"],
                "conversation_id": info["conversation_id"]}
        with patch.object(web, "dispatch_pending"):
            self.assertEqual(self.client.post("/generate", data=form).status_code, 200)
            self.assertEqual(self.client.post("/generate", data=form).status_code, 409)
        with patch.object(main, "run_incremental_modify_pages", side_effect=self.fake_generation):
            self.dispatch(self.root)
        self.assertEqual(self.client.post("/generate", data=form).status_code, 409)

    def test_api_owner_filter_and_hashed_download(self):
        previous = self.seed_cache()
        info = self.client.get("/state").json
        download = info["download_url"]
        response = self.client.get(download)
        self.assertEqual(response.data, (previous / "original.pptx").read_bytes())
        response.close()
        with patch.object(web, "_owner", return_value="different-owner"):
            self.assertEqual(self.client.get("/api/conversations").json["conversations"], [])
            self.assertEqual(self.client.post(f"/api/conversations/{info['conversation_id']}/select").status_code, 404)
            self.assertEqual(self.client.get(download).status_code, 404)
        (previous / "original.pptx").write_bytes(b"corrupt")
        self.assertEqual(self.client.get(download).status_code, 409)

    def test_reopening_store_restores_current_version(self):
        self.seed_cache()
        before = self.client.get("/state").json
        web._cached_store.cache_clear()
        after = self.client.get("/state").json
        self.assertEqual(before, after)

    def test_unknown_retry_requires_explicit_confirmation(self):
        with patch.object(web, "dispatch_pending"):
            self.client.post("/generate", data={"topic": "主题"})
            info = self.client.get("/state").json
            store = web._store()
            job = store.claim_job(info["job_id"], "token")
            store.fail_job(job["id"], "token", json.loads(job["snapshot_json"]), "timeout", status="unknown")
            self.assertEqual(self.client.post(f"/api/jobs/{job['id']}/retry", json={}).status_code, 409)
            self.assertEqual(self.client.post(f"/api/jobs/{job['id']}/retry", json={"confirm_unknown": True}).status_code, 202)
            self.assertEqual(self.client.get("/state").json["generation_status"], "queued")

    def test_archive_import_is_explicit_and_does_not_change_original(self):
        previous = self.seed_cache()
        state_file = previous / "state.json"
        before = state_file.read_bytes()
        self.client.post("/api/reset")
        archive = self.client.get("/api/archives").json["archives"][0]
        self.assertIsNone(self.client.get("/state").json["conversation_id"])
        response = self.client.post("/api/archives/previous/import", json={"source_hash": archive["source_hash"]})
        self.assertEqual(response.status_code, 200)
        self.assertEqual(state_file.read_bytes(), before)
        self.assertEqual(len(web._store().conversations()), 1)
        self.assertEqual(self.client.post("/api/archives/previous/import", json={"source_hash": "wrong"}).status_code, 409)

    def test_explicit_and_natural_page_count_is_persisted_and_exported(self):
        for raw, requirements, expected in [("6", "请做三页", 6), ("", "请做成一页", 1), ("", "用于汇报", 3)]:
            with self.subTest(raw=raw, requirements=requirements):
                response = self.client.post("/generate", data={"topic": "页数约束", "page_range": raw, "requirements": requirements})
                self.assertEqual(response.status_code, 200)
                state = self.client.get("/state").json
                self.assertEqual(state["generation_status"], "succeeded")
                self.assertEqual(len(Presentation(state["last_file"]).slides), expected)
                version = web._store().version(state["version_id"])
                saved = json.loads(version["snapshot_json"])
                self.assertTrue(saved["data"]["page_policy"]["min_slides"] <= expected <= saved["data"]["page_policy"]["max_slides"])

    def test_invalid_generation_count_creates_no_job(self):
        for raw in ("0", "51", "3-1", "abc", "2.5"):
            with self.subTest(raw=raw), patch.object(web, "dispatch_pending") as dispatch:
                self.assertEqual(self.client.post("/generate", data={"topic": "主题", "page_range": raw}).status_code, 400)
                dispatch.assert_not_called()
        self.assertEqual(web._store().conversations(), [])

    def test_page_edit_cannot_change_total_count_or_current_version(self):
        self.seed_cache()
        before = self.client.get("/state").json
        with patch.object(main, "run_incremental_modify_pages") as pipeline:
            response = self.client.post("/generate", data={"topic": "修改文字", "mode": "modify_text", "page_range": "6", "modify_page": "2"})
        self.assertEqual(response.status_code, 400)
        pipeline.assert_not_called()
        self.assertEqual(self.client.get("/state").json["version_id"], before["version_id"])

    def test_page_count_is_part_of_idempotency_fingerprint(self):
        form = {"topic": "主题", "request_id": "page-count-key-123", "page_range": "3"}
        self.assertEqual(self.client.post("/generate", data=form).status_code, 200)
        self.assertEqual(self.client.post("/generate", data=form).status_code, 200)
        self.assertEqual(self.client.post("/generate", data={**form, "page_range": "4"}).status_code, 409)


if __name__ == "__main__":
    unittest.main()
