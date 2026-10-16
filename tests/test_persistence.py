import copy
import builtins
import json
import os
import subprocess
import sys
import tempfile
import time
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

import main
from config.archives import create_backup, read_archive, restore_backup
from config.jobs import StageOutcomeUnknown, StageRunner, dispatch_pending, execute_job, task_queue
from config.llm_client import MODEL_REQUEST_GUARD, chat_completion_with_retry, record_model_response
from config.schema import PPTFinalBlueprint, PPTOutline
from config.state import PPTGenerationState
from config.storage import PPTStore, StorageConflict, content_hash
from openai import APIConnectionError, APIStatusError
import httpx
from test_pipeline import blueprint_fixture, outline_fixture


class PersistenceTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory(prefix="ppt_persistence_test_")
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name) / "outputs"
        self.store = PPTStore(self.root)
        for replacement in (
            patch.dict(os.environ, {"DASHSCOPE_API_KEY": "test-key-not-a-credential", "PPT_PREVIEW_ENABLED": "0"}),
            patch("config_manager.load_api_key", return_value=""),
            patch("config_manager.load_base_url", return_value=""),
        ):
            replacement.start()
            self.addCleanup(replacement.stop)

    def job(self, owner="local", conversation=None, base=None, initial=None):
        run_dir = self.root / f"run_{time.time_ns()}"
        run_dir.mkdir()
        state = initial or PPTGenerationState(topic="固定主题")
        state.append_message("user", "固定输入")
        request = {"user_input": "固定输入", "filename": "result.pptx", "requirements": "",
                   "page_numbers": None, "use_incremental": False, "mode": "new"}
        job, _ = self.store.create_job(request, state.snapshot(), run_dir, run_dir.name,
                                       content_hash(request), conversation, base, owner)
        return job

    def runner(self, job, token="worker-token"):
        claimed = self.store.claim_job(job["id"], token, owner=job["owner_id"])
        state = PPTGenerationState.from_snapshot(json.loads(claimed["initial_json"]))
        state.bind_persistence(lambda updated: self.store.save_job_state(job["id"], token, updated.snapshot()), strict=True)
        state.stage_executor = StageRunner(self.store, claimed, token, state)
        return state

    def clean_worker_environment(self):
        environment = os.environ.copy()
        if os.name == "nt":
            environment["PATH"] = str(Path(os.environ.get("SystemRoot", "C:/Windows")) / "System32")
        for key in ("CONDA_PREFIX", "CAIROCFFI_DLL_DIRECTORIES", "PPT_CAIRO_BIN"):
            environment.pop(key, None)
        return environment

    def test_render_import_failure_is_committed_and_can_retry(self):
        from config_manager import RenderDependencyError
        job = self.job()
        original_import = builtins.__import__
        def failing_import(name, *args, **kwargs):
            if name == "main":
                raise RenderDependencyError("SVG 渲染运行时不可用；请检查 Cairo")
            return original_import(name, *args, **kwargs)
        with patch("builtins.__import__", side_effect=failing_import):
            execute_job(self.root, job["id"])
        failed = self.store.job(job["id"])
        self.assertEqual(failed["status"], "failed")
        self.assertEqual(failed["error"], "RenderDependencyError")
        self.assertFalse(self.store.pending_dispatch(stale_seconds=0))
        self.assertEqual(self.store.stages(job["id"]), [])
        self.assertIn("请检查 Cairo", json.loads(failed["snapshot_json"])["status"])
        self.assertEqual(self.store.versions(job["conversation_id"]), [])
        self.store.retry_job(job["id"])
        with patch.dict(os.environ, {"DASHSCOPE_API_KEY": ""}):
            execute_job(self.root, job["id"])
        self.assertEqual(self.store.job(job["id"])["status"], "succeeded")

    def test_runtime_check_has_no_queue_or_model_side_effect(self):
        job = self.job()
        result = subprocess.run([sys.executable, "-X", "utf8", "worker.py", "--root", str(self.root), "--check-runtime"],
                                capture_output=True, text=True, encoding="utf-8", timeout=15,
                                env=self.clean_worker_environment())
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(json.loads(result.stdout)["status"], "ready")
        self.assertEqual(self.store.job(job["id"])["status"], "queued")
        self.assertEqual(self.store.stages(job["id"]), [])
        self.assertFalse((self.root / "queue.sqlite3").exists())

    def fake_agents(self, image=False):
        def planner(state, **kwargs):
            value = outline_fixture()
            if image:
                value["slides"][0]["image_prompt"] = "固定配图"
            result = PPTOutline.model_validate(value)
            state.append_message("assistant", "模拟大纲")
            state.set("outline", result.model_dump())
            record_model_response("qwen3.8-flash", {"total_tokens": 123}, "synthetic-request")
            return result
        def writer(state, raw):
            result = PPTOutline.model_validate(copy.deepcopy(raw.model_dump()))
            state.append_message("assistant", "模拟文案")
            state.set("outline", result.model_dump())
            return result
        def generate(**kwargs):
            from PIL import Image
            path = Path(kwargs["output_dir"]) / kwargs["filename"]
            path.parent.mkdir(parents=True, exist_ok=True)
            Image.new("RGB", (16, 9), "white").save(path)
            return str(path)
        def layout(state, payload):
            state.append_message("assistant", payload["slides"][0]["title"])
            canvas = blueprint_fixture()["canvases"][0]
            canvas['elements'][0]['paragraphs'][0]['font_size'] = 32
            if payload["slides"][0].get("image_path"):
                canvas["elements"].append({"type": "image", "left": 8, "top": 2, "width": 4, "height": 2.25,
                                            "image_path": payload["slides"][0]["image_path"]})
            return PPTFinalBlueprint.model_validate({"canvases": [canvas]})
        agents = [Mock(side_effect=function) for function in (planner, writer, generate, layout)]
        for name, agent in zip(("get_planner_agent", "get_writer_agent", "get_image_generator", "get_layout_agent"), agents):
            replacement = patch.object(main, name, return_value=agent)
            replacement.start()
            self.addCleanup(replacement.stop)
        replacement = patch.object(main, "review_ppt_outline", return_value=(True, []))
        replacement.start()
        self.addCleanup(replacement.stop)
        return agents

    def test_render_retry_reuses_models_images_and_usage(self):
        job = self.job()
        planner, writer, image, layout = self.fake_agents(image=True)
        with patch.object(main, "_render_blueprint", side_effect=ValueError("模拟本地渲染失败")):
            execute_job(self.root, job["id"])
        self.assertEqual(self.store.job(job["id"])["status"], "failed")
        self.store.retry_job(job["id"])
        execute_job(self.root, job["id"])
        result = self.store.job(job["id"])
        self.assertEqual(result["status"], "succeeded")
        self.assertEqual((planner.call_count, writer.call_count, image.call_count, layout.call_count), (1, 1, 1, 3))
        self.assertEqual(len(self.store.versions(job["conversation_id"])), 1)
        version = self.store.version(result["result_version_id"])
        snapshot = json.loads(version["snapshot_json"])
        self.assertTrue(snapshot["data"]["outline"]["slides"][0]["image_path"])
        self.assertTrue(self.store.verify_assets(json.loads(version["manifest_json"])))
        calls = [json.loads(stage["usage_json"]) for stage in self.store.stages(job["id"]) if stage["usage_json"]]
        self.assertEqual(calls[0][0]["usage"]["total_tokens"], 123)
        self.assertEqual(len(self.store.messages(job["conversation_id"])), 1)

    def test_fallback_survives_render_retry_without_repeating_models_and_is_exposed_by_web(self):
        job = self.job()
        agents = self.fake_agents()
        def bad_layout(state, payload):
            canvas = blueprint_fixture()['canvases'][0]
            canvas['elements'][0]['top'] = 8
            return PPTFinalBlueprint.model_validate({'canvases':[canvas]})
        agents[3].side_effect = bad_layout
        with patch.object(main,'_render_blueprint',side_effect=ValueError('本地导出故障')):
            execute_job(self.root,job['id'])
        self.assertEqual(self.store.job(job['id'])['status'],'failed')
        counts = tuple(agent.call_count for agent in agents)
        self.store.retry_job(job['id'])
        execute_job(self.root,job['id'])
        done = self.store.job(job['id'])
        self.assertEqual(done['status'],'succeeded')
        self.assertEqual(tuple(agent.call_count for agent in agents),counts)
        snapshot = json.loads(self.store.version(done['result_version_id'])['snapshot_json'])
        self.assertTrue(snapshot['data']['layout_fallbacks'])
        local = [stage for stage in self.store.stages(job['id']) if stage['stage_key'].endswith('/fallback')]
        self.assertTrue(all(not stage['external'] and not stage['usage_json'] for stage in local))
        import app as web
        with patch.object(web,'OUTPUT_ROOT',self.root):
            response = web.app.test_client().get('/state?conversation_id='+job['conversation_id'])
        self.assertEqual(response.json['layout_fallbacks'],snapshot['data']['layout_fallbacks'])
        self.assertEqual(len(self.store.versions(job['conversation_id'])),1)

    def test_locked_text_revision_commits_parent_and_failed_revision_keeps_current(self):
        state = PPTGenerationState(topic="固定主题")
        state.set("outline", outline_fixture())
        state.set("blueprint", blueprint_fixture())
        job = self.job(initial=state)
        claimed = self.runner(job)
        original = self.root / job["run_dir"] / "result.pptx"
        main._render_blueprint(PPTFinalBlueprint.model_validate(state.get("blueprint")), str(original))
        claimed.set("last_file", str(original))
        first = self.store.finish_job(job["id"], "worker-token", claimed.snapshot(), [self.store.describe_asset(original, "pptx")])
        first_snapshot = self.store.version(first)["snapshot_json"]
        first_bytes = original.read_bytes()

        def revision(base, too_long=False):
            draft = PPTGenerationState.from_snapshot(json.loads(self.store.version(base)["snapshot_json"]))
            draft.set_operation("incremental_page", "text", pages=[2])
            run_dir = self.root / f"revision_{time.time_ns()}"
            run_dir.mkdir()
            request = {"user_input": "更新文字", "filename": "revision.pptx", "requirements": "保持布局", "use_incremental": True, "page_numbers": [2], "mode": "modify_text"}
            pending, _ = self.store.create_job(request, draft.snapshot(), run_dir, run_dir.name, content_hash(request), job["conversation_id"], base)
            def writer(state, payload):
                payload = copy.deepcopy(payload)
                payload["slides"][0]["full_content"] = ["中文长文本" * 200 if too_long else "版本更新后的正文"]
                return PPTOutline.model_validate(payload)
            with patch.object(main, "get_writer_agent", return_value=writer), patch.object(main, "get_layout_agent") as layout, patch.object(main, "get_image_generator") as image:
                execute_job(self.root, pending["id"])
                layout.assert_not_called(); image.assert_not_called()
            return self.store.job(pending["id"])

        success = revision(first)
        self.assertEqual(success["status"], "succeeded")
        second = success["result_version_id"]
        self.assertEqual(self.store.version(second)["parent_id"], first)
        self.assertEqual(self.store.version(first)["snapshot_json"], first_snapshot)
        self.assertEqual(original.read_bytes(), first_bytes)
        failed = revision(second, too_long=True)
        self.assertEqual(failed["status"], "failed")
        self.assertEqual(self.store.conversation(job["conversation_id"])["current_version_id"], second)
        self.assertEqual(len(self.store.versions(job["conversation_id"])), 2)
        self.assertTrue(any("超出容量" in message for message in json.loads(failed["snapshot_json"])["progress_log"]))

    def test_missing_cached_image_needs_confirmation_before_regeneration(self):
        job = self.job()
        _, _, image, _ = self.fake_agents(image=True)
        with patch.object(main, "_render_blueprint", side_effect=ValueError("模拟渲染失败")):
            execute_job(self.root, job["id"])
        image_stage = next(stage for stage in self.store.stages(job["id"]) if stage["stage_key"] == "image/page1")
        self.store.asset_path(json.loads(image_stage["files_json"])[0]["path"]).unlink()
        self.store.retry_job(job["id"])
        execute_job(self.root, job["id"])
        self.assertEqual(self.store.job(job["id"])["status"], "unknown")
        self.assertEqual(image.call_count, 1)
        with self.assertRaises(StorageConflict):
            self.store.retry_job(job["id"])
        self.store.retry_job(job["id"], confirm_unknown=True)
        execute_job(self.root, job["id"])
        self.assertEqual(self.store.job(job["id"])["status"], "succeeded")
        self.assertEqual(image.call_count, 2)

    def test_transaction_rejects_changed_baseline_and_expired_worker(self):
        job = self.job()
        state = self.runner(job)
        path = self.root / job["run_dir"] / "result.pptx"
        main._render_offline_demo("固定主题", str(path))
        manifest = [self.store.describe_asset(path, "pptx")]
        first = self.store.finish_job(job["id"], "worker-token", state.snapshot(), manifest)
        next_job = self.job(conversation=job["conversation_id"], base=first)
        draft = self.runner(next_job)
        with self.store.connection(write=True) as db:
            db.execute("UPDATE decks SET current_version_id=NULL WHERE id=?", (job["deck_id"],))
        with self.assertRaises(StorageConflict):
            self.store.finish_job(next_job["id"], "worker-token", draft.snapshot(), manifest)
        self.assertEqual(len(self.store.versions(job["conversation_id"])), 1)
        with self.store.connection(write=True) as db:
            db.execute("UPDATE jobs SET lease_until=0 WHERE id=?", (next_job["id"],))
        with self.assertRaises(StorageConflict):
            self.store.save_job_state(next_job["id"], "worker-token", draft.snapshot())
        self.assertTrue(path.is_file())

    def test_scope_is_enforced_for_versions_jobs_assets_and_messages(self):
        job = self.job(owner="other-owner")
        with patch.dict(os.environ, {"DASHSCOPE_API_KEY": ""}):
            execute_job(self.root, job["id"])
        result = self.store.job(job["id"], "other-owner")
        self.assertEqual(result["status"], "succeeded")
        artifact = self.store.version_artifacts(result["result_version_id"], "other-owner")[0]
        self.assertIsNone(self.store.job(job["id"]))
        self.assertIsNone(self.store.conversation(job["conversation_id"]))
        self.assertIsNone(self.store.version(result["result_version_id"]))
        self.assertIsNone(self.store.artifact(artifact["id"]))
        self.assertEqual(self.store.messages(job["conversation_id"]), [])

    def test_real_process_interruption_distinguishes_local_and_external_stages(self):
        for external in (False, True):
            with self.subTest(external=external):
                job = self.job()
                marker = self.root / f"marker_{external}"
                script = """
import json, sys, time
from pathlib import Path
from config.storage import PPTStore
from config.state import PPTGenerationState
from config.jobs import StageRunner
root, job_id, marker, external = sys.argv[1:]
store = PPTStore(root)
job = store.claim_job(job_id, 'terminated-worker')
state = PPTGenerationState.from_snapshot(json.loads(job['initial_json']))
state.stage_executor = StageRunner(store, job, 'terminated-worker', state)
state.execute_stage('planner/completed', {'input': 1}, lambda: {'planned': True})
store.start_stage(job_id, 'terminated-worker', 'image/pending' if external == '1' else 'render', 'hash', None, external == '1')
Path(marker).write_text('ready')
time.sleep(30)
"""
                process = subprocess.Popen([sys.executable, "-X", "utf8", "-c", script, str(self.root), job["id"], str(marker), str(int(external))],
                                           stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
                try:
                    deadline = time.monotonic() + 10
                    while not marker.exists() and process.poll() is None and time.monotonic() < deadline:
                        time.sleep(0.05)
                    self.assertTrue(marker.exists(), "测试子进程未进入目标阶段")
                    process.kill()
                    process.communicate(timeout=5)
                finally:
                    if process.poll() is None:
                        process.kill()
                        process.communicate(timeout=5)
                with self.store.connection(write=True) as db:
                    db.execute("UPDATE jobs SET lease_until=0 WHERE id=?", (job["id"],))
                recovered = self.store.recover_expired()
                self.assertEqual(recovered[0]["status"], "unknown" if external else "queued")
                with self.assertRaises(StorageConflict):
                    self.store.heartbeat(job["id"], "terminated-worker")
                if external:
                    with self.assertRaises(StorageConflict):
                        self.store.retry_job(job["id"])
                    self.store.retry_job(job["id"], confirm_unknown=True)
                state = self.runner(job, token="new-worker")
                operation = Mock(side_effect=AssertionError("已完成模型阶段不能重复调用"))
                self.assertEqual(state.execute_stage("planner/completed", {"input": 1}, operation), {"planned": True})
                operation.assert_not_called()
                self.store.fail_job(job["id"], "new-worker", state.snapshot(), "end-test")

    def test_durable_network_timeout_does_not_retry(self):
        create = Mock(side_effect=APIConnectionError(request=httpx.Request("POST", "https://example.invalid")))
        client = SimpleNamespace(chat=SimpleNamespace(completions=SimpleNamespace(create=create)))
        token = MODEL_REQUEST_GUARD.set(lambda: None)
        try:
            with patch("config.llm_client.time.sleep") as sleep:
                with self.assertRaises(APIConnectionError):
                    chat_completion_with_retry(client, model="test", messages=[])
                sleep.assert_not_called()
        finally:
            MODEL_REQUEST_GUARD.reset(token)
        self.assertEqual(create.call_count, 1)

    def test_failed_provider_response_keeps_available_request_id(self):
        job = self.job()
        state = self.runner(job)
        response = httpx.Response(400, headers={"x-request-id": "synthetic-failed-request"},
                                  request=httpx.Request("POST", "https://example.invalid"))
        operation = Mock(side_effect=APIStatusError("synthetic error", response=response, body=None))
        with self.assertRaises(APIStatusError):
            state.execute_stage("planner/failure", {}, operation, model="qwen3.8-flash")
        stage = self.store.stages(job["id"])[0]
        self.assertEqual(stage["status"], "failed")
        self.assertEqual(stage["request_id"], "synthetic-failed-request")

    def test_persistent_huey_queue_and_separate_worker_once(self):
        job = self.job()
        queue = dispatch_pending(self.root)
        queue.storage.close()
        script = """
import os, runpy, sys
from pathlib import Path
import config_manager
config_manager.CONFIG_FILE = Path(sys.argv[1]) / 'no-config.json'
os.environ['DASHSCOPE_API_KEY'] = ''
sys.argv = ['worker.py', '--root', sys.argv[1], '--once']
runpy.run_path('worker.py', run_name='__main__')
"""
        result = subprocess.run([sys.executable, "-X", "utf8", "-c", script, str(self.root)], capture_output=True, text=True, encoding="utf-8", timeout=15,
                                env=self.clean_worker_environment())
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(self.store.job(job["id"])["status"], "succeeded")
        execute_job(self.root, job["id"])
        self.assertEqual(len(self.store.versions(job["conversation_id"])), 1)

    def test_lost_queue_delivery_is_compensated_without_duplicate_version(self):
        job = self.job()
        queue = dispatch_pending(self.root)
        try:
            self.assertIsNotNone(queue.dequeue())
            with self.store.connection(write=True) as db:
                db.execute("UPDATE jobs SET dispatched_at=0 WHERE id=?", (job["id"],))
            dispatch_pending(self.root, recover=True)
            task = queue.dequeue()
            self.assertIsNotNone(task)
            with patch.dict(os.environ, {"DASHSCOPE_API_KEY": ""}):
                queue.execute(task)
            execute_job(self.root, job["id"])
            self.assertEqual(len(self.store.versions(job["conversation_id"])), 1)
        finally:
            queue.storage.close()

    def test_continuous_consumer_executes_a_persisted_task(self):
        job = self.job()
        dialogue_dir = self.root / "discussion"; dialogue_dir.mkdir()
        request = {"kind": "dialogue", "user_input": "如何安排四页？", "intent": "discuss", "page_numbers": []}
        dialogue, _ = self.store.create_job(request, PPTGenerationState(topic="讨论").snapshot(), dialogue_dir,
                                           "consumer-discussion", content_hash(request))
        queue = dispatch_pending(self.root)
        queue.storage.close()
        script = """
import os, runpy, sys
from pathlib import Path
import config_manager
config_manager.CONFIG_FILE = Path(sys.argv[1]) / 'no-config.json'
os.environ['DASHSCOPE_API_KEY'] = ''
sys.argv = ['worker.py', '--root', sys.argv[1]]
runpy.run_path('worker.py', run_name='__main__')
"""
        process = subprocess.Popen([sys.executable, "-X", "utf8", "-c", script, str(self.root)], stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                                   env=self.clean_worker_environment())
        try:
            deadline = time.monotonic() + 10
            while any(self.store.job(item["id"])["status"] != "succeeded" for item in (job, dialogue)) and process.poll() is None and time.monotonic() < deadline:
                time.sleep(0.05)
            self.assertEqual(self.store.job(job["id"])["status"], "succeeded")
            self.assertEqual(self.store.job(dialogue["id"])["status"], "succeeded")
            self.assertEqual(len(self.store.dialogue_records(dialogue["conversation_id"])), 2)
        finally:
            if process.poll() is None:
                process.terminate()
            try:
                process.communicate(timeout=5)
            except subprocess.TimeoutExpired:
                process.kill()
                process.communicate(timeout=5)

    def test_schema_upgrade_preserves_jobs(self):
        job = self.job()
        with self.store.connection(write=True) as db:
            db.execute("ALTER TABLE stages DROP COLUMN prompt_version")
            db.execute("PRAGMA user_version=1")
        upgraded = PPTStore(self.root)
        self.assertEqual(upgraded.job(job["id"])["initial_json"], job["initial_json"])
        with upgraded.connection() as db:
            self.assertEqual(db.execute("PRAGMA user_version").fetchone()[0], 2)
            self.assertIn("prompt_version", {row[1] for row in db.execute("PRAGMA table_info(stages)")})

    def test_backup_restore_rebases_assets_and_preserves_user_text(self):
        job = self.job()
        with patch.dict(os.environ, {"DASHSCOPE_API_KEY": ""}):
            execute_job(self.root, job["id"])
        result = self.store.job(job["id"])
        raw_text = f"原始路径文本 {self.root} 保持不变"
        pending_state = PPTGenerationState(topic=raw_text)
        pending = self.job(initial=pending_state)
        queue = dispatch_pending(self.root)
        queue.storage.close()
        backup = Path(self.tmp.name) / "backup"
        restored = Path(self.tmp.name) / "restored"
        create_backup(self.store, backup)
        restore_backup(backup, restored)
        recovered = PPTStore(restored)
        state = recovered.load_state(job["conversation_id"])
        self.assertTrue(Path(state.get("last_file")).is_relative_to(restored))
        self.assertTrue(recovered.verify_assets(json.loads(recovered.version(result["result_version_id"])["manifest_json"])))
        self.assertEqual(recovered.load_state(pending["conversation_id"]).topic, raw_text)
        self.assertEqual(recovered.job(pending["id"])["status"], "queued")
        with self.assertRaises(ValueError):
            restore_backup(backup, restored)
        with self.assertRaises(ValueError):
            create_backup(self.store, self.root / "nested-backup")
        asset = next(path for path in backup.rglob("*.pptx"))
        asset.write_bytes(b"corrupt")
        with self.assertRaises(ValueError):
            restore_backup(backup, Path(self.tmp.name) / "invalid-restore")

    def test_archive_and_asset_paths_cannot_escape_root(self):
        for path in ("../outside", str(Path(self.tmp.name) / "outside")):
            with self.subTest(path=path), self.assertRaises(ValueError):
                self.store.asset_path(path)
        with self.assertRaises(ValueError):
            read_archive(self.store, "../outside")

    def test_recovery_keeps_already_unknown_external_call_pending(self):
        job = self.job()
        self.runner(job)
        stage = self.store.start_stage(job["id"], "worker-token", "image/pending", "hash", "image", True)
        self.store.fail_stage(stage, job["id"], "worker-token", "unknown", "timeout")
        with self.store.connection(write=True) as db:
            db.execute("UPDATE jobs SET lease_until=0 WHERE id=?", (job["id"],))
        self.assertEqual(self.store.recover_expired()[0]["status"], "unknown")

    def test_restore_expires_old_worker_lease_without_repeating_external_call(self):
        job = self.job()
        self.runner(job)
        self.store.start_stage(job["id"], "worker-token", "image/pending", "hash", "image", True)
        backup, restored = Path(self.tmp.name) / "backup", Path(self.tmp.name) / "restored"
        create_backup(self.store, backup)
        restore_backup(backup, restored)
        recovered = PPTStore(restored)
        self.assertEqual(recovered.job(job["id"])["status"], "unknown")
        self.assertEqual(self.store.job(job["id"])["status"], "running")
        with self.assertRaises(StorageConflict):
            recovered.claim_job(job["id"], "worker-token") or recovered.heartbeat(job["id"], "worker-token")


if __name__ == "__main__":
    unittest.main()
