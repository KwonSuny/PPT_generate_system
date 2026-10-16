import copy
import builtins
import json
import os
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

import app as web
import main
from agents.dialogue import DialogueReply, Proposal, get_dialogue_agent, page_catalog, validate_reply, retrieve_pages
from config.context import bounded_messages, token_bound, ContextBudgetError
from config.state import PPTGenerationState
from config.schema import PPTFinalBlueprint
from config.storage import PPTStore, content_hash
from config.jobs import execute_job, StageOutcomeUnknown
from config.archives import read_archive
from test_content import table_slide, canvas_for, diagram_slide


class DialogueTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory(prefix="ppt_dialogue_")
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name)
        self.sent = []
        for replacement in (patch.object(web, "OUTPUT_ROOT", self.root),
                            patch.object(web, "dispatch_pending", side_effect=self.dispatch),
                            patch("config_manager.load_api_key", return_value=""),
                            patch("config_manager.load_base_url", return_value=""),
                            patch.dict(os.environ, {"DASHSCOPE_API_KEY": "", "PPT_PREVIEW_ENABLED": "0"})):
            replacement.start(); self.addCleanup(replacement.stop)
        self.client = web.app.test_client()
        self.store = web._store()

    def dispatch(self, root):
        for job in self.store.pending_dispatch():
            execute_job(root, job["id"])

    def test_discussion_does_not_import_ppt_render_dependencies(self):
        original_import = builtins.__import__
        def guarded_import(name, *args, **kwargs):
            if name in {"main", "renderer", "cairosvg"}:
                raise AssertionError("讨论不能依赖 PPT 渲染模块")
            return original_import(name, *args, **kwargs)
        with patch("builtins.__import__", side_effect=guarded_import):
            response = self.client.post('/api/dialogue', json={"message": "四页汇报如何安排？", "intent": "discuss", "request_id": "render-independent-discussion"})
        self.assertEqual(response.status_code, 202)
        self.assertEqual(self.store.job(response.json["job_id"])["status"], "succeeded")
        data = self.client.get(f'/api/conversations/{response.json["conversation_id"]}/dialogue').json
        self.assertTrue(any(item["role"] == "assistant" for item in data["messages"]))

    def seed(self):
        directory = self.root / "seed"; directory.mkdir()
        slide = table_slide(); canvas = main.PPTXRenderer.resolve_canvas(canvas_for(slide))
        blueprint = PPTFinalBlueprint(canvases=[canvas])
        ppt = directory / "table.pptx"; main._render_blueprint(blueprint, ppt)
        state = PPTGenerationState(topic="指标对比", requirements="数据是合成演示")
        state.set("outline", {"schema_version": 2, "topic": "指标对比", "slides": [slide]})
        state.set("blueprint", blueprint.model_dump()); state.set("last_file", str(ppt))
        (directory / "state.json").write_text(json.dumps(state.snapshot(), ensure_ascii=False), encoding="utf8")
        snapshot, manifest, digest = read_archive(self.store, directory.name)
        conversation = self.store.import_legacy(directory.name, digest, snapshot, manifest)
        self.client.post(f"/api/conversations/{conversation}/select")
        version = self.store.conversation(conversation)["current_version_id"]
        return conversation, version

    def answer(self, state, payload):
        self.sent.append(payload)
        from config.llm_client import record_model_response
        record_model_response("qwen3.8-flash", {"total_tokens": 42}, "mock-request")
        return DialogueReply(answer="基准准确率为 87.65%。", answerable=True,
            citations=[{"page": 1, "ref": "table/rows/accuracy/cells/before", "quote": "87.65%"}])

    def post_question(self, conversation, version, key="dialogue-key-001"):
        return self.client.post("/api/dialogue", json={"intent": "question", "message": "基准准确率是多少？",
            "conversation_id": conversation, "version_id": version, "request_id": key, "pages": [1]})

    def test_question_persists_reply_usage_without_creating_version_or_running_generation(self):
        conversation, version = self.seed()
        initial_status = self.client.get("/state").json["generation_status"]
        with patch.dict(os.environ, {"DASHSCOPE_API_KEY": "test-key"}), patch("agents.dialogue.get_dialogue_agent", return_value=self.answer), patch("main.run_ppt_system") as generate:
            response = self.post_question(conversation, version)
            duplicate = self.post_question(conversation, version)
        self.assertEqual(response.status_code, 202)
        self.assertEqual(response.json["job_id"], duplicate.json["job_id"])
        generate.assert_not_called()
        self.assertEqual(len(self.store.versions(conversation)), 1)
        self.assertEqual(self.store.conversation(conversation)["current_version_id"], version)
        records = self.client.get(f"/api/conversations/{conversation}/dialogue").json["messages"]
        self.assertEqual(records[-1]["response"]["version_id"], version)
        self.assertNotIn("canvases", records[-1]["content"])
        stage = self.store.stages(response.json["job_id"])[0]
        self.assertEqual(json.loads(stage["usage_json"])[0]["usage"]["total_tokens"], 42)
        self.assertEqual(self.client.get('/state').json["generation_status"], initial_status)
        self.assertTrue(self.client.get(f"/api/versions/{version}/preview").json["preview"])

    def test_question_uses_committed_outline_after_failed_generation_draft(self):
        conversation, version = self.seed()
        state = self.store.load_state(conversation); state.set("outline", {"slides": [{"title": "错误草稿"}]})
        run = self.root / "failed"; run.mkdir()
        request = {"user_input": "修改", "filename": "failed.pptx", "requirements": "", "use_incremental": False, "page_numbers": None}
        job, _ = self.store.create_job(request, state.snapshot(), run, "failed-key", content_hash(request), conversation, version)
        self.store.claim_job(job["id"], "worker"); self.store.fail_job(job["id"], "worker", state.snapshot(), "Synthetic")
        with patch.dict(os.environ, {"DASHSCOPE_API_KEY": "test-key"}), patch("agents.dialogue.get_dialogue_agent", return_value=self.answer):
            self.assertEqual(self.post_question(conversation, version).status_code, 202)
        self.assertEqual(self.sent[0]["outline"]["slides"][0]["title"], "指标对比")
        self.assertEqual(self.client.get('/state').json["generation_status"], "failed")

    def test_pinned_question_survives_a_new_current_version(self):
        conversation, version = self.seed()
        with patch.object(web, "dispatch_pending"):
            response = self.post_question(conversation, version)
        state = self.store.load_state(conversation)
        run = self.root / "new-version"; run.mkdir()
        request = {"user_input": "新版本", "filename": "new.pptx", "requirements": "", "use_incremental": False, "page_numbers": None}
        job, _ = self.store.create_job(request, state.snapshot(), run, "new-version-key", content_hash(request), conversation, version)
        self.store.claim_job(job["id"], "new-worker")
        updated = table_slide(); updated["table"]["rows"][0]["cells"][1]["value"] = 0.5
        canvas = main.PPTXRenderer.resolve_canvas(canvas_for(updated))
        blueprint = PPTFinalBlueprint(canvases=[canvas]); output = run / "changed.pptx"
        main._render_blueprint(blueprint, output)
        state.set("outline", {"slides": [updated]}); state.set("blueprint", blueprint.model_dump()); state.set("last_file", str(output))
        new_version = self.store.finish_job(job["id"], "new-worker", state.snapshot(), [self.store.describe_asset(output, "pptx")])
        with patch.dict(os.environ, {"DASHSCOPE_API_KEY": "test-key"}), patch("agents.dialogue.get_dialogue_agent", return_value=self.answer):
            execute_job(self.root, response.json["job_id"])
        self.assertEqual(self.store.job(response.json["job_id"])["status"], "succeeded")
        self.assertEqual(self.sent[0]["version_id"], version)
        self.assertEqual(self.sent[0]["outline"]["slides"][0]["table"]["rows"][0]["cells"][1]["value"], 0.8765)
        self.assertFalse(any(item["content"] == "新版本" for item in self.sent[0]["history"]))
        self.assertEqual(self.store.conversation(conversation)["current_version_id"], new_version)

    def test_discussion_draft_and_generation_share_history_and_survive_restart(self):
        reply = DialogueReply(answer="可以生成四页。", proposal=Proposal(action="new", topic="咖啡", requirements="完整加工过程，不生图", page_range="4"))
        with patch.dict(os.environ, {"DASHSCOPE_API_KEY": "test-key"}), patch("agents.dialogue.get_dialogue_agent", return_value=lambda *args: reply):
            response = self.client.post('/api/dialogue', json={"message": "咖啡四页，不生图", "intent": "discuss", "request_id": "discussion-draft"})
        conversation = response.json["conversation_id"]
        generated = self.client.post('/generate', data={"topic": "咖啡", "requirements": "完整加工过程，不生图", "page_range": "4", "mode": "new", "conversation_id": conversation, "request_id": "draft-generate"})
        self.assertEqual(generated.status_code, 200)
        self.assertEqual(len(self.store.conversations()), 1)
        self.assertEqual(len(PPTStore(self.root).versions(conversation)), 1)
        self.assertEqual(self.store.conversation(conversation)["title"], "咖啡")
        records = PPTStore(self.root).dialogue_records(conversation)
        self.assertEqual(records[0]["content"], "咖啡四页，不生图")
        self.assertEqual(records[1]["response"]["proposal"]["page_range"], "4")

    def test_unknown_dialogue_requires_confirmation_and_preserves_version(self):
        conversation, version = self.seed()
        def unknown(*args): raise StageOutcomeUnknown("未知模型结果")
        with patch.dict(os.environ, {"DASHSCOPE_API_KEY": "test-key"}), patch("agents.dialogue.get_dialogue_agent", return_value=unknown):
            response = self.post_question(conversation, version)
            job_id = response.json["job_id"]
            self.assertEqual(self.store.job(job_id)["status"], "unknown")
            self.assertEqual(self.client.post(f'/api/jobs/{job_id}/retry', json={}).status_code, 409)
        self.assertEqual(self.store.conversation(conversation)["current_version_id"], version)
        self.assertFalse(any(item["role"] == "assistant" for item in self.store.dialogue_records(conversation)))

    def test_invalid_question_and_other_owner_are_rejected_before_task_creation(self):
        self.assertEqual(self.client.post('/api/dialogue', json={"message": "是什么", "intent": "question", "request_id": "missing-ppt-key"}).status_code, 400)
        conversation, version = self.seed()
        with patch.object(web, "_owner", return_value="another-owner"):
            self.assertEqual(self.client.get(f'/api/conversations/{conversation}/dialogue').status_code, 404)
            self.assertEqual(self.client.get(f'/api/versions/{version}/preview').status_code, 404)
            self.assertEqual(self.post_question(conversation, version).status_code, 404)
        self.assertEqual(self.client.post('/api/dialogue', json={"message": "页码", "intent": "question", "conversation_id": conversation, "version_id": version, "pages": [True], "request_id": "invalid-pages-key"}).status_code, 400)

    def test_scoped_state_and_history_cursor_ignore_other_selected_cookie(self):
        conversation, version = self.seed()
        with self.store.connection(write=True) as db:
            for i in range(120):
                db.execute("INSERT INTO messages (conversation_id,role,content,created_at) VALUES (?, 'user', ?, '2026-10-13')", (conversation, f"记录 {i}"))
        with self.client.session_transaction() as session:
            session.pop(web.SESSION_ID_KEY, None)
        self.assertIsNone(self.client.get('/state').json['conversation_id'])
        scoped = self.client.get('/state', query_string={'conversation_id': conversation})
        self.assertEqual(scoped.json['version_id'], version)
        self.assertEqual(self.client.get('/', query_string={'conversation_id': conversation}).status_code, 200)
        recent = self.client.get(f'/api/conversations/{conversation}/dialogue').json
        self.assertEqual(len(recent['messages']), 100); self.assertTrue(recent['has_older'])
        older = self.client.get(f'/api/conversations/{conversation}/dialogue', query_string={'before_id': recent['before_id']}).json
        self.assertEqual(len(older['messages']), 20)
        self.assertLess(older['messages'][-1]['id'], recent['messages'][0]['id'])
        self.assertFalse(older['has_older'])
        with patch.object(web, '_owner', return_value='another-owner'):
            self.assertEqual(self.client.get('/state', query_string={'conversation_id': conversation}).status_code, 404)

    def test_invalid_generation_keeps_user_form_input(self):
        response = self.client.post('/generate', data={"topic": "保留我的主题", "requirements": "保留原始要求", "page_range": "bad", "mode": "new", "filename": "custom"})
        self.assertEqual(response.status_code, 400)
        self.assertIn('value="保留我的主题"', response.get_data(as_text=True))
        self.assertIn('保留原始要求', response.get_data(as_text=True))
        self.assertIn('value="bad"', response.get_data(as_text=True))


class ContextAndEvidenceTests(unittest.TestCase):
    def test_reply_rejects_unverifiable_citations_and_question_proposal(self):
        evidence = page_catalog({"slides": [table_slide()]})
        valid = DialogueReply(answer="87.65%", answerable=True, citations=[{"page": 1, "ref": "table/rows/accuracy/cells/before", "quote": "87.65%"}])
        validate_reply(valid, "question", evidence, 1)
        wrong = valid.model_copy(deep=True); wrong.citations[0].quote = "99.9%"
        with self.assertRaises(ValueError): validate_reply(wrong, "question", evidence, 1)
        wrong = valid.model_copy(deep=True); wrong.proposal = Proposal(action="new", topic="错误操作", requirements="不应生成")
        with self.assertRaises(ValueError): validate_reply(wrong, "question", evidence, 1)
        no_evidence = DialogueReply(answer="模型猜测的价格是100元", answerable=False)
        self.assertNotIn("100", validate_reply(no_evidence, "question", evidence, 1).answer)

    def test_explicit_text_page_selection_cannot_drop_evidence_to_fit_budget(self):
        slides = [dict(table_slide(), full_content=["不可省略的指定页证据" * 800]) for _ in range(2)]
        with patch.dict(os.environ, {"PPT_CONTEXT_INPUT_TOKENS": "24000"}), patch("agents.dialogue.create_client"), patch("agents.dialogue.chat_completion_with_retry") as call:
            with self.assertRaises(ContextBudgetError):
                get_dialogue_agent("test-key")(PPTGenerationState(topic="证据"), {"intent": "question", "message": "对比第1页和第2页", "outline": {"slides": slides}})
        call.assert_not_called()

    def test_legacy_malformed_graph_does_not_invent_or_break_relations(self):
        slide = dict(table_slide(), diagram={"nodes": None, "edges": None})
        self.assertIn("指标对比", page_catalog({"slides": [slide]})[0]["evidence"].values())
        slide["diagram"] = {"nodes": [{"id": "a", "label": "节点"}], "edges": [{"id": "bad", "source": [], "target": "a"}]}
        evidence = page_catalog({"slides": [slide]})[0]["evidence"]
        self.assertNotIn("diagram/edges/bad/relation", evidence)

    def test_latest_explicit_page_count_overrides_model_range(self):
        raw = {"answer": "这是待应用的计划", "proposal": {"action": "new", "topic": "咖啡", "requirements": "制作4页，使用SVG", "page_range": "1-4", "pages": []}}
        completion = SimpleNamespace(choices=[SimpleNamespace(message=SimpleNamespace(content=json.dumps(raw, ensure_ascii=False)), finish_reason="stop")])
        with patch("agents.dialogue.create_client"), patch("agents.dialogue.chat_completion_with_retry", return_value=completion):
            reply = get_dialogue_agent("test-key")(PPTGenerationState(topic="咖啡"), {"intent": "discuss", "message": "做4页，不生图", "outline": None})
        self.assertEqual(reply.proposal.page_range, "4")
        self.assertIn("严禁生图", reply.proposal.requirements)

    def test_compaction_preserves_raw_history_quotes_and_budget(self):
        records = [{"id": i, "role": "user", "content": f"第{i}次要求：必须保留数字，不生图。" + "以前的讨论内容。" * 12} for i in range(40)]
        before = copy.deepcopy(records)
        with patch.dict(os.environ, {"PPT_CONTEXT_INPUT_TOKENS": "3000"}):
            messages, summary = bounded_messages("系统规则", "本次问题", records, "qwen3.8-flash")
        self.assertLessEqual(token_bound(messages), 3000)
        self.assertEqual(summary["strategy"], "extractive")
        self.assertTrue(summary["source_hash"])
        self.assertEqual(before, records)
        self.assertTrue(summary["extracts"])
        for entry in summary["extracts"]:
            self.assertIn("不生图", entry["quote"])
            self.assertIn(entry["quote"], records[entry["source_id"]]["content"])
        with patch.dict(os.environ, {"PPT_CONTEXT_INPUT_TOKENS": "2000"}):
            with self.assertRaises(ContextBudgetError): bounded_messages("规则", "超大资料" * 3000, [], "qwen3.8-flash")

    def test_conflicting_page_request_clarifies_without_model_call(self):
        with patch("agents.dialogue.create_client"), patch("agents.dialogue.chat_completion_with_retry") as call:
            reply = get_dialogue_agent("test-key")(PPTGenerationState(topic="咖啡"), {"intent": "discuss", "message": "做3页还是4页", "outline": None})
        call.assert_not_called()
        self.assertIsNone(reply.proposal)
        self.assertIn("确认最终页数", reply.answer)

    def test_overview_uses_all_page_summaries_and_partial_refusal_is_scoped(self):
        slides = [dict(table_slide(), title=f"主题 {i}", punchline=f"第{i}部分结论") for i in range(1, 13)]
        pages = retrieve_pages({"slides": slides}, "总结整份PPT")
        self.assertEqual(len(pages), 12)
        self.assertTrue(all(set(p["evidence"]) == {"title", "punchline"} for p in pages))
        reply = validate_reply(DialogueReply(answer="猜测"), "question", pages[:5], 12)
        self.assertIn("本次检索", reply.answer)
        self.assertNotIn("当前版本的 PPT 中没有", reply.answer)

    def test_stage_cache_includes_prompt_history(self):
        from config.jobs import StageRunner
        with tempfile.TemporaryDirectory(prefix="ppt_history_cache_") as name:
            store = PPTStore(name); state = PPTGenerationState(topic="缓存")
            run = Path(name) / "run"; run.mkdir()
            request = {"user_input": "缓存", "filename": "test.pptx"}
            job, _ = store.create_job(request, state.snapshot(), run, "history-cache-key", content_hash(request))
            store.claim_job(job["id"], "worker")
            runner = StageRunner(store, job, "worker", state)
            calls = []
            def produce():
                calls.append(1); return {"result": len(calls)}
            state.set("dialogue_history", [{"role": "user", "content": "蓝色"}])
            runner("writer/probe", {}, produce, "qwen3.8-flash", None, True)
            runner("writer/probe", {}, produce, "qwen3.8-flash", None, True)
            state.set("dialogue_history", [{"role": "user", "content": "改成棕色"}])
            runner("writer/probe", {}, produce, "qwen3.8-flash", None, True)
            self.assertEqual(len(calls), 2)

    def test_writer_context_excludes_old_blueprint_and_agent_json(self):
        state = PPTGenerationState(topic="新主题", requirements="保留单位")
        state.set("blueprint", {"marker_old_geometry": "旧几何" * 30000})
        state.messages = [{"role": "assistant", "content": '{"huge_agent_json":' + '0' * 20000 + '}'}, {"role": "user", "content": "不生图"}]
        joined = json.dumps(state.build_messages("Writer规则", "本次单页内容"), ensure_ascii=False)
        self.assertNotIn("marker_old_geometry", joined)
        self.assertNotIn("huge_agent_json", joined)
        self.assertIn("不生图", joined)
        self.assertTrue(state.get("context_compaction")["writer"]["input_upper_bound"])

    def test_svg_text_is_evidence_and_window_size_respects_small_desktop(self):
        slide = table_slide(); slide["svg_code"] = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 50"><text>种植到采摘</text></svg>'
        self.assertIn("种植到采摘", page_catalog({"slides": [slide]})[0]["evidence"].values())
        graph = page_catalog({"slides": [diagram_slide()]})[0]["evidence"]
        self.assertEqual(graph["diagram/edges/first/relation"], "用户需求 → 内容规划")
        import importlib, sys
        fake = SimpleNamespace(screens=[SimpleNamespace(width=1366, height=768)])
        with patch.dict(sys.modules, {"webview": fake}):
            launcher = importlib.import_module("launcher")
        with patch.object(launcher, "webview", fake):
            self.assertEqual(launcher._window_size(), (1302, 688))


if __name__ == "__main__":
    unittest.main()
