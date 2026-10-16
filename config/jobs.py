import json
import logging
import os
import threading
from functools import lru_cache
from pathlib import Path
from uuid import uuid4

from huey import SqliteHuey, crontab
from openai import APIConnectionError, APITimeoutError
from requests.exceptions import ConnectionError as RequestsConnectionError, Timeout

from config.llm_client import MODEL_CALL_OBSERVER, MODEL_REQUEST_GUARD, resolve_model_name
from config.state import PPTGenerationState
from config.storage import PPTStore, StorageConflict, content_hash, json_text


logger = logging.getLogger(__name__)


class StageOutcomeUnknown(RuntimeError):
    pass


def atomic_json(path, payload):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(f".{path.name}.{uuid4().hex}.tmp")
    try:
        with temporary.open("w", encoding="utf-8") as stream:
            stream.write(json_text(payload))
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temporary, path)
    finally:
        temporary.unlink(missing_ok=True)


class StageRunner:
    def __init__(self, store, job, worker_token, state):
        self.store, self.job, self.token, self.state = store, job, worker_token, state

    def _context(self):
        snapshot = self.state.snapshot()
        return {key: snapshot[key] for key in (
            "topic", "requirements", "messages", "feedback", "operation_mode", "modify_scope",
            "modify_page", "modify_pages", "requirement_history"
        )} | {"outline": self.state.get("outline"), "blueprint": self.state.get("blueprint"),
              "page_policy": self.state.get("page_policy"), "dialogue_history": self.state.get("dialogue_history")}

    def _fingerprint(self, payload):
        payload = self.store.rebase_snapshot(payload)
        files = []
        def visit(value):
            if isinstance(value, dict):
                for key, item in value.items():
                    if key in {"image_path", "filename"} and isinstance(item, str):
                        candidate = Path(item)
                        if candidate.is_file() and candidate.resolve().is_relative_to(self.store.root):
                            asset = self.store.describe_asset(candidate, "input")
                            files.append({"path": asset["path"], "sha256": asset["sha256"]})
                        try:
                            value[key] = candidate.resolve().relative_to(self.store.root).as_posix()
                        except ValueError:
                            pass
                    else:
                        visit(item)
            elif isinstance(value, list):
                for item in value:
                    visit(item)
        visit(payload)
        return content_hash({"payload": payload, "files": sorted(files, key=lambda item: item["path"])})

    def __call__(self, key, inputs, operation, model, result_model, external):
        source = Path(__file__).resolve().parents[1]
        role = key.split("/", 1)[0]
        names = {"planner": "planner", "writer": "writer", "layout": "template", "image": "image_generator", "dialogue": "dialogue"}
        prompt_files = [source / "config" / name for name in ("schema.py", "sanitizers.py", "content.py")]
        if "/candidate-" in key or "/design-" in key:
            prompt_files.extend([source / "config" / "design.py", source / "config" / "icon_library.py"])
        if role in {"planner", "writer", "dialogue"}:
            prompt_files.append(source / "config" / "context.py")
        if role in names:
            prompt_files.append(source / "agents" / f"{names[role]}.py")
            prompt_files.append(source / 'config/llm_client.py')
        if role in {"layout", "render", "vision"}:
            prompt_files.append(source / "config/layout_budget.py")
        if key.endswith("/fallback"):
            prompt_files.extend([source / "config/layout_fallback.py", source / "renderer.py", source / "main.py"])
        elif role in {"render", "demo"}:
            prompt_files.extend([source / "main.py", source / "renderer.py"])
        elif role == "preview":
            prompt_files.append(source / "config" / "preview.py")
        elif role == "vision":
            prompt_files.extend([source / "agents" / "visual_review.py", source / "config" / "visual.py"])
        from hashlib import sha256
        source_bytes = b"".join(path.read_bytes() for path in prompt_files)
        if "/candidate-" in key or "/design-" in key:
            source_bytes += b"".join(path.name.encode()+b"\0"+path.read_bytes() for path in sorted((source / "assets/icons").glob("*.svg")))
        prompt_version = sha256(source_bytes).hexdigest()
        fingerprint = self._fingerprint({"inputs": inputs, "context": self._context(),
                                         "model": model, "prompt_version": prompt_version,
                                         'thinking_mode':os.environ.get(f'PPT_{role.upper()}_THINKING','auto'),
                                         'temperature':os.environ.get(f'PPT_{role.upper()}_TEMPERATURE','auto'),
                                         'enable_search':os.environ.get('PPT_ENABLE_SEARCH','1'),
                                         'thinking_budget':os.environ.get('PPT_THINKING_BUDGET','2048')})
        cached = self.store.cached_stage(self.job["id"], key, fingerprint)
        if cached:
            files = json.loads(cached["files_json"])
            if not self.store.verify_assets(files):
                self.store.fail_stage(cached["id"], self.job["id"], self.token,
                                      "unknown" if external else "invalidated", "CachedAssetChanged")
                if external:
                    raise StageOutcomeUnknown("已完成阶段的文件丢失或被修改，确认后才可重新调用模型")
            else:
                self.store.heartbeat(self.job["id"], self.token)
                self.state.restore_snapshot(self.store.rebase_snapshot(json.loads(cached["snapshot_json"])))
                self.state.append_progress(f"Resume: 复用已完成阶段 {key}")
                value = json.loads(cached["value_json"])
                if result_model:
                    return result_model.model_validate(self.store.rebase_snapshot(value))
                if files and isinstance(value, str):
                    return str(self.store.asset_path(files[0]["path"]))
                return value
        stage_id = self.store.start_stage(self.job["id"], self.token, key, fingerprint, model, external, prompt_version)
        guard_token = MODEL_REQUEST_GUARD.set(lambda: self.store.heartbeat(self.job["id"], self.token))
        observer_token = MODEL_CALL_OBSERVER.set(lambda actual_model, usage, request_id: self.store.record_stage_usage(
            stage_id, self.job["id"], self.token, actual_model, usage, request_id
        ))
        try:
            result = operation()
            if result is None:
                raise StageOutcomeUnknown("外部阶段未返回可用结果") if external else RuntimeError("阶段未返回结果")
            value = result.model_dump() if hasattr(result, "model_dump") else result
            files = []
            if isinstance(result, str) and Path(result).is_file():
                files = [self.store.describe_asset(result, "stage-output")]
            elif role == "preview" and isinstance(result, dict):
                files = [self.store.describe_asset(result["pdf"], "pdf")]
                files.extend(self.store.describe_asset(path, "preview") for path in result["pages"])
            elif role in {"image", "render", "demo"}:
                raise StageOutcomeUnknown("外部阶段的文件未落盘") if external else ValueError("本地阶段的文件未落盘")
            self.store.complete_stage(stage_id, self.job["id"], self.token, value, self.state.snapshot(), files)
            return result
        except Exception as exc:
            request_id = getattr(exc, "request_id", None)
            if request_id:
                self.store.record_stage_usage(stage_id, self.job["id"], self.token, model, None, request_id)
            unknown = external and isinstance(exc, (
                StageOutcomeUnknown, APIConnectionError, APITimeoutError, Timeout, RequestsConnectionError
            ))
            self.store.fail_stage(stage_id, self.job["id"], self.token, "unknown" if unknown else "failed", type(exc).__name__)
            raise
        finally:
            MODEL_CALL_OBSERVER.reset(observer_token)
            MODEL_REQUEST_GUARD.reset(guard_token)


def generation_manifest(store, state, filename, run_dir):
    from pptx import Presentation
    if not Path(filename).is_file() or not Presentation(filename).slides:
        raise ValueError("未生成有效的 PPTX 文件")
    manifest = [store.describe_asset(filename, "pptx")]
    assets = {}
    snapshot_dir = run_dir / "generated" / "snapshots" / Path(filename).stem
    for role in ("outline", "blueprint"):
        value = state.get(role)
        if value:
            path = snapshot_dir / f"{role}.json"
            atomic_json(path, value)
            manifest.append(store.describe_asset(path, role))
            assets[role] = str(path)
    image_paths = set()
    for slide in (state.get("outline") or {}).get("slides", []):
        if slide.get("image_path"):
            image_paths.add(slide["image_path"])
    for canvas in (state.get("blueprint") or {}).get("canvases", []):
        for element in canvas.get("elements", []):
            if element.get("image_path"):
                image_paths.add(element["image_path"])
    for path in sorted(image_paths):
        manifest.append(store.describe_asset(path, "image"))
    preview = state.get("preview_assets") or {}
    if preview.get("pdf"):
        manifest.append(store.describe_asset(preview["pdf"], "pdf"))
        manifest.extend(store.describe_asset(path, "preview") for path in preview.get("pages", []))
    if state.get("visual_review_path"):
        manifest.append(store.describe_asset(state.get("visual_review_path"), "visual-review"))
    for path in state.get("visual_candidates") or []:
        if Path(path) != Path(filename):
            manifest.append(store.describe_asset(path, "candidate-pptx"))
    return manifest, assets


def execute_job(root, job_id):
    from config_manager import load_api_key, load_base_url, RenderDependencyError
    store = PPTStore(root)
    job = store.job(job_id, owner=None)
    if not job:
        return
    token = uuid4().hex
    job = store.claim_job(job_id, token, owner=job["owner_id"])
    if not job:
        return
    state = PPTGenerationState.from_snapshot(store.rebase_snapshot(json.loads(job["initial_json"])))
    run_dir = store.asset_path(job["run_dir"])
    run_dir.mkdir(parents=True, exist_ok=True)
    state.data["_run_dir"] = str(run_dir)
    state.bind_persistence(lambda updated: store.save_job_state(job_id, token, updated.snapshot()), strict=True)
    state.stage_executor = StageRunner(store, job, token, state)
    stopped = threading.Event()

    def heartbeat():
        while not stopped.wait(20):
            try:
                store.heartbeat(job_id, token)
            except Exception:
                stopped.set()

    heartbeat_thread = threading.Thread(target=heartbeat, daemon=True)
    heartbeat_thread.start()
    original_environment = {key: os.environ.get(key) for key in ("DASHSCOPE_API_KEY", "DASHSCOPE_BASE_URL")}
    try:
        state.append_progress("Job: worker 已领取任务")
        if job["base_version_id"]:
            base = store.version(job["base_version_id"], job["owner_id"])
            if not base or not store.verify_assets(json.loads(base["manifest_json"])):
                raise ValueError("基准版本资产缺失或被修改，请恢复原文件后再试")
        if not os.environ.get("DASHSCOPE_API_KEY"):
            saved_key = load_api_key()
            if saved_key:
                os.environ["DASHSCOPE_API_KEY"] = saved_key
        if not os.environ.get("DASHSCOPE_BASE_URL"):
            saved_url = load_base_url()
            if saved_url:
                os.environ["DASHSCOPE_BASE_URL"] = saved_url
        payload = json.loads(job["request_json"])
        if payload.get("kind") == "design_candidates":
            from config.design import build_candidates
            build_candidates(state, store, job, payload, run_dir, os.environ.get("DASHSCOPE_API_KEY", ""))
            atomic_json(run_dir / "state.json", state.snapshot())
            store.finish_design_candidates(job_id, token, state.snapshot())
            return
        if payload.get("kind") == "design_adopt":
            from config.design import prepare_adoption
            output = prepare_adoption(state, store, job, payload)
            manifest, assets = generation_manifest(store, state, output, run_dir)
            state.set("generation_status", "succeeded")
            state.add_history_snapshot(Path(output).name, outline_path=assets.get("outline"), blueprint_path=assets.get("blueprint"))
            store.finish_job(job_id, token, state.snapshot(), manifest)
            return
        if payload.get("kind") == "dialogue":
            from agents.dialogue import get_dialogue_agent, DialogueReply
            version = store.version(job["base_version_id"], job["owner_id"]) if job["base_version_id"] else None
            committed = json.loads(version["snapshot_json"]).get("data", {}) if version else {}
            history = [item for item in store.dialogue_records(job["conversation_id"], job["owner_id"])
                       if item["job_id"] != job_id and (payload.get("history_through_id") is None or item["id"] <= payload["history_through_id"])]
            context = {"intent": payload["intent"], "message": payload["user_input"], "pages": payload.get("page_numbers"),
                       "outline": committed.get("outline"), "requirements": committed.get("requirements", ""),
                       "version_id": job["base_version_id"], "history": history}
            state.append_progress("Dialogue: 正在读取页面证据与会话记录")
            if os.environ.get("DASHSCOPE_API_KEY"):
                agent = get_dialogue_agent(os.environ["DASHSCOPE_API_KEY"])
                reply = state.execute_stage("dialogue/reply", context, lambda: agent(state, context),
                                            model=resolve_model_name("dialogue"), result_model=DialogueReply)
            else:
                reply = DialogueReply(answer="请先在设置中连接模型服务，再讨论需求或提问 PPT 内容。")
            response = {**reply.model_dump(), "version_id": job["base_version_id"], "intent": payload["intent"],
                        "memory": state.get("context_compaction", {}).get("dialogue", {}), "retrieval": state.get("dialogue_retrieval", {})}
            state.set("dialogue_response", response)
            atomic_json(run_dir / "dialogue.json", response)
            atomic_json(run_dir / "state.json", state.snapshot())
            store.finish_dialogue(job_id, token, state.snapshot(), response)
            return
        from main import run_incremental_modify_pages, run_ppt_system
        state.set("dialogue_history", [item for item in store.dialogue_records(job["conversation_id"], job["owner_id"])
                                      if payload.get("history_through_id") is None or item["id"] <= payload["history_through_id"]])
        filename = str(run_dir / payload["filename"])
        kwargs = {"filename": filename, "state": state, "requirements": payload["requirements"], "run_dir": str(run_dir)}
        if payload["use_incremental"]:
            output = run_incremental_modify_pages(payload["user_input"], page_numbers=payload["page_numbers"], **kwargs)
        else:
            output = run_ppt_system(payload["user_input"], **kwargs)
        if not output or state.get("generation_status") == "failed":
            raise RuntimeError("生成管线未返回成功结果")
        from config.visual import run_visual_pipeline
        used_calls = sum(len(json.loads(stage["usage_json"] or "[]")) or int(stage["status"] in {"running", "unknown", "succeeded"})
                         for stage in store.stages(job_id) if stage["stage_key"].startswith("vision/"))
        output = run_visual_pipeline(state, output, run_dir, os.environ.get("DASHSCOPE_API_KEY", ""),
                                     page_numbers=payload["page_numbers"] if payload["use_incremental"] else None,
                                     used_calls=used_calls)
        manifest, assets = generation_manifest(store, state, output, run_dir)
        committed = PPTGenerationState.from_snapshot(state.snapshot())
        committed.set("last_file", str(output))
        committed.set("generation_status", "succeeded")
        if committed.get("result_kind") != "offline_demo":
            committed.add_history_snapshot(Path(output).name, outline_path=assets.get("outline"), blueprint_path=assets.get("blueprint"))
        store.finish_job(job_id, token, committed.snapshot(), manifest)
    except Exception as exc:
        stages = store.stages(job_id)
        status = "unknown" if isinstance(exc, StageOutcomeUnknown) or any(stage["status"] == "unknown" for stage in stages) else (
            "conflict" if isinstance(exc, StorageConflict) else "failed"
        )
        failed = PPTGenerationState.from_snapshot(state.snapshot())
        failed.set("generation_status", status)
        logger.error("Job %s failed (%s); status=%s", job_id, type(exc).__name__, status)
        failed.append_progress("Job: 本轮任务未提交新版本；" + (
            "外部调用结果或缓存资产需要确认" if status == "unknown" else
            (str(exc) if isinstance(exc, RenderDependencyError) else "详情见本次运行日志")
        ))
        try:
            store.fail_job(job_id, token, failed.snapshot(), type(exc).__name__, status=status)
        except StorageConflict:
            logger.warning("Job %s lease expired; result was not committed", job_id)
        log_dir = run_dir / "generated"
        log_dir.mkdir(parents=True, exist_ok=True)
        atomic_json(log_dir / "job_error.json", {"error_type": type(exc).__name__, "status": status})
    finally:
        stopped.set()
        heartbeat_thread.join(timeout=2)
        for key, value in original_environment.items():
            if value is None:
                os.environ.pop(key, None)
            else:
                os.environ[key] = value
    current = store.job(job_id, job["owner_id"])
    if current:
        atomic_json(run_dir / "state.json", json.loads(current["snapshot_json"]))


@lru_cache(maxsize=8)
def task_queue(root):
    root = str(Path(root).resolve())
    queue = SqliteHuey("ppt-generation", filename=str(Path(root) / "queue.sqlite3"),
                       results=False, fsync=True, immediate=False)

    @queue.task(name="ppt.generate", retries=0)
    def generate(job_id):
        execute_job(root, job_id)

    @queue.periodic_task(crontab(minute="*"), name="ppt.recover", retries=0)
    def recover():
        dispatch_pending(root, recover=True)

    return queue, generate


def dispatch_pending(root, recover=False):
    store = PPTStore(root)
    if recover:
        store.recover_expired()
    queue, generate = task_queue(str(store.root))
    for job in store.pending_dispatch():
        generate(job["id"])
        store.mark_dispatched(job["id"])
    return queue
