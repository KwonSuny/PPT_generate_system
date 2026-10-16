import hashlib
import json
import sqlite3
import time
from contextlib import contextmanager
from datetime import datetime, timezone
from pathlib import Path
from uuid import uuid4

from config.state import PPTGenerationState


LOCAL_OWNER = "local"


class StorageConflict(ValueError):
    pass


def json_text(value):
    return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"), allow_nan=False)


def content_hash(value):
    return hashlib.sha256(json_text(value).encode("utf-8")).hexdigest()


def file_hash(path):
    digest = hashlib.sha256()
    with open(path, "rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def now_text():
    return datetime.now(timezone.utc).isoformat(timespec="milliseconds")


class PPTStore:
    def __init__(self, root):
        self.root = Path(root).resolve()
        self.root.mkdir(parents=True, exist_ok=True)
        self.path = self.root / "app.sqlite3"
        with self.connection() as db:
            db.execute("PRAGMA journal_mode=WAL")
            try:
                db.execute("SELECT json_extract('{}', '$.kind')")
            except sqlite3.OperationalError as exc:
                raise ValueError("SQLite 运行时缺少 JSON 函数，请使用支持 JSON 的 Python/SQLite 版本") from exc
            version = db.execute("PRAGMA user_version").fetchone()[0]
            if version not in (0, 1, 2):
                raise ValueError("数据库版本不受当前程序支持")
            db.executescript("""
                CREATE TABLE IF NOT EXISTS metadata (key TEXT PRIMARY KEY, value TEXT NOT NULL);
                CREATE TABLE IF NOT EXISTS conversations (
                    id TEXT PRIMARY KEY, owner_id TEXT NOT NULL, title TEXT NOT NULL,
                    created_at TEXT NOT NULL, updated_at TEXT NOT NULL
                );
                CREATE TABLE IF NOT EXISTS decks (
                    id TEXT PRIMARY KEY, conversation_id TEXT NOT NULL UNIQUE
                        REFERENCES conversations(id), current_version_id TEXT
                );
                CREATE TABLE IF NOT EXISTS versions (
                    id TEXT PRIMARY KEY, deck_id TEXT NOT NULL REFERENCES decks(id),
                    parent_id TEXT REFERENCES versions(id), created_at TEXT NOT NULL,
                    snapshot_json TEXT NOT NULL, manifest_json TEXT NOT NULL
                );
                CREATE TABLE IF NOT EXISTS jobs (
                    id TEXT PRIMARY KEY, owner_id TEXT NOT NULL,
                    conversation_id TEXT NOT NULL REFERENCES conversations(id),
                    deck_id TEXT NOT NULL REFERENCES decks(id),
                    base_version_id TEXT REFERENCES versions(id),
                    idempotency_key TEXT NOT NULL, request_hash TEXT NOT NULL,
                    request_json TEXT NOT NULL, initial_json TEXT NOT NULL,
                    snapshot_json TEXT NOT NULL, run_dir TEXT NOT NULL,
                    status TEXT NOT NULL, worker_token TEXT, lease_until REAL,
                    dispatched_at REAL, result_version_id TEXT REFERENCES versions(id),
                    error TEXT, created_at TEXT NOT NULL, updated_at TEXT NOT NULL,
                    UNIQUE(owner_id, idempotency_key)
                );
                CREATE INDEX IF NOT EXISTS jobs_conversation ON jobs(conversation_id, created_at);
                CREATE TABLE IF NOT EXISTS stages (
                    id INTEGER PRIMARY KEY AUTOINCREMENT, job_id TEXT NOT NULL REFERENCES jobs(id),
                    stage_key TEXT NOT NULL, input_hash TEXT NOT NULL, model TEXT,
                    prompt_version TEXT,
                    status TEXT NOT NULL, external INTEGER NOT NULL,
                    value_json TEXT, snapshot_json TEXT, files_json TEXT,
                    usage_json TEXT, request_id TEXT, error TEXT,
                    created_at TEXT NOT NULL, updated_at TEXT NOT NULL
                );
                CREATE INDEX IF NOT EXISTS stages_lookup ON stages(job_id, stage_key, input_hash, id);
                CREATE TABLE IF NOT EXISTS artifacts (
                    id TEXT PRIMARY KEY, version_id TEXT NOT NULL REFERENCES versions(id),
                    role TEXT NOT NULL, path TEXT NOT NULL, sha256 TEXT NOT NULL, size INTEGER NOT NULL
                );
                CREATE TABLE IF NOT EXISTS messages (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    conversation_id TEXT NOT NULL REFERENCES conversations(id),
                    role TEXT NOT NULL, content TEXT NOT NULL, job_id TEXT REFERENCES jobs(id),
                    created_at TEXT NOT NULL
                );
                CREATE INDEX IF NOT EXISTS messages_conversation ON messages(conversation_id, id);
                CREATE TABLE IF NOT EXISTS legacy_imports (
                    run_id TEXT PRIMARY KEY, source_hash TEXT NOT NULL,
                    conversation_id TEXT NOT NULL REFERENCES conversations(id)
                );
            """)
            db.execute("BEGIN IMMEDIATE")
            if "prompt_version" not in {row[1] for row in db.execute("PRAGMA table_info(stages)")}:
                db.execute("ALTER TABLE stages ADD COLUMN prompt_version TEXT")
            db.execute("PRAGMA user_version=2")
            row = db.execute("SELECT value FROM metadata WHERE key='roots'").fetchone()
            self.source_roots = json.loads(row[0]) if row else []
            if str(self.root) not in self.source_roots:
                self.source_roots.append(str(self.root))
                db.execute("INSERT OR REPLACE INTO metadata VALUES ('roots', ?)", (json_text(self.source_roots),))

    @contextmanager
    def connection(self, write=False):
        db = sqlite3.connect(self.path, timeout=10)
        db.row_factory = sqlite3.Row
        db.execute("PRAGMA foreign_keys=ON")
        db.execute("PRAGMA busy_timeout=10000")
        try:
            if write:
                db.execute("BEGIN IMMEDIATE")
            yield db
            db.commit()
        except BaseException:
            db.rollback()
            raise
        finally:
            db.close()

    def asset_path(self, relative_path):
        path = (self.root / relative_path).resolve()
        if not path.is_relative_to(self.root) or path == self.root:
            raise ValueError("资产路径超出产物根目录")
        return path

    def describe_asset(self, path, role):
        path = Path(path).resolve()
        if not path.is_relative_to(self.root) or not path.is_file():
            raise ValueError("资产文件缺失或超出产物根目录")
        return {"id": str(uuid4()), "role": role, "path": path.relative_to(self.root).as_posix(),
                "sha256": file_hash(path), "size": path.stat().st_size}

    def rebase_path(self, value):
        if not isinstance(value, str):
            return value
        normalized = value.replace("\\", "/")
        for old_root in sorted(self.source_roots, key=len, reverse=True):
            prefix = old_root.replace("\\", "/").rstrip("/")
            candidate, comparison = normalized, prefix
            if len(prefix) > 2 and prefix[1:3] == ":/":
                candidate, comparison = candidate.casefold(), comparison.casefold()
            if candidate.startswith(comparison + "/"):
                return str(self.asset_path(normalized[len(prefix) + 1:]))
        return value

    def rebase_snapshot(self, snapshot):
        snapshot = json.loads(json_text(snapshot))
        def visit(value):
            if isinstance(value, dict):
                for key, item in value.items():
                    if key in {"image_path", "last_file", "_run_dir", "outline_path", "blueprint_path"}:
                        value[key] = self.rebase_path(item)
                    else:
                        visit(item)
            elif isinstance(value, list):
                for item in value:
                    visit(item)
        visit(snapshot)
        return snapshot

    def verify_assets(self, manifest):
        try:
            return all(self.asset_path(item["path"]).is_file()
                       and file_hash(self.asset_path(item["path"])) == item["sha256"]
                       and self.asset_path(item["path"]).stat().st_size == item["size"] for item in manifest)
        except (OSError, ValueError, KeyError, TypeError):
            return False

    def conversation(self, conversation_id, owner=LOCAL_OWNER):
        with self.connection() as db:
            row = db.execute("""SELECT c.*, d.id AS deck_id, d.current_version_id
                FROM conversations c JOIN decks d ON d.conversation_id=c.id
                WHERE c.id=? AND c.owner_id=?""", (conversation_id, owner)).fetchone()
            return dict(row) if row else None

    def conversations(self, owner=LOCAL_OWNER):
        with self.connection() as db:
            return [dict(row) for row in db.execute("""SELECT c.*, d.current_version_id
                FROM conversations c JOIN decks d ON d.conversation_id=c.id
                WHERE c.owner_id=? ORDER BY c.updated_at DESC, c.id""", (owner,))]

    def versions(self, conversation_id, owner=LOCAL_OWNER):
        with self.connection() as db:
            rows = db.execute("""SELECT v.id, v.parent_id, v.created_at, v.snapshot_json
                FROM versions v JOIN decks d ON d.id=v.deck_id
                JOIN conversations c ON c.id=d.conversation_id
                WHERE c.id=? AND c.owner_id=? ORDER BY v.created_at DESC, v.rowid DESC""", (conversation_id, owner))
            result = []
            for row in rows:
                snapshot = json.loads(row["snapshot_json"])
                result.append({"id": row["id"], "parent_id": row["parent_id"], "created_at": row["created_at"],
                               "mode": snapshot.get("operation_mode"),
                               "file_name": Path(snapshot.get("data", {}).get("last_file", "")).name})
            return result

    def version(self, version_id, owner=LOCAL_OWNER):
        with self.connection() as db:
            row = db.execute("""SELECT v.*, c.id AS conversation_id FROM versions v
                JOIN decks d ON d.id=v.deck_id JOIN conversations c ON c.id=d.conversation_id
                WHERE v.id=? AND c.owner_id=?""", (version_id, owner)).fetchone()
            return dict(row) if row else None

    def job(self, job_id, owner=LOCAL_OWNER):
        with self.connection() as db:
            if owner is None:
                # 仅独立 worker 通过内部任务 ID 使用；Web 必须传入所有者。
                row = db.execute("SELECT * FROM jobs WHERE id=?", (job_id,)).fetchone()
            else:
                row = db.execute("SELECT * FROM jobs WHERE id=? AND owner_id=?", (job_id, owner)).fetchone()
            return dict(row) if row else None

    def job_by_key(self, key, owner=LOCAL_OWNER):
        with self.connection() as db:
            row = db.execute("SELECT * FROM jobs WHERE owner_id=? AND idempotency_key=?", (owner, key)).fetchone()
            return dict(row) if row else None

    def latest_job(self, conversation_id, owner=LOCAL_OWNER):
        with self.connection() as db:
            row = db.execute("""SELECT * FROM jobs WHERE conversation_id=? AND owner_id=? AND COALESCE(json_extract(request_json, '$.kind'), 'generation') NOT IN ('dialogue','design_candidates')
                ORDER BY created_at DESC, rowid DESC LIMIT 1""", (conversation_id, owner)).fetchone()
            return dict(row) if row else None

    def load_state(self, conversation_id, owner=LOCAL_OWNER):
        conversation = self.conversation(conversation_id, owner)
        if not conversation:
            return PPTGenerationState(topic="")
        job = self.latest_job(conversation_id, owner)
        if job:
            state = PPTGenerationState.from_snapshot(self.rebase_snapshot(json.loads(job["snapshot_json"])))
            state.data["generation_status"] = job["status"]
            state.data["_job_id"] = job["id"]
        elif conversation["current_version_id"]:
            state = PPTGenerationState.from_snapshot(self.rebase_snapshot(json.loads(self.version(
                conversation["current_version_id"], owner
            )["snapshot_json"])))
        else:
            state = PPTGenerationState(topic=conversation["title"])
        state.data.update({"_conversation_id": conversation_id, "_deck_id": conversation["deck_id"],
                           "_version_id": conversation["current_version_id"]})
        if conversation["current_version_id"]:
            committed = self.rebase_snapshot(json.loads(self.version(conversation["current_version_id"], owner)["snapshot_json"]))
            state.data["last_file"] = committed.get("data", {}).get("last_file", "")
        return state

    def create_job(self, request_data, snapshot, run_dir, idempotency_key,
                   request_hash, conversation_id=None, base_version_id=None, owner=LOCAL_OWNER):
        job_id, timestamp = str(uuid4()), now_text()
        relative_run = Path(run_dir).resolve().relative_to(self.root).as_posix()
        with self.connection(write=True) as db:
            existing = db.execute("SELECT * FROM jobs WHERE owner_id=? AND idempotency_key=?",
                                  (owner, idempotency_key)).fetchone()
            if existing:
                if existing["request_hash"] != request_hash:
                    raise StorageConflict("同一请求标识不能用于不同输入")
                return dict(existing), False
            if conversation_id is None:
                conversation_id, deck_id = str(uuid4()), str(uuid4())
                db.execute("INSERT INTO conversations VALUES (?, ?, ?, ?, ?)",
                           (conversation_id, owner, snapshot["topic"], timestamp, timestamp))
                db.execute("INSERT INTO decks VALUES (?, ?, NULL)", (deck_id, conversation_id))
            else:
                conversation = db.execute("""SELECT d.* FROM decks d JOIN conversations c ON c.id=d.conversation_id
                    WHERE c.id=? AND c.owner_id=?""", (conversation_id, owner)).fetchone()
                if not conversation:
                    raise StorageConflict("会话不存在")
                deck_id = conversation["id"]
                dialogue = request_data.get("kind") in {"dialogue", "design_candidates"}
                if dialogue:
                    if base_version_id and not db.execute("SELECT 1 FROM versions WHERE id=? AND deck_id=?", (base_version_id, deck_id)).fetchone():
                        raise StorageConflict("问答版本不属于当前文稿")
                elif conversation["current_version_id"] != base_version_id:
                    raise StorageConflict("当前版本已改变，请刷新后重新提交")
                active_kind = request_data["kind"] if dialogue else "generation"
                if db.execute("SELECT 1 FROM jobs WHERE deck_id=? AND status IN ('queued', 'running') AND CASE WHEN json_extract(request_json, '$.kind') IN ('dialogue','design_candidates') THEN json_extract(request_json, '$.kind') ELSE 'generation' END=?", (deck_id, active_kind)).fetchone():
                    raise StorageConflict("当前会话已有同类待执行或正在运行的任务")
            if request_data.get("kind") not in {"dialogue", "design_candidates"} and base_version_id is None:
                db.execute("UPDATE conversations SET title=? WHERE id=? AND owner_id=?", (snapshot["topic"], conversation_id, owner))
            snapshot = json.loads(json_text(snapshot))
            snapshot["data"].update({"_job_id": job_id, "_conversation_id": conversation_id,
                                     "_deck_id": deck_id, "_version_id": base_version_id,
                                     "_run_dir": str(self.asset_path(relative_run)), "generation_status": "queued"})
            db.execute("""INSERT INTO jobs (id, owner_id, conversation_id, deck_id, base_version_id,
                idempotency_key, request_hash, request_json, initial_json, snapshot_json, run_dir,
                status, created_at, updated_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'queued', ?, ?)""",
                (job_id, owner, conversation_id, deck_id, base_version_id, idempotency_key, request_hash,
                 json_text(request_data), json_text(snapshot), json_text(snapshot), relative_run, timestamp, timestamp))
            db.execute("INSERT INTO messages (conversation_id, role, content, job_id, created_at) VALUES (?, 'user', ?, ?, ?)",
                       (conversation_id, request_data["user_input"], job_id, timestamp))
            row = db.execute("SELECT * FROM jobs WHERE id=?", (job_id,)).fetchone()
            return dict(row), True

    def pending_dispatch(self, stale_seconds=60):
        with self.connection() as db:
            return [dict(row) for row in db.execute("""SELECT * FROM jobs WHERE status='queued'
                AND (dispatched_at IS NULL OR dispatched_at<?) ORDER BY created_at""", (time.time() - stale_seconds,))]

    def mark_dispatched(self, job_id):
        with self.connection(write=True) as db:
            db.execute("UPDATE jobs SET dispatched_at=? WHERE id=? AND status='queued'", (time.time(), job_id))

    def claim_job(self, job_id, worker_token, lease_seconds=120, owner=LOCAL_OWNER):
        with self.connection(write=True) as db:
            row = db.execute("SELECT * FROM jobs WHERE id=? AND owner_id=?", (job_id, owner)).fetchone()
            if not row or row["status"] != "queued":
                return None
            current = db.execute("SELECT current_version_id FROM decks WHERE id=?", (row["deck_id"],)).fetchone()[0]
            if current != row["base_version_id"] and json.loads(row["request_json"]).get("kind") not in {"dialogue", "design_candidates"}:
                db.execute("UPDATE jobs SET status='conflict', error=?, updated_at=? WHERE id=?",
                           ("基准版本已改变，未开始模型调用", now_text(), job_id))
                return None
            if db.execute("SELECT 1 FROM jobs WHERE deck_id=? AND status='running' AND id<>?",
                          (row["deck_id"], job_id)).fetchone():
                return None
            db.execute("""UPDATE jobs SET status='running', worker_token=?, lease_until=?, updated_at=?
                WHERE id=?""", (worker_token, time.time() + lease_seconds, now_text(), job_id))
            return dict(db.execute("SELECT * FROM jobs WHERE id=?", (job_id,)).fetchone())

    def _assert_lease(self, db, job_id, worker_token):
        row = db.execute("SELECT * FROM jobs WHERE id=? AND status='running' AND worker_token=? AND lease_until>?",
                         (job_id, worker_token, time.time())).fetchone()
        if not row:
            raise StorageConflict("任务租约已失效，停止写入")
        return row

    def heartbeat(self, job_id, worker_token, lease_seconds=120):
        with self.connection(write=True) as db:
            self._assert_lease(db, job_id, worker_token)
            db.execute("UPDATE jobs SET lease_until=? WHERE id=?", (time.time() + lease_seconds, job_id))

    def save_job_state(self, job_id, worker_token, snapshot):
        with self.connection(write=True) as db:
            self._assert_lease(db, job_id, worker_token)
            db.execute("UPDATE jobs SET snapshot_json=?, updated_at=? WHERE id=?",
                       (json_text(snapshot), now_text(), job_id))

    def start_stage(self, job_id, worker_token, stage_key, input_hash, model, external, prompt_version=None):
        with self.connection(write=True) as db:
            self._assert_lease(db, job_id, worker_token)
            timestamp = now_text()
            cursor = db.execute("""INSERT INTO stages (job_id, stage_key, input_hash, model, prompt_version, status,
                external, created_at, updated_at) VALUES (?, ?, ?, ?, ?, 'running', ?, ?, ?)""",
                (job_id, stage_key, input_hash, model, prompt_version, int(external), timestamp, timestamp))
            return cursor.lastrowid

    def cached_stage(self, job_id, stage_key, input_hash):
        with self.connection() as db:
            row = db.execute("""SELECT * FROM stages WHERE job_id=? AND stage_key=? AND input_hash=?
                AND status='succeeded' ORDER BY id DESC LIMIT 1""", (job_id, stage_key, input_hash)).fetchone()
            return dict(row) if row else None

    def complete_stage(self, stage_id, job_id, worker_token, value, snapshot, files):
        with self.connection(write=True) as db:
            self._assert_lease(db, job_id, worker_token)
            db.execute("""UPDATE stages SET status='succeeded', value_json=?, snapshot_json=?,
                files_json=?, updated_at=? WHERE id=? AND job_id=?""",
                (json_text(value), json_text(snapshot), json_text(files), now_text(), stage_id, job_id))

    def fail_stage(self, stage_id, job_id, worker_token, status, error):
        with self.connection(write=True) as db:
            self._assert_lease(db, job_id, worker_token)
            db.execute("UPDATE stages SET status=?, error=?, updated_at=? WHERE id=? AND job_id=?",
                       (status, error, now_text(), stage_id, job_id))

    def record_stage_usage(self, stage_id, job_id, worker_token, model, usage, request_id):
        with self.connection(write=True) as db:
            self._assert_lease(db, job_id, worker_token)
            row = db.execute("SELECT usage_json FROM stages WHERE id=? AND job_id=?", (stage_id, job_id)).fetchone()
            calls = json.loads(row[0]) if row and row[0] else []
            calls.append({"model": model, "usage": usage, "request_id": request_id})
            db.execute("UPDATE stages SET usage_json=?, request_id=?, updated_at=? WHERE id=? AND job_id=?",
                       (json_text(calls), request_id, now_text(), stage_id, job_id))

    def stages(self, job_id):
        with self.connection() as db:
            return [dict(row) for row in db.execute("SELECT * FROM stages WHERE job_id=? ORDER BY id", (job_id,))]

    def finish_job(self, job_id, worker_token, snapshot, manifest):
        if not self.verify_assets(manifest):
            raise ValueError("版本资产校验失败")
        version_id, timestamp = str(uuid4()), now_text()
        with self.connection(write=True) as db:
            job = self._assert_lease(db, job_id, worker_token)
            current = db.execute("SELECT current_version_id FROM decks WHERE id=?", (job["deck_id"],)).fetchone()[0]
            if current != job["base_version_id"]:
                raise StorageConflict("基准版本已改变，未提交本轮导出文件")
            snapshot = json.loads(json_text(snapshot))
            snapshot["data"].update({"_version_id": version_id, "generation_status": "succeeded"})
            db.execute("INSERT INTO versions VALUES (?, ?, ?, ?, ?, ?)",
                       (version_id, job["deck_id"], job["base_version_id"], timestamp, json_text(snapshot), json_text(manifest)))
            for item in manifest:
                db.execute("INSERT INTO artifacts VALUES (?, ?, ?, ?, ?, ?)",
                           (item["id"], version_id, item["role"], item["path"], item["sha256"], item["size"]))
            db.execute("UPDATE decks SET current_version_id=? WHERE id=?", (version_id, job["deck_id"]))
            db.execute("UPDATE conversations SET updated_at=? WHERE id=?", (timestamp, job["conversation_id"]))
            db.execute("""UPDATE jobs SET status='succeeded', result_version_id=?, snapshot_json=?,
                updated_at=?, worker_token=NULL, lease_until=NULL WHERE id=?""",
                (version_id, json_text(snapshot), timestamp, job_id))
        return version_id

    def fail_job(self, job_id, worker_token, snapshot, error, status="failed"):
        with self.connection(write=True) as db:
            self._assert_lease(db, job_id, worker_token)
            snapshot = json.loads(json_text(snapshot))
            snapshot["data"]["generation_status"] = status
            db.execute("""UPDATE jobs SET status=?, snapshot_json=?, error=?, updated_at=?,
                worker_token=NULL, lease_until=NULL WHERE id=?""", (status, json_text(snapshot), error, now_text(), job_id))

    def recover_expired(self):
        recovered = []
        with self.connection(write=True) as db:
            for job in db.execute("SELECT * FROM jobs WHERE status='running' AND lease_until<=?", (time.time(),)).fetchall():
                external_pending = db.execute("SELECT 1 FROM stages WHERE job_id=? AND external=1 AND status IN ('running','unknown')",
                                              (job["id"],)).fetchone()
                status = "unknown" if external_pending else "queued"
                db.execute("UPDATE stages SET status=? WHERE job_id=? AND status='running'",
                           ("unknown" if external_pending else "interrupted", job["id"]))
                db.execute("""UPDATE jobs SET status=?, worker_token=NULL, lease_until=NULL,
                    dispatched_at=NULL, updated_at=?, error=? WHERE id=?""",
                    (status, now_text(), "worker 中断，外部调用结果待确认" if external_pending else "worker 中断，等待恢复", job["id"]))
                recovered.append({"id": job["id"], "status": status})
        return recovered

    def retry_job(self, job_id, confirm_unknown=False, owner=LOCAL_OWNER):
        with self.connection(write=True) as db:
            job = db.execute("SELECT * FROM jobs WHERE id=? AND owner_id=?", (job_id, owner)).fetchone()
            if not job or job["status"] not in {"failed", "unknown", "conflict"}:
                raise StorageConflict("任务不处于可重试状态")
            if job["status"] == "unknown" and not confirm_unknown:
                raise StorageConflict("外部调用结果未知，确认可能重复消耗额度后才能重试")
            current = db.execute("SELECT current_version_id FROM decks WHERE id=?", (job["deck_id"],)).fetchone()[0]
            if current != job["base_version_id"] and json.loads(job["request_json"]).get("kind") not in {"dialogue", "design_candidates"}:
                raise StorageConflict("基准版本已改变，请创建新的修改任务")
            kind = json.loads(job["request_json"]).get("kind", "generation")
            kind = kind if kind in {"dialogue", "design_candidates"} else "generation"
            if db.execute("SELECT 1 FROM jobs WHERE deck_id=? AND status IN ('queued','running') AND CASE WHEN json_extract(request_json, '$.kind') IN ('dialogue','design_candidates') THEN json_extract(request_json, '$.kind') ELSE 'generation' END=?", (job["deck_id"], kind)).fetchone():
                raise StorageConflict("当前文稿已有待执行任务")
            if confirm_unknown:
                db.execute("UPDATE stages SET status='acknowledged' WHERE job_id=? AND status='unknown'", (job_id,))
                for stage in db.execute("SELECT * FROM stages WHERE job_id=? AND external=1 AND status='succeeded'", (job_id,)).fetchall():
                    if not self.verify_assets(json.loads(stage["files_json"])):
                        db.execute("UPDATE stages SET status='invalidated' WHERE id=?", (stage["id"],))
            db.execute("UPDATE jobs SET status='queued', dispatched_at=NULL, error=NULL, updated_at=? WHERE id=?",
                       (now_text(), job_id))

    def artifact(self, artifact_id, owner=LOCAL_OWNER):
        with self.connection() as db:
            row = db.execute("""SELECT a.*, c.id AS conversation_id FROM artifacts a
                JOIN versions v ON v.id=a.version_id JOIN decks d ON d.id=v.deck_id
                JOIN conversations c ON c.id=d.conversation_id WHERE a.id=? AND c.owner_id=?""", (artifact_id, owner)).fetchone()
            return dict(row) if row else None

    def version_artifacts(self, version_id, owner=LOCAL_OWNER):
        with self.connection() as db:
            return [dict(row) for row in db.execute("""SELECT a.* FROM artifacts a
                JOIN versions v ON v.id=a.version_id JOIN decks d ON d.id=v.deck_id
                JOIN conversations c ON c.id=d.conversation_id WHERE v.id=? AND c.owner_id=?""", (version_id, owner))]

    def messages(self, conversation_id, owner=LOCAL_OWNER):
        with self.connection() as db:
            return [dict(row) for row in db.execute("""SELECT m.* FROM messages m
                JOIN conversations c ON c.id=m.conversation_id WHERE c.id=? AND c.owner_id=? ORDER BY m.id""", (conversation_id, owner))]

    def finish_design_candidates(self, job_id, worker_token, snapshot):
        with self.connection(write=True) as db:
            job = self._assert_lease(db, job_id, worker_token)
            if json.loads(job["request_json"]).get("kind") != "design_candidates":
                raise StorageConflict("任务不是设计候选任务")
            db.execute("UPDATE jobs SET status='succeeded', snapshot_json=?, updated_at=?, worker_token=NULL, lease_until=NULL WHERE id=?",
                       (json_text(snapshot), now_text(), job_id))

    def design_jobs(self, conversation_id, owner=LOCAL_OWNER):
        with self.connection() as db:
            return [dict(row) for row in db.execute("""SELECT * FROM jobs WHERE conversation_id=? AND owner_id=?
                AND json_extract(request_json, '$.kind')='design_candidates' ORDER BY created_at DESC, rowid DESC LIMIT 10""", (conversation_id, owner))]

    def finish_dialogue(self, job_id, worker_token, snapshot, response):
        timestamp = now_text()
        with self.connection(write=True) as db:
            job = self._assert_lease(db, job_id, worker_token)
            if json.loads(job["request_json"]).get("kind") != "dialogue":
                raise StorageConflict("任务不是对话任务")
            content = {"type": "dialogue_reply", "version_id": job["base_version_id"], "response": {key: value for key, value in response.items() if key != "memory"}}
            db.execute("INSERT INTO messages (conversation_id, role, content, job_id, created_at) VALUES (?, 'assistant', ?, ?, ?)",
                       (job["conversation_id"], json_text(content), job_id, timestamp))
            db.execute("UPDATE jobs SET status='succeeded', snapshot_json=?, updated_at=?, worker_token=NULL, lease_until=NULL WHERE id=?",
                       (json_text(snapshot), timestamp, job_id))
            db.execute("UPDATE conversations SET updated_at=? WHERE id=?", (timestamp, job["conversation_id"]))

    def dialogue_records(self, conversation_id, owner=LOCAL_OWNER, limit=None, before_id=None):
        with self.connection() as db:
            query = """SELECT m.*, COALESCE(json_extract(j.request_json, '$.kind'), 'generation') AS kind
                FROM messages m JOIN conversations c ON c.id=m.conversation_id LEFT JOIN jobs j ON j.id=m.job_id
                WHERE c.id=? AND c.owner_id=? AND (? IS NULL OR m.id<?)"""
            parameters = [conversation_id, owner, before_id, before_id]
            if limit:
                query += " ORDER BY m.id DESC LIMIT ?"; parameters.append(limit)
            else:
                query += " ORDER BY m.id"
            rows = db.execute(query, parameters).fetchall()
            if limit:
                rows = list(reversed(rows))
        records = []
        for row in rows:
            item = dict(row)
            if item["role"] == "user":
                records.append(item)
            elif item["role"] == "assistant" and item["kind"] == "dialogue":
                try:
                    envelope = json.loads(item["content"])
                    if envelope.get("type") != "dialogue_reply":
                        continue
                    item.update(content=envelope["response"]["answer"], response=envelope["response"], version_id=envelope.get("version_id"))
                    records.append(item)
                except (ValueError, KeyError, TypeError):
                    continue
        return records

    def latest_message_id(self, conversation_id, owner=LOCAL_OWNER):
        with self.connection() as db:
            return db.execute("""SELECT MAX(m.id) FROM messages m JOIN conversations c ON c.id=m.conversation_id
                WHERE c.id=? AND c.owner_id=?""", (conversation_id, owner)).fetchone()[0] or 0

    def dialogue_jobs(self, conversation_id, owner=LOCAL_OWNER):
        with self.connection() as db:
            return [dict(row) for row in db.execute("""SELECT id, status, base_version_id, error, updated_at FROM jobs
                WHERE conversation_id=? AND owner_id=? AND json_extract(request_json, '$.kind')='dialogue'
                ORDER BY created_at DESC, rowid DESC LIMIT 20""", (conversation_id, owner))]

    def imported_runs(self, owner=LOCAL_OWNER):
        with self.connection() as db:
            return {row["run_id"]: row["conversation_id"] for row in db.execute("""SELECT l.* FROM legacy_imports l
                JOIN conversations c ON c.id=l.conversation_id WHERE c.owner_id=?""", (owner,))}

    def import_legacy(self, run_id, source_hash, snapshot, manifest, owner=LOCAL_OWNER):
        if not self.verify_assets(manifest):
            raise ValueError("归档资产校验失败")
        with self.connection(write=True) as db:
            existing = db.execute("""SELECT l.*, c.owner_id FROM legacy_imports l
                JOIN conversations c ON c.id=l.conversation_id WHERE l.run_id=?""", (run_id,)).fetchone()
            if existing:
                if existing["source_hash"] != source_hash or existing["owner_id"] != owner:
                    raise StorageConflict("归档已导入或内容已改变，不能覆盖既有索引")
                return existing["conversation_id"]
            conversation_id, deck_id, version_id, timestamp = str(uuid4()), str(uuid4()), str(uuid4()), now_text()
            snapshot = json.loads(json_text(snapshot))
            snapshot["data"].update({"_conversation_id": conversation_id, "_deck_id": deck_id,
                                     "_version_id": version_id, "generation_status": "succeeded"})
            db.execute("INSERT INTO conversations VALUES (?, ?, ?, ?, ?)",
                       (conversation_id, owner, snapshot.get("topic") or run_id, timestamp, timestamp))
            db.execute("INSERT INTO decks VALUES (?, ?, ?)", (deck_id, conversation_id, version_id))
            db.execute("INSERT INTO versions VALUES (?, ?, NULL, ?, ?, ?)",
                       (version_id, deck_id, timestamp, json_text(snapshot), json_text(manifest)))
            for item in manifest:
                db.execute("INSERT INTO artifacts VALUES (?, ?, ?, ?, ?, ?)",
                           (item["id"], version_id, item["role"], item["path"], item["sha256"], item["size"]))
            for message in snapshot.get("messages", []):
                if message.get("role") == "user":
                    db.execute("INSERT INTO messages (conversation_id, role, content, created_at) VALUES (?, 'user', ?, ?)",
                               (conversation_id, message.get("content", ""), timestamp))
            db.execute("INSERT INTO legacy_imports VALUES (?, ?, ?)", (run_id, source_hash, conversation_id))
            return conversation_id

    def backup(self, destination):
        destination = Path(destination).resolve()
        if destination.exists() or destination == self.path:
            raise ValueError("备份目标必须是新的文件")
        destination.parent.mkdir(parents=True, exist_ok=True)
        with self.connection() as source:
            target = sqlite3.connect(destination)
            try:
                source.backup(target)
            finally:
                target.close()
