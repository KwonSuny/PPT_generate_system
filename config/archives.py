import json
import os
import re
import shutil
import sqlite3
from contextlib import closing
from pathlib import Path
from tempfile import mkdtemp

from pptx import Presentation

from config.jobs import atomic_json
from config.storage import LOCAL_OWNER, PPTStore, content_hash, file_hash, now_text


def read_archive(store, run_id):
    if not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9_-]{0,127}", run_id):
        raise ValueError("归档目录名无效")
    run_dir = store.asset_path(run_id)
    if run_dir.parent != store.root:
        raise ValueError("归档目录超出输出根目录")
    state_file = run_dir / "state.json"
    if not state_file.is_file() or state_file.stat().st_size > 32 * 1024 * 1024:
        raise ValueError("归档状态文件缺失或过大")
    snapshot = json.loads(state_file.read_text(encoding="utf-8-sig"))
    if not isinstance(snapshot, dict) or not isinstance(snapshot.get("data"), dict):
        raise ValueError("归档状态格式无效")
    if snapshot["data"].get("_job_id"):
        raise ValueError("这是新任务目录，请从会话列表选择")
    paths = {}

    def resolve(value):
        value = store.rebase_path(value)
        normalized = value.replace("\\", "/")
        if "/debug_logs/" in normalized:
            path = store.asset_path(normalized.split("/debug_logs/", 1)[1])
        elif normalized.startswith("debug_logs/"):
            path = store.asset_path(normalized[len("debug_logs/"):])
        elif Path(value).is_absolute():
            path = Path(value).resolve()
        else:
            path = (run_dir / normalized).resolve()
        if not path.is_relative_to(store.root) or not path.is_file():
            raise ValueError("归档引用的资产缺失或超出输出根目录")
        return path

    def visit(value):
        if isinstance(value, dict):
            for key, item in value.items():
                if key in {"last_file", "image_path", "outline_path", "blueprint_path"} and item:
                    if not isinstance(item, str):
                        raise ValueError("归档资产路径格式无效")
                    path = resolve(item)
                    value[key] = str(path)
                    role = "pptx" if key == "last_file" else key.removesuffix("_path")
                    paths[(str(path), role)] = path
                else:
                    visit(item)
        elif isinstance(value, list):
            for item in value:
                visit(item)

    visit(snapshot)
    ppt = snapshot["data"].get("last_file")
    if not ppt:
        raise ValueError("归档没有可验证的成功 PPTX")
    try:
        if not Presentation(ppt).slides:
            raise ValueError("empty")
    except Exception as exc:
        raise ValueError("归档 PPTX 无法读取或没有页面") from exc
    for key in ("messages", "progress_log", "feedback", "requirement_history", "history_snapshots"):
        if key in snapshot and not isinstance(snapshot[key], list):
            raise ValueError("归档历史字段格式无效")
    if any(not isinstance(message, dict) for message in snapshot.get("messages", [])):
        raise ValueError("归档消息格式无效")
    # 一个旧状态可能引用多个历史快照，当前下载只对应 last_file。
    manifest = [store.describe_asset(state_file, "legacy-state")]
    manifest.extend(store.describe_asset(path, role) for (_, role), path in sorted(paths.items()))
    snapshot["data"]["_run_dir"] = str(run_dir)
    evidence = [{key: item[key] for key in ("role", "path", "sha256", "size")} for item in manifest]
    return snapshot, manifest, content_hash(evidence)


def scan_archives(store, owner=LOCAL_OWNER):
    imported = store.imported_runs(owner)
    with store.connection() as db:
        new_runs = {row[0] for row in db.execute("SELECT run_dir FROM jobs")}
    result = []
    for run_dir in sorted(store.root.iterdir(), key=lambda path: path.name, reverse=True):
        if not run_dir.is_dir() or run_dir.name in new_runs or not (run_dir / "state.json").is_file():
            continue
        record = {"run_id": run_dir.name, "conversation_id": imported.get(run_dir.name)}
        try:
            snapshot, manifest, digest = read_archive(store, run_dir.name)
            record.update({"status": "imported" if record["conversation_id"] else "ready",
                           "title": snapshot.get("topic", run_dir.name), "source_hash": digest,
                           "assets": [{key: item[key] for key in ("role", "path", "sha256", "size")} for item in manifest]})
        except Exception as exc:
            record.update({"status": "invalid", "message": str(exc) if isinstance(exc, ValueError)
                           else f"归档无法验证（{type(exc).__name__}）"})
        result.append(record)
    return result


def _copy_checked(source, destination):
    digest, size = file_hash(source), source.stat().st_size
    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(source, destination)
    if file_hash(destination) != digest or destination.stat().st_size != size:
        raise ValueError("文件在备份过程中发生变化，请稍后重新备份")
    return {"path": destination, "sha256": digest, "size": size}


def _validate_database(root):
    with closing(sqlite3.connect(root / "app.sqlite3")) as db:
        if db.execute("PRAGMA integrity_check").fetchone()[0] != "ok" or db.execute("PRAGMA foreign_key_check").fetchone():
            raise ValueError("数据库完整性校验失败")
        manifests = [json.loads(row[0]) for row in db.execute("SELECT manifest_json FROM versions")]
        manifests += [json.loads(row[0]) for row in db.execute("SELECT files_json FROM stages WHERE status='succeeded'")]
    for manifest in manifests:
        for item in manifest:
            path = (root / item["path"]).resolve()
            if not path.is_relative_to(root.resolve()) or not path.is_file() or file_hash(path) != item["sha256"]:
                raise ValueError("版本或阶段资产校验失败")


def create_backup(store, destination):
    destination = Path(destination).expanduser().resolve()
    if destination.exists() or destination.is_relative_to(store.root):
        raise ValueError("备份必须写入输出根目录以外的新目录")
    destination.parent.mkdir(parents=True, exist_ok=True)
    temporary = Path(mkdtemp(prefix=".ppt_backup_", dir=destination.parent)).resolve()
    try:
        store.backup(temporary / "app.sqlite3")
        queue_path = store.root / "queue.sqlite3"
        if queue_path.is_file():
            with closing(sqlite3.connect(queue_path)) as source, closing(sqlite3.connect(temporary / "queue.sqlite3")) as target:
                source.backup(target)
        database_files = {name + suffix for name in ("app.sqlite3", "queue.sqlite3") for suffix in ("", "-wal", "-shm", "-journal")}
        for path in store.root.rglob("*"):
            if not path.is_file() or (path.parent == store.root and path.name in database_files) or path.suffix == ".tmp":
                continue
            if not path.resolve().is_relative_to(store.root):
                raise ValueError("备份目录包含指向根目录外的文件")
            _copy_checked(path, temporary / path.relative_to(store.root))
        _validate_database(temporary)
        files = [{"path": path.relative_to(temporary).as_posix(), "sha256": file_hash(path), "size": path.stat().st_size}
                 for path in sorted(temporary.rglob("*")) if path.is_file()]
        atomic_json(temporary / "backup.json", {"format": 1, "created_at": now_text(), "source_root": str(store.root), "files": files})
        if destination.exists():
            raise ValueError("备份目标已存在，未覆盖")
        os.replace(temporary, destination)
    finally:
        if temporary.exists() and temporary.parent == destination.parent and temporary.name.startswith(".ppt_backup_"):
            shutil.rmtree(temporary)
    return destination


def restore_backup(source, destination):
    source, destination = Path(source).resolve(), Path(destination).expanduser().resolve()
    if destination.exists() or destination.is_relative_to(source) or source.is_relative_to(destination):
        raise ValueError("恢复必须使用独立的新目录，不能覆盖已有数据")
    manifest = json.loads((source / "backup.json").read_text(encoding="utf-8"))
    if manifest.get("format") != 1:
        raise ValueError("备份格式不受支持")
    destination.parent.mkdir(parents=True, exist_ok=True)
    temporary = Path(mkdtemp(prefix=".ppt_restore_", dir=destination.parent)).resolve()
    try:
        seen = set()
        for item in manifest["files"]:
            relative = item["path"]
            src, dst = (source / relative).resolve(), (temporary / relative).resolve()
            if not src.is_relative_to(source) or not dst.is_relative_to(temporary) or relative in seen:
                raise ValueError("备份清单包含非法或重复路径")
            seen.add(relative)
            if file_hash(src) != item["sha256"] or src.stat().st_size != item["size"]:
                raise ValueError("备份文件校验失败")
            copied = _copy_checked(src, dst)
            if copied["sha256"] != item["sha256"]:
                raise ValueError("备份文件在恢复时发生变化")
        _validate_database(temporary)
        if destination.exists():
            raise ValueError("恢复目标已存在，未覆盖")
        os.replace(temporary, destination)
    finally:
        if temporary.exists() and temporary.parent == destination.parent and temporary.name.startswith(".ppt_restore_"):
            shutil.rmtree(temporary)
    store = PPTStore(destination)
    # 备份时运行中的任务不保留旧 worker 的租约。
    with store.connection(write=True) as db:
        db.execute("UPDATE jobs SET lease_until=0 WHERE status='running'")
        db.execute("UPDATE jobs SET dispatched_at=NULL WHERE status='queued'")
    store.recover_expired()
    return destination
