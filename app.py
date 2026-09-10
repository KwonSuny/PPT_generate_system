import os
import re
import json
import threading
from datetime import datetime
from pathlib import Path
from typing import Optional
from flask import Flask, render_template, request, session, jsonify, has_request_context
from main import run_ppt_system, run_incremental_modify_pages
from config.state import PPTGenerationState
from config_manager import get_output_root, load_api_key, save_api_key, has_api_key, clear_api_key, load_base_url, save_base_url
from pptx import Presentation

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "ppt-secret-key")
BASE_DIR = Path(__file__).resolve().parent
# 输出根目录（可被测试覆盖）：所有生成物统一收纳在大时间戳文件夹内
OUTPUT_ROOT = get_output_root()

STATE_SESSION_KEY = "ppt_state"


def _create_run_dir() -> Path:
    """为本次生成创建统一的大时间戳文件夹，PPT 与中间信息都收纳其中。"""
    run_id = datetime.now().strftime("%Y%m%d_%H%M%S")
    run_dir = OUTPUT_ROOT / run_id
    run_dir.mkdir(parents=True, exist_ok=True)
    return run_dir


def _session_pointer_path() -> Path:
    return OUTPUT_ROOT / "current_session.txt"


def _read_session_pointer() -> Optional[str]:
    try:
        sp = _session_pointer_path()
        if sp.exists():
            rid = sp.read_text(encoding="utf-8").strip()
            return rid or None
    except Exception:
        pass
    return None


def _write_session_pointer(run_id: str) -> None:
    try:
        _session_pointer_path().write_text(run_id, encoding="utf-8")
    except Exception as e:
        print(f"[Session] 写入会话指针失败: {e}")


def _clear_session_pointer() -> None:
    try:
        sp = _session_pointer_path()
        if sp.exists():
            sp.unlink()
    except Exception as e:
        print(f"[Session] 清除会话指针失败: {e}")


def _active_run_dir() -> Optional[Path]:
    rid = _read_session_pointer()
    if rid:
        d = OUTPUT_ROOT / rid
        if d.is_dir():
            return d
    return None


def _ensure_active_run_dir() -> Path:
    d = _active_run_dir()
    if d is None:
        d = _create_run_dir()
        _write_session_pointer(d.name)
    return d


def _state_path_for(state: PPTGenerationState) -> Path:
    """返回当前 state 应写入的 state.json 路径（优先取 state 记录的 _run_dir）。"""
    rd = state.get("_run_dir")
    if rd:
        return Path(rd) / "state.json"
    return _ensure_active_run_dir() / "state.json"


def _parse_page_numbers(raw: str):
    """解析页码输入，支持单页/逗号分隔/范围/混合：'2'、'2,3'、'2-4'、'1,3-5'。

    :return: 去重排序后的页码列表（从 1 开始）；无法解析时返回 None。
    """
    raw = (raw or "").strip()
    if not raw:
        return None
    pages = set()
    for part in re.split(r"[,\s，;；]+", raw):
        part = part.strip()
        if not part:
            continue
        if re.fullmatch(r"\d+", part):
            pages.add(int(part))
        elif re.fullmatch(r"\d+\s*-\s*\d+", part):
            start, end = (int(x) for x in re.split(r"\s*-\s*", part))
            if start > end:
                start, end = end, start
            pages.update(range(start, end + 1))
        else:
            return None
    if not pages:
        return None
    return sorted(pages)


def _incremental_label(use_incremental: bool, page_numbers, previous_state, modify_scope: str) -> str:
    """生成增量修改的可读标签，区分单页/多页/全局。"""
    if not use_incremental:
        return "增量修改"
    pages = page_numbers or []
    cached_slides = (previous_state.get("outline") or {}).get("slides") or []
    total = len(cached_slides)
    scope_txt = {"text": "仅改文本", "design": "仅改设计", "both": "整体修改"}.get(modify_scope, "修改")
    if pages and total and len(pages) == total:
        return f"全局{scope_txt}（全部 {total} 页）"
    if len(pages) == 1:
        return f"单页{scope_txt}（第 {pages[0]} 页）"
    if len(pages) > 1:
        return f"多页{scope_txt}（第 {', '.join(map(str, pages))} 页）"
    return "增量修改"


def _api_key() -> str:
    """优先读环境变量（launcher 已注入），其次读本地配置文件。"""
    env_key = os.environ.get("DASHSCOPE_API_KEY", "").strip()
    return env_key or load_api_key()


def _build_ppt_preview(ppt_path: str):
    if not os.path.exists(ppt_path):
        return []
    prs = Presentation(ppt_path)
    preview = []
    for slide in prs.slides:
        texts = []
        for shape in slide.shapes:
            if hasattr(shape, "text") and shape.text:
                texts.append(shape.text.strip())
        title = texts[0] if texts else "未命名页"
        body = " | ".join(texts[1:]) if len(texts) > 1 else ""
        preview.append({"title": title, "text": body})
    return preview


@app.get("/")
def index():
    state = _load_state()
    return render_template(
        "index.html",
        state=state,
        current_outline=state.get("outline"),
        has_previous_state=bool(state.get("outline") or state.get("blueprint"))
    )


@app.get("/setup")
def setup_page():
    return render_template("setup.html", has_key=has_api_key())


@app.get("/api/has_key")
def api_has_key():
    return jsonify({"has_key": has_api_key(), "base_url": load_base_url()})


@app.post("/api/save_key")
def api_save_key():
    payload = request.get_json(silent=True) or {}
    key = (payload.get("api_key") or request.form.get("api_key") or "").strip()
    base_url = (payload.get("base_url") or request.form.get("base_url") or "").strip()
    if not key:
        return jsonify({"status": "error", "message": "请输入 API Key"}), 400
    try:
        save_api_key(key)
    except Exception as e:
        return jsonify({"status": "error", "message": f"密钥加密保存失败：{e}"}), 500
    save_base_url(base_url)
    # 同步到当前进程环境，使后续后台生成线程立即可用
    os.environ["DASHSCOPE_API_KEY"] = key
    if base_url:
        os.environ["DASHSCOPE_BASE_URL"] = base_url
    else:
        os.environ.pop("DASHSCOPE_BASE_URL", None)
    return jsonify({"status": "success"})


@app.post("/api/logout")
def api_logout():
    clear_api_key()
    os.environ.pop("DASHSCOPE_API_KEY", None)
    return jsonify({"status": "success"})


def _load_state() -> PPTGenerationState:
    payload = None
    run_dir = _active_run_dir()
    if run_dir is not None:
        disk_path = run_dir / "state.json"
        if disk_path.exists():
            try:
                with open(disk_path, "r", encoding="utf-8") as f:
                    payload = json.load(f)
            except Exception as e:
                print(f"[Disk Load] 试图读取磁盘状态失败: {e}")
                payload = None

    if not payload:
        payload = session.get(STATE_SESSION_KEY)

    if not payload:
        return PPTGenerationState(topic="", requirements="")

    return PPTGenerationState(
        topic=payload.get("topic", ""),
        requirements=payload.get("requirements", ""),
        messages=payload.get("messages", []),
        data=payload.get("data", {}),
        feedback=payload.get("feedback", []),
        status=payload.get("status", "None"),
        progress_log=payload.get("progress_log", []),
        operation_mode=payload.get("operation_mode", "new"),
        modify_scope=payload.get("modify_scope", "both"),
        modify_page=payload.get("modify_page"),
        modify_pages=payload.get("modify_pages", []),
        requirement_history=payload.get("requirement_history", []),
        history_snapshots=payload.get("history_snapshots", []),
    )

def _save_state(state: PPTGenerationState) -> None:
    snap = state.snapshot()
    try:
        if has_request_context():
            session[STATE_SESSION_KEY] = snap
    except Exception:
        pass

    try:
        disk_path = _state_path_for(state)
        disk_path.parent.mkdir(parents=True, exist_ok=True)
        with open(disk_path, "w", encoding="utf-8") as f:
            json.dump(snap, f, ensure_ascii=False, indent=5)
    except Exception as e:
        print(f"无法写入: {e}")


def _write_snapshot_disk(state: PPTGenerationState) -> None:
    try:
        snap = state.snapshot()
        disk_path = _state_path_for(state)
        disk_path.parent.mkdir(parents=True, exist_ok=True)
        with open(disk_path, "w", encoding="utf-8") as f:
            json.dump(snap, f, ensure_ascii=False, indent=5)
    except Exception as e:
        print(f"无法写入state.json: {e}")


def _persist_safe(state: PPTGenerationState) -> None:
    if has_request_context():
        _save_state(state)
    else:
        _write_snapshot_disk(state)


def _persist_generation_assets(state: PPTGenerationState, stem: str, run_dir: Path) -> dict:
    """把本轮生成的语义大纲与物理蓝图落盘为版本资产，供后续增量修改/回放使用。"""
    result = {}
    try:
        snap_dir = run_dir / "generated" / "snapshots" / stem
        snap_dir.mkdir(parents=True, exist_ok=True)
        outline = state.get("outline")
        blueprint = state.get("blueprint")
        if outline:
            outline_path = snap_dir / "outline.json"
            with open(outline_path, "w", encoding="utf-8") as f:
                json.dump(outline, f, ensure_ascii=False, indent=2)
            result["outline"] = str(outline_path)
        if blueprint:
            blueprint_path = snap_dir / "blueprint.json"
            with open(blueprint_path, "w", encoding="utf-8") as f:
                json.dump(blueprint, f, ensure_ascii=False, indent=2)
            result["blueprint"] = str(blueprint_path)
    except Exception as e:
        print(f"[Snapshot] 写入生成资产失败: {e}")
    return result


@app.get("/state")
def state_api():
    state = _load_state()
    last_file = state.get("last_file") or ""
    preview_path = str(Path(last_file)) if last_file else ""
    return jsonify({
        "status": state.status,
        "progress_log": state.progress_log,
        "last_updated": state.last_updated,
        "preview": _build_ppt_preview(preview_path) if preview_path else [],
        "last_file": last_file,
    })


@app.post("/generate")
def generate():
    topic = (request.form.get("topic") or "").strip()
    requirements = (request.form.get("requirements") or "").strip()
    filename = (request.form.get("filename") or "generated_presentation.pptx").strip()
    mode = (request.form.get("mode") or "new").strip()
    page_number = (request.form.get("modify_page") or "").strip()

    if not topic:
        return render_template("index.html", error="请输入主题内容。")

    user_input = topic
    if requirements:
        user_input = f"{topic}\n要求：{requirements}"

    page_numbers = _parse_page_numbers(page_number)

    previous_state = _load_state()
    has_cache = bool(previous_state.get("outline"))
    # 页面级增量路由
    use_incremental = (
        mode != "new"
        and has_cache
        and (page_numbers is not None or mode in ("modify_text", "modify_design"))
    )
    if use_incremental and not page_numbers:
        cached_slides = (previous_state.get("outline") or {}).get("slides") or []
        page_numbers = list(range(1, len(cached_slides) + 1))
    effective_mode = "incremental_page" if use_incremental else mode
    modify_page = page_numbers[0] if page_numbers else None

    clean_name = re.sub(r"[^A-Za-z0-9_\-]+", "_", filename).strip("_") or "generated_presentation"
    safe_name = f"{clean_name}.pptx" if not clean_name.endswith(".pptx") else clean_name

    run_dir = _create_run_dir()
    run_id = run_dir.name
    _write_session_pointer(run_id)
    output_path = run_dir / safe_name
    output_path_str = str(output_path)

    if mode == "new":
        # 新对话不继承任何历史信息（包括历史快照）
        state = PPTGenerationState(topic=topic, requirements=requirements)
    else:
        state = previous_state
        state.topic = topic
        state.requirements = requirements
    # 让 state 知道自己所属的对话文件夹，持久化时写入对应 run_dir 下的 state.json
    state.set("_run_dir", str(run_dir))
    state.bind_persistence(_persist_safe)

    modify_scope = "both"
    if mode == "modify_text":
        modify_scope = "text"
    elif mode == "modify_design":
        modify_scope = "design"

    state.set_operation(effective_mode, scope=modify_scope, page=modify_page, pages=page_numbers)
    if requirements:
        state.add_requirement_history(requirements)
    if mode == "continue":
        state.append_message("user", f"继续优化当前 PPT：{requirements or topic}")
    elif mode == "modify_text":
        state.append_message("user", f"仅按照最新要求修改文本内容：{requirements or topic}")
    elif mode == "modify_design":
        state.append_message("user", f"仅调整设计风格和版式，不改变文案主体：{requirements or topic}")
    else:
        state.append_message("user", user_input)

    mode_labels = {
        "new": "重新生成",
        "continue": "基于历史继续修改",
        "modify_text": "仅修改文本内容",
        "modify_design": "仅调整设计风格",
        "incremental_page": _incremental_label(use_incremental, page_numbers, previous_state, modify_scope),
    }

    state.set("last_file", output_path_str)

    def run_generation():
        try:
            generated_sub = run_dir / "generated"
            generated_sub.mkdir(parents=True, exist_ok=True)
            run_log_path = generated_sub / f"run_{Path(output_path_str).stem}.log"
            with open(run_log_path, "a", encoding="utf-8") as logf:
                def log(msg: str):
                    line = f"[{__name__}] {msg}\n"
                    print(line, end="")
                    logf.write(line)
                    logf.flush()

                log("Generation thread started")
                _save_state(state)
                if use_incremental:
                    log(f"Routing to incremental modify pages={page_numbers}...")
                    generated_path = run_incremental_modify_pages(
                        user_input,
                        page_numbers=page_numbers,
                        filename=output_path_str,
                        state=state,
                        requirements=requirements,
                        run_dir=str(run_dir),
                    )
                else:
                    log("Calling run_ppt_system...")
                    generated_path = run_ppt_system(
                        user_input,
                        filename=output_path_str,
                        state=state,
                        requirements=requirements,
                        run_dir=str(run_dir),
                    )
                log(f"Pipeline returned: {generated_path}")
                if generated_path:
                    state.set("last_file", generated_path)
                    assets = _persist_generation_assets(state, Path(generated_path).stem, run_dir)
                    state.add_history_snapshot(
                        Path(generated_path).name,
                        outline_path=assets.get("outline"),
                        blueprint_path=assets.get("blueprint"),
                    )
                    state.append_progress("Rendering: 完成，PPT 导出成功")
                _save_state(state)
                log("Generation thread finished normally")
        except Exception as exc:
            try:
                state.append_progress(f"Error: {exc}")
                _save_state(state)
            except Exception:
                pass
            print(f"出错: {exc}")

    thread = threading.Thread(target=run_generation, daemon=True)
    thread.start()

    return render_template(
        "index.html",
        state=state,
        current_outline=state.get("outline"),
        success=True,
        file_name=safe_name,
        save_path=output_path_str,
        has_previous_state=True,
        mode_label=mode_labels.get(effective_mode, "继续修改"),
    )


@app.route("/api/reset", methods=["POST"])
def reset_conversation():
    # 清空当前 Session
    session.clear()

    # 清除当前会话指针：下一次加载将不再携带上一次对话的任何信息
    _clear_session_pointer()

    return jsonify({"status": "success", "message": "已成功重置，新对话已准备就绪！"})


if __name__ == "__main__":
    # 仅监听本机回环地址，避免局域网暴露 /api/save_key 等接口
    app.run(host="127.0.0.1", port=8000, debug=False)
