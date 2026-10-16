# -*- coding: utf-8 -*-
"""桌面启动器：后台启动 Flask 服务，用 pywebview 套独立原生窗口。

用法：python launcher.py
- 首次运行（未配置 API Key）：打开设置页，填写后自动进入主界面。
- 已配置：直接进入主界面。
"""

import os
import socket
import subprocess
import sys
import threading
import time
import urllib.request
from pathlib import Path

import webview

import app as app_module
from config_manager import load_api_key


def _free_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def _start_flask(port: int) -> None:
    app_module.app.run(host="127.0.0.1", port=port, debug=False, use_reloader=False)


def _wait_until_ready(base_url: str, timeout: float = 10.0) -> bool:
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            urllib.request.urlopen(f"{base_url}/api/has_key", timeout=0.5)
            return True
        except Exception:
            time.sleep(0.1)
    return False


def _window_size():
    try:
        screen = webview.screens[0]
        return max(720, min(1440, screen.width - 64)), max(560, min(900, screen.height - 80))
    except Exception:
        return 1280, 780


def main() -> None:
    key = os.environ.get("DASHSCOPE_API_KEY") or load_api_key()
    output_root = app_module.OUTPUT_ROOT
    output_root.mkdir(parents=True, exist_ok=True)
    worker_log = (output_root / "worker.log").open("a", encoding="utf-8")
    worker = subprocess.Popen(
        [sys.executable, "-X", "utf8", str(Path(__file__).with_name("worker.py")), "--root", str(output_root)],
        stdout=worker_log, stderr=subprocess.STDOUT,
        creationflags=subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0,
    )
    try:
        port = _free_port()
        threading.Thread(target=_start_flask, args=(port,), daemon=True).start()
        base_url = f"http://127.0.0.1:{port}"
        if not _wait_until_ready(base_url):
            raise RuntimeError("Web 服务启动失败，请检查本机端口与应用日志")
        first_url = base_url if key else f"{base_url}/setup"
        width, height = _window_size()
        webview.create_window(
            "PPT 工作台", first_url, width=width, height=height,
            min_size=(min(960, width), min(640, height)), confirm_close=False,
        )
        # pywebview 默认取消下载；交由内嵌浏览器的保存对话框处理本地文件。
        webview.settings['ALLOW_DOWNLOADS'] = True
        webview.start()
    finally:
        worker.terminate()
        try:
            worker.wait(timeout=10)
        except subprocess.TimeoutExpired:
            worker.kill()
            worker.wait()
        worker_log.close()


if __name__ == "__main__":
    main()
