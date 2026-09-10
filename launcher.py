# -*- coding: utf-8 -*-
"""桌面启动器：后台启动 Flask 服务，用 pywebview 套独立原生窗口。

用法：python launcher.py
- 首次运行（未配置 API Key）：打开设置页，填写后自动进入主界面。
- 已配置：直接进入主界面。
"""

import os
import socket
import threading
import time
import urllib.request

import webview

import app as app_module
from config_manager import load_api_key, load_base_url


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


def main() -> None:
    key = load_api_key()
    if key:
        os.environ["DASHSCOPE_API_KEY"] = key
    base_url = load_base_url()
    if base_url:
        os.environ["DASHSCOPE_BASE_URL"] = base_url

    port = _free_port()
    threading.Thread(target=_start_flask, args=(port,), daemon=True).start()

    base_url = f"http://127.0.0.1:{port}"
    _wait_until_ready(base_url)

    first_url = base_url if key else f"{base_url}/setup"

    webview.create_window(
        "PPT 生成器",
        first_url,
        width=1280,
        height=820,
        min_size=(960, 640),
        confirm_close=False,
    )
    webview.start()


if __name__ == "__main__":
    main()
