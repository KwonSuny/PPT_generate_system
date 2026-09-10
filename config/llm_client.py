# -*- coding: utf-8 -*-
"""OpenAI 兼容客户端的健壮性封装：显式超时 + 连接异常指数退避重试。

背景：Layout Agent 曾因 ``httpx.RemoteProtocolError: Server disconnected
without sending a response``（网关/网络在返回响应前主动断开连接）导致整条
生成管线失败。这里统一封装两点：

1. 显式超时：避免请求无限挂起；
2. 对「连接类异常」（APIConnectionError / APITimeoutError）做指数退避重试，
   降低瞬时网络抖动的影响；鉴权失败、参数错误等 4xx 错误不重试、直接抛出。
"""

import time
from typing import Optional

from openai import APIConnectionError, APITimeoutError, OpenAI

DEFAULT_BASE_URL = ""

# 可安全重试的异常：连接被重置、超时。鉴权/参数类错误不在此列。
_RETRYABLE_EXCEPTIONS = (APIConnectionError, APITimeoutError)


def create_client(
    api_key: str,
    base_url: str = DEFAULT_BASE_URL,
    timeout: float = 600.0,
) -> OpenAI:
    """创建带显式超时、关闭内置重试（由 chat_completion_with_retry 统一控制）的客户端。"""
    return OpenAI(api_key=api_key, base_url=base_url, timeout=timeout, max_retries=0)


def chat_completion_with_retry(
    client: OpenAI,
    *,
    model: str,
    messages,
    max_retries: int = 3,
    **kwargs,
):
    """带指数退避重试的 chat.completions.create 封装。

    仅对连接/超时类异常重试；其余异常直接抛出。
    """
    last_exc: Optional[Exception] = None
    for attempt in range(1, max_retries + 1):
        try:
            return client.chat.completions.create(model=model, messages=messages, **kwargs)
        except _RETRYABLE_EXCEPTIONS as exc:
            last_exc = exc
            if attempt >= max_retries:
                break
            wait = min(2 ** attempt, 8)
            print(f"[LLM] 连接异常（第 {attempt}/{max_retries} 次），{wait}s 后重试：{exc}")
            time.sleep(wait)
    raise last_exc
