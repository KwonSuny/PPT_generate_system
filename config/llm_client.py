"""OpenAI 兼容客户端的健壮性封装：显式超时 + 连接异常指数退避重试。

背景：Layout Agent 曾因 ``httpx.RemoteProtocolError: Server disconnected
without sending a response``（网关/网络在返回响应前主动断开连接）导致整条
生成管线失败。这里统一封装两点：

1. 显式超时：避免请求无限挂起；
2. 对「连接类异常」（APIConnectionError / APITimeoutError）做指数退避重试，
   降低瞬时网络抖动的影响；鉴权失败、参数错误等 4xx 错误不重试、直接抛出。
"""

import os
import time
from contextvars import ContextVar
from typing import Optional

from openai import APIConnectionError, APITimeoutError, OpenAI

# 默认地址：阿里云百炼（DashScope）OpenAI 兼容模式。
DEFAULT_BASE_URL = "https://dashscope.aliyuncs.com/compatible-mode/v1"

DEFAULT_MODEL_NAMES = {
    "planner": "qwen3.8-27b",
    "writer": "qwen3.8-27b",
    "layout": "qwen3.8-27b",
    "image": "qwen-image-2.1-pro",
    "vision": "qwen3.8-27b",
    "dialogue": "qwen3.8-27b",
}

MODEL_REQUEST_GUARD = ContextVar("ppt_model_request_guard", default=None)
MODEL_CALL_OBSERVER = ContextVar("ppt_model_call_observer", default=None)


def guard_model_request():
    guard = MODEL_REQUEST_GUARD.get()
    if guard:
        guard()


def record_model_response(model, usage=None, request_id=None):
    observer = MODEL_CALL_OBSERVER.get()
    if observer:
        if hasattr(usage, "model_dump"):
            usage = usage.model_dump()
        observer(model, usage, request_id)


def resolve_model_name(role: str) -> str:
    default = DEFAULT_MODEL_NAMES[role]
    return (os.environ.get(f"PPT_{role.upper()}_MODEL") or "").strip() or default


def model_request_options(role, model, search=False):
    """按阶段声明思考策略；未知兼容模型不注入供应商特有开关。"""
    mode = os.environ.get(f'PPT_{role.upper()}_THINKING', 'auto').strip().lower()
    if mode not in {'auto', 'on', 'off'}:
        raise ValueError(f'PPT_{role.upper()}_THINKING 只能为 auto/on/off')
    name = model.lower()
    only_thinking = name.startswith(('kimi-k3', 'glm-5.3'))
    extra = {}
    if not only_thinking and (name.startswith('qwen') or mode!='auto'):
        extra['enable_thinking'] = mode=='on'
    if mode=='on' and name.startswith('qwen'):
        extra['thinking_budget'] = min(8192, max(256, int(os.environ.get('PPT_THINKING_BUDGET','2048'))))
    if search and name.startswith('qwen') and os.environ.get('PPT_ENABLE_SEARCH','1')!='0':
        extra['enable_search'] = True
    limits = {'planner':6144, 'writer':12288, 'layout':8192, 'dialogue':4096, 'vision':2048}
    if role=='vision' and only_thinking:
        limits[role]=4096
    result={'extra_body':extra, 'max_tokens':limits[role]}
    # Kimi K3 等模型限制采样参数；未显式覆盖时采用模型服务默认值。
    temperature=os.environ.get(f'PPT_{role.upper()}_TEMPERATURE','').strip()
    if temperature or (name.startswith('qwen') and not only_thinking):
        default={'planner':.5,'writer':.4,'layout':.35,'dialogue':.2,'vision':.1}[role]
        result['temperature']=float(temperature) if temperature else default
        if not 0 <= result['temperature'] <= 2:
            raise ValueError('采样温度必须在0–2之间')
    return result

# 仅重试连接和超时异常；供应商可能已执行请求，重试仍可能重复消耗额度。
_RETRYABLE_EXCEPTIONS = (APIConnectionError, APITimeoutError)


def create_client(
    api_key: str,
    base_url: Optional[str] = None,
    timeout: float = 600.0,
) -> OpenAI:
    """创建带显式超时、关闭内置重试（由 chat_completion_with_retry 统一控制）的客户端。

    base_url 三级解析：显式传参 > 环境变量 DASHSCOPE_BASE_URL > DEFAULT_BASE_URL。
    空串/None 视为未指定，防止把空协议传给 httpx 触发 UnsupportedProtocol。
    """
    resolved = (base_url or "").strip() or (os.environ.get("DASHSCOPE_BASE_URL") or "").strip() or DEFAULT_BASE_URL
    return OpenAI(api_key=api_key, base_url=resolved, timeout=timeout, max_retries=0)


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
            guard_model_request()
            response = client.chat.completions.create(model=model, messages=messages, **kwargs)
            record_model_response(model, getattr(response, "usage", None), getattr(response, "_request_id", None))
            return response
        except _RETRYABLE_EXCEPTIONS as exc:
            last_exc = exc
            # 持久 worker 交由阶段记录处理未知结果，避免同一超时请求被盲目重发。
            if MODEL_REQUEST_GUARD.get() is not None or attempt >= max_retries:
                break
            wait = min(2 ** attempt, 8)
            print(f"[LLM] 连接异常（第 {attempt}/{max_retries} 次），{wait}s 后重试：{exc}")
            time.sleep(wait)
    raise last_exc
