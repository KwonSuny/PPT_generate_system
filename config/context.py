"""保守输入预算与可追溯的摘录摘要；原始消息不删改。"""
import json
import os
import re
from hashlib import sha256


class ContextBudgetError(ValueError):
    pass


# 2026-10-13 核实百炼文档；业务预算与供应商窗口分开。
KNOWN_WINDOWS = {"qwen3.8-flash": 1_000_000, "qwen3.8-27b": 1_000_000}


def token_bound(value):
    text = value if isinstance(value, str) else json.dumps(value, ensure_ascii=False, separators=(",", ":"))
    return len(text.encode("utf-8")) + 64


def input_budget(model):
    try:
        configured = int(os.environ.get("PPT_CONTEXT_INPUT_TOKENS", "24000"))
        window = int(os.environ.get("PPT_CONTEXT_WINDOW", str(KNOWN_WINDOWS.get(model, 0))))
    except ValueError as exc:
        raise ContextBudgetError("上下文预算配置必须是整数") from exc
    if not 2000 <= configured <= 500000:
        raise ContextBudgetError("输入预算必须在 2000–500000 之间")
    if window:
        configured = min(configured, window - 16384)
    if configured < 2000:
        raise ContextBudgetError("模型窗口没有足够的输入与输出空间")
    return configured


def compact_history(records, budget):
    records = [{"id": item.get("id", i + 1), "role": item["role"], "content": item.get("content", "")}
               for i, item in enumerate(records) if item.get("role") in {"user", "assistant"}]
    digest = sha256(json.dumps(records, ensure_ascii=False, sort_keys=True).encode()).hexdigest()
    if token_bound(records) <= budget:
        return records, {"strategy": "recent", "source_hash": digest, "source_ids": [r["id"] for r in records], "omitted": 0}
    recent, used = [], 64
    for record in reversed(records):
        size = token_bound(record)
        if len(recent) >= 6 or used + size > budget * 0.55:
            break
        recent.insert(0, record); used += size
    old = records[:len(records) - len(recent)] if recent else records
    extracts = []
    remaining = max(0, budget - token_bound(recent) - 256)
    # 摘录不改写否定项和数字；较新的要求优先，遗漏明确计数。
    for record in reversed(old):
        if record["role"] != "user":
            continue
        sentences = re.split(r"(?<=[。！？\n])", record["content"])
        quoted = "".join(s for s in sentences if re.search(r"必须|禁止|不要|不能|保留|不生图|页|风格|颜色|预算|数据|来源|仅|只|SVG|表格", s, re.I))
        quoted = quoted or record["content"]
        entry = {"source_id": record["id"], "quote": quoted}
        size = token_bound(entry)
        if size <= remaining:
            extracts.insert(0, entry); remaining -= size
    retained = {entry["source_id"] for entry in extracts}
    quotes = {entry["source_id"]: entry["quote"] for entry in extracts}
    omitted = sum(record["id"] not in retained or quotes.get(record["id"]) != record["content"] for record in old)
    summary = {"strategy": "extractive", "source_hash": digest, "source_range": [old[0]["id"], old[-1]["id"]] if old else [], "source_count": len(old),
               "extracts": extracts, "omitted": omitted,
               "notice": "历史摘录只用于恢复需求与偏好，不能作为 PPT 事实。遗漏历史时应澄清，不能声称完整记忆。"}
    return ([{"role": "user", "content": json.dumps({"history_summary": summary}, ensure_ascii=False)}] + recent), summary


def bounded_messages(system, task, records, model):
    budget = input_budget(model)
    mandatory = [{"role": "system", "content": system}, {"role": "user", "content": task}]
    used = token_bound(mandatory)
    if used > budget:
        raise ContextBudgetError("本次内容超出输入预算，请缩小页码范围或精简本次资料；原始记录仍保留")
    remaining = budget - used - 256
    history, summary = compact_history(records, max(0, remaining)) if records and remaining > 512 else ([], {"strategy": "none", "omitted": len(records)})
    messages = mandatory[:1] + [{"role": r["role"], "content": r["content"]} for r in history] + mandatory[1:]
    # JSON 封装也占空间；摘要过大时退回最近记录，不裁剪当前任务。
    while token_bound(messages) > budget and len(messages) > 2:
        messages.pop(1)
        summary["omitted"] = summary.get("omitted", 0) + 1
    if summary.get("omitted"):
        messages[-1]["content"] += "\n历史上下文有省略；涉及未展示的旧要求时先澄清，不得假装已完整记住。"
        while token_bound(messages) > budget and len(messages) > 2:
            messages.pop(1)
        if token_bound(messages) > budget:
            raise ContextBudgetError("当前任务已占满输入预算，请精简本次资料")
    summary.update(input_upper_bound=token_bound(messages), input_budget=budget, model_window=int(os.environ["PPT_CONTEXT_WINDOW"]) if os.environ.get("PPT_CONTEXT_WINDOW") else KNOWN_WINDOWS.get(model),
                   estimator="utf8_bytes_upper_bound", human_review="not_reviewed")
    return messages, summary


def generation_messages(state, system, task, role):
    from config.llm_client import resolve_model_name
    context = {"topic": state.topic, "requirements": state.requirements, "page_policy": state.get("page_policy"),
               "operation_mode": state.operation_mode, "modify_scope": state.modify_scope,
               "modify_pages": state.modify_pages, "feedback": state.feedback}
    if role == "planner" and state.should_use_incremental_mode():
        context["current_outline"] = state.get("outline")
    records = state.get("dialogue_history") or [m for m in state.messages if m.get("role") == "user"]
    messages, summary = bounded_messages(system, task + "\n当前任务：" + json.dumps(context, ensure_ascii=False), records, resolve_model_name(role))
    state.set("context_compaction", {**(state.get("context_compaction") or {}), role: summary})
    return messages
