from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from typing import Any, Callable, Dict, List, Optional


@dataclass
class PPTGenerationState:
    topic: str
    requirements: str = ""
    messages: List[Dict[str, str]] = field(default_factory=list)
    data: Dict[str, Any] = field(default_factory=dict)
    feedback: List[str] = field(default_factory=list)
    status: str = "idle"
    progress_log: List[str] = field(default_factory=list)
    last_updated: Optional[str] = None
    operation_mode: str = "new"
    modify_scope: str = "both"
    modify_page: Optional[int] = None
    modify_pages: List[int] = field(default_factory=list)
    requirement_history: List[str] = field(default_factory=list)
    history_snapshots: List[Dict[str, Any]] = field(default_factory=list)
    persistence_callback: Optional[Callable[["PPTGenerationState"], None]] = None
    stage_executor: Optional[Callable[..., Any]] = field(default=None, repr=False)
    strict_persistence: bool = field(default=False, repr=False)

    def __post_init__(self) -> None:
        self.data.setdefault("topic", self.topic)
        self.data.setdefault("requirements", self.requirements)
        self.data.setdefault("operation_mode", self.operation_mode)
        self.data.setdefault("modify_scope", self.modify_scope)
        self.data.setdefault("modify_page", self.modify_page)
        self.data.setdefault("modify_pages", self.modify_pages)
        self.data.setdefault("progress_log", self.progress_log)
        self.data.setdefault("status", self.status)
        self.data.setdefault("last_updated", self.last_updated)
        self.data.setdefault("requirement_history", self.requirement_history)
        self.data.setdefault("history_snapshots", self.history_snapshots)

    def bind_persistence(self, callback: Callable[["PPTGenerationState"], None], strict=False) -> None:
        self.persistence_callback = callback
        self.strict_persistence = strict

    @classmethod
    def from_snapshot(cls, payload):
        state = cls(topic=payload.get("topic", ""))
        state.restore_snapshot(payload)
        return state

    def restore_snapshot(self, payload):
        for name in ("topic", "requirements", "messages", "data", "feedback", "status", "progress_log",
                     "last_updated", "operation_mode", "modify_scope", "modify_page", "modify_pages",
                     "requirement_history", "history_snapshots"):
            if name in payload:
                setattr(self, name, payload[name])

    def execute_stage(self, key, inputs, operation, model=None, result_model=None, external=True):
        if self.stage_executor:
            return self.stage_executor(key, inputs, operation, model, result_model, external)
        return operation()

    def _persist(self) -> None:
        if self.persistence_callback is not None:
            try:
                self.persistence_callback(self)
            except Exception as e:
                if self.strict_persistence:
                    raise
                print(f"[PPTGenerationState] Persistence callback raised: {e}")

    def append_message(self, role: str, content: str) -> None:
        self.messages.append({"role": role, "content": content})
        self._persist()

    def append_progress(self, message: str) -> None:
        self.progress_log.append(message)
        self.status = message
        self.last_updated = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self._persist()

    def set_operation(
        self,
        mode: str,
        scope: str = "both",
        page: Optional[int] = None,
        pages: Optional[List[int]] = None,
    ) -> None:
        self.operation_mode = mode
        self.modify_scope = scope
        self.modify_page = page
        self.modify_pages = list(pages) if pages else ([page] if page is not None else [])
        self.data["operation_mode"] = mode
        self.data["modify_scope"] = scope
        self.data["modify_page"] = page
        self.data["modify_pages"] = self.modify_pages
        self._persist()

    def page_label(self) -> str:
        """返回目标页的可读描述，供 Agent 提示词使用。"""
        pages = self.modify_pages or ([self.modify_page] if self.modify_page is not None else [])
        if not pages:
            return ""
        if len(pages) == 1:
            return f"第 {pages[0]} 页"
        return f"第 {', '.join(str(p) for p in pages)} 页"

    def add_requirement_history(self, requirement: str) -> None:
        if requirement and requirement not in self.requirement_history:
            self.requirement_history.append(requirement)
            self._persist()

    def set(self, key: str, value: Any) -> None:
        self.data[key] = value
        self._persist()

    def add_history_snapshot(
        self,
        file_name: str,
        outline_path: Optional[str] = None,
        blueprint_path: Optional[str] = None,
    ) -> None:
        outline = self.get("outline") or {}
        slides = outline.get("slides") if isinstance(outline, dict) else None
        slide_titles = []
        if isinstance(slides, list):
            for slide in slides:
                title = slide.get("title") if isinstance(slide, dict) else None
                punchline = slide.get("punchline") if isinstance(slide, dict) else None
                slide_titles.append({
                    "title": title,
                    "punchline": punchline,
                })

        entry = {
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "topic": self.topic,
            "requirements": self.requirements,
            "page_policy": self.get("page_policy"),
            "mode": self.operation_mode,
            "modify_page": self.modify_page,
            "modify_pages": self.modify_pages,
            "file_name": file_name,
            "slide_summary": slide_titles,
        }
        if outline_path:
            entry["outline_path"] = outline_path
        if blueprint_path:
            entry["blueprint_path"] = blueprint_path
        self.history_snapshots.append(entry)
        self.data["history_snapshots"] = self.history_snapshots
        self._persist()

    def get(self, key: str, default: Any = None) -> Any:
        if hasattr(self, key):
            return getattr(self, key)
        return self.data.get(key, default)

    def should_use_incremental_mode(self) -> bool:
        return bool(self.get("outline")) or self.operation_mode != "new"

    def build_messages(self, system_prompt: str, user_prompt: str, role="writer") -> List[Dict[str, str]]:
        from config.context import generation_messages
        return generation_messages(self, system_prompt, user_prompt, role)

    def snapshot(self) -> Dict[str, Any]:
        return {
            "topic": self.topic,
            "requirements": self.requirements,
            "messages": list(self.messages),
            "data": dict(self.data),
            "feedback": list(self.feedback),
            "status": self.status,
            "progress_log": list(self.progress_log),
            "last_updated": self.last_updated,
            "operation_mode": self.operation_mode,
            "modify_scope": self.modify_scope,
            "modify_page": self.modify_page,
            "modify_pages": list(self.modify_pages),
            "requirement_history": list(self.requirement_history),
            "history_snapshots": list(self.history_snapshots),
        }
