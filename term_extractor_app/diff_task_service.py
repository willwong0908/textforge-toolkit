"""Background task state for large Excel comparisons."""

from __future__ import annotations

import threading
import traceback
from dataclasses import dataclass, field
from time import time
from typing import Any, Callable
from uuid import uuid4

from .logging_utils import configure_file_logger


@dataclass
class DiffTask:
    task_id: str
    status: str = "running"
    message: str = "正在准备比对"
    result: dict[str, Any] | None = None
    error: str = ""
    traceback_text: str = ""
    created_at: float = field(default_factory=time)
    updated_at: float = field(default_factory=time)

    def to_dict(self) -> dict[str, Any]:
        return {
            "task_id": self.task_id,
            "status": self.status,
            "message": self.message,
            "result": self.result,
            "error": self.error,
            "created_at": self.created_at,
            "updated_at": self.updated_at,
        }


class DiffTaskService:
    def __init__(self) -> None:
        self._tasks: dict[str, DiffTask] = {}
        self._lock = threading.Lock()
        self._logger = configure_file_logger(with_console=True)

    def start(self, operation: Callable[[Callable[[str], None]], dict[str, Any]]) -> DiffTask:
        task = DiffTask(task_id=uuid4().hex)
        with self._lock:
            self._tasks[task.task_id] = task

        def update(message: str) -> None:
            with self._lock:
                task.message = str(message or "正在比对")
                task.updated_at = time()

        def run() -> None:
            try:
                self._logger.info("[DIFF] task=%s started", task.task_id)
                result = operation(update)
                with self._lock:
                    task.status = "completed"
                    task.message = "比对完成"
                    task.result = result
                    task.updated_at = time()
                self._logger.info(
                    "[DIFF] task=%s completed diffs=%s",
                    task.task_id,
                    int((result or {}).get("total_count", 0) or 0),
                )
            except Exception as exc:
                trace = traceback.format_exc()
                with self._lock:
                    task.status = "failed"
                    task.error = str(exc)
                    task.message = "比对失败"
                    task.traceback_text = trace
                    task.updated_at = time()
                self._logger.error("[DIFF] task=%s failed: %s\n%s", task.task_id, exc, trace)

        threading.Thread(target=run, name=f"diff-task-{task.task_id[:8]}", daemon=True).start()
        return task

    def get(self, task_id: str) -> DiffTask | None:
        with self._lock:
            return self._tasks.get(task_id)
