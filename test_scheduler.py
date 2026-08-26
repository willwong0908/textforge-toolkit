from __future__ import annotations

import asyncio
import logging
import sqlite3
import tempfile
import time
import unittest
from contextlib import contextmanager
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from term_extractor_app.ai_review import review_service
from term_extractor_app.models import LLMRequest, LLMResponse
from term_extractor_app.scheduler import AdaptiveConcurrencyController, AsyncRequestScheduler


def _request(task_id: str = "request-1") -> LLMRequest:
    return LLMRequest(task_id=task_id, task_type="review", prompt="test")


class _NeverEndingAdapter:
    provider_name = "test-provider"
    settings = SimpleNamespace(model="test-model")

    def __init__(self) -> None:
        self.calls = 0

    async def send_prompt(self, request: LLMRequest, attempt: int = 1) -> LLMResponse:
        self.calls += 1
        await asyncio.Event().wait()
        raise AssertionError("unreachable")


class _ImmediateAdapter:
    provider_name = "test-provider"
    settings = SimpleNamespace(model="test-model")

    async def send_prompt(self, request: LLMRequest, attempt: int = 1) -> LLMResponse:
        return LLMResponse(
            task_id=request.task_id,
            task_type=request.task_type,
            content="{}",
            provider=self.provider_name,
            model=self.settings.model,
            latency_ms=1,
            attempts=attempt,
            success=True,
        )


class SchedulerRecoveryTests(unittest.TestCase):
    def test_wall_clock_timeout_triggers_retry_and_finishes(self) -> None:
        adapter = _NeverEndingAdapter()
        retries: list[tuple[int, str]] = []
        scheduler = AsyncRequestScheduler(
            adapter=adapter,
            controller=AdaptiveConcurrencyController("固定", 1, 1),
            max_retries=1,
            stop_requested=lambda: False,
            attempt_timeout_seconds=0.1,
            on_retry=lambda request, response, next_attempt, backoff: retries.append(
                (next_attempt, response.error_type)
            ),
        )

        async def run_test():
            scheduler._sleep_with_cancellation = lambda seconds: asyncio.sleep(0)
            return await scheduler.run([_request()])

        started_at = time.monotonic()
        results = asyncio.run(run_test())
        self.assertLess(time.monotonic() - started_at, 1.0)
        self.assertEqual(adapter.calls, 2)
        self.assertEqual(retries, [(2, "hard_timeout")])
        self.assertFalse(results["request-1"].success)
        self.assertEqual(results["request-1"].error_type, "hard_timeout")
        self.assertEqual(results["request-1"].attempts, 2)

    def test_callback_failure_does_not_leave_queue_join_hanging(self) -> None:
        scheduler = AsyncRequestScheduler(
            adapter=_ImmediateAdapter(),
            controller=AdaptiveConcurrencyController("固定", 1, 1),
            max_retries=0,
            stop_requested=lambda: False,
        )

        def fail_callback(request, response, snapshot):
            raise ValueError("database write failed")

        with self.assertRaisesRegex(RuntimeError, "database write failed"):
            asyncio.run(scheduler.run([_request()], on_result=fail_callback))

    def test_ai_review_log_is_written_to_file_logger(self) -> None:
        class _Connection:
            def execute(self, *args, **kwargs):
                return None

        @contextmanager
        def fake_connection():
            yield _Connection()

        with tempfile.TemporaryDirectory() as temp_dir:
            log_path = Path(temp_dir) / "log.txt"
            logger = logging.getLogger("test-ai-review-file-log")
            logger.handlers.clear()
            logger.propagate = False
            logger.setLevel(logging.INFO)
            handler = logging.FileHandler(log_path, encoding="utf-8")
            handler.setFormatter(logging.Formatter("%(message)s"))
            logger.addHandler(handler)
            try:
                with patch.object(review_service, "get_connection", fake_connection), patch.object(
                    review_service.logging, "getLogger", return_value=logger
                ):
                    review_service._add_log("task-safe-id", "warning", "请求超时，准备重试")
                handler.flush()
                content = log_path.read_text(encoding="utf-8")
                self.assertIn("[AI_REVIEW][task-safe-id] 请求超时，准备重试", content)
            finally:
                handler.close()
                logger.handlers.clear()

    def test_startup_recovery_closes_orphaned_tasks_and_requests(self) -> None:
        connection = sqlite3.connect(":memory:")
        connection.row_factory = sqlite3.Row
        connection.execute(
            "CREATE TABLE review_tasks (id TEXT PRIMARY KEY, status TEXT, updated_at TEXT)"
        )
        connection.execute(
            "CREATE TABLE review_request_states (task_id TEXT, status TEXT, updated_at TEXT)"
        )
        connection.execute("INSERT INTO review_tasks VALUES ('stale', 'running', '')")
        connection.execute("INSERT INTO review_tasks VALUES ('done', 'completed', '')")
        connection.execute("INSERT INTO review_request_states VALUES ('stale', 'thinking', '')")

        @contextmanager
        def fake_connection():
            yield connection
            connection.commit()

        try:
            with patch.object(review_service, "get_connection", fake_connection), patch.object(
                review_service, "_add_log"
            ) as add_log:
                recovered = review_service.recover_interrupted_review_tasks()
            self.assertEqual(recovered, 1)
            self.assertEqual(
                connection.execute("SELECT status FROM review_tasks WHERE id = 'stale'").fetchone()[0],
                "failed",
            )
            self.assertEqual(
                connection.execute(
                    "SELECT status FROM review_request_states WHERE task_id = 'stale'"
                ).fetchone()[0],
                "failed",
            )
            self.assertEqual(
                connection.execute("SELECT status FROM review_tasks WHERE id = 'done'").fetchone()[0],
                "completed",
            )
            add_log.assert_called_once()
        finally:
            connection.close()


if __name__ == "__main__":
    unittest.main()
