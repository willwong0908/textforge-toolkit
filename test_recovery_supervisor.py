from __future__ import annotations

import asyncio
import unittest
from types import SimpleNamespace
from unittest.mock import patch

from term_extractor_app.models import LLMRequest, LLMResponse, ProviderSettings
from term_extractor_app.ai_review import review_service
from term_extractor_app.recovery_supervisor import (
    TerminalFailureContext,
    is_ai_correctable_response_failure,
    is_terminal_ai_recovery_allowed,
    request_terminal_recovery,
    retry_llm_with_corrective_prompt,
)


class _RecoveryAdapter:
    def __init__(self, content: str) -> None:
        self.content = content
        self.calls: list[tuple[LLMRequest, int]] = []
        self.closed = False

    async def send_prompt(self, request: LLMRequest, attempt: int = 1) -> LLMResponse:
        self.calls.append((request, attempt))
        return LLMResponse(
            task_id=request.task_id,
            task_type=request.task_type,
            content=self.content,
            provider="test",
            model="test-model",
            latency_ms=1,
            attempts=attempt,
            success=True,
        )

    async def close(self) -> None:
        self.closed = True


class RecoverySupervisorTests(unittest.TestCase):
    def test_configuration_failures_do_not_recurse_into_ai(self) -> None:
        self.assertFalse(is_terminal_ai_recovery_allowed("auth", "API Key 无效"))
        self.assertTrue(is_terminal_ai_recovery_allowed("response_validation", "缺少 id 3"))
        self.assertTrue(is_ai_correctable_response_failure("response_validation"))
        self.assertFalse(is_ai_correctable_response_failure("hard_timeout"))

    def test_terminal_decision_accepts_only_allow_listed_action(self) -> None:
        context = TerminalFailureContext(
            workflow="text_preprocessing",
            task_id="task-1",
            phase="review",
            error_type="ValueError",
            message="schema mismatch",
            has_checkpoint=True,
        )
        adapter = _RecoveryAdapter('{"action":"resume_from_checkpoint","reason":"检查点完整"}')
        provider = ProviderSettings(api_key="key", base_url="https://example.invalid", model="model", timeout_seconds=10)
        with patch(
            "term_extractor_app.recovery_supervisor._load_active_provider",
            return_value=("test", provider),
        ), patch(
            "term_extractor_app.recovery_supervisor.ProviderRegistry.create_adapter",
            return_value=adapter,
        ):
            decision = asyncio.run(
                request_terminal_recovery(context, allowed_actions=("resume_from_checkpoint", "abort"))
            )
        self.assertTrue(decision.should_retry)
        self.assertEqual(decision.action, "resume_from_checkpoint")
        self.assertTrue(adapter.closed)

        invalid_adapter = _RecoveryAdapter('{"action":"delete_cache","reason":"重来"}')
        with patch(
            "term_extractor_app.recovery_supervisor._load_active_provider",
            return_value=("test", provider),
        ), patch(
            "term_extractor_app.recovery_supervisor.ProviderRegistry.create_adapter",
            return_value=invalid_adapter,
        ):
            decision = asyncio.run(request_terminal_recovery(context, allowed_actions=("abort",)))
        self.assertFalse(decision.should_retry)
        self.assertIn("未授权动作", decision.error)

    def test_corrective_request_contains_failure_and_runs_only_once(self) -> None:
        adapter = _RecoveryAdapter('{"items":[{"id":"1"}]}')
        request = LLMRequest(
            task_id="package-1",
            task_type="candidate_review_batch",
            prompt="original",
            messages=[{"role": "user", "content": "original"}],
            metadata={"package": [{"id": "1"}]},
        )
        failed = LLMResponse(
            task_id=request.task_id,
            task_type=request.task_type,
            content="bad",
            provider="test",
            model="model",
            latency_ms=1,
            attempts=3,
            success=False,
            error="missing id 1",
            error_type="response_validation",
            retryable=True,
        )

        def validator(recovery_request, response):
            self.assertTrue(recovery_request.metadata["terminal_ai_recovery"])
            self.assertIn("missing id 1", recovery_request.messages[-1]["content"])
            return response

        result = asyncio.run(
            retry_llm_with_corrective_prompt(
                adapter=adapter,
                request=request,
                failed_response=failed,
                response_validator=validator,
                timeout_seconds=10,
            )
        )
        self.assertTrue(result.success)
        self.assertEqual(result.attempts, 4)
        self.assertEqual(len(adapter.calls), 1)

    def test_review_task_terminal_recovery_is_bounded_to_one_retry(self) -> None:
        task = {"config": {}, "request_states": [], "target_language": "en"}
        with patch.object(
            review_service,
            "_run_review_task_impl",
            side_effect=[ValueError("first"), ValueError("second")],
        ) as run_impl, patch.object(
            review_service,
            "get_review_task",
            return_value=task,
        ), patch.object(
            review_service,
            "_attempt_review_task_terminal_recovery",
            return_value=True,
        ) as recovery, patch.object(
            review_service,
            "_mark_unfinished_review_requests_failed",
        ) as mark_failed, patch.object(
            review_service,
            "_update_task",
        ) as update_task, patch.object(
            review_service,
            "_track_ai_review_finish",
        ), patch.object(
            review_service,
            "_add_log",
        ):
            review_service._run_review_task("task-1")
        self.assertEqual(run_impl.call_count, 2)
        recovery.assert_called_once()
        mark_failed.assert_called_once_with("task-1")
        update_task.assert_called_once_with("task-1", status="failed")


if __name__ == "__main__":
    unittest.main()
