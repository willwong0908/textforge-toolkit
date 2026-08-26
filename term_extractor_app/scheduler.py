"""Adaptive async request scheduler."""

import asyncio
import contextlib
import inspect
import math
import random
import time
from typing import Awaitable, Callable, Dict, List, Optional

from .models import LLMRequest, LLMResponse, SchedulerSnapshot


ResultCallback = Callable[[LLMRequest, LLMResponse, SchedulerSnapshot], Optional[Awaitable[None]]]
ResponseValidator = Callable[[LLMRequest, LLMResponse], LLMResponse]
RetryCallback = Callable[[LLMRequest, LLMResponse, int, float], Optional[Awaitable[None]]]


class SchedulerCancelledError(Exception):
    """Raised when request scheduling is cancelled by the user."""


class AdaptiveConcurrencyController:
    def __init__(
        self,
        mode: str,
        user_max: int,
        provider_max: int,
        local_max: int = 8,
        start_concurrency: int = 2,
    ):
        effective_max = max(1, min(int(user_max or 1), int(provider_max or 1), int(local_max or 1)))
        self.mode = mode
        self.effective_max = effective_max
        self.current_concurrency = min(max(1, int(start_concurrency)), effective_max) if mode == "自动" else effective_max
        self.success_streak = 0
        self.success_window = 1

    def observe_success(self) -> None:
        if self.mode != "自动":
            return
        self.success_streak += 1
        if self.success_streak >= self.success_window and self.current_concurrency < self.effective_max:
            self.current_concurrency += 1
            self.success_streak = 0

    def observe_failure(self) -> None:
        if self.mode != "自动":
            return
        self.current_concurrency = max(1, int(math.ceil(self.current_concurrency / 2.0)))
        self.success_streak = 0


class AsyncRequestScheduler:
    def __init__(
        self,
        adapter,
        controller: AdaptiveConcurrencyController,
        max_retries: int,
        stop_requested: Callable[[], bool],
        response_validator: Optional[ResponseValidator] = None,
        attempt_timeout_seconds: Optional[float] = None,
        on_retry: Optional[RetryCallback] = None,
    ):
        self.adapter = adapter
        self.controller = controller
        self.max_retries = max(0, int(max_retries))
        self.stop_requested = stop_requested
        self.response_validator = response_validator
        self.attempt_timeout_seconds = (
            max(0.1, float(attempt_timeout_seconds)) if attempt_timeout_seconds is not None else None
        )
        self.on_retry = on_retry
        self.processed_count = 0
        self.success_count = 0
        self.failure_count = 0
        self.retry_count = 0

    async def run(
        self,
        requests: List[LLMRequest],
        on_result: Optional[ResultCallback] = None,
        on_request_started: Optional[Callable[[LLMRequest], Optional[Awaitable[None]]]] = None,
    ) -> Dict[str, LLMResponse]:
        results: Dict[str, LLMResponse] = {}
        if not requests:
            return results

        queue: asyncio.Queue = asyncio.Queue()
        for request in requests:
            await queue.put(request)

        worker_count = self.controller.effective_max

        fatal_errors: List[Exception] = []

        async def worker(worker_index: int) -> None:
            while True:
                if self.stop_requested():
                    self._drain_queue(queue)
                    return
                if worker_index >= self.controller.current_concurrency:
                    await asyncio.sleep(0.1)
                    continue
                try:
                    request = await asyncio.wait_for(queue.get(), timeout=0.2)
                except asyncio.TimeoutError:
                    if queue.empty():
                        return
                    continue

                try:
                    if on_request_started is not None:
                        started_result = on_request_started(request)
                        if inspect.isawaitable(started_result):
                            await started_result
                    try:
                        response = await self._execute_request(request)
                    except SchedulerCancelledError:
                        self._drain_queue(queue)
                        return
                    results[request.task_id] = response
                    self.processed_count += 1
                    if response.success:
                        self.success_count += 1
                    else:
                        self.failure_count += 1
                    snapshot = SchedulerSnapshot(
                        processed_count=self.processed_count,
                        total_count=len(requests),
                        current_concurrency=self.controller.current_concurrency,
                        success_count=self.success_count,
                        failure_count=self.failure_count,
                        retry_count=self.retry_count,
                    )
                    if on_result is not None:
                        callback_result = on_result(request, response, snapshot)
                        if inspect.isawaitable(callback_result):
                            await callback_result
                except SchedulerCancelledError:
                    self._drain_queue(queue)
                    return
                except Exception as exc:
                    # Never let a worker disappear while queue.join() is waiting.
                    # The caller receives the failure and can persist a terminal task state.
                    fatal_errors.append(exc)
                    self._drain_queue(queue)
                    return
                finally:
                    queue.task_done()

        workers = [asyncio.create_task(worker(index)) for index in range(worker_count)]
        await queue.join()
        for task in workers:
            task.cancel()
        await asyncio.gather(*workers, return_exceptions=True)
        if fatal_errors:
            raise RuntimeError(f"请求调度器内部异常：{fatal_errors[0]}") from fatal_errors[0]
        return results

    async def _execute_request(self, request: LLMRequest) -> LLMResponse:
        attempt = 1
        while True:
            try:
                response = await self._send_with_cancellation(request, attempt)
            except Exception as exc:  # pragma: no cover - defensive runtime safeguard
                response = LLMResponse(
                    task_id=request.task_id,
                    task_type=request.task_type,
                    content="",
                    provider=getattr(self.adapter, "provider_name", ""),
                    model=getattr(getattr(self.adapter, "settings", None), "model", ""),
                    latency_ms=0,
                    attempts=attempt,
                    success=False,
                    error=str(exc) or repr(exc),
                    error_type="adapter_exception",
                    retryable=True,
                )
            if self.response_validator is not None:
                response = self.response_validator(request, response)
            if response.success:
                self.controller.observe_success()
                return response

            self.controller.observe_failure()
            if attempt > self.max_retries or not response.retryable or self.stop_requested():
                return response

            self.retry_count += 1
            backoff = min(6.0, 0.8 * (2 ** (attempt - 1))) + random.uniform(0.0, 0.4)
            if self.on_retry is not None:
                callback_result = self.on_retry(request, response, attempt + 1, backoff)
                if inspect.isawaitable(callback_result):
                    await callback_result
            await self._sleep_with_cancellation(backoff)
            attempt += 1

    async def _send_with_cancellation(self, request: LLMRequest, attempt: int) -> LLMResponse:
        task = asyncio.create_task(self.adapter.send_prompt(request, attempt=attempt))
        started_at = time.monotonic()
        try:
            while not task.done():
                if self.stop_requested():
                    task.cancel()
                    with contextlib.suppress(asyncio.CancelledError):
                        await task
                    raise SchedulerCancelledError()
                if (
                    self.attempt_timeout_seconds is not None
                    and time.monotonic() - started_at >= self.attempt_timeout_seconds
                ):
                    task.cancel()
                    with contextlib.suppress(asyncio.CancelledError):
                        await task
                    timeout_seconds = self.attempt_timeout_seconds
                    return LLMResponse(
                        task_id=request.task_id,
                        task_type=request.task_type,
                        content="",
                        provider=getattr(self.adapter, "provider_name", ""),
                        model=getattr(getattr(self.adapter, "settings", None), "model", ""),
                        latency_ms=max(0, int((time.monotonic() - started_at) * 1000)),
                        attempts=attempt,
                        success=False,
                        error=f"单次请求超过 {timeout_seconds:g} 秒，已取消",
                        error_type="hard_timeout",
                        retryable=True,
                    )
                await asyncio.sleep(0.1)
            return await task
        except asyncio.CancelledError:
            task.cancel()
            raise

    async def _sleep_with_cancellation(self, seconds: float) -> None:
        remaining = max(0.0, float(seconds))
        while remaining > 0:
            if self.stop_requested():
                raise SchedulerCancelledError()
            interval = min(0.2, remaining)
            await asyncio.sleep(interval)
            remaining -= interval

    @staticmethod
    def _drain_queue(queue: asyncio.Queue) -> None:
        while True:
            try:
                queue.get_nowait()
            except asyncio.QueueEmpty:
                return
            else:
                queue.task_done()
