"""Bounded AI fallback used immediately before a workflow becomes terminal."""

from __future__ import annotations

import asyncio
import json
import re
from dataclasses import dataclass, field
from typing import Any, Callable, Iterable, Optional

from .models import LLMRequest, LLMResponse, ProviderSettings
from .providers import ProviderRegistry
from .storage import SettingsStore, get_app_paths


_NON_RECOVERABLE_ERROR_MARKERS = (
    "api key",
    "apikey",
    "authentication",
    "unauthorized",
    "invalid key",
    "余额不足",
    "额度不足",
    "鉴权",
    "未配置模型",
    "模型名称为空",
)

_AI_CORRECTABLE_RESPONSE_ERRORS = {
    "response_validation",
    "invalid_json",
    "invalid_json_response_format",
    "semantic_validation",
    "empty_response",
}


@dataclass(frozen=True)
class TerminalFailureContext:
    workflow: str
    task_id: str
    phase: str
    error_type: str
    message: str
    has_checkpoint: bool
    completed_units: int = 0
    remaining_units: int = 0
    metadata: dict[str, Any] = field(default_factory=dict)


@dataclass(frozen=True)
class TerminalRecoveryDecision:
    attempted: bool
    action: str = "abort"
    reason: str = ""
    error: str = ""

    @property
    def should_retry(self) -> bool:
        return self.attempted and self.action != "abort" and not self.error


def is_terminal_ai_recovery_allowed(error_type: str, message: str) -> bool:
    """Do not recursively ask the model to repair invalid model credentials."""

    combined = f"{error_type} {message}".strip().lower()
    return not any(marker in combined for marker in _NON_RECOVERABLE_ERROR_MARKERS)


def is_ai_correctable_response_failure(error_type: str) -> bool:
    """Only content/validation failures benefit from another corrective model call."""

    return str(error_type or "").strip().lower() in _AI_CORRECTABLE_RESPONSE_ERRORS


def _clip(value: object, limit: int = 4000) -> str:
    text = str(value or "").strip()
    if len(text) <= limit:
        return text
    return text[:limit] + "\n...[truncated]"


def _compact_metadata(value: object, depth: int = 0) -> object:
    if depth >= 3:
        return _clip(value, 500)
    if isinstance(value, dict):
        return {
            _clip(key, 80): _compact_metadata(item, depth + 1)
            for key, item in list(value.items())[:20]
        }
    if isinstance(value, (list, tuple)):
        return [_compact_metadata(item, depth + 1) for item in list(value)[:10]]
    if isinstance(value, str):
        return _clip(value, 500)
    if value is None or isinstance(value, (int, float, bool)):
        return value
    return _clip(value, 500)


def _parse_json_object(content: str) -> dict[str, Any]:
    text = str(content or "").strip()
    fenced = re.search(r"```(?:json)?\s*(\{.*?\})\s*```", text, flags=re.IGNORECASE | re.DOTALL)
    if fenced:
        text = fenced.group(1)
    else:
        start = text.find("{")
        end = text.rfind("}")
        if start >= 0 and end > start:
            text = text[start : end + 1]
    parsed = json.loads(text)
    if not isinstance(parsed, dict):
        raise ValueError("Recovery Agent 未返回 JSON 对象。")
    return parsed


def _load_active_provider() -> tuple[str, ProviderSettings]:
    settings = SettingsStore(get_app_paths()).load()
    provider_name = str(settings.provider_name or "DeepSeek")
    provider = settings.provider_settings.get(provider_name)
    if provider is None:
        raise ValueError("当前没有可用的模型配置。")
    provider_copy = ProviderSettings.from_dict(provider.to_dict())
    if not provider_copy.api_key.strip():
        raise ValueError("当前模型 API Key 为空。")
    if not provider_copy.model.strip():
        raise ValueError("当前模型名称为空。")
    return provider_name, provider_copy


async def request_terminal_recovery(
    context: TerminalFailureContext,
    *,
    allowed_actions: Iterable[str],
) -> TerminalRecoveryDecision:
    """Ask the configured model for one allow-listed recovery action."""

    actions = tuple(dict.fromkeys(str(item or "").strip() for item in allowed_actions if str(item or "").strip()))
    if not actions:
        return TerminalRecoveryDecision(attempted=False, error="没有可执行的恢复动作。")
    if not is_terminal_ai_recovery_allowed(context.error_type, context.message):
        return TerminalRecoveryDecision(attempted=False, error="该错误属于模型配置或鉴权问题，跳过 AI 兜底。")

    try:
        provider_name, provider = _load_active_provider()
    except Exception as exc:
        return TerminalRecoveryDecision(attempted=False, error=str(exc))

    payload = {
        "workflow": context.workflow,
        "task_id": context.task_id,
        "phase": context.phase,
        "error_type": context.error_type,
        "error_message": _clip(context.message),
        "has_checkpoint": bool(context.has_checkpoint),
        "completed_units": max(0, int(context.completed_units)),
        "remaining_units": max(0, int(context.remaining_units)),
        "available_actions": list(actions),
        "metadata": _compact_metadata(dict(context.metadata or {})),
    }
    system_prompt = (
        "你是长工作流的最终恢复决策器。常规重试已经耗尽，任务即将终止。"
        "只能从 available_actions 中选择一个动作，不得建议执行代码、Shell、删除文件、修改配置或无限重试。"
        "优先保护已经完成的结果和检查点。只返回严格 JSON："
        '{"action":"动作名","reason":"不超过120字的原因"}。'
    )
    user_prompt = "请根据以下结构化故障决定最后一次恢复动作：\n" + json.dumps(payload, ensure_ascii=False)
    request = LLMRequest(
        task_id=f"terminal-recovery-{context.task_id}",
        task_type="workflow_terminal_recovery",
        prompt=user_prompt,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        metadata={
            "enable_thinking": True,
            "reasoning_effort": "low",
            "force_prompt_only_json": True,
        },
    )
    adapter = ProviderRegistry.create_adapter(provider_name, provider)
    timeout_seconds = min(60, max(10, int(provider.timeout_seconds or 90)))
    try:
        response = await asyncio.wait_for(adapter.send_prompt(request, attempt=1), timeout=timeout_seconds)
    except asyncio.TimeoutError:
        return TerminalRecoveryDecision(attempted=True, error=f"Recovery Agent 超过 {timeout_seconds} 秒未返回。")
    except Exception as exc:
        return TerminalRecoveryDecision(attempted=True, error=f"Recovery Agent 请求失败：{exc}")
    finally:
        await adapter.close()

    if not response.success:
        return TerminalRecoveryDecision(attempted=True, error=response.error or "Recovery Agent 请求失败。")
    try:
        parsed = _parse_json_object(response.content)
    except Exception as exc:
        return TerminalRecoveryDecision(attempted=True, error=f"Recovery Agent 输出无效：{exc}")
    action = str(parsed.get("action") or "").strip()
    reason = _clip(parsed.get("reason"), 240)
    if action not in actions:
        return TerminalRecoveryDecision(attempted=True, error=f"Recovery Agent 返回了未授权动作：{action or '空'}")
    return TerminalRecoveryDecision(attempted=True, action=action, reason=reason)


def request_terminal_recovery_sync(
    context: TerminalFailureContext,
    *,
    allowed_actions: Iterable[str],
) -> TerminalRecoveryDecision:
    return asyncio.run(request_terminal_recovery(context, allowed_actions=allowed_actions))


async def retry_llm_with_corrective_prompt(
    *,
    adapter: Any,
    request: LLMRequest,
    failed_response: LLMResponse,
    response_validator: Optional[Callable[[LLMRequest, LLMResponse], LLMResponse]] = None,
    timeout_seconds: int,
) -> LLMResponse:
    """Make one final model call with the concrete validation/request failure attached."""

    if not is_terminal_ai_recovery_allowed(failed_response.error_type, failed_response.error):
        return failed_response
    expected_ids = [
        str(item.get("id") or "")
        for item in list(request.metadata.get("package") or [])
        if isinstance(item, dict) and str(item.get("id") or "").strip()
    ]
    corrective_prompt = (
        "常规请求与重试均未得到可接受结果。请作为最终纠错 Agent 重新完成原任务。\n"
        f"失败类型：{_clip(failed_response.error_type, 120)}\n"
        f"校验或请求错误：{_clip(failed_response.error, 1200)}\n"
        f"必须完整返回的 ID：{json.dumps(expected_ids, ensure_ascii=False)}\n"
        "不要解释失败原因，不要省略条目，只输出原任务要求的严格 JSON。"
    )
    metadata = dict(request.metadata or {})
    metadata.update(
        {
            "terminal_ai_recovery": True,
            "enable_thinking": True,
            "reasoning_effort": "low",
            "force_prompt_only_json": True,
        }
    )
    recovery_request = LLMRequest(
        task_id=request.task_id,
        task_type=request.task_type,
        prompt=request.prompt + "\n\n" + corrective_prompt,
        messages=list(request.messages or []) + [{"role": "user", "content": corrective_prompt}],
        metadata=metadata,
    )
    attempt = max(1, int(failed_response.attempts or 1)) + 1
    try:
        response = await asyncio.wait_for(
            adapter.send_prompt(recovery_request, attempt=attempt),
            timeout=max(10, int(timeout_seconds)),
        )
    except asyncio.TimeoutError:
        return LLMResponse(
            task_id=request.task_id,
            task_type=request.task_type,
            content="",
            provider=failed_response.provider,
            model=failed_response.model,
            latency_ms=0,
            attempts=attempt,
            success=False,
            error=f"最终 AI 兜底超过 {max(10, int(timeout_seconds))} 秒未返回。",
            error_type="terminal_recovery_timeout",
            retryable=False,
        )
    except Exception as exc:
        return LLMResponse(
            task_id=request.task_id,
            task_type=request.task_type,
            content="",
            provider=failed_response.provider,
            model=failed_response.model,
            latency_ms=0,
            attempts=attempt,
            success=False,
            error=f"最终 AI 兜底请求失败：{exc}",
            error_type="terminal_recovery_failed",
            retryable=False,
        )
    response.attempts = attempt
    response.retryable = False
    if response_validator is not None:
        response = response_validator(recovery_request, response)
        response.attempts = attempt
        response.retryable = False
    return response
