from __future__ import annotations

import asyncio
from typing import Any

from ..models import LLMRequest, ProviderSettings
from ..providers import ProviderRegistry
from ..storage import SettingsStore, get_app_paths


class SharedProviderError(Exception):
    pass


def _load_provider_settings(api_key_override: str | None = None) -> tuple[str, ProviderSettings]:
    settings = SettingsStore(get_app_paths()).load()
    provider_name = str(settings.provider_name or "DeepSeek")
    provider = settings.provider_settings.get(provider_name)
    if provider is None:
        raise SharedProviderError("当前没有可用的模型配置。")
    provider_copy = ProviderSettings.from_dict(provider.to_dict())
    if api_key_override is not None:
        provider_copy.api_key = str(api_key_override or "").strip()
    return provider_name, provider_copy


def get_shared_ai_settings() -> dict[str, Any]:
    provider_name, provider = _load_provider_settings()
    settings = SettingsStore(get_app_paths()).load()
    ai_review_stage = dict(settings.input_defaults.get("ai_review_stage_settings", {}) or {})
    return {
        "provider": provider_name,
        "api_key": provider.api_key,
        "selected_model": provider.model,
        "models": [],
        "max_concurrency": int(provider.max_concurrency or 6),
        "max_chars_per_request": int(ai_review_stage.get("batch_request_char_limit") or 3000),
        "enable_thinking": bool(ai_review_stage.get("enable_thinking", False)),
        "disable_system_proxy": bool(provider.disable_system_proxy),
        "timeout_seconds": int(provider.timeout_seconds or 90),
        "base_url": provider.base_url,
    }


def list_models(api_key: str) -> list[str]:
    async def _run() -> list[str]:
        provider_name, provider = _load_provider_settings(api_key_override=api_key)
        adapter = ProviderRegistry.create_adapter(provider_name, provider)
        try:
            ok, message, model_names = await adapter.list_models()
        finally:
            await adapter.close()
        if not ok:
            raise SharedProviderError(message)
        return model_names

    return asyncio.run(_run())


def test_chat(api_key: str, model: str, enable_thinking: bool = False) -> str:
    async def _run() -> str:
        provider_name, provider = _load_provider_settings(api_key_override=api_key)
        provider.model = model
        adapter = ProviderRegistry.create_adapter(provider_name, provider)
        try:
            request = LLMRequest(
                task_id="ai-review-test",
                task_type="test",
                prompt="Return exactly: OK",
                messages=[
                    {"role": "system", "content": "You are a concise API health checker."},
                    {"role": "user", "content": "Return exactly: OK"},
                ],
                metadata={"enable_thinking": bool(enable_thinking)},
            )
            response = await adapter.send_prompt(request)
        finally:
            await adapter.close()
        if not response.success:
            raise SharedProviderError(response.error or "模型连接测试失败")
        return str(response.content or "").strip()

    return asyncio.run(_run())


def review_chat(
    api_key: str,
    model: str,
    system_prompt: str,
    user_prompt: str,
    enable_thinking: bool = False,
) -> str:
    async def _run() -> str:
        provider_name, provider = _load_provider_settings(api_key_override=api_key)
        provider.model = model
        adapter = ProviderRegistry.create_adapter(provider_name, provider)
        try:
            request = LLMRequest(
                task_id="ai-review-batch",
                task_type="candidate_review_batch",
                prompt=user_prompt,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt},
                ],
                metadata={"enable_thinking": bool(enable_thinking)},
            )
            response = await adapter.send_prompt(request)
        finally:
            await adapter.close()
        if not response.success:
            raise SharedProviderError(response.error or "审校请求失败")
        return str(response.content or "")

    return asyncio.run(_run())


def followup_chat(
    task_id: str,
    messages: list[dict[str, str]],
    enable_thinking: bool = False,
) -> str:
    async def _run() -> str:
        provider_name, provider = _load_provider_settings()
        if not provider.api_key:
            raise SharedProviderError("请先加载 DeepSeek API Key")
        if not provider.model:
            raise SharedProviderError("请先选择 DeepSeek 模型")
        adapter = ProviderRegistry.create_adapter(provider_name, provider)
        try:
            request = LLMRequest(
                task_id=task_id,
                task_type="ai_review_followup_chat",
                prompt=messages[-1]["content"] if messages else "",
                messages=messages,
                metadata={"enable_thinking": bool(enable_thinking)},
            )
            response = await adapter.send_prompt(request)
        finally:
            await adapter.close()
        if not response.success:
            raise SharedProviderError(response.error or "追问请求失败")
        return str(response.content or "")

    return asyncio.run(_run())
