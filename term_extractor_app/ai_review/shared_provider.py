from __future__ import annotations

import asyncio
import json
from typing import Any

import httpx

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
    configured_limit = int(ai_review_stage.get("batch_request_char_limit") or 0)
    if configured_limit <= 6000:
        configured_limit = 20000
    reasoning_effort = str(ai_review_stage.get("reasoning_effort") or "low").strip().lower()
    if reasoning_effort not in {"low", "high", "max"}:
        reasoning_effort = "low"
    return {
        "provider": provider_name,
        "api_key": provider.api_key,
        "selected_model": provider.model,
        "models": [],
        "max_concurrency": int(provider.max_concurrency or 6),
        "max_chars_per_request": configured_limit,
        "enable_thinking": True,
        "max_items_per_request": int(ai_review_stage.get("max_items_per_request") or 80),
        "workspace_enable_thinking": True,
        "reasoning_effort": reasoning_effort,
        "auto_start_after_inspection": bool(ai_review_stage.get("auto_start_after_inspection", False)),
        "debug_payload_logging": bool(ai_review_stage.get("debug_payload_logging", False)),
        "disable_system_proxy": bool(provider.disable_system_proxy),
        "timeout_seconds": int(provider.timeout_seconds or 90),
        "base_url": provider.base_url,
    }


def workspace_chat(messages: list[dict[str, str]], on_delta: Any | None = None) -> str:
    if on_delta is not None:
        return _workspace_chat_stream(messages, on_delta)

    async def _run() -> str:
        provider_name, provider = _load_provider_settings()
        settings = get_shared_ai_settings()
        if not provider.api_key:
            raise SharedProviderError("请先配置模型 API Key")
        if not provider.model:
            raise SharedProviderError("请先在模型设置中选择模型")
        adapter = ProviderRegistry.create_adapter(provider_name, provider)
        try:
            request = LLMRequest(
                task_id="ai-review-workspace",
                task_type="ai_review_workspace",
                prompt=messages[-1]["content"] if messages else "",
                messages=messages,
                metadata=_thinking_metadata(provider_name, provider, str(settings.get("reasoning_effort") or "low")),
            )
            response = await adapter.send_prompt(request)
        finally:
            await adapter.close()
        if not response.success:
            raise SharedProviderError(response.error or "Workspace Agent 请求失败")
        return str(response.content or "")

    return asyncio.run(_run())


def _workspace_chat_stream(messages: list[dict[str, str]], on_delta: Any) -> str:
    provider_name, provider = _load_provider_settings()
    settings = get_shared_ai_settings()
    if not provider.api_key:
        raise SharedProviderError("请先配置模型 API Key")
    if not provider.model:
        raise SharedProviderError("请先在模型设置中选择模型")
    headers = {"Content-Type": "application/json", "Authorization": f"Bearer {provider.api_key}"}
    headers.update(dict(provider.extra_headers or {}))
    payload: dict[str, Any] = {
        "model": provider.model,
        "messages": messages,
        "temperature": 0.0,
        "stream": True,
    }
    payload.update(_thinking_metadata(provider_name, provider, str(settings.get("reasoning_effort") or "low")))
    url = provider.base_url.rstrip("/") + "/chat/completions"
    chunks: list[str] = []
    reasoning_started = False
    try:
        with httpx.Client(
            timeout=float(max(10, provider.timeout_seconds)),
            verify=False,
            trust_env=not provider.disable_system_proxy,
        ) as client:
            with client.stream("POST", url, headers=headers, json=payload) as response:
                if response.status_code >= 400:
                    response.read()
                    raise SharedProviderError(
                        f"Workspace Agent 请求失败：HTTP {response.status_code} {response.text[:500]}"
                    )
                content_type = str(response.headers.get("content-type") or "").lower()
                if "text/event-stream" not in content_type:
                    response.read()
                    data = response.json()
                    content = _stream_response_content(data)
                    if content:
                        on_delta(content)
                        return content
                    raise SharedProviderError("Workspace Agent 返回内容为空")
                for line in response.iter_lines():
                    if not line or not line.startswith("data:"):
                        continue
                    raw = line[5:].strip()
                    if raw == "[DONE]":
                        break
                    try:
                        data = json.loads(raw)
                    except json.JSONDecodeError:
                        continue
                    reasoning = _stream_reasoning_content(data)
                    if reasoning:
                        prefix = "正在思考…\n" if not reasoning_started else ""
                        reasoning_started = True
                        _send_delta(on_delta, prefix + reasoning, "reasoning")
                    delta = _stream_delta_content(data)
                    if delta:
                        _send_delta(
                            on_delta,
                            ("\n正在生成识别方案…\n" if reasoning_started and not chunks else "") + delta,
                            "content",
                        )
                        chunks.append(delta)
    except httpx.HTTPError as exc:
        raise SharedProviderError(str(exc) or "Workspace Agent 网络请求失败") from exc
    content = "".join(chunks)
    if not content.strip():
        raise SharedProviderError("Workspace Agent 返回内容为空")
    return content


def _stream_delta_content(payload: dict[str, Any]) -> str:
    choices = payload.get("choices") or []
    if not choices or not isinstance(choices[0], dict):
        return ""
    delta = choices[0].get("delta") or {}
    if not isinstance(delta, dict):
        return ""
    content = delta.get("content") or ""
    if isinstance(content, list):
        return "".join(
            str(item.get("text") or "") for item in content
            if isinstance(item, dict) and item.get("type") == "text"
        )
    return str(content)


def _stream_reasoning_content(payload: dict[str, Any]) -> str:
    choices = payload.get("choices") or []
    if not choices or not isinstance(choices[0], dict):
        return ""
    delta = choices[0].get("delta") or {}
    if not isinstance(delta, dict):
        return ""
    value = delta.get("reasoning_content") or delta.get("reasoning") or ""
    return str(value)


def _stream_response_content(payload: dict[str, Any]) -> str:
    choices = payload.get("choices") or []
    if not choices or not isinstance(choices[0], dict):
        return ""
    message = choices[0].get("message") or {}
    return str(message.get("content") or "") if isinstance(message, dict) else ""


def _send_delta(callback: Any, value: str, phase: str) -> None:
    try:
        callback(value, phase)
    except TypeError:
        callback(value)


def _thinking_metadata(provider_name: str, provider: ProviderSettings, effort: str) -> dict[str, Any]:
    identity = f"{provider_name} {provider.base_url}".lower()
    if "deepseek" in identity:
        return {
            "thinking": {"type": "enabled"},
            "reasoning_effort": effort if effort in {"low", "high", "max"} else "low",
        }
    return {"enable_thinking": True}


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
            settings = get_shared_ai_settings()
            request = LLMRequest(
                task_id="ai-review-batch",
                task_type="candidate_review_batch",
                prompt=user_prompt,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt},
                ],
                metadata=_thinking_metadata(provider_name, provider, str(settings.get("reasoning_effort") or "low")),
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
