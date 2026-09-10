"""Provider registry and adapters."""

from __future__ import annotations

import json
import time
from typing import Dict, List, Optional, Tuple

import httpx

from .constants import PROVIDER_PRESETS
from .models import LLMRequest, LLMResponse, ProviderSettings


def _join_url(base_url: str, suffix: str) -> str:
    return base_url.rstrip("/") + "/" + suffix.lstrip("/")


class ProviderRegistry:
    @staticmethod
    def names():
        return list(PROVIDER_PRESETS.keys())

    @staticmethod
    def get_preset(provider_name: str) -> Dict[str, object]:
        return dict(PROVIDER_PRESETS[provider_name])

    @staticmethod
    def create_adapter(provider_name: str, settings: ProviderSettings) -> "OpenAICompatibleAdapter":
        return OpenAICompatibleAdapter(provider_name, settings)


class OpenAICompatibleAdapter:
    def __init__(self, provider_name: str, settings: ProviderSettings):
        self.provider_name = provider_name
        self.settings = settings
        self.client = self._build_client(trust_env=not settings.disable_system_proxy)

    def _build_client(self, trust_env: bool) -> httpx.AsyncClient:
        return httpx.AsyncClient(
            timeout=float(max(10, self.settings.timeout_seconds)),
            verify=False,
            trust_env=trust_env,
        )

    async def close(self) -> None:
        await self.client.aclose()

    @staticmethod
    def _describe_http_error(exc: Exception) -> str:
        message = str(exc).strip()
        if message:
            return message
        cause = getattr(exc, "__cause__", None)
        if cause:
            cause_message = str(cause).strip()
            if cause_message:
                return cause_message
            return repr(cause)
        return repr(exc)

    @staticmethod
    def _json_response_format() -> Dict[str, str]:
        return {"type": "json_object"}

    @staticmethod
    def _is_json_batch_task(request: Optional[LLMRequest]) -> bool:
        if request is None:
            return False
        return request.task_type in {
            "candidate_recall_batch",
            "candidate_review_batch",
            "nontrans_discovery_batch",
            "nontrans_regex_generation_batch",
            "nontrans_regex_reorder",
        }

    @staticmethod
    def _request_forces_prompt_only_json(request: Optional[LLMRequest]) -> bool:
        if request is None:
            return False
        return bool(dict(getattr(request, "metadata", {}) or {}).get("force_prompt_only_json", False))

    @classmethod
    def _supports_json_response_format(cls, request: Optional[LLMRequest]) -> bool:
        return cls._is_json_batch_task(request) and not cls._request_forces_prompt_only_json(request)

    @classmethod
    def _initial_json_mode(cls, request: Optional[LLMRequest]) -> str:
        if not cls._is_json_batch_task(request):
            return "prompt_only"
        if cls._request_forces_prompt_only_json(request):
            return "prompt_only_recovery"
        return "response_format"

    @staticmethod
    def _looks_like_response_format_unsupported(message: str) -> bool:
        normalized = str(message or "").lower()
        markers = (
            "response_format",
            "json_object",
            "unsupported",
            "not support",
            "invalid parameter",
            "unknown parameter",
            "unrecognized request argument",
        )
        return any(marker in normalized for marker in markers)

    def _headers(self) -> Dict[str, str]:
        headers = {"Content-Type": "application/json"}
        if self.settings.api_key:
            headers["Authorization"] = "Bearer {0}".format(self.settings.api_key)
        if self.settings.extra_headers:
            headers.update(self.settings.extra_headers)
        return headers

    def _chat_url(self) -> str:
        return _join_url(self.settings.base_url or "", "chat/completions")

    def _models_url(self) -> str:
        return _join_url(self.settings.base_url or "", "models")

    def _normalize_error(self, status_code: Optional[int], message: str) -> Tuple[str, bool]:
        normalized = message or "未知错误"
        if status_code in (400, 404):
            return normalized, False
        if status_code in (401, 403):
            return normalized, False
        if status_code == 429:
            return normalized, True
        if status_code and status_code >= 500:
            return normalized, True
        return normalized, True

    def _parse_content(self, payload: Dict[str, object]) -> str:
        choices = payload.get("choices", [])
        if not choices:
            return ""
        first_choice = choices[0]
        if not isinstance(first_choice, dict):
            return ""
        message = first_choice.get("message", {})
        if isinstance(message, dict):
            content = message.get("content", "")
            if isinstance(content, list):
                texts = []
                for item in content:
                    if isinstance(item, dict) and item.get("type") == "text":
                        texts.append(str(item.get("text", "")))
                return "\n".join(texts).strip()
            return str(content or "")
        return ""

    async def send_prompt(self, request: LLMRequest, attempt: int = 1) -> LLMResponse:
        phase_callback = dict(getattr(request, "metadata", {}) or {}).get("stream_phase_callback")
        if callable(phase_callback):
            return await self._send_prompt_streaming(request, attempt, phase_callback)
        start = time.perf_counter()
        wants_json_mode = self._supports_json_response_format(request)
        json_mode = self._initial_json_mode(request)

        try:
            response, json_mode = await self._post_with_json_mode_fallback(
                self.client,
                request,
                wants_json_mode=wants_json_mode,
                current_json_mode=json_mode,
            )
            latency_ms = int((time.perf_counter() - start) * 1000)
            if response.status_code >= 400:
                error_text = self._extract_error_text(response)
                normalized, retryable = self._normalize_error(response.status_code, error_text)
                return LLMResponse(
                    task_id=request.task_id,
                    task_type=request.task_type,
                    content="",
                    provider=self.provider_name,
                    model=self.settings.model,
                    latency_ms=latency_ms,
                    attempts=attempt,
                    success=False,
                    error=normalized,
                    error_type="http_{0}".format(response.status_code),
                    retryable=retryable,
                    response_metadata={"json_mode": json_mode},
                )

            payload = response.json()
            content = self._parse_content(payload)
            if not content:
                return LLMResponse(
                    task_id=request.task_id,
                    task_type=request.task_type,
                    content="",
                    provider=self.provider_name,
                    model=self.settings.model,
                    latency_ms=latency_ms,
                    attempts=attempt,
                    success=False,
                    error="模型返回空内容",
                    error_type="empty_response",
                    retryable=True,
                    response_metadata={"json_mode": json_mode},
                )

            response_metadata = {"json_mode": json_mode}
            usage = payload.get("usage")
            if isinstance(usage, dict):
                response_metadata["usage"] = usage
            return LLMResponse(
                task_id=request.task_id,
                task_type=request.task_type,
                content=content,
                provider=self.provider_name,
                model=self.settings.model,
                latency_ms=latency_ms,
                attempts=attempt,
                success=True,
                response_metadata=response_metadata,
            )
        except httpx.TimeoutException as exc:
            latency_ms = int((time.perf_counter() - start) * 1000)
            return LLMResponse(
                task_id=request.task_id,
                task_type=request.task_type,
                content="",
                provider=self.provider_name,
                model=self.settings.model,
                latency_ms=latency_ms,
                attempts=attempt,
                success=False,
                error=str(exc) or "请求超时",
                error_type="timeout",
                retryable=True,
                response_metadata={"json_mode": json_mode},
            )
        except httpx.HTTPError as exc:
            if not self.settings.disable_system_proxy:
                fallback_start = time.perf_counter()
                try:
                    fallback_client = self._build_client(trust_env=False)
                    try:
                        response, json_mode = await self._post_with_json_mode_fallback(
                            fallback_client,
                            request,
                            wants_json_mode=wants_json_mode,
                            current_json_mode=json_mode,
                        )
                    finally:
                        await fallback_client.aclose()
                    latency_ms = int((time.perf_counter() - fallback_start) * 1000)
                    if response.status_code < 400:
                        payload = response.json()
                        content = self._parse_content(payload)
                        if content:
                            return LLMResponse(
                                task_id=request.task_id,
                                task_type=request.task_type,
                                content=content,
                                provider=self.provider_name,
                                model=self.settings.model,
                                latency_ms=latency_ms,
                                attempts=attempt,
                                success=True,
                                response_metadata={"json_mode": json_mode},
                            )
                    else:
                        fallback_error, retryable = self._normalize_error(
                            response.status_code, self._extract_error_text(response)
                        )
                        return LLMResponse(
                            task_id=request.task_id,
                            task_type=request.task_type,
                            content="",
                            provider=self.provider_name,
                            model=self.settings.model,
                            latency_ms=latency_ms,
                            attempts=attempt,
                            success=False,
                            error="系统代理请求失败，直连后仍然失败：{0}".format(fallback_error),
                            error_type="proxy_then_http_{0}".format(response.status_code),
                            retryable=retryable,
                            response_metadata={"json_mode": json_mode},
                        )
                except httpx.HTTPError as fallback_exc:
                    latency_ms = int((time.perf_counter() - start) * 1000)
                    return LLMResponse(
                        task_id=request.task_id,
                        task_type=request.task_type,
                        content="",
                        provider=self.provider_name,
                        model=self.settings.model,
                        latency_ms=latency_ms,
                        attempts=attempt,
                        success=False,
                        error="系统代理请求失败，直连也失败：{0}".format(self._describe_http_error(fallback_exc)),
                        error_type="network",
                        retryable=True,
                        response_metadata={"json_mode": json_mode},
                    )
            latency_ms = int((time.perf_counter() - start) * 1000)
            return LLMResponse(
                task_id=request.task_id,
                task_type=request.task_type,
                content="",
                provider=self.provider_name,
                model=self.settings.model,
                latency_ms=latency_ms,
                attempts=attempt,
                success=False,
                error=self._describe_http_error(exc),
                error_type="network",
                retryable=True,
                response_metadata={"json_mode": json_mode},
            )

    async def _send_prompt_streaming(self, request: LLMRequest, attempt: int, phase_callback) -> LLMResponse:
        """Stream DeepSeek-compatible deltas so the review UI can expose the real model phase."""
        start = time.perf_counter()
        messages = list(getattr(request, "messages", []) or []) or [{"role": "user", "content": request.prompt}]
        metadata = dict(getattr(request, "metadata", {}) or {})
        provider_identity = f"{self.provider_name} {self.settings.base_url}".lower()
        is_deepseek = "deepseek" in provider_identity
        payload: Dict[str, object] = {
            "model": self.settings.model,
            "messages": messages,
            "temperature": 0.0 if self._is_json_batch_task(request) else 0.2,
            "stream": True,
        }
        if "enable_thinking" in metadata:
            enabled = bool(metadata.get("enable_thinking"))
            payload["thinking" if is_deepseek else "enable_thinking"] = (
                {"type": "enabled" if enabled else "disabled"} if is_deepseek else enabled
            )
        if "thinking" in metadata:
            payload["thinking"] = metadata.get("thinking")
        if is_deepseek and "reasoning_effort" in metadata:
            payload["reasoning_effort"] = str(metadata.get("reasoning_effort") or "low")
        if self._supports_json_response_format(request):
            payload["response_format"] = self._json_response_format()

        content_chunks: List[str] = []
        current_phase = ""

        def report_phase(phase: str) -> None:
            nonlocal current_phase
            if phase == current_phase:
                return
            current_phase = phase
            phase_callback(phase)

        try:
            async with self.client.stream("POST", self._chat_url(), headers=self._headers(), json=payload) as response:
                if response.status_code >= 400:
                    raw = await response.aread()
                    message = raw.decode("utf-8", errors="replace")
                    normalized, retryable = self._normalize_error(response.status_code, message)
                    return LLMResponse(
                        task_id=request.task_id, task_type=request.task_type, content="", provider=self.provider_name,
                        model=self.settings.model, latency_ms=int((time.perf_counter() - start) * 1000), attempts=attempt,
                        success=False, error=normalized, error_type=f"http_{response.status_code}", retryable=retryable,
                        response_metadata={"json_mode": "stream"},
                    )
                async for line in response.aiter_lines():
                    if not line or not line.startswith("data:"):
                        continue
                    raw = line[5:].strip()
                    if raw == "[DONE]":
                        break
                    try:
                        data = json.loads(raw)
                    except json.JSONDecodeError:
                        continue
                    choices = data.get("choices") or []
                    delta = choices[0].get("delta") if choices and isinstance(choices[0], dict) else {}
                    if not isinstance(delta, dict):
                        continue
                    if delta.get("reasoning_content") or delta.get("reasoning"):
                        report_phase("thinking")
                    content = delta.get("content") or ""
                    if isinstance(content, list):
                        content = "".join(
                            str(item.get("text") or "") for item in content
                            if isinstance(item, dict) and item.get("type") == "text"
                        )
                    if content:
                        report_phase("output")
                        content_chunks.append(str(content))
        except httpx.TimeoutException as exc:
            return LLMResponse(
                task_id=request.task_id, task_type=request.task_type, content="", provider=self.provider_name,
                model=self.settings.model, latency_ms=int((time.perf_counter() - start) * 1000), attempts=attempt,
                success=False, error=str(exc) or "请求超时", error_type="timeout", retryable=True,
                response_metadata={"json_mode": "stream"},
            )
        except httpx.HTTPError as exc:
            return LLMResponse(
                task_id=request.task_id, task_type=request.task_type, content="", provider=self.provider_name,
                model=self.settings.model, latency_ms=int((time.perf_counter() - start) * 1000), attempts=attempt,
                success=False, error=self._describe_http_error(exc), error_type="network", retryable=True,
                response_metadata={"json_mode": "stream"},
            )
        content = "".join(content_chunks)
        if not content.strip():
            return LLMResponse(
                task_id=request.task_id, task_type=request.task_type, content="", provider=self.provider_name,
                model=self.settings.model, latency_ms=int((time.perf_counter() - start) * 1000), attempts=attempt,
                success=False, error="模型返回空内容", error_type="empty_response", retryable=True,
                response_metadata={"json_mode": "stream"},
            )
        return LLMResponse(
            task_id=request.task_id, task_type=request.task_type, content=content, provider=self.provider_name,
            model=self.settings.model, latency_ms=int((time.perf_counter() - start) * 1000), attempts=attempt,
            success=True, response_metadata={"json_mode": "stream"},
        )

    async def _post_with_json_mode_fallback(
        self,
        client: httpx.AsyncClient,
        request: LLMRequest,
        wants_json_mode: bool,
        current_json_mode: str,
    ) -> Tuple[httpx.Response, str]:
        response = await self._post_once(
            client,
            request.prompt,
            request=request,
            use_json_response_format=wants_json_mode,
        )
        if not wants_json_mode or response.status_code < 400:
            return response, current_json_mode

        error_text = self._extract_error_text(response)
        if not self._looks_like_response_format_unsupported(error_text):
            return response, current_json_mode

        fallback_response = await self._post_once(
            client,
            request.prompt,
            request=request,
            use_json_response_format=False,
        )
        return fallback_response, "prompt_only_fallback"

    def _extract_error_text(self, response: httpx.Response) -> str:
        error_text = response.text
        try:
            data = response.json()
            if isinstance(data, dict):
                if isinstance(data.get("error"), dict):
                    error_text = json.dumps(data.get("error", {}), ensure_ascii=False)
                elif data.get("error"):
                    error_text = str(data.get("error"))
        except Exception:
            pass
        return error_text

    async def _post_once(
        self,
        client: httpx.AsyncClient,
        prompt: str,
        request: Optional[LLMRequest] = None,
        use_json_response_format: bool = False,
    ) -> httpx.Response:
        messages = list(getattr(request, "messages", []) or [])
        if not messages:
            messages = [{"role": "user", "content": prompt}]
        payload = {
            "model": self.settings.model,
            "messages": messages,
            "temperature": 0.0 if self._is_json_batch_task(request) else 0.2,
        }
        metadata = dict(getattr(request, "metadata", {}) or {})
        provider_identity = f"{self.provider_name} {self.settings.base_url}".lower()
        is_deepseek = "deepseek" in provider_identity
        if "enable_thinking" in metadata:
            enabled = bool(metadata.get("enable_thinking"))
            if is_deepseek:
                payload["thinking"] = {"type": "enabled" if enabled else "disabled"}
            else:
                payload["enable_thinking"] = enabled
        if "thinking" in metadata:
            payload["thinking"] = metadata.get("thinking")
        if is_deepseek and "reasoning_effort" in metadata:
            payload["reasoning_effort"] = str(metadata.get("reasoning_effort") or "low")
        if use_json_response_format and self._is_json_batch_task(request):
            payload["response_format"] = self._json_response_format()
        return await client.post(self._chat_url(), headers=self._headers(), json=payload)

    async def test_connection(self) -> LLMResponse:
        request = LLMRequest(
            task_id="connection-test",
            task_type="test",
            prompt="请只回复：API连接测试成功",
        )
        return await self.send_prompt(request, attempt=1)

    async def list_models(self) -> Tuple[bool, str, list[str]]:
        try:
            response = await self.client.get(self._models_url(), headers=self._headers())
        except httpx.TimeoutException as exc:
            return False, str(exc) or "加载模型超时", []
        except httpx.HTTPError as exc:
            return False, self._describe_http_error(exc), []

        if response.status_code >= 400:
            return False, self._extract_error_text(response), []

        try:
            payload = response.json()
        except Exception as exc:
            return False, "模型列表响应不是合法 JSON：{0}".format(exc), []

        raw_items = payload.get("data", []) if isinstance(payload, dict) else []
        if not isinstance(raw_items, list):
            return False, "模型列表响应缺少 data 数组", []

        model_names = []
        for item in raw_items:
            if not isinstance(item, dict):
                continue
            model_id = str(item.get("id", "")).strip()
            if model_id:
                model_names.append(model_id)

        model_names = sorted(set(model_names))
        if not model_names:
            return False, "当前接口没有返回可用模型", []
        return True, "已加载 {0} 个模型".format(len(model_names)), model_names
