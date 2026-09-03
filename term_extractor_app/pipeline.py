"""End-to-end term extraction pipeline."""

from __future__ import annotations

import logging
import json
import re
import uuid
from collections import defaultdict
from pathlib import Path
from typing import Callable, Dict, List, Optional

from .constants import STAGE_LABELS
from .core import (
    aggregate_reviewed_terms,
    allowed_term_types_from_recall_scopes,
    build_candidate_recall_batches,
    build_candidate_terms_from_chunk_recall,
    build_chunk_term_recall_batches,
    build_failure_guidance,
    build_review_retry_batches,
    build_source_location,
    build_term_recall_clean_records,
    candidate_has_traceable_evidence,
    candidate_terms_from_runtime,
    format_enabled_recall_scopes,
    format_recall_retry_feedback,
    format_review_retry_feedback,
    has_meaningful_clean_text,
    merge_candidate_terms,
    parse_candidate_batch_response,
    parse_chunk_term_recall_response,
    parse_review_batch_response,
    preserve_unique,
    read_source_records,
    reviewed_terms_from_runtime,
    save_nontrans_regex_to_excel,
    save_results_to_excel,
    scan_folder,
    segment_source_records,
    segments_from_runtime,
    source_records_from_runtime,
)
from .models import (
    CandidateTerm,
    FailureRecord,
    LLMRequest,
    NonTransElement,
    NonTransRegexRow,
    RuntimeTaskState,
    SourceRecord,
    TaskInput,
    TermRecallCleanRecord,
    TextSegment,
    normalize_extraction_mode,
    now_iso,
    sync_extraction_flags,
)
from .ai_review.memoq_service import MemoQError, lookup_terms
from .nontrans import (
    attach_runtime_examples_to_nontrans_rows,
    build_missing_regex_generation_batches,
    build_nontrans_discovery_batches,
    build_nontrans_regex_sheet_rows,
    collect_covering_rows_for_elements,
    deduplicate_nontrans_regex_rows,
    estimate_prompt_tokens,
    expand_nontrans_regex_rows,
    filter_candidate_nontrans_records,
    load_builtin_nontrans_rules,
    normalize_ascii_candidate_patterns,
    normalize_nontrans_placeholder_format,
    parse_missing_regex_generation_response,
    parse_nontrans_discovery_response,
    placeholder_pattern_from_format,
    protect_nontrans_text,
    resolve_nontrans_regex_order,
    update_builtin_rule_examples_from_rows,
)
from .providers import ProviderRegistry
from .scheduler import AdaptiveConcurrencyController, AsyncRequestScheduler
from .storage import AppPaths, RuntimeCacheStore
from .storage import SettingsStore, append_pending_nontrans_rule_imports, load_pending_nontrans_rule_imports
from .telemetry import track_event


class TaskCancelledError(Exception):
    """Raised when the user stops the running task."""


class TermExtractionService:
    def __init__(
        self,
        settings,
        runtime_store: RuntimeCacheStore,
        paths: AppPaths,
        logger: logging.Logger,
        log_callback: Optional[Callable[[str], None]] = None,
        progress_callback: Optional[Callable[[Dict[str, object]], None]] = None,
        stop_requested: Optional[Callable[[], bool]] = None,
    ):
        self.settings = settings
        self.runtime_store = runtime_store
        self.paths = paths
        self.logger = logger
        self.log_callback = log_callback or (lambda message: None)
        self.progress_callback = progress_callback or (lambda payload: None)
        self.stop_requested = stop_requested or (lambda: False)
        self.settings_store = SettingsStore(paths)
        self._task_log_dir: Optional[Path] = None

    def _check_cancelled(self) -> None:
        if self.stop_requested():
            raise TaskCancelledError("任务已停止。")

    @staticmethod
    def _infer_resume_stage(runtime: RuntimeTaskState) -> str:
        cancelled_from_stage = str(runtime.stats.get("cancelled_from_stage", "") or "").strip()
        if cancelled_from_stage and cancelled_from_stage != "CANCELLED":
            return cancelled_from_stage

        current_stage = str(runtime.stage or "").strip()
        if current_stage and current_stage != "CANCELLED":
            return current_stage

        if runtime.output_file and Path(runtime.output_file).exists():
            return "EXPORTING"
        if runtime.approved_results or runtime.review_results:
            return "EXPORTING"
        if runtime.reviewed_terms:
            return "AGGREGATING_TERMS"
        if runtime.candidate_terms:
            return "REVIEWING_CANDIDATES"
        if runtime.segments:
            return "RECALLING_CANDIDATES"
        if runtime.source_records:
            return "SEGMENTING_TEXT"
        return "READING_FILES"

    def _prepare_runtime_for_resume(self, runtime: RuntimeTaskState) -> RuntimeTaskState:
        if runtime.stage != "CANCELLED":
            return runtime

        resume_stage = self._infer_resume_stage(runtime)
        runtime.stage = resume_stage
        runtime.last_error = ""
        runtime.stats["cancelled_from_stage"] = resume_stage
        return runtime

    def _persist_cancelled_runtime(self, runtime: RuntimeTaskState, exc: TaskCancelledError) -> None:
        cancelled_from_stage = runtime.stage if runtime.stage and runtime.stage != "CANCELLED" else self._infer_resume_stage(runtime)
        runtime.stats["cancelled_from_stage"] = cancelled_from_stage
        runtime.stage = "CANCELLED"
        runtime.last_error = str(exc)
        self.runtime_store.save(runtime)

    def _log(self, message: str, level: str = "info") -> None:
        getattr(self.logger, level)(message)
        self.log_callback(message)

    def _ensure_task_log_dir(self, runtime: RuntimeTaskState) -> Path:
        target = self.paths.task_logs_dir / str(runtime.task_id or "unknown_task")
        target.mkdir(parents=True, exist_ok=True)
        self._task_log_dir = target
        return target

    def _write_task_trace(self, runtime: RuntimeTaskState, name: str, payload: Dict[str, object]) -> Optional[Path]:
        try:
            target_dir = self._ensure_task_log_dir(runtime)
            target_path = target_dir / name
            target_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
            return target_path
        except Exception as exc:
            self._log("写入任务日志失败：{0}".format(exc), level="warning")
            return None

    @staticmethod
    def _truncate_text(value: object, max_chars: int = 4000) -> str:
        text = str(value or "")
        if len(text) <= max_chars:
            return text
        return text[: max(200, int(max_chars or 200) - 3)] + "..."

    def _summarize_messages(self, messages, max_content_chars: int = 2000):
        summarized = []
        for message in list(messages or []):
            summarized.append(
                {
                    "role": str(message.get("role", "") or ""),
                    "content": self._truncate_text(message.get("content", ""), max_content_chars),
                }
            )
        return summarized

    @staticmethod
    def _full_messages(messages):
        return [
            {
                "role": str(message.get("role", "") or ""),
                "content": str(message.get("content", "") or ""),
            }
            for message in list(messages or [])
        ]

    @staticmethod
    def _pretty_json_text(payload: object) -> str:
        try:
            return json.dumps(payload, ensure_ascii=False, indent=2)
        except Exception:
            return str(payload)

    def _log_llm_exchange_detail(self, request, response) -> None:
        stage = str(request.metadata.get("stage", request.task_type) or request.task_type)
        request_messages = self._full_messages(getattr(request, "messages", []) or [])
        request_prompt = str(getattr(request, "prompt", "") or "")
        request_detail = {
            "stage": stage,
            "task_id": str(getattr(request, "task_id", "") or ""),
            "task_type": str(getattr(request, "task_type", "") or ""),
            "attempt": int(getattr(response, "attempts", 1) or 1),
            "messages": request_messages,
            "prompt": request_prompt,
            "metadata": dict(getattr(request, "metadata", {}) or {}),
        }
        response_detail = {
            "stage": stage,
            "task_id": str(getattr(response, "task_id", "") or ""),
            "task_type": str(getattr(response, "task_type", "") or ""),
            "provider": str(getattr(response, "provider", "") or ""),
            "model": str(getattr(response, "model", "") or ""),
            "latency_ms": int(getattr(response, "latency_ms", 0) or 0),
            "attempt": int(getattr(response, "attempts", 1) or 1),
            "success": bool(getattr(response, "success", False)),
            "error": str(getattr(response, "error", "") or ""),
            "error_type": str(getattr(response, "error_type", "") or ""),
            "retryable": bool(getattr(response, "retryable", False)),
            "response_metadata": dict(getattr(response, "response_metadata", {}) or {}),
            "content": str(getattr(response, "content", "") or ""),
        }
        self._log(
            "REQUEST DETAIL [{0}]\n{1}".format(
                str(getattr(request, "task_id", "") or ""),
                self._pretty_json_text(request_detail),
            )
        )
        self._log(
            "RESPONSE DETAIL [{0}]\n{1}".format(
                str(getattr(response, "task_id", "") or ""),
                self._pretty_json_text(response_detail),
            )
        )

    def _emit_progress(self, runtime: RuntimeTaskState, **extra) -> None:
        payload = {
            "stage": runtime.stage,
            "stage_label": STAGE_LABELS.get(runtime.stage, runtime.stage),
            "current_file": runtime.stats.get("current_file", ""),
            "current_batch": runtime.stats.get("current_batch", 0),
            "total_batches": runtime.stats.get("total_batches", 0),
            "success_count": runtime.stats.get("success_count", 0),
            "failure_count": runtime.stats.get("failure_count", 0),
            "retry_count": runtime.stats.get("retry_count", 0),
            "current_concurrency": runtime.stats.get("current_concurrency", 1),
            "progress_current": runtime.stats.get("progress_current", 0),
            "progress_total": runtime.stats.get("progress_total", 0),
            "last_error": runtime.last_error,
        }
        payload.update(runtime.stats)
        payload.update(extra)
        self.progress_callback(payload)

    def _record_failure(
        self,
        runtime: RuntimeTaskState,
        task_type: str,
        item_id: str,
        reason: str,
        stage: str,
        source_excerpt: str = "",
        attempt_count: int = 0,
        guidance: str = "",
    ) -> None:
        key = (stage, task_type, item_id, reason)
        for existing_item in runtime.failed_items:
            if (
                existing_item.stage == stage
                and existing_item.task_type == task_type
                and existing_item.item_id == item_id
                and existing_item.reason == reason
            ):
                existing_item.attempt_count = max(int(existing_item.attempt_count or 0), int(attempt_count or 0))
                if source_excerpt:
                    existing_item.source_excerpt = source_excerpt[:500]
                if guidance:
                    existing_item.guidance = str(guidance or "").strip()
                return
        normalized_guidance = str(guidance or "").strip()
        if not normalized_guidance:
            normalized_guidance = build_failure_guidance(
                stage=stage,
                task_type=task_type,
                reason=reason,
                allowed_term_types=allowed_term_types_from_recall_scopes(
                    self.settings.input_defaults.get("recall_scopes", [])
                ),
            )
        runtime.failed_items.append(
            FailureRecord(
                task_type=task_type,
                item_id=item_id,
                reason=reason,
                guidance=normalized_guidance,
                stage=stage,
                source_excerpt=source_excerpt[:500],
                attempt_count=attempt_count,
            )
        )

    @staticmethod
    def _clear_failure_records(runtime: RuntimeTaskState, *, stage: str, task_type: str, item_ids: Sequence[str]) -> None:
        target_ids = {str(item_id or "").strip() for item_id in list(item_ids or []) if str(item_id or "").strip()}
        if not target_ids:
            return
        runtime.failed_items = [
            item
            for item in runtime.failed_items
            if not (
                item.stage == stage
                and item.task_type == task_type
                and str(item.item_id or "").strip() in target_ids
            )
        ]

    @staticmethod
    def _failure_reason(response, fallback: str) -> str:
        reason = (response.error or "").strip() if response is not None else ""
        json_mode = ""
        if response is not None:
            json_mode = str(getattr(response, "response_metadata", {}).get("json_mode", "") or "").strip()
        if json_mode == "prompt_only_fallback":
            suffix = "（未启用接口级JSON强约束）"
        elif json_mode == "prompt_only_recovery":
            suffix = "（接口级JSON强约束失败后已自动降级）"
        else:
            suffix = ""
        if reason:
            return reason + suffix
        error_type = (response.error_type or "").strip() if response is not None else ""
        if error_type:
            return "{0} ({1}){2}".format(fallback, error_type, suffix)
        return fallback + suffix

    @staticmethod
    def _current_json_mode(response) -> str:
        if response is None:
            return ""
        return str(getattr(response, "response_metadata", {}).get("json_mode", "") or "").strip()

    @staticmethod
    def _mark_request_prompt_only_recovery(request) -> None:
        metadata = dict(getattr(request, "metadata", {}) or {})
        metadata["force_prompt_only_json"] = True
        request.metadata = metadata

    @staticmethod
    def _llm_stage_key(request) -> str:
        task_type = str(getattr(request, "task_type", "") or "")
        metadata = dict(getattr(request, "metadata", {}) or {})
        if task_type == "nontrans_discovery_batch":
            return "nontrans_discovery"
        if task_type == "nontrans_regex_generation_batch":
            return "nontrans_regex_generation"
        if task_type == "nontrans_regex_reorder":
            return "nontrans_reorder"
        if task_type == "candidate_review_batch":
            return "term_review"
        if task_type == "candidate_recall_batch" and metadata.get("recall_mode") == "chunk":
            return "term_recall_chunk"
        if task_type == "candidate_recall_batch":
            return "term_recall_segment"
        return task_type or "unknown"

    @staticmethod
    def _estimate_prompt_chars(request) -> int:
        messages = list(getattr(request, "messages", []) or [])
        if messages:
            return sum(len(str(message.get("content", "") or "")) for message in messages)
        return len(str(getattr(request, "prompt", "") or ""))

    @staticmethod
    def _extract_usage_tokens(response) -> Dict[str, int]:
        metadata = dict(getattr(response, "response_metadata", {}) or {})
        usage = metadata.get("usage", {})
        if not isinstance(usage, dict):
            return {}
        aliases = {
            "prompt": ("prompt_tokens", "input_tokens"),
            "completion": ("completion_tokens", "output_tokens"),
            "total": ("total_tokens",),
        }
        extracted: Dict[str, int] = {}
        for key, names in aliases.items():
            for name in names:
                value = usage.get(name)
                if value is None:
                    continue
                try:
                    extracted[key] = int(value)
                    break
                except (TypeError, ValueError):
                    continue
        return extracted

    def _record_llm_response_stats(self, runtime: RuntimeTaskState, request, response) -> None:
        stage_key = self._llm_stage_key(request)
        prompt_chars = self._estimate_prompt_chars(request)
        latency_ms = max(0, int(getattr(response, "latency_ms", 0) or 0))
        attempts = max(1, int(getattr(response, "attempts", 1) or 1))

        runtime.stats["llm_request_count"] = int(runtime.stats.get("llm_request_count", 0) or 0) + 1
        if getattr(response, "success", False):
            runtime.stats["llm_success_response_count"] = int(runtime.stats.get("llm_success_response_count", 0) or 0) + 1
        else:
            runtime.stats["llm_failure_response_count"] = int(runtime.stats.get("llm_failure_response_count", 0) or 0) + 1
        runtime.stats["llm_attempt_count"] = int(runtime.stats.get("llm_attempt_count", 0) or 0) + attempts
        runtime.stats["llm_retry_attempt_count"] = int(runtime.stats.get("llm_retry_attempt_count", 0) or 0) + max(0, attempts - 1)
        prompt_chars_total = prompt_chars * attempts
        runtime.stats["llm_latency_ms_total"] = int(runtime.stats.get("llm_latency_ms_total", 0) or 0) + latency_ms
        runtime.stats["llm_prompt_char_count"] = int(runtime.stats.get("llm_prompt_char_count", 0) or 0) + prompt_chars_total

        total_responses = int(runtime.stats.get("llm_success_response_count", 0) or 0) + int(runtime.stats.get("llm_failure_response_count", 0) or 0)
        runtime.stats["llm_latency_ms_avg"] = int(round(runtime.stats["llm_latency_ms_total"] / total_responses)) if total_responses else 0

        stage_prefix = "llm_{0}".format(stage_key)
        runtime.stats["{0}_request_count".format(stage_prefix)] = int(runtime.stats.get("{0}_request_count".format(stage_prefix), 0) or 0) + 1
        runtime.stats["{0}_prompt_char_count".format(stage_prefix)] = int(runtime.stats.get("{0}_prompt_char_count".format(stage_prefix), 0) or 0) + prompt_chars_total

        usage = self._extract_usage_tokens(response)
        if "prompt" in usage:
            runtime.stats["llm_prompt_token_count"] = int(runtime.stats.get("llm_prompt_token_count", 0) or 0) + usage["prompt"]
            runtime.stats["{0}_prompt_token_count".format(stage_prefix)] = int(runtime.stats.get("{0}_prompt_token_count".format(stage_prefix), 0) or 0) + usage["prompt"]
        if "completion" in usage:
            runtime.stats["llm_completion_token_count"] = int(runtime.stats.get("llm_completion_token_count", 0) or 0) + usage["completion"]
        if "total" in usage:
            runtime.stats["llm_total_token_count"] = int(runtime.stats.get("llm_total_token_count", 0) or 0) + usage["total"]

    def _finalize_invalid_json_response(self, request, response, error_message: str, source_excerpt: str) -> object:
        json_mode = self._current_json_mode(response)
        if json_mode == "response_format":
            self._mark_request_prompt_only_recovery(request)
            response.retryable = True
            response.error_type = "invalid_json_response_format"
            response.error = "{0} 已切换为非强约束 JSON 模式重试。".format(error_message)
        else:
            response.retryable = False
            response.error_type = "invalid_json"
            response.error = error_message
        self._write_failed_response_debug(request, response, source_excerpt)
        return response

    def _create_runtime(self, task_input: TaskInput, provider_name: str, model_name: str) -> RuntimeTaskState:
        timestamp = now_iso()
        return RuntimeTaskState(
            cache_version=3,
            task_id=uuid.uuid4().hex[:12],
            stage="INITIALIZED",
            created_at=timestamp,
            updated_at=timestamp,
            input_config=task_input.to_dict(),
            provider_name=provider_name,
            model_name=model_name,
            stats={
                "success_count": 0,
                "failure_count": 0,
                "retry_count": 0,
                "semantic_retry_count": 0,
                "current_concurrency": 1,
                "progress_current": 0,
                "progress_total": 0,
                "current_file": "",
                "current_batch": 0,
                "total_batches": 0,
                "source_record_count": 0,
                "segment_count": 0,
                "candidate_count": 0,
                "approved_count": 0,
                "review_count": 0,
                "llm_request_count": 0,
                "llm_success_response_count": 0,
                "llm_failure_response_count": 0,
                "llm_attempt_count": 0,
                "llm_retry_attempt_count": 0,
                "llm_latency_ms_total": 0,
                "llm_latency_ms_avg": 0,
                "llm_prompt_char_count": 0,
                "llm_prompt_token_count": 0,
                "llm_completion_token_count": 0,
                "llm_total_token_count": 0,
            },
        )

    def _stage_settings(self, stage_key: str) -> Dict[str, object]:
        defaults = dict(self.settings.input_defaults.get("term_stage_settings", {}) or {})
        stage_settings = dict(self.settings.input_defaults.get(stage_key, {}) or {})
        merged = defaults
        merged.update(stage_settings)
        return merged

    def _stage_int_setting(self, stage_key: str, key: str, fallback: int) -> int:
        settings = self._stage_settings(stage_key)
        try:
            return int(settings.get(key, fallback) or fallback)
        except (TypeError, ValueError):
            return int(fallback)

    def _request_limit_int(self, key: str, fallback: int) -> int:
        try:
            return int(self.settings.request_limits.get(key, fallback) or fallback)
        except (TypeError, ValueError):
            return int(fallback)

    def _stage_enable_thinking(self, stage_key: str) -> bool:
        return bool(self._stage_settings(stage_key).get("enable_thinking", False))

    def _stage_reasoning_effort(self, stage_key: str) -> str:
        effort = str(self._stage_settings(stage_key).get("reasoning_effort") or "low").strip().lower()
        return effort if effort in {"low", "medium", "high"} else "low"

    async def test_connection(self) -> str:
        provider_name = self.settings.provider_name
        provider_settings = self.settings.provider_settings[provider_name]
        adapter = ProviderRegistry.create_adapter(provider_name, provider_settings)
        try:
            response = await adapter.test_connection()
        finally:
            await adapter.close()

    async def list_available_models(self) -> List[str]:
        provider_name = self.settings.provider_name
        provider_settings = self.settings.provider_settings[provider_name]
        adapter = ProviderRegistry.create_adapter(provider_name, provider_settings)
        try:
            success, message, models = await adapter.list_models()
            if not success:
                raise ValueError(message or "加载模型失败")
            return models
        finally:
            await adapter.close()

    async def run(self, task_input: Optional[TaskInput], resume: bool = False) -> str:
        provider_name = self.settings.provider_name
        provider_settings = self.settings.provider_settings[provider_name]
        if not provider_settings.model.strip():
            raise ValueError("当前模型名称为空，请先加载或填写模型。")
        if not provider_settings.api_key.strip():
            raise ValueError("当前还没有填写 API Key。")

        if resume:
            runtime = self.runtime_store.load()
            if not runtime:
                raise ValueError("没有可继续的任务缓存。")
            runtime = self._prepare_runtime_for_resume(runtime)
            task_input = TaskInput.from_dict(runtime.input_config)
            self._log("检测到任务缓存，继续执行上次任务。")
        else:
            if task_input is None:
                raise ValueError("缺少任务输入配置。")
            runtime = self._create_runtime(task_input, provider_name, provider_settings.model)
            self.runtime_store.save(runtime)
            self._log("已创建新的任务缓存。")

        self._ensure_task_log_dir(runtime)

        self._check_cancelled()
        assert task_input is not None
        task_input.extraction_mode = normalize_extraction_mode(task_input.extraction_mode)
        sync_extraction_flags(self.settings.input_defaults)
        self.settings.input_defaults["extraction_mode"] = task_input.extraction_mode
        sync_extraction_flags(self.settings.input_defaults)

        scan_result = scan_folder(task_input.folder_path)
        task_input.file_type = scan_result.file_type
        runtime.input_config = task_input.to_dict()
        runtime.provider_name = provider_name
        runtime.model_name = provider_settings.model

        source_records = source_records_from_runtime(runtime.source_records)
        if not source_records:
            runtime.stage = "READING_FILES"
            runtime.stats["progress_total"] = scan_result.file_count
            runtime.stats["progress_current"] = 0
            self.runtime_store.save(runtime)
            self._emit_progress(runtime, message="开始读取输入文件。")

            source_records, processed_files = read_source_records(
                folder_path=task_input.folder_path,
                file_type=scan_result.file_type,
                header_name=task_input.header_name,
                sheet_selections=task_input.column_selections,
                progress_callback=lambda payload: self._on_read_progress(runtime, payload),
            )
            runtime.source_records = [item.to_dict() for item in source_records]
            runtime.processed_files = processed_files
            runtime.stats["source_record_count"] = len(source_records)
            self.runtime_store.save(runtime)
            self._log("已读取 {0} 条源文本。".format(len(source_records)))

        if not runtime.source_records:
            raise ValueError("没有读取到可处理文本，请检查列名或源文件内容。")

        if task_input.extraction_mode == "nontrans_only":
            return await self._run_nontrans_only_mode(runtime, task_input, provider_name, provider_settings)

        nontrans_regex_rows = []
        nontrans_runtime_rows = []
        use_chunk_recall = True
        can_run_nontrans_preface = (
            not runtime.segments
            and not runtime.candidate_terms
            and not runtime.reviewed_terms
            and runtime.stage not in {"RECALLING_CANDIDATES", "REVIEWING_CANDIDATES", "AGGREGATING_TERMS", "EXPORTING"}
        )
        if can_run_nontrans_preface:
            nontrans_result = await self._run_nontrans_regex_stage(
                runtime,
                provider_name,
                provider_settings,
                source_records_from_runtime(runtime.source_records),
            )
            nontrans_regex_rows = list(nontrans_result["sheet_rows"])
            nontrans_runtime_rows = list(nontrans_result["regex_rows"])

        self._check_cancelled()
        segments = segments_from_runtime(runtime.segments)
        if not segments:
            runtime.stage = "SEGMENTING_TEXT"
            runtime.stats["progress_current"] = 0
            runtime.stats["progress_total"] = len(runtime.source_records)
            self.runtime_store.save(runtime)
            self._emit_progress(runtime, message="正在按单条文本切分片段。")

            source_records = source_records_from_runtime(runtime.source_records)
            recall_single_item_limit = self._stage_int_setting(
                "term_recall_stage_settings",
                "single_item_char_limit",
                task_input.single_item_char_limit,
            )
            segments = segment_source_records(
                source_records,
                recall_single_item_limit,
                ascii_filter_blacklist=self.settings.input_defaults.get("ascii_filter_blacklist", []),
                ascii_filter_whitelist=self.settings.input_defaults.get("ascii_filter_whitelist", []),
            )
            runtime.segments = [item.to_dict() for item in segments]
            runtime.stats["segment_count"] = len(segments)
            runtime.stats["progress_current"] = len(runtime.source_records)
            self.runtime_store.save(runtime)
            self._emit_progress(runtime, message="文本分段完成。")
            self._log("共生成 {0} 个文本片段。".format(len(segments)))

        self._check_cancelled()
        candidate_terms = candidate_terms_from_runtime(runtime.candidate_terms)
        if not candidate_terms or runtime.stage == "RECALLING_CANDIDATES":
            if use_chunk_recall:
                await self._run_chunk_candidate_recall_stage(
                    runtime,
                    task_input,
                    provider_name,
                    provider_settings,
                    nontrans_runtime_rows,
                )
            else:
                await self._run_candidate_recall_stage(runtime, task_input, provider_name, provider_settings)
            candidate_terms = candidate_terms_from_runtime(runtime.candidate_terms)

        if not runtime.candidate_terms:
            raise ValueError("未召回到任何候选词，请调整提示词或输入文本。")

        self._check_cancelled()
        reviewed_terms = reviewed_terms_from_runtime(runtime.reviewed_terms)
        if not reviewed_terms or runtime.stage == "REVIEWING_CANDIDATES":
            await self._run_review_stage(runtime, task_input, provider_name, provider_settings)
            reviewed_terms = reviewed_terms_from_runtime(runtime.reviewed_terms)

        self._check_cancelled()
        runtime.stage = "AGGREGATING_TERMS"
        runtime.stats["progress_current"] = 0
        runtime.stats["progress_total"] = len(runtime.reviewed_terms)
        self.runtime_store.save(runtime)
        self._emit_progress(runtime, message="正在聚合术语结果。")

        reviewed_terms = reviewed_terms_from_runtime(runtime.reviewed_terms)
        approved_rows, review_rows = aggregate_reviewed_terms(
            reviewed_terms,
            single_occurrence_approved_policy=str(
                self.settings.input_defaults.get("single_occurrence_approved_policy", "allow_to_library")
                or "allow_to_library"
            ),
        )
        runtime.approved_results = approved_rows
        runtime.review_results = review_rows
        runtime.extracted_terms = preserve_unique([row["术语原文"] for row in approved_rows])
        runtime.stats["approved_count"] = len(approved_rows)
        runtime.stats["review_count"] = len(review_rows)
        runtime.stats["progress_current"] = len(runtime.reviewed_terms)
        self.runtime_store.save(runtime)
        self._log("聚合完成：正式术语 {0} 条。".format(len(approved_rows)))

        self._check_cancelled()
        runtime.stage = "EXPORTING"
        runtime.stats["progress_current"] = 1
        runtime.stats["progress_total"] = 1
        self.runtime_store.save(runtime)
        self._emit_progress(runtime, message="正在导出 Excel。")

        output_file = save_results_to_excel(
            approved_rows=runtime.approved_results,
            review_rows=runtime.review_results,
            failed_items=self._build_failure_export_rows(runtime.failed_items),
            output_folder=self.paths.output_dir,
            export_review_sheet=task_input.export_review_sheet,
            nontrans_regex_rows=nontrans_regex_rows,
        )
        runtime.output_file = str(output_file)
        runtime.stage = "COMPLETED"
        self.runtime_store.save(runtime)
        self.runtime_store.clear()
        self._log("结果已导出到 {0}".format(output_file))
        self._emit_progress(runtime, message="处理完成。")
        return str(output_file)

    async def _run_nontrans_only_mode(self, runtime, task_input, provider_name, provider_settings) -> str:
        nontrans_result = await self._run_nontrans_regex_stage(
            runtime,
            provider_name,
            provider_settings,
            source_records_from_runtime(runtime.source_records),
        )
        sheet_rows = list(nontrans_result["sheet_rows"])

        self._check_cancelled()
        runtime.stage = "EXPORTING"
        runtime.stats["progress_current"] = 1
        runtime.stats["progress_total"] = 1
        self.runtime_store.save(runtime)
        self._emit_progress(runtime, message="正在导出非译元素正则 Excel。")
        output_file = save_nontrans_regex_to_excel(sheet_rows, self.paths.output_dir)
        runtime.output_file = str(output_file)
        runtime.stage = "COMPLETED"
        self.runtime_store.save(runtime)
        self.runtime_store.clear()
        self._log("非译元素正则已导出到 {0}".format(output_file))
        self._emit_progress(runtime, message="处理完成。")
        return str(output_file)

    async def _run_nontrans_regex_stage(self, runtime, provider_name, provider_settings, source_records):
        runtime.stage = "NONTRANS_DISCOVERY"
        runtime.stats["progress_current"] = 0
        runtime.stats["progress_total"] = len(source_records)
        self.runtime_store.save(runtime)
        self._emit_progress(runtime, message="正在筛选并提取非译元素。")

        nontrans_settings = self._stage_settings("nontrans_stage_settings")
        chunk_char_limit = self._stage_int_setting("nontrans_stage_settings", "chunk_char_limit", 3000)
        internal_token_limit = self._request_limit_int("internal_nontrans_prompt_token_limit", 56000)
        regex_batch_item_limit = self._request_limit_int("internal_nontrans_regex_batch_item_limit", 50)
        enable_thinking = bool(nontrans_settings.get("enable_thinking", False))
        reasoning_effort = self._stage_reasoning_effort("nontrans_stage_settings")
        builtin_regex_enabled = bool(nontrans_settings.get("builtin_regex_enabled", True))
        ai_discovery_enabled = bool(nontrans_settings.get("ai_discovery_enabled", True))
        ai_regex_generation_enabled = bool(nontrans_settings.get("ai_regex_generation_enabled", True))
        ascii_patterns = normalize_ascii_candidate_patterns(
            self.settings.input_defaults.get("ascii_candidate_patterns", [])
        )
        candidate_records = filter_candidate_nontrans_records(source_records, ascii_patterns)
        self._log("非译元素候选条目 {0} 条。".format(len(candidate_records)))

        discovery_requests = []
        if ai_discovery_enabled:
            discovery_user_prompt = self.settings.prompt_templates["nontrans_discovery_user_prompt_template"]
            discovery_system_prompt = self.settings.prompt_templates["nontrans_discovery_system_prompt_template"]
            discovery_batches = build_nontrans_discovery_batches(
                candidate_records,
                discovery_user_prompt,
                chunk_char_limit,
                internal_token_limit=internal_token_limit,
            )
            for batch_index, batch in enumerate(discovery_batches, start=1):
                discovery_requests.append(
                    LLMRequest(
                        task_id="nontrans_discovery_batch{0}".format(batch_index),
                        task_type="nontrans_discovery_batch",
                        prompt=str(batch.get("prompt", "")),
                        messages=[
                            {"role": "system", "content": discovery_system_prompt},
                            {"role": "user", "content": str(batch.get("prompt", ""))},
                        ],
                        metadata={
                            "stage": "NONTRANS_DISCOVERY",
                            "batch_index": batch_index,
                            "enable_thinking": enable_thinking,
                            "reasoning_effort": reasoning_effort,
                            "items": list(batch.get("items", []) or []),
                        },
                    )
                )

        discovery_responses = await self._run_nontrans_requests(
            runtime,
            provider_name,
            provider_settings,
            discovery_requests,
            progress_total=len(discovery_requests),
            message="非译元素发现请求进行中。",
        )
        discovered_elements = []
        records_by_id = {record.record_id: record for record in candidate_records}
        for request in discovery_requests:
            request_records = {
                str(item["id"]): records_by_id[str(item["record_id"])]
                for item in list(request.metadata.get("items", []) or [])
                if str(item.get("record_id", "")) in records_by_id
            }
            parsed, _ = await self._resolve_nontrans_parse_with_retries(
                runtime=runtime,
                provider_name=provider_name,
                provider_settings=provider_settings,
                request=request,
                initial_response=discovery_responses[request.task_id],
                parser=lambda content, request_records=request_records: parse_nontrans_discovery_response(
                    content,
                    request_records,
                ),
                fallback="非译元素发现响应校验失败",
                retry_message="非译元素发现纠偏重试进行中。",
            )
            discovered_elements.extend(parsed["resolved"])

        merged_elements = self._merge_nontrans_elements_for_pipeline(discovered_elements)
        builtin_rows = (
            deduplicate_nontrans_regex_rows(expand_nontrans_regex_rows(load_builtin_nontrans_rules()))
            if builtin_regex_enabled
            else []
        )
        used_builtin_rows, missing_elements = collect_covering_rows_for_elements(builtin_rows, merged_elements)

        runtime.stage = "NONTRANS_REGEX_GENERATION"
        runtime.stats["progress_current"] = 0
        runtime.stats["progress_total"] = len(missing_elements)
        self.runtime_store.save(runtime)
        self._emit_progress(runtime, message="正在为缺失的非译元素生成正则。")

        generated_rules = []
        regex_rows = deduplicate_nontrans_regex_rows(list(used_builtin_rows))
        unresolved_elements = [NonTransElement.from_dict(item.to_dict()) for item in missing_elements]
        max_regen_rounds = max(1, int(self.settings.request_limits.get("max_retries", 3) or 3))
        runtime.stats["nontrans_regex_regen_round_count"] = 0

        if regex_rows:
            order_result = resolve_nontrans_regex_order(regex_rows, merged_elements)
            regex_rows = list(order_result["ordered_rows"])
            unresolved_elements = list(order_result["unresolved_elements"])
            runtime.stats["nontrans_regex_order_round_count"] = int(order_result.get("round_count", 0) or 0)

        regex_user_prompt = self.settings.prompt_templates["nontrans_regex_user_prompt_template"]
        regex_system_prompt = self.settings.prompt_templates["nontrans_regex_system_prompt_template"]
        generation_round = 0
        previous_unresolved_signature = set()
        while unresolved_elements and ai_regex_generation_enabled and generation_round < max_regen_rounds:
            generation_round += 1
            runtime.stats["nontrans_regex_regen_round_count"] = generation_round
            runtime.stats["progress_current"] = 0
            runtime.stats["progress_total"] = len(unresolved_elements)
            self.runtime_store.save(runtime)
            self._emit_progress(runtime, message="正在为剩余非译元素生成正则。")

            regex_requests = []
            regex_batches = build_missing_regex_generation_batches(
                unresolved_elements,
                regex_user_prompt,
                chunk_char_limit,
                internal_token_limit=internal_token_limit,
                max_items_per_batch=regex_batch_item_limit,
            )
            for batch_index, batch in enumerate(regex_batches, start=1):
                regex_requests.append(
                    LLMRequest(
                        task_id="nontrans_regex_generation_round{0}_batch{1}".format(generation_round, batch_index),
                        task_type="nontrans_regex_generation_batch",
                        prompt=str(batch.get("prompt", "")),
                        messages=[
                            {"role": "system", "content": regex_system_prompt},
                            {"role": "user", "content": str(batch.get("prompt", ""))},
                        ],
                        metadata={
                            "stage": "NONTRANS_REGEX_GENERATION",
                            "batch_index": batch_index,
                            "generation_round": generation_round,
                            "enable_thinking": enable_thinking,
                            "reasoning_effort": reasoning_effort,
                            "items": list(batch.get("items", []) or []),
                        },
                    )
                )

            regex_responses = await self._run_nontrans_requests(
                runtime,
                provider_name,
                provider_settings,
                regex_requests,
                progress_total=len(regex_requests),
                message="非译元素正则生成请求进行中。",
            )

            round_rules = []
            for request in regex_requests:
                def parser_for_regex_request(target_request):
                    elements_by_request_id = {}
                    for item in list(target_request.metadata.get("items", []) or []):
                        request_id = str(item.get("id", ""))
                        element_text = str(item.get("element", ""))
                        element_type = str(item.get("element_type", "") or "")
                        matched = next(
                            (
                                element
                                for element in unresolved_elements
                                if element.element == element_text and str(element.element_type or "") == element_type
                            ),
                            None,
                        )
                        if matched is not None:
                            elements_by_request_id[request_id] = matched
                    return lambda content, elements_by_request_id=elements_by_request_id: parse_missing_regex_generation_response(
                        content,
                        elements_by_request_id,
                    )

                parsed, _ = await self._resolve_nontrans_parse_with_retries(
                    runtime=runtime,
                    provider_name=provider_name,
                    provider_settings=provider_settings,
                    request=request,
                    initial_response=regex_responses[request.task_id],
                    parser=parser_for_regex_request(request),
                    fallback="非译元素正则生成响应校验失败",
                    retry_message="非译元素正则生成纠偏重试进行中。",
                    parser_for_request=parser_for_regex_request,
                )
                round_rules.extend(parsed["resolved"])

            if round_rules:
                generated_rules.extend(round_rules)
                latest_settings = self.settings_store.load()
                pending_before = len(load_pending_nontrans_rule_imports(latest_settings))
                merged_pending = append_pending_nontrans_rule_imports(latest_settings, round_rules)
                self.settings_store.save(latest_settings)
                if len(merged_pending) > pending_before:
                    track_event("feature_used.nontrans_regex_generated")

            regex_rows = deduplicate_nontrans_regex_rows(used_builtin_rows + expand_nontrans_regex_rows(generated_rules))
            if regex_rows:
                order_result = resolve_nontrans_regex_order(regex_rows, merged_elements)
                regex_rows = list(order_result["ordered_rows"])
                unresolved_elements = list(order_result["unresolved_elements"])
                runtime.stats["nontrans_regex_order_round_count"] = int(
                    runtime.stats.get("nontrans_regex_order_round_count", 0) or 0
                ) + int(order_result.get("round_count", 0) or 0)

            unresolved_signature = {
                (str(item.element or "").strip(), str(item.element_type or "").strip())
                for item in unresolved_elements
            }
            if not round_rules or unresolved_signature == previous_unresolved_signature:
                break
            previous_unresolved_signature = unresolved_signature

        if unresolved_elements:
            failure_reason = "仍有非译元素无法被任何正则完整匹配，共 {0} 条。".format(len(unresolved_elements))
            self._record_failure(
                runtime,
                task_type="nontrans_regex_resolution",
                item_id="unresolved_nontrans_elements",
                reason=failure_reason,
                stage="NONTRANS_REGEX_GENERATION",
                source_excerpt="\n".join(str(item.element or "") for item in unresolved_elements[:20]),
                attempt_count=int(runtime.stats.get("nontrans_regex_regen_round_count", 0) or 0),
                guidance="请检查剩余非译元素，必要时补充更合适的正则生成提示词或手动维护规则库。",
            )
            runtime.last_error = failure_reason
            self.runtime_store.save(runtime)
            raise ValueError(failure_reason)

        regex_rows = attach_runtime_examples_to_nontrans_rows(regex_rows, source_records)
        try:
            update_builtin_rule_examples_from_rows(regex_rows, max_examples_per_rule=3)
        except Exception as exc:
            self._log("更新内置规则库样例失败，已跳过，不影响本次结果导出：{0}".format(exc), level="warning")
        sheet_rows = build_nontrans_regex_sheet_rows(regex_rows)

        self._check_cancelled()
        runtime.stats["nontrans_candidate_record_count"] = len(candidate_records)
        runtime.stats["nontrans_element_count"] = len(merged_elements)
        runtime.stats["nontrans_regex_row_count"] = len(regex_rows)
        self.runtime_store.save(runtime)
        return {
            "sheet_rows": sheet_rows,
            "regex_rows": regex_rows,
        }

    async def _run_nontrans_reorder_stage(
        self,
        runtime,
        provider_name,
        provider_settings,
        regex_rows,
        enable_thinking: bool,
    ):
        return list(regex_rows or [])

    async def _run_nontrans_requests(
        self,
        runtime,
        provider_name,
        provider_settings,
        requests,
        progress_total: int,
        message: str,
    ):
        if not requests:
            return {}
        responses = {}
        runtime.stats["total_batches"] = len(requests)
        runtime.stats["current_batch"] = 0
        runtime.stats["progress_total"] = progress_total
        self.runtime_store.save(runtime)

        def handle_response(rt, request, response, snapshot):
            responses[request.task_id] = response
            rt.stats["current_batch"] = snapshot.processed_count
            rt.stats["progress_current"] = min(int(progress_total or snapshot.total_count), snapshot.processed_count)
            rt.stats["current_concurrency"] = snapshot.current_concurrency
            rt.stats["retry_count"] = snapshot.retry_count + int(rt.stats.get("semantic_retry_count", 0) or 0)
            rt.completed_batch_ids.append(request.task_id)
            if response.success:
                rt.stats["success_count"] = rt.stats.get("success_count", 0) + 1
            else:
                rt.stats["failure_count"] = rt.stats.get("failure_count", 0) + 1
                rt.last_error = self._failure_reason(response, "非译元素请求失败")
                self._write_failed_response_debug(request, response, self._nontrans_request_excerpt(request))
                self._record_failure(
                    rt,
                    task_type=request.task_type,
                    item_id=request.task_id,
                    reason=self._failure_reason(response, "非译元素请求失败"),
                    stage=str(request.metadata.get("stage", "") or runtime.stage),
                    source_excerpt=self._nontrans_request_excerpt(request),
                    attempt_count=response.attempts,
                )
            self.runtime_store.save(rt)
            self._emit_progress(rt, message=message)

        await self._run_scheduler(
            runtime,
            provider_name,
            provider_settings,
            requests,
            handle_response,
            response_validator=self._validate_nontrans_batch_response,
        )
        failed = [response for response in responses.values() if not response.success]
        if failed:
            raise ValueError(failed[0].error or "非译元素请求失败。")
        return responses

    def _raise_on_nontrans_parse_issues(self, runtime, request, parsed, fallback: str) -> None:
        reason = self._build_nontrans_parse_failure_reason(parsed, fallback)
        self._record_failure(
            runtime,
            task_type=request.task_type,
            item_id=request.task_id,
            reason=reason,
            stage=str(request.metadata.get("stage", "") or runtime.stage),
            source_excerpt=self._nontrans_request_excerpt(request),
            attempt_count=0,
        )
        self.runtime_store.save(runtime)
        raise ValueError(reason)

    @staticmethod
    def _build_nontrans_parse_failure_reason(parsed, fallback: str) -> str:
        batch_issues = list(parsed.get("batch_issues", []) or [])
        item_issues = dict(parsed.get("item_issues", {}) or {})
        reason_parts = []
        if batch_issues:
            reason_parts.append("批次问题={0}".format(batch_issues))
        if item_issues:
            reason_parts.append("条目问题={0}".format(item_issues))
        if not reason_parts:
            return fallback
        return "{0}：{1}".format(fallback, "；".join(reason_parts))

    @staticmethod
    def _format_nontrans_parse_retry_guidance(parsed, fallback: str) -> str:
        batch_issues = [str(item or "").strip() for item in list(parsed.get("batch_issues", []) or []) if str(item or "").strip()]
        item_issues = dict(parsed.get("item_issues", {}) or {})
        issue_lines = []
        if batch_issues:
            issue_lines.extend("- {0}".format(issue) for issue in batch_issues[:6])
        for item_id, issues in list(item_issues.items())[:8]:
            clean_issues = [str(issue or "").strip() for issue in list(issues or []) if str(issue or "").strip()]
            if clean_issues:
                issue_lines.append("- id={0}: {1}".format(item_id, "；".join(clean_issues[:3])))
        joined = "\n".join(issue_lines) if issue_lines else "- 返回内容未通过本地校验。"
        return (
            "{0}\n"
            "请只返回纠正后的 JSON，并严格遵守以下要求：\n"
            "- 所有 element / regex 都必须与原始文本字符级一致。\n"
            "- 不要把 JSON 转义后的形式当作实际文本返回，例如不要返回 \\\\\" 来代替 \"，不要返回 <font=\\\\\"...\\\\\"> 来代替 <font=\"...\">。\n"
            "- 不要遗漏 id，不要新增 id，不要返回原文中不存在的元素。\n"
            "本地校验发现的问题：\n{1}"
        ).format(fallback, joined)

    def _build_nontrans_regex_retry_request(self, request, parsed, attempt: int, guidance: str):
        if str(getattr(request, "task_type", "") or "") != "nontrans_regex_generation_batch":
            return None
        item_issues = dict(parsed.get("item_issues", {}) or {})
        failed_ids = {str(item_id or "").strip() for item_id in item_issues.keys() if str(item_id or "").strip()}
        if not failed_ids:
            return None
        original_items = list(getattr(request, "metadata", {}).get("items", []) or [])
        retry_items = [dict(item) for item in original_items if str(item.get("id", "")).strip() in failed_ids]
        if not retry_items or len(retry_items) >= len(original_items):
            return None
        regex_user_prompt = self.settings.prompt_templates["nontrans_regex_user_prompt_template"]
        retry_prompt = regex_user_prompt.format(items_json=json.dumps(retry_items, ensure_ascii=False, indent=2))
        retry_prompt = retry_prompt + "\n\nRetry guidance:\n" + guidance
        metadata = dict(getattr(request, "metadata", {}) or {})
        metadata["items"] = retry_items
        metadata["retry_failed_ids_only"] = True
        metadata["original_task_id"] = str(getattr(request, "task_id", "") or "")
        return LLMRequest(
            task_id="{0}_failed_items_retry{1}".format(request.task_id, attempt),
            task_type=request.task_type,
            prompt=retry_prompt,
            messages=[
                {"role": "system", "content": self.settings.prompt_templates["nontrans_regex_system_prompt_template"]},
                {"role": "user", "content": retry_prompt},
            ],
            metadata=metadata,
        )

    async def _resolve_nontrans_parse_with_retries(
        self,
        runtime,
        provider_name,
        provider_settings,
        request,
        initial_response,
        parser,
        fallback: str,
        retry_message: str,
        parser_for_request=None,
    ):
        stage = str(request.metadata.get("stage", "") or runtime.stage)
        task_type = str(request.task_type or "")
        response = initial_response
        parsed = parser(response.content)
        if not list(parsed.get("batch_issues", []) or []) and not dict(parsed.get("item_issues", {}) or {}):
            self._clear_failure_records(runtime, stage=stage, task_type=task_type, item_ids=[request.task_id])
            return parsed, response

        accumulated_resolved = list(parsed.get("resolved", []) or [])
        last_parsed = parsed
        last_request = request
        semantic_attempt_count = 0
        self._write_semantic_validation_trace(
            runtime,
            request,
            response,
            {
                "batch_issues": list(parsed.get("batch_issues", []) or []),
                "item_issues": dict(parsed.get("item_issues", {}) or {}),
            },
            fallback,
            guidance=self._format_nontrans_parse_retry_guidance(parsed, fallback),
            semantic_attempt_count=semantic_attempt_count,
            source_excerpt=self._nontrans_request_excerpt(request),
        )
        for attempt in range(2, 4):
            guidance = self._format_nontrans_parse_retry_guidance(last_parsed, fallback)
            retry_request = self._build_nontrans_regex_retry_request(last_request, last_parsed, attempt, guidance)
            if retry_request is None:
                retry_messages = [dict(message) for message in list(request.messages or [])]
                if retry_messages and retry_messages[-1].get("role") == "user":
                    retry_messages[-1]["content"] = str(retry_messages[-1].get("content", "") or "") + "\n\nRetry guidance:\n" + guidance
                else:
                    retry_messages.append({"role": "user", "content": guidance})
                retry_request = LLMRequest(
                    task_id="{0}_parse_retry{1}".format(request.task_id, attempt),
                    task_type=request.task_type,
                    prompt=str(request.prompt or ""),
                    messages=retry_messages,
                    metadata=dict(request.metadata or {}),
                )
            retry_responses = await self._run_nontrans_requests(
                runtime,
                provider_name,
                provider_settings,
                [retry_request],
                progress_total=3,
                message=retry_message,
            )
            response = retry_responses[retry_request.task_id]
            last_request = retry_request
            retry_parser = parser_for_request(retry_request) if parser_for_request is not None else parser
            last_parsed = retry_parser(response.content)
            accumulated_resolved.extend(list(last_parsed.get("resolved", []) or []))
            semantic_attempt_count += 1
            runtime.stats["semantic_retry_count"] = int(runtime.stats.get("semantic_retry_count", 0) or 0) + 1
            runtime.stats["retry_count"] = int(runtime.stats.get("retry_count", 0) or 0) + 1
            self.runtime_store.save(runtime)
            self._emit_progress(runtime, message=retry_message)
            if not list(last_parsed.get("batch_issues", []) or []) and not dict(last_parsed.get("item_issues", {}) or {}):
                if str(getattr(request, "task_type", "") or "") == "nontrans_regex_generation_batch":
                    merged_resolved = []
                    seen_rule_ids = set()
                    for item in accumulated_resolved:
                        rule_id = str(getattr(item, "rule_id", "") or "")
                        key = rule_id or str(getattr(item, "pattern", "") or "")
                        if key and key in seen_rule_ids:
                            continue
                        if key:
                            seen_rule_ids.add(key)
                        merged_resolved.append(item)
                    last_parsed = dict(last_parsed)
                    last_parsed["resolved"] = merged_resolved
                self._clear_failure_records(
                    runtime,
                    stage=stage,
                    task_type=task_type,
                    item_ids=[request.task_id, last_request.task_id],
                )
                self.runtime_store.save(runtime)
                return last_parsed, response
            self._log(
                "{0} 第 {1} 次校验仍未通过，继续重试。".format(fallback, attempt),
                level="warning",
            )
            self._write_semantic_validation_trace(
                runtime,
                last_request,
                response,
                {
                    "batch_issues": list(last_parsed.get("batch_issues", []) or []),
                    "item_issues": dict(last_parsed.get("item_issues", {}) or {}),
                },
                fallback,
                guidance=self._format_nontrans_parse_retry_guidance(last_parsed, fallback),
                semantic_attempt_count=semantic_attempt_count,
                source_excerpt=self._nontrans_request_excerpt(last_request),
            )

        reason = self._build_nontrans_parse_failure_reason(last_parsed, fallback)
        self._record_failure(
            runtime,
            task_type=task_type,
            item_id=str(request.task_id or ""),
            reason=reason,
            stage=stage,
            source_excerpt=self._nontrans_request_excerpt(request),
            attempt_count=semantic_attempt_count,
        )
        self.runtime_store.save(runtime)
        raise ValueError(reason)
        return last_parsed, response

    @staticmethod
    def _merge_nontrans_elements_for_pipeline(elements):
        merged = {}
        ordered = []
        for element in elements:
            key = (element.element, element.element_type)
            existing = merged.get(key)
            if existing is None:
                clone = type(element).from_dict(element.to_dict())
                merged[key] = clone
                ordered.append(clone)
                continue
            existing.source_record_ids = preserve_unique(existing.source_record_ids + list(element.source_record_ids or []))
            existing.sample_contexts = preserve_unique(existing.sample_contexts + list(element.sample_contexts or []))
            existing.occurrence_count += max(1, int(element.occurrence_count or 1))
        return ordered

    @staticmethod
    def _nontrans_request_excerpt(request) -> str:
        excerpts = []
        for item in list(request.metadata.get("items", []) or [])[:3]:
            text = str(item.get("text", item.get("element", item.get("regex", ""))) or "")
            if text:
                excerpts.append(text)
        return "\n".join(excerpts)

    async def _run_candidate_recall_stage(self, runtime, task_input, provider_name, provider_settings) -> None:
        runtime.stage = "RECALLING_CANDIDATES"
        runtime.candidate_terms = []
        runtime.completed_batch_ids = []
        runtime.stats["progress_current"] = 0
        runtime.stats["progress_total"] = len(runtime.segments)
        runtime.stats["total_batches"] = 0
        runtime.stats["current_batch"] = 0
        runtime.stats["semantic_retry_count"] = 0
        self.runtime_store.save(runtime)
        self._emit_progress(runtime, message="开始批量召回术语候选词。")

        segments = segments_from_runtime(runtime.segments)
        segment_lookup = {segment.segment_id: segment for segment in segments}
        merged_candidates: Dict[tuple, CandidateTerm] = {}

        runtime.candidate_terms = [item.to_dict() for item in merged_candidates.values()]
        runtime.stats["candidate_count"] = len(runtime.candidate_terms)
        self.runtime_store.save(runtime)

        candidate_system_prompt_template = self.settings.prompt_templates["candidate_system_prompt_template"]
        candidate_user_prompt_template = self.settings.prompt_templates["candidate_user_prompt_template"]
        enabled_recall_scopes = format_enabled_recall_scopes(self.settings.input_defaults.get("recall_scopes", []))
        recall_batch_limit = self._stage_int_setting(
            "term_recall_stage_settings",
            "batch_request_char_limit",
            task_input.batch_request_char_limit,
        )
        recall_enable_thinking = self._stage_enable_thinking("term_recall_stage_settings")
        recall_reasoning_effort = self._stage_reasoning_effort("term_recall_stage_settings")

        if task_input.memoq_term_base_ids:
            candidate_user_prompt_template += (
                "\n\n术语库增量规则：仅提取 memoQ 术语库中尚未收录的增量术语；"
                "不要重复返回已有术语，也不要为了凑数量改写已有术语。"
            )
        total_items = len(segments)
        filtered_segment_ids = set()
        for segment in segments:
            if has_meaningful_clean_text(segment.llm_text):
                continue
            filtered_segment_ids.add(segment.segment_id)
            self._record_failure(
                runtime,
                task_type="candidate_recall_item",
                item_id=segment.segment_id,
                reason="过滤后无有效文本",
                stage="RECALLING_CANDIDATES",
                source_excerpt=segment.text,
                attempt_count=0,
            )

        pending_ids = [segment.segment_id for segment in segments if segment.segment_id not in filtered_segment_ids]
        attempt_counts = {item_id: 0 for item_id in pending_ids}
        retry_feedback_by_segment_id: Dict[str, List[str]] = {}
        cycle = 1

        while pending_ids:
            self._check_cancelled()
            pending_segments = [segment_lookup[item_id] for item_id in pending_ids]
            batches = build_candidate_recall_batches(
                pending_segments,
                candidate_user_prompt_template,
                task_input.source_language,
                recall_batch_limit,
                system_prompt_template=candidate_system_prompt_template,
                enabled_recall_scopes=enabled_recall_scopes,
                retry_feedback_by_segment_id=retry_feedback_by_segment_id,
            )
            base_completed = total_items - len(pending_ids)
            runtime.stats["progress_total"] = total_items
            runtime.stats["progress_current"] = base_completed
            runtime.stats["total_batches"] = len(batches)
            runtime.stats["current_batch"] = 0
            runtime.stats["semantic_retry_count"] = max(0, cycle - 1)
            self.runtime_store.save(runtime)
            if cycle == 1:
                self._emit_progress(runtime, message="候选词批量召回进行中。")
            else:
                self._emit_progress(runtime, message="正在重试未正确映射的候选召回条目。")

            for item_id in pending_ids:
                attempt_counts[item_id] += 1

            completed_ids_this_cycle = set()
            unresolved_reasons: Dict[str, List[str]] = defaultdict(list)
            requests = []
            for batch_index, batch in enumerate(batches, start=1):
                requests.append(
                    LLMRequest(
                        task_id="candidate_recall_cycle{0}_batch{1}".format(cycle, batch_index),
                        task_type="candidate_recall_batch",
                        prompt=batch["prompt"],
                        messages=batch.get("messages", []),
                        metadata={
                            "stage": "RECALLING_CANDIDATES",
                            "cycle": cycle,
                            "batch_index": batch_index,
                            "enable_thinking": recall_enable_thinking,
                            "reasoning_effort": recall_reasoning_effort,
                            "items": {
                                item["request_id"]: item["item"].to_dict()
                                for item in batch["items"]
                            },
                        },
                    )
                )

            await self._run_scheduler(
                runtime,
                provider_name,
                provider_settings,
                requests,
                lambda rt, req, resp, snap: self._handle_candidate_batch_result(
                    rt,
                    req,
                    resp,
                    snap,
                    merged_candidates,
                    completed_ids_this_cycle,
                    unresolved_reasons,
                    base_completed,
                ),
                response_validator=self._validate_candidate_batch_response,
            )
            self._check_cancelled()

            next_pending = []
            for item_id in pending_ids:
                reasons = preserve_unique(unresolved_reasons.get(item_id, []))
                if item_id in completed_ids_this_cycle and not reasons:
                    continue
                if not reasons and item_id not in completed_ids_this_cycle:
                    reasons = ["该条目本轮没有得到可映射结果"]

                if attempt_counts[item_id] >= 3:
                    segment = segment_lookup[item_id]
                    self._record_failure(
                        runtime,
                        task_type="candidate_recall_item",
                        item_id=item_id,
                        reason="；".join(reasons) or "连续 3 次批处理仍未成功映射",
                        stage="RECALLING_CANDIDATES",
                        source_excerpt=segment.text,
                        attempt_count=attempt_counts[item_id],
                    )
                else:
                    next_pending.append(item_id)

            pending_ids = next_pending
            retry_feedback_by_segment_id = {
                item_id: format_recall_retry_feedback(unresolved_reasons.get(item_id, []))
                for item_id in pending_ids
                if unresolved_reasons.get(item_id)
            }
            if pending_ids:
                cycle += 1
                self._log(
                    "候选召回有 {0} 条未正确映射，开始第 {1} 轮批量重试。".format(len(pending_ids), cycle)
                )

        candidate_terms = sorted(
            merged_candidates.values(),
            key=lambda item: (item.surface_form, item.segment_id, item.recall_source),
        )
        merged_candidate_terms = merge_candidate_terms(candidate_terms)
        if task_input.memoq_term_base_ids and merged_candidate_terms:
            try:
                existing = lookup_terms(
                    task_input.memoq_term_base_ids,
                    task_input.source_language,
                    "",
                    [item.surface_form for item in merged_candidate_terms],
                )
                existing_sources = {str(item.get("source") or "").strip().casefold() for item in existing}
                before_count = len(merged_candidate_terms)
                merged_candidate_terms = [item for item in merged_candidate_terms if item.surface_form.strip().casefold() not in existing_sources]
                self._log("memoQ 术语库去重完成：移除已有术语 {0} 条，保留增量术语 {1} 条。".format(before_count - len(merged_candidate_terms), len(merged_candidate_terms)))
            except MemoQError as exc:
                self._log("memoQ 术语库去重失败，继续审核召回结果：{0}".format(exc))
        runtime.candidate_terms = [item.to_dict() for item in merged_candidate_terms]
        runtime.stats["candidate_count"] = len(merged_candidate_terms)
        runtime.stats["progress_current"] = total_items
        self.runtime_store.save(runtime)
        self._log(
            "候选词召回完成，合并去重后保留 {0} 条候选术语。".format(len(merged_candidate_terms))
        )

    async def _run_chunk_candidate_recall_stage(
        self,
        runtime,
        task_input,
        provider_name,
        provider_settings,
        nontrans_regex_rows,
    ) -> None:
        runtime.stage = "RECALLING_CANDIDATES"
        runtime.candidate_terms = []
        runtime.completed_batch_ids = []
        runtime.stats["progress_current"] = 0
        runtime.stats["total_batches"] = 0
        runtime.stats["current_batch"] = 0
        runtime.stats["semantic_retry_count"] = 0
        self.runtime_store.save(runtime)
        self._emit_progress(runtime, message="正在基于清洗去重文本召回术语候选词。")

        source_records = source_records_from_runtime(runtime.source_records)
        source_records_by_id = {record.record_id: record for record in source_records}
        placeholder_format = normalize_nontrans_placeholder_format(
            self.settings.input_defaults.get("nontrans_placeholder_format", "<{n}>")
        )
        placeholder_pattern = placeholder_pattern_from_format(placeholder_format)
        source_segments = segment_source_records(
            source_records,
            self._stage_int_setting(
                "term_recall_stage_settings",
                "single_item_char_limit",
                task_input.single_item_char_limit,
            ),
            ascii_filter_blacklist=self.settings.input_defaults.get("ascii_filter_blacklist", []),
            ascii_filter_whitelist=self.settings.input_defaults.get("ascii_filter_whitelist", []),
        )
        runtime.segments = [item.to_dict() for item in source_segments]
        runtime.stats["segment_count"] = len(source_segments)
        self.runtime_store.save(runtime)

        protected_items = [
            (record.record_id, protect_nontrans_text(record.text, nontrans_regex_rows, placeholder_format))
            for record in source_records
        ]
        protected_changed_count = sum(
            1
            for record_id, protected in protected_items
            if protected.protected_text != str(source_records_by_id[record_id].text or "")
        )
        numeric_normalization_enabled = bool(
            self.settings.input_defaults.get("numeric_normalization_enabled", True)
        )
        clean_records, dedupe_map = build_term_recall_clean_records(
            protected_items,
            placeholder_pattern=placeholder_pattern,
            numeric_normalization_enabled=numeric_normalization_enabled,
            ascii_filter_blacklist=self.settings.input_defaults.get("ascii_filter_blacklist", []),
            ascii_filter_whitelist=self.settings.input_defaults.get("ascii_filter_whitelist", []),
        )
        clean_records_by_source_id = {record.source_record_id: record for record in clean_records}
        clean_changed_count = sum(
            1
            for record in clean_records
            if record.term_recall_clean_text.strip() != str(record.original_text or "").strip()
        )
        clean_lost_meaningful_count = sum(
            1
            for record in clean_records
            if has_meaningful_clean_text(record.original_text)
            and not has_meaningful_clean_text(record.term_recall_clean_text)
        )
        numeric_normalized_count = sum(
            1
            for record in clean_records
            if record.term_recall_clean_text
            and record.term_recall_clean_text != re.sub(r"<\d+>", "", record.nontrans_protected_text).strip()
        )

        recallable_records = []
        filtered_record_ids = set()
        for record in clean_records:
            if has_meaningful_clean_text(record.term_recall_clean_text):
                recallable_records.append(record)
                continue
            filtered_record_ids.add(record.source_record_id)
            source_record = source_records_by_id.get(record.source_record_id)
            self._record_failure(
                runtime,
                task_type="candidate_recall_item",
                item_id=record.source_record_id,
                reason="清洗后无有效召回文本",
                stage="RECALLING_CANDIDATES",
                source_excerpt=source_record.text if source_record else record.original_text,
                attempt_count=0,
            )

        unique_recallable_count = len({record.dedupe_key for record in recallable_records if record.dedupe_key})
        deduped_record_count = max(0, len(recallable_records) - unique_recallable_count)
        dedupe_savings_percent = (
            int(round((deduped_record_count / len(recallable_records)) * 100))
            if recallable_records
            else 0
        )
        runtime.stats["term_recall_clean_record_count"] = len(clean_records)
        runtime.stats["term_recall_recallable_record_count"] = len(recallable_records)
        runtime.stats["term_recall_unique_text_count"] = unique_recallable_count
        runtime.stats["term_recall_deduped_record_count"] = deduped_record_count
        runtime.stats["term_recall_dedupe_savings_percent"] = dedupe_savings_percent
        runtime.stats["term_recall_estimated_item_mode_text_count"] = len(recallable_records)
        runtime.stats["term_recall_filtered_record_count"] = len(filtered_record_ids)
        runtime.stats["term_recall_numeric_normalized_count"] = numeric_normalized_count
        runtime.stats["nontrans_protected_changed_record_count"] = protected_changed_count
        runtime.stats["term_recall_clean_changed_record_count"] = clean_changed_count
        runtime.stats["term_recall_clean_lost_meaningful_record_count"] = clean_lost_meaningful_count
        runtime.stats["progress_total"] = unique_recallable_count
        self.runtime_store.save(runtime)

        merged_candidates: Dict[tuple, CandidateTerm] = {}

        candidate_system_prompt_template = self.settings.prompt_templates["candidate_system_prompt_template"]
        candidate_user_prompt_template = self.settings.prompt_templates["candidate_user_prompt_template"]
        enabled_recall_scopes = format_enabled_recall_scopes(self.settings.input_defaults.get("recall_scopes", []))
        recall_batch_limit = self._stage_int_setting(
            "term_recall_stage_settings",
            "batch_request_char_limit",
            task_input.batch_request_char_limit,
        )
        recall_enable_thinking = self._stage_enable_thinking("term_recall_stage_settings")
        recall_reasoning_effort = self._stage_reasoning_effort("term_recall_stage_settings")

        batches = build_chunk_term_recall_batches(
            recallable_records,
            candidate_user_prompt_template,
            task_input.source_language,
            recall_batch_limit,
            system_prompt_template=candidate_system_prompt_template,
            enabled_recall_scopes=enabled_recall_scopes,
        )
        runtime.stats["total_batches"] = len(batches)
        runtime.stats["term_recall_chunk_batch_count"] = len(batches)
        runtime.stats["current_batch"] = 0
        runtime.stats["candidate_count"] = len(merged_candidates)
        self.runtime_store.save(runtime)

        requests = []
        request_records_by_task_id = {}
        for batch_index, batch in enumerate(batches, start=1):
            request_records = {
                item["request_id"]: item["item"]
                for item in batch["items"]
            }
            request_records_by_task_id["chunk_candidate_recall_batch{0}".format(batch_index)] = request_records
            requests.append(
                LLMRequest(
                    task_id="chunk_candidate_recall_batch{0}".format(batch_index),
                    task_type="candidate_recall_batch",
                    prompt=batch["prompt"],
                    messages=batch.get("messages", []),
                    metadata={
                        "stage": "RECALLING_CANDIDATES",
                        "recall_mode": "chunk",
                        "batch_index": batch_index,
                            "enable_thinking": recall_enable_thinking,
                            "reasoning_effort": recall_reasoning_effort,
                        "items": {
                            item["request_id"]: item["item"].to_dict()
                            for item in batch["items"]
                        },
                    },
                )
            )

        completed_request_ids = set()
        unresolved_reasons: Dict[str, List[str]] = defaultdict(list)
        await self._run_scheduler(
            runtime,
            provider_name,
            provider_settings,
            requests,
            lambda rt, req, resp, snap: self._handle_chunk_candidate_batch_result(
                rt,
                req,
                resp,
                snap,
                merged_candidates,
                completed_request_ids,
                unresolved_reasons,
                clean_records_by_source_id,
                source_records_by_id,
                dedupe_map,
            ),
            response_validator=self._validate_candidate_batch_response,
        )
        self._check_cancelled()

        for request in requests:
            request_records = request_records_by_task_id.get(request.task_id, {})
            for request_id, clean_record in request_records.items():
                if request_id in completed_request_ids:
                    continue
                reasons = preserve_unique(unresolved_reasons.get(request_id, [])) or ["该清洗文本没有得到可解析召回结果"]
                for source_record_id in dedupe_map.get(clean_record.dedupe_key, [clean_record.source_record_id]):
                    source_record = source_records_by_id.get(source_record_id)
                    self._record_failure(
                        runtime,
                        task_type="candidate_recall_item",
                        item_id=source_record_id,
                        reason="；".join(reasons),
                        stage="RECALLING_CANDIDATES",
                        source_excerpt=source_record.text if source_record else clean_record.original_text,
                        attempt_count=1,
                    )

        candidate_terms = sorted(
            merged_candidates.values(),
            key=lambda item: (item.surface_form, item.source_record_id, item.recall_source),
        )
        merged_candidate_terms = merge_candidate_terms(candidate_terms)
        runtime.candidate_terms = [item.to_dict() for item in merged_candidate_terms]
        runtime.stats["candidate_count"] = len(merged_candidate_terms)
        runtime.stats["progress_current"] = unique_recallable_count
        self.runtime_store.save(runtime)
        self._log(
            "清洗去重召回完成：原始文本 {0} 条，去重后召回文本 {1} 条，候选术语 {2} 条。".format(
                len(source_records),
                unique_recallable_count,
                len(merged_candidate_terms),
            )
        )

    async def _run_review_stage(self, runtime, task_input, provider_name, provider_settings) -> None:
        runtime.stage = "REVIEWING_CANDIDATES"
        runtime.reviewed_terms = []
        runtime.completed_batch_ids = []
        runtime.stats["progress_current"] = 0
        runtime.stats["progress_total"] = len(runtime.candidate_terms)
        runtime.stats["total_batches"] = 0
        runtime.stats["current_batch"] = 0
        runtime.stats["semantic_retry_count"] = 0
        self.runtime_store.save(runtime)
        self._emit_progress(runtime, message="开始批量判定术语候选词。")

        segments = {segment.segment_id: segment for segment in segments_from_runtime(runtime.segments)}
        candidates = candidate_terms_from_runtime(runtime.candidate_terms)
        review_system_prompt_template = self.settings.prompt_templates["classification_system_prompt_template"]
        review_user_prompt_template = self.settings.prompt_templates["classification_user_prompt_template"]
        enabled_recall_scopes = ""
        allowed_term_types: List[str] = []
        review_batch_limit = self._stage_int_setting(
            "term_review_stage_settings",
            "batch_request_char_limit",
            task_input.batch_request_char_limit,
        )
        review_context_limit = self._stage_int_setting(
            "term_review_stage_settings",
            "max_context_chars",
            220,
        )
        review_enable_thinking = self._stage_enable_thinking("term_review_stage_settings")
        review_reasoning_effort = self._stage_reasoning_effort("term_review_stage_settings")

        reviewable_candidates = []
        for candidate in candidates:
            if not candidate_has_traceable_evidence(candidate):
                self._record_failure(
                    runtime,
                    task_type="candidate_review_item",
                    item_id=candidate.candidate_id,
                    reason="候选词无法逐字追溯到来源文本或清洗文本，已跳过 AI 校验。",
                    stage="REVIEWING_CANDIDATES",
                    source_excerpt=candidate.context_text,
                    attempt_count=0,
                )
                continue
            segment = segments.get(candidate.segment_id)
            if segment is None and candidate.source_locations:
                reviewable_candidates.append(candidate)
                continue
            if segment is None:
                self._record_failure(
                    runtime,
                    task_type="candidate_review_item",
                    item_id=candidate.candidate_id,
                    reason="候选词缺少关联片段，无法判定。",
                    stage="REVIEWING_CANDIDATES",
                    source_excerpt=candidate.context_text,
                )
                continue
            reviewable_candidates.append(candidate)

        total_items = len(reviewable_candidates)
        pending_ids = [candidate.candidate_id for candidate in reviewable_candidates]
        attempt_counts = {candidate.candidate_id: 0 for candidate in reviewable_candidates}
        candidate_lookup = {candidate.candidate_id: candidate for candidate in reviewable_candidates}
        retry_feedback_by_candidate_id: Dict[str, List[str]] = {}
        cycle = 1

        while pending_ids:
            self._check_cancelled()
            pending_candidates = [candidate_lookup[item_id] for item_id in pending_ids]
            force_single_retry = cycle >= 3 or any(
                self._should_force_single_review_retry(retry_feedback_by_candidate_id.get(item_id, []))
                for item_id in pending_ids
            )
            max_items_per_batch = 1 if force_single_retry else (2 if cycle >= 2 else None)
            max_context_chars = review_context_limit if cycle == 1 else (min(160, review_context_limit) if cycle == 2 else min(120, review_context_limit))
            batches = build_review_retry_batches(
                pending_candidates,
                review_user_prompt_template,
                review_batch_limit,
                system_prompt_template=review_system_prompt_template,
                enabled_recall_scopes=enabled_recall_scopes,
                term_type_choices=allowed_term_types,
                retry_feedback_by_candidate_id=retry_feedback_by_candidate_id,
                max_items_per_batch=max_items_per_batch,
                max_context_chars=max_context_chars,
            )
            base_completed = total_items - len(pending_ids)
            runtime.stats["progress_total"] = total_items
            runtime.stats["progress_current"] = base_completed
            runtime.stats["total_batches"] = len(batches)
            runtime.stats["current_batch"] = 0
            runtime.stats["semantic_retry_count"] = max(0, cycle - 1)
            self.runtime_store.save(runtime)
            if cycle == 1:
                self._emit_progress(runtime, message="候选词批量判定进行中。")
            else:
                self._emit_progress(runtime, message="正在重试未正确映射的术语判定条目。")

            for item_id in pending_ids:
                attempt_counts[item_id] += 1

            completed_ids_this_cycle = set()
            unresolved_reasons: Dict[str, List[str]] = defaultdict(list)
            requests = []
            for batch_index, batch in enumerate(batches, start=1):
                items = {}
                for item in batch["items"]:
                    candidate = item["item"]
                    segment = segments.get(candidate.segment_id)
                    source_location = build_source_location(segment) if segment is not None else (
                        dict(candidate.source_locations[0]) if candidate.source_locations else {}
                    )
                    items[item["request_id"]] = {
                        "candidate": candidate.to_dict(),
                        "source_location": source_location,
                    }
                requests.append(
                    LLMRequest(
                        task_id="candidate_review_cycle{0}_batch{1}".format(cycle, batch_index),
                        task_type="candidate_review_batch",
                        prompt=batch["prompt"],
                        messages=batch.get("messages", []),
                        metadata={
                            "stage": "REVIEWING_CANDIDATES",
                            "cycle": cycle,
                            "batch_index": batch_index,
                            "allowed_term_types": allowed_term_types,
                            "enable_thinking": review_enable_thinking,
                            "reasoning_effort": review_reasoning_effort,
                            "items": items,
                        },
                    )
                )

            await self._run_scheduler(
                runtime,
                provider_name,
                provider_settings,
                requests,
                lambda rt, req, resp, snap: self._handle_review_batch_result(
                    rt,
                    req,
                    resp,
                    snap,
                    completed_ids_this_cycle,
                    unresolved_reasons,
                    base_completed,
                ),
                response_validator=self._validate_review_batch_response,
            )
            self._check_cancelled()

            next_pending = []
            for item_id in pending_ids:
                reasons = preserve_unique(unresolved_reasons.get(item_id, []))
                if item_id in completed_ids_this_cycle and not reasons:
                    continue
                if not reasons and item_id not in completed_ids_this_cycle:
                    reasons = ["该条目本轮没有得到可映射结果"]

                if attempt_counts[item_id] >= 3:
                    candidate = candidate_lookup[item_id]
                    self._record_failure(
                        runtime,
                        task_type="candidate_review_item",
                        item_id=item_id,
                        reason="；".join(reasons) or "连续 3 次批处理仍未成功映射",
                        stage="REVIEWING_CANDIDATES",
                        source_excerpt=candidate.context_text,
                        attempt_count=attempt_counts[item_id],
                    )
                else:
                    next_pending.append(item_id)

            pending_ids = next_pending
            retry_feedback_by_candidate_id = {
                item_id: format_review_retry_feedback(
                    unresolved_reasons.get(item_id, []),
                    allowed_term_types=allowed_term_types,
                )
                for item_id in pending_ids
                if unresolved_reasons.get(item_id)
            }
            if pending_ids:
                cycle += 1
                self._log(
                    "术语判定有 {0} 条未正确映射，开始第 {1} 轮批量重试。".format(len(pending_ids), cycle)
                )

        runtime.stats["progress_current"] = total_items
        self.runtime_store.save(runtime)
        self._log("候选词判定完成，共记录 {0} 条判定结果。".format(len(runtime.reviewed_terms)))

    async def _run_scheduler(
        self,
        runtime,
        provider_name,
        provider_settings,
        requests,
        result_handler,
        response_validator=None,
    ) -> None:
        adapter = ProviderRegistry.create_adapter(provider_name, provider_settings)
        controller = AdaptiveConcurrencyController(
            # Keep every AI stage within one predictable, self-adjusting limit.
            # Model connection settings must not make a task unexpectedly serial.
            mode="自动",
            user_max=4,
            provider_max=4,
        )
        scheduler = AsyncRequestScheduler(
            adapter=adapter,
            controller=controller,
            max_retries=int(self.settings.request_limits.get("max_retries", 3)),
            stop_requested=self.stop_requested,
            response_validator=response_validator,
        )
        try:
            await scheduler.run(
                requests,
                on_result=lambda request, response, snapshot: self._handle_scheduler_result(
                    runtime,
                    request,
                    response,
                    snapshot,
                    result_handler,
                ),
            )
        finally:
            await adapter.close()

    def _handle_scheduler_result(self, runtime, request, response, snapshot, result_handler):
        self._record_llm_response_stats(runtime, request, response)
        self._write_task_trace(
            runtime,
            "request_{0}_{1}.json".format(str(request.task_id), int(getattr(response, "attempts", 1) or 1)),
            {
                "task_id": request.task_id,
                "task_type": request.task_type,
                "metadata": dict(getattr(request, "metadata", {}) or {}),
                "messages": self._full_messages(getattr(request, "messages", []) or []),
                "prompt": str(getattr(request, "prompt", "") or ""),
                "summary": {
                    "prompt_chars": self._estimate_prompt_chars(request),
                    "message_count": len(list(getattr(request, "messages", []) or [])),
                },
            },
        )
        self._write_task_trace(
            runtime,
            "response_{0}_{1}.json".format(str(request.task_id), int(getattr(response, "attempts", 1) or 1)),
            {
                "task_id": response.task_id,
                "task_type": response.task_type,
                "provider": response.provider,
                "model": response.model,
                "latency_ms": response.latency_ms,
                "attempts": response.attempts,
                "success": response.success,
                "error": response.error,
                "error_type": response.error_type,
                "retryable": response.retryable,
                "response_metadata": dict(getattr(response, "response_metadata", {}) or {}),
                "content": str(getattr(response, "content", "") or ""),
                "summary": {
                    "response_chars": len(str(getattr(response, "content", "") or "")),
                    "json_mode": self._current_json_mode(response),
                },
            },
        )
        self._log(
            "[{0}] {1} attempt={2} success={3} latency={4}ms".format(
                str(request.metadata.get("stage", request.task_type) or request.task_type),
                request.task_id,
                int(getattr(response, "attempts", 1) or 1),
                bool(getattr(response, "success", False)),
                int(getattr(response, "latency_ms", 0) or 0),
            )
        )
        request_prompt = str(getattr(request, "prompt", "") or "")
        request_messages = list(getattr(request, "messages", []) or [])
        response_content = str(getattr(response, "content", "") or "")
        self._log(
            "批次摘要[{0}] prompt_chars={1} messages={2} response_chars={3}".format(
                request.task_id,
                len(request_prompt),
                len(request_messages),
                len(response_content),
            )
        )
        self._log_llm_exchange_detail(request, response)
        if str(getattr(response, "error", "") or "").strip():
            self._log("Response error: {0}".format(str(getattr(response, "error", "") or "")), level="warning")
        return result_handler(runtime, request, response, snapshot)

    def _handle_candidate_batch_result(
        self,
        runtime,
        request,
        response,
        snapshot,
        merged_candidates,
        completed_ids_this_cycle,
        unresolved_reasons,
        base_completed,
    ) -> None:
        runtime.stats["current_batch"] = snapshot.processed_count
        runtime.stats["progress_total"] = runtime.stats.get("progress_total", 0)
        runtime.stats["current_concurrency"] = snapshot.current_concurrency
        runtime.stats["retry_count"] = snapshot.retry_count + int(runtime.stats.get("semantic_retry_count", 0) or 0)
        runtime.completed_batch_ids.append(request.task_id)

        batch_segments = {
            request_id: TextSegment.from_dict(payload)
            for request_id, payload in dict(request.metadata.get("items", {}) or {}).items()
        }

        if response.success:
            runtime.stats["success_count"] = runtime.stats.get("success_count", 0) + 1
            parsed = parse_candidate_batch_response(response.content, batch_segments)
            if parsed["batch_issues"] or parsed["item_issues"]:
                self._write_semantic_response_debug(
                    request,
                    response,
                    self._batch_source_excerpt(batch_segments),
                    {
                        "batch_issues": parsed["batch_issues"],
                        "item_issues": parsed["item_issues"],
                        "item_warnings": parsed["item_warnings"],
                    },
                )

            for batch_issue in parsed["batch_issues"]:
                self._record_failure(
                    runtime,
                    task_type="candidate_recall_batch",
                    item_id=request.task_id,
                    reason=batch_issue,
                    stage="RECALLING_CANDIDATES",
                    source_excerpt=self._batch_source_excerpt(batch_segments),
                    attempt_count=response.attempts,
                )

            for request_id, issues in parsed["item_issues"].items():
                segment = batch_segments.get(request_id)
                if segment is not None:
                    unresolved_reasons[segment.segment_id].extend(issues)

            for request_id, terms in parsed["resolved"].items():
                segment = batch_segments[request_id]
                if request_id not in parsed["item_issues"]:
                    completed_ids_this_cycle.add(segment.segment_id)

                for term in terms:
                    key = (segment.segment_id, term)
                    existing = merged_candidates.get(key)
                    if existing is None:
                        merged_candidates[key] = CandidateTerm(
                            candidate_id="{0}:llm:{1}".format(segment.segment_id, term),
                            surface_form=term,
                            source_record_id=segment.source_record_id,
                            segment_id=segment.segment_id,
                            recall_source="llm",
                            context_text=segment.context_text,
                            evidence_text=term,
                            source_record_ids=[segment.source_record_id],
                            segment_ids=[segment.segment_id],
                            source_locations=[build_source_location(segment)],
                            sample_contexts=[segment.context_text],
                            occurrence_count=1,
                        )
        else:
            runtime.stats["failure_count"] = runtime.stats.get("failure_count", 0) + 1
            runtime.last_error = self._failure_reason(response, "候选词召回失败")
            self._write_failed_response_debug(request, response, self._batch_source_excerpt(batch_segments))
            self._record_failure(
                runtime,
                task_type="candidate_recall_batch",
                item_id=request.task_id,
                reason=self._failure_reason(response, "候选词召回失败"),
                stage="RECALLING_CANDIDATES",
                source_excerpt=self._batch_source_excerpt(batch_segments),
                attempt_count=response.attempts,
            )
            for segment in batch_segments.values():
                unresolved_reasons[segment.segment_id].append(self._failure_reason(response, "候选词召回失败"))
            self._log(
                "候选词召回批次失败：{0} -> {1}".format(
                    request.task_id, self._failure_reason(response, "候选词召回失败")
                ),
                level="warning",
            )

        runtime.candidate_terms = [item.to_dict() for item in merged_candidates.values()]
        runtime.stats["candidate_count"] = len(runtime.candidate_terms)
        runtime.stats["progress_current"] = base_completed + len(completed_ids_this_cycle)
        self.runtime_store.save(runtime)
        self._emit_progress(runtime, message="候选词批量召回进行中。")

    def _handle_chunk_candidate_batch_result(
        self,
        runtime,
        request,
        response,
        snapshot,
        merged_candidates,
        completed_request_ids,
        unresolved_reasons,
        clean_records_by_source_id,
        source_records_by_id,
        dedupe_map,
    ) -> None:
        runtime.stats["current_batch"] = snapshot.processed_count
        runtime.stats["current_concurrency"] = snapshot.current_concurrency
        runtime.stats["retry_count"] = snapshot.retry_count
        runtime.completed_batch_ids.append(request.task_id)

        clean_records_by_request_id = {
            request_id: TermRecallCleanRecord.from_dict(payload)
            for request_id, payload in dict(request.metadata.get("items", {}) or {}).items()
        }

        if response.success:
            runtime.stats["success_count"] = runtime.stats.get("success_count", 0) + 1
            parsed = parse_chunk_term_recall_response(response.content, clean_records_by_request_id)
            accepted_debug_count = 0
            rejected_debug_count = 0
            rejected_debug_samples = []
            accepted_request_ids_seen = set()
            for debug_item in list(parsed.get("debug_items", []) or []):
                request_id = str(debug_item.get("request_id", "") or "")
                raw_surface = str(debug_item.get("raw_surface", "") or "")
                sanitized = str(debug_item.get("sanitized", "") or "")
                accepted_request_ids = ",".join(list(debug_item.get("accepted_request_ids", []) or []))
                rejected_reason = str(debug_item.get("rejected_reason", "") or "")
                noise_request_ids = ",".join(list(debug_item.get("noise_request_ids", []) or []))
                sample_source_text = str(debug_item.get("sample_source_text", "") or "")
                if accepted_request_ids:
                    accepted_debug_count += 1
                    if request_id:
                        accepted_request_ids_seen.add(request_id)
                else:
                    rejected_debug_count += 1
                    if len(rejected_debug_samples) < 5:
                        rejected_debug_samples.append(
                            "request_id={0} raw={1} sanitized={2} reason={3}".format(
                                request_id,
                                raw_surface,
                                sanitized,
                                rejected_reason,
                            )
                        )
            self._log(
                "召回映射汇总[{0}] 请求 {1} 条，成功候选 {2} 条，命中请求 {3} 条，拒绝 {4} 条。".format(
                    request.task_id,
                    len(clean_records_by_request_id),
                    accepted_debug_count,
                    len(accepted_request_ids_seen),
                    rejected_debug_count,
                )
            )
            if rejected_debug_samples:
                self._log(
                    "召回拒绝示例[{0}] {1}".format(
                        request.task_id,
                        " | ".join(rejected_debug_samples),
                    ),
                    level="warning",
                )
            for request_id in clean_records_by_request_id:
                completed_request_ids.add(request_id)

            expanded_terms_by_request_id = {}
            for request_id, terms in parsed["resolved"].items():
                clean_record = clean_records_by_request_id.get(request_id)
                if clean_record is None:
                    continue
                source_record_ids = dedupe_map.get(clean_record.dedupe_key, [clean_record.source_record_id])
                for source_record_id in source_record_ids:
                    source_clean_record = clean_records_by_source_id.get(source_record_id)
                    if source_clean_record is None:
                        continue
                    expanded_terms_by_request_id[source_record_id] = terms

            chunk_candidates = build_candidate_terms_from_chunk_recall(
                expanded_terms_by_request_id,
                clean_records_by_source_id,
                source_records_by_id,
            )
            new_candidate_count = 0
            merged_candidate_count = 0
            for candidate in chunk_candidates:
                key = (candidate.source_record_id, candidate.surface_form, candidate.recall_source)
                existing = merged_candidates.get(key)
                if existing is None:
                    merged_candidates[key] = candidate
                    new_candidate_count += 1
                else:
                    merged_candidate_count += 1
                    existing.occurrence_count += max(1, int(candidate.occurrence_count or 1))
                    existing.source_record_ids = preserve_unique(existing.source_record_ids + candidate.source_record_ids)
                    existing.segment_ids = preserve_unique(existing.segment_ids + candidate.segment_ids)
                    existing.source_locations = existing.source_locations + candidate.source_locations
                    existing.sample_contexts = preserve_unique(existing.sample_contexts + candidate.sample_contexts)
            self._log(
                "候选写入汇总[{0}] 新增 {1} 条，合并 {2} 条，累计候选 {3} 条。".format(
                    request.task_id,
                    new_candidate_count,
                    merged_candidate_count,
                    len(merged_candidates),
                )
            )
        else:
            runtime.stats["failure_count"] = runtime.stats.get("failure_count", 0) + 1
            runtime.last_error = self._failure_reason(response, "清洗文本候选词召回失败")
            self._write_failed_response_debug(request, response, self._chunk_recall_request_excerpt(clean_records_by_request_id))
            self._record_failure(
                runtime,
                task_type="candidate_recall_batch",
                item_id=request.task_id,
                reason=self._failure_reason(response, "清洗文本候选词召回失败"),
                stage="RECALLING_CANDIDATES",
                source_excerpt=self._chunk_recall_request_excerpt(clean_records_by_request_id),
                attempt_count=response.attempts,
            )
            for request_id in clean_records_by_request_id:
                unresolved_reasons[request_id].append(self._failure_reason(response, "清洗文本候选词召回失败"))

        runtime.candidate_terms = [item.to_dict() for item in merged_candidates.values()]
        runtime.stats["candidate_count"] = len(runtime.candidate_terms)
        runtime.stats["progress_current"] = len(completed_request_ids)
        self.runtime_store.save(runtime)
        self._emit_progress(runtime, message="清洗文本候选词召回进行中。")

    def _handle_review_batch_result(
        self,
        runtime,
        request,
        response,
        snapshot,
        completed_ids_this_cycle,
        unresolved_reasons,
        base_completed,
    ) -> None:
        runtime.stats["current_batch"] = snapshot.processed_count
        runtime.stats["current_concurrency"] = snapshot.current_concurrency
        runtime.stats["retry_count"] = snapshot.retry_count + int(runtime.stats.get("semantic_retry_count", 0) or 0)
        runtime.completed_batch_ids.append(request.task_id)

        batch_items = {}
        for request_id, payload in dict(request.metadata.get("items", {}) or {}).items():
            batch_items[request_id] = {
                "candidate": CandidateTerm.from_dict(payload.get("candidate", {})),
                "source_location": dict(payload.get("source_location", {}) or {}),
            }

        if response.success:
            runtime.stats["success_count"] = runtime.stats.get("success_count", 0) + 1
            parsed = parse_review_batch_response(
                response.content,
                batch_items,
                allowed_term_types=request.metadata.get("allowed_term_types"),
            )
            if parsed["batch_issues"] or parsed["item_issues"]:
                self._write_semantic_response_debug(
                    request,
                    response,
                    self._batch_candidate_excerpt(batch_items),
                    {
                        "batch_issues": parsed["batch_issues"],
                        "item_issues": parsed["item_issues"],
                    },
                )

            for batch_issue in parsed["batch_issues"]:
                self._record_failure(
                    runtime,
                    task_type="candidate_review_batch",
                    item_id=request.task_id,
                    reason=batch_issue,
                    stage="REVIEWING_CANDIDATES",
                    source_excerpt=self._batch_candidate_excerpt(batch_items),
                    attempt_count=response.attempts,
                )

            for request_id, issues in parsed["item_issues"].items():
                candidate = batch_items.get(request_id, {}).get("candidate")
                if candidate is not None:
                    unresolved_reasons[candidate.candidate_id].extend(issues)

            for request_id, reviewed_term in parsed["resolved"].items():
                candidate = batch_items[request_id]["candidate"]
                if request_id in parsed["item_issues"]:
                    continue
                completed_ids_this_cycle.add(candidate.candidate_id)
                runtime.reviewed_terms.append(reviewed_term.to_dict())
        else:
            runtime.stats["failure_count"] = runtime.stats.get("failure_count", 0) + 1
            runtime.last_error = self._failure_reason(response, "候选词判定失败")
            self._write_failed_response_debug(request, response, self._batch_candidate_excerpt(batch_items))
            self._record_failure(
                runtime,
                task_type="candidate_review_batch",
                item_id=request.task_id,
                reason=self._failure_reason(response, "候选词判定失败"),
                stage="REVIEWING_CANDIDATES",
                source_excerpt=self._batch_candidate_excerpt(batch_items),
                attempt_count=response.attempts,
            )
            for item in batch_items.values():
                unresolved_reasons[item["candidate"].candidate_id].append(
                    self._failure_reason(response, "候选词判定失败")
                )
            self._log(
                "候选词判定批次失败：{0} -> {1}".format(
                    request.task_id, self._failure_reason(response, "候选词判定失败")
                ),
                level="warning",
            )

        runtime.stats["progress_current"] = base_completed + len(completed_ids_this_cycle)
        self.runtime_store.save(runtime)
        self._emit_progress(runtime, message="候选词批量判定进行中。")

    @staticmethod
    def _should_force_single_review_retry(reasons: List[str]) -> bool:
        combined = " ".join(str(reason or "") for reason in reasons)
        markers = (
            "JSON",
            "缺少 id",
            "重复 id",
            "未请求的 id",
            "空响应",
            "返回为空",
            "不是 JSON 对象",
            "遗漏了该编号",
        )
        return any(marker in combined for marker in markers)

    @staticmethod
    def _response_excerpt(content: str, max_chars: int = 160) -> str:
        compact = " ".join(str(content or "").split())
        if not compact:
            return "(空响应)"
        if len(compact) <= max_chars:
            return compact
        return compact[: max(20, int(max_chars or 20) - 3)].strip() + "..."

    def _write_failed_response_debug(self, request, response, source_excerpt: str) -> None:
        try:
            self.paths.debug_failed_responses_dir.mkdir(parents=True, exist_ok=True)
            timestamp = now_iso().replace(":", "").replace("-", "")
            target_path = self.paths.debug_failed_responses_dir / (
                "{0}_{1}_{2}_attempt{3}.json".format(
                    timestamp,
                    str(request.metadata.get("stage", "") or "unknown"),
                    request.task_id,
                    int(getattr(response, "attempts", 0) or 0),
                )
            )
            payload = {
                "provider": getattr(response, "provider", ""),
                "model": getattr(response, "model", ""),
                "stage": request.metadata.get("stage", ""),
                "task_id": request.task_id,
                "task_type": request.task_type,
                "attempt": int(getattr(response, "attempts", 0) or 0),
                "error": str(getattr(response, "error", "") or ""),
                "error_type": str(getattr(response, "error_type", "") or ""),
                "response_metadata": dict(getattr(response, "response_metadata", {}) or {}),
                "request_messages": self._summarize_messages(getattr(request, "messages", []) or [], 1500),
                "request_prompt": self._truncate_text(getattr(request, "prompt", "") or "", 5000),
                "source_excerpt": self._truncate_text(source_excerpt or "", 2000),
                "raw_response": self._truncate_text(getattr(response, "content", "") or "", 5000),
            }
            target_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
            self._log("失败响应诊断已写入：{0}".format(target_path), level="warning")
        except Exception as exc:  # pragma: no cover - debugging fallback
            self._log("写入失败响应调试文件失败：{0}".format(exc), level="warning")

    def _write_semantic_response_debug(self, request, response, source_excerpt: str, parsed_issues: Dict[str, object]) -> None:
        original_error = getattr(response, "error", "")
        original_error_type = getattr(response, "error_type", "")
        original_metadata = dict(getattr(response, "response_metadata", {}) or {})
        try:
            response.error = "模型返回 JSON 通过解析，但语义校验未通过。"
            response.error_type = "semantic_validation"
            metadata = dict(original_metadata)
            metadata["parsed_issues"] = parsed_issues
            response.response_metadata = metadata
            self._write_failed_response_debug(request, response, source_excerpt)
        finally:
            response.error = original_error
            response.error_type = original_error_type
            response.response_metadata = original_metadata

    def _write_semantic_validation_trace(
        self,
        runtime: RuntimeTaskState,
        request,
        response,
        parsed_issues: Dict[str, object],
        fallback: str,
        guidance: str = "",
        semantic_attempt_count: int = 0,
        source_excerpt: str = "",
    ) -> None:
        task_id = str(getattr(request, "task_id", "") or "unknown")
        if not source_excerpt:
            source_excerpt = self._nontrans_request_excerpt(request)
        target_path = self._write_task_trace(
            runtime,
            "semantic_validation_{0}.json".format(task_id),
            {
                "stage": str(getattr(request, "metadata", {}).get("stage", "") or runtime.stage),
                "task_id": task_id,
                "task_type": str(getattr(request, "task_type", "") or ""),
                "fallback": fallback,
                "semantic_attempt_count": int(semantic_attempt_count or 0),
                "parsed_issues": parsed_issues,
                "retry_guidance": guidance,
                "request_metadata": dict(getattr(request, "metadata", {}) or {}),
                "request_messages": self._full_messages(getattr(request, "messages", []) or []),
                "raw_response": str(getattr(response, "content", "") or ""),
                "source_excerpt": source_excerpt,
            },
        )
        if target_path is not None:
            self._log("语义校验诊断已写入：{0}".format(target_path), level="warning")

    def _validate_candidate_batch_response(self, request, response):
        if not response.success:
            return response
        if request.metadata.get("recall_mode") == "chunk":
            clean_records_by_request_id = {
                request_id: TermRecallCleanRecord.from_dict(payload)
                for request_id, payload in dict(request.metadata.get("items", {}) or {}).items()
            }
            try:
                parse_chunk_term_recall_response(response.content, clean_records_by_request_id)
            except Exception as exc:
                source_excerpt = self._chunk_recall_request_excerpt(clean_records_by_request_id)
                response.success = False
                self._finalize_invalid_json_response(request, response, str(exc), source_excerpt)
            return response
        batch_segments = {
            request_id: TextSegment.from_dict(payload)
            for request_id, payload in dict(request.metadata.get("items", {}) or {}).items()
        }
        try:
            parse_candidate_batch_response(response.content, batch_segments)
        except Exception as exc:
            source_excerpt = self._batch_source_excerpt(batch_segments)
            response.success = False
            self._finalize_invalid_json_response(request, response, str(exc), source_excerpt)
        return response

    def _validate_review_batch_response(self, request, response):
        if not response.success:
            return response
        batch_items = {}
        for request_id, payload in dict(request.metadata.get("items", {}) or {}).items():
            batch_items[request_id] = {
                "candidate": CandidateTerm.from_dict(payload.get("candidate", {})),
                "source_location": dict(payload.get("source_location", {}) or {}),
            }
        try:
            parse_review_batch_response(
                response.content,
                batch_items,
                allowed_term_types=request.metadata.get("allowed_term_types"),
            )
        except Exception as exc:
            response.success = False
            source_excerpt = self._batch_candidate_excerpt(batch_items)
            self._finalize_invalid_json_response(
                request,
                response,
                "{0} 响应摘录：{1}".format(str(exc), self._response_excerpt(response.content)),
                source_excerpt,
            )
        return response

    def _validate_nontrans_batch_response(self, request, response):
        if not response.success:
            return response
        stage = str(request.metadata.get("stage", "") or "")
        if stage == "NONTRANS_DISCOVERY":
            request_records = {}
            for item in list(request.metadata.get("items", []) or []):
                request_id = str(item.get("id", "")).strip()
                record_id = str(item.get("record_id", "")).strip()
                text = str(item.get("text", "") or item.get("llm_text", "") or "").strip()
                if not request_id:
                    continue
                request_records[request_id] = SourceRecord(
                    record_id=record_id or request_id,
                    file_name="",
                    source_type="",
                    sheet_or_unit="",
                    row_index=0,
                    column_name="",
                    text=text,
                )
            try:
                parse_nontrans_discovery_response(response.content, request_records)
            except Exception as exc:
                response.success = False
                self._finalize_invalid_json_response(
                    request,
                    response,
                    "{0} 响应摘录：{1}".format(str(exc), self._response_excerpt(response.content)),
                    self._nontrans_request_excerpt(request),
                )
            return response

        if stage == "NONTRANS_REGEX_GENERATION":
            elements_by_request_id = {}
            for item in list(request.metadata.get("items", []) or []):
                request_id = str(item.get("id", "")).strip()
                element_text = str(item.get("element", "") or "").strip()
                element_type = str(item.get("element_type", item.get("category", "other")) or "other").strip() or "other"
                if not request_id or not element_text:
                    continue
                elements_by_request_id[request_id] = NonTransElement(
                    element_id=request_id,
                    element=element_text,
                    element_type=element_type,
                    source_record_ids=[],
                    sample_contexts=[],
                    occurrence_count=1,
                )
            try:
                parse_missing_regex_generation_response(response.content, elements_by_request_id)
            except Exception as exc:
                response.success = False
                self._finalize_invalid_json_response(
                    request,
                    response,
                    "{0} 响应摘录：{1}".format(str(exc), self._response_excerpt(response.content)),
                    self._nontrans_request_excerpt(request),
                )
            return response

        if stage == "NONTRANS_REGEX_REORDER":
            items = list(request.metadata.get("items", []) or [])
            regex_rows = []
            for index, item in enumerate(items, start=1):
                regex_rows.append(
                    NonTransRegexRow(
                        row_id=str(item.get("row_id", "") or "row_{0}".format(index)),
                        rule_id=str(item.get("rule_id", "") or ""),
                        name=str(item.get("name", "") or ""),
                        regex=str(item.get("regex", "") or ""),
                        role=str(item.get("role", "") or ""),
                        element_type=str(item.get("element_type", "") or "other"),
                        order_index=index,
                        examples=list(item.get("examples", []) or []),
                    )
                )
            return response

        return response

    def _build_failure_export_rows(self, failed_items: List[FailureRecord]) -> List[Dict[str, object]]:
        rows = []
        for item in failed_items:
            rows.append(
                {
                    "阶段": item.stage,
                    "任务类型": item.task_type,
                    "项目ID": item.item_id,
                    "失败原因": item.reason,
                    "纠正建议": item.guidance,
                    "重试次数": item.attempt_count,
                    "原文摘录": item.source_excerpt,
                }
            )
        return rows

    def _batch_source_excerpt(self, batch_segments: Dict[str, TextSegment]) -> str:
        excerpts = [segment.text for segment in list(batch_segments.values())[:3]]
        return "\n\n".join(excerpts)[:500]

    def _chunk_recall_request_excerpt(self, clean_records_by_request_id: Dict[str, TermRecallCleanRecord]) -> str:
        excerpts = [
            record.term_recall_clean_text
            for record in list(clean_records_by_request_id.values())[:3]
            if record.term_recall_clean_text
        ]
        return "\n\n".join(excerpts)[:500]

    def _batch_candidate_excerpt(self, batch_items: Dict[str, Dict[str, object]]) -> str:
        excerpts = []
        for item in list(batch_items.values())[:3]:
            candidate = item.get("candidate")
            if candidate is not None:
                excerpts.append(str(getattr(candidate, "context_text", "")))
        return "\n\n".join(excerpts)[:500]

    def _on_read_progress(self, runtime: RuntimeTaskState, payload: Dict[str, object]) -> None:
        runtime.stats["current_file"] = payload.get("current_file", "")
        runtime.stats["progress_current"] = int(payload.get("current", 0) or 0)
        runtime.stats["progress_total"] = int(payload.get("total", 0) or 0)
        self.runtime_store.save(runtime)
        self._emit_progress(runtime, message=str(payload.get("message", "")))


