"""Typed application models."""

from dataclasses import asdict, dataclass, field
import json
from datetime import datetime
from typing import Any, Dict, List, Optional


def normalize_extraction_mode(value: Any) -> str:
    """Keep the app on the two agreed modes only."""
    return "nontrans_only" if str(value or "").strip() == "nontrans_only" else "terms"


def sync_extraction_flags(input_defaults: Dict[str, Any]) -> Dict[str, Any]:
    mode = normalize_extraction_mode(input_defaults.get("extraction_mode", "terms"))
    input_defaults["extraction_mode"] = mode
    input_defaults["enable_nontrans_extraction"] = True
    input_defaults["enable_term_extraction"] = mode == "terms"
    return input_defaults


def now_iso() -> str:
    return datetime.now().isoformat(timespec="seconds")


@dataclass
class ProviderSettings:
    api_key: str = ""
    base_url: str = ""
    model: str = ""
    timeout_seconds: int = 90
    max_concurrency: int = 6
    extra_headers: Dict[str, str] = field(default_factory=dict)
    disable_system_proxy: bool = True

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: Optional[Dict[str, Any]]) -> "ProviderSettings":
        data = data or {}
        return cls(
            api_key=str(data.get("api_key", "")),
            base_url=str(data.get("base_url", "")),
            model=str(data.get("model", "")),
            timeout_seconds=int(data.get("timeout_seconds", 90) or 90),
            max_concurrency=int(data.get("max_concurrency", 6) or 6),
            extra_headers=dict(data.get("extra_headers", {}) or {}),
            disable_system_proxy=bool(data.get("disable_system_proxy", True)),
        )


@dataclass
class AppSettings:
    config_version: int
    input_defaults: Dict[str, Any]
    provider_name: str
    provider_settings: Dict[str, ProviderSettings]
    prompt_templates: Dict[str, str]
    request_limits: Dict[str, Any]
    ui_preferences: Dict[str, Any]

    def to_dict(self) -> Dict[str, Any]:
        return {
            "config_version": self.config_version,
            "input_defaults": dict(self.input_defaults),
            "provider_name": self.provider_name,
            "provider_settings": {
                name: settings.to_dict()
                for name, settings in self.provider_settings.items()
            },
            "prompt_templates": dict(self.prompt_templates),
            "request_limits": dict(self.request_limits),
            "ui_preferences": dict(self.ui_preferences),
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "AppSettings":
        return cls(
            config_version=int(data.get("config_version", 1)),
            input_defaults=dict(data.get("input_defaults", {}) or {}),
            provider_name=str(data.get("provider_name", "DeepSeek")),
            provider_settings={
                name: ProviderSettings.from_dict(item)
                for name, item in dict(data.get("provider_settings", {}) or {}).items()
            },
            prompt_templates=dict(data.get("prompt_templates", {}) or {}),
            request_limits=dict(data.get("request_limits", {}) or {}),
            ui_preferences=dict(data.get("ui_preferences", {}) or {}),
        )


@dataclass
class TaskInput:
    folder_path: str
    header_name: str
    source_language: str
    single_item_char_limit: int
    batch_request_char_limit: int
    file_type: str = ""
    export_review_sheet: bool = True
    extraction_mode: str = "terms"
    memoq_term_base_ids: List[str] = field(default_factory=list)
    column_selections: Dict[str, List[str]] = field(default_factory=dict)
    input_files: List[str] = field(default_factory=list)
    file_mappings: Dict[str, Dict[str, List[str]]] = field(default_factory=dict)

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "TaskInput":
        export_review_sheet = data.get("export_review_sheet", True)
        legacy_batch_limit = int(data.get("max_chars_per_request", 3000) or 3000)
        return cls(
            folder_path=str(data.get("folder_path", "")),
            header_name=(json.dumps(data.get("header_name"), ensure_ascii=False) if isinstance(data.get("header_name"), list) else str(data.get("header_name", ""))),
            source_language=str(data.get("source_language", "")),
            single_item_char_limit=int(data.get("single_item_char_limit", 500) or 500),
            batch_request_char_limit=int(data.get("batch_request_char_limit", legacy_batch_limit) or legacy_batch_limit),
            file_type=str(data.get("file_type", "")),
            export_review_sheet=bool(export_review_sheet),
            extraction_mode=normalize_extraction_mode(data.get("extraction_mode", "terms")),
            memoq_term_base_ids=[str(x) for x in (data.get("memoq_term_base_ids") or []) if str(x).strip()],
            column_selections={str(k): [str(x) for x in (v or []) if str(x).strip()] for k, v in (data.get("column_selections") or {}).items()},
            input_files=[str(x) for x in (data.get("input_files") or []) if str(x).strip()],
            file_mappings={str(k): {str(sk): [str(x) for x in (sv or []) if str(x).strip()] for sk, sv in (v or {}).items()} for k, v in (data.get("file_mappings") or {}).items()},
        )


@dataclass
class SourceRecord:
    record_id: str
    file_name: str
    source_type: str
    sheet_or_unit: str
    row_index: int
    column_name: str
    text: str

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "SourceRecord":
        return cls(
            record_id=str(data.get("record_id", "")),
            file_name=str(data.get("file_name", "")),
            source_type=str(data.get("source_type", "")),
            sheet_or_unit=str(data.get("sheet_or_unit", "")),
            row_index=int(data.get("row_index", 0) or 0),
            column_name=str(data.get("column_name", "")),
            text=str(data.get("text", "")),
        )


@dataclass
class TextSegment:
    segment_id: str
    source_record_id: str
    file_name: str
    source_type: str
    sheet_or_unit: str
    row_index: int
    column_name: str
    segment_index: int
    total_segments: int
    text: str
    context_text: str
    llm_text: str = ""
    llm_context_text: str = ""

    def to_dict(self) -> Dict[str, Any]:
        data = asdict(self)
        if data.get("llm_text", "") == data.get("text", ""):
            data["llm_text"] = ""
        llm_context_text = str(data.get("llm_context_text", "") or "")
        context_text = str(data.get("context_text", "") or "")
        llm_text = str(data.get("llm_text", "") or data.get("text", "") or "")
        if llm_context_text == context_text or llm_context_text == llm_text:
            data["llm_context_text"] = ""
        return data

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "TextSegment":
        return cls(
            segment_id=str(data.get("segment_id", "")),
            source_record_id=str(data.get("source_record_id", "")),
            file_name=str(data.get("file_name", "")),
            source_type=str(data.get("source_type", "")),
            sheet_or_unit=str(data.get("sheet_or_unit", "")),
            row_index=int(data.get("row_index", 0) or 0),
            column_name=str(data.get("column_name", "")),
            segment_index=int(data.get("segment_index", 1) or 1),
            total_segments=int(data.get("total_segments", 1) or 1),
            text=str(data.get("text", "")),
            context_text=str(data.get("context_text", "")),
            llm_text=str(data.get("llm_text", data.get("text", "")) or ""),
            llm_context_text=str(data.get("llm_context_text", data.get("context_text", "")) or ""),
        )


@dataclass
class CandidateTerm:
    candidate_id: str
    surface_form: str
    source_record_id: str
    segment_id: str
    recall_source: str
    context_text: str
    evidence_text: str = ""
    source_record_ids: List[str] = field(default_factory=list)
    segment_ids: List[str] = field(default_factory=list)
    source_locations: List[Dict[str, Any]] = field(default_factory=list)
    sample_contexts: List[str] = field(default_factory=list)
    occurrence_count: int = 1

    def to_dict(self) -> Dict[str, Any]:
        data = asdict(self)
        data["source_record_ids"] = list(data.get("source_record_ids", []) or [])[:20]
        data["segment_ids"] = list(data.get("segment_ids", []) or [])[:20]
        data["source_locations"] = list(data.get("source_locations", []) or [])[:10]
        data["sample_contexts"] = list(data.get("sample_contexts", []) or [])[:3]
        return data

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "CandidateTerm":
        return cls(
            candidate_id=str(data.get("candidate_id", "")),
            surface_form=str(data.get("surface_form", "")),
            source_record_id=str(data.get("source_record_id", "")),
            segment_id=str(data.get("segment_id", "")),
            recall_source=str(data.get("recall_source", "")),
            context_text=str(data.get("context_text", "")),
            evidence_text=str(data.get("evidence_text", "")),
            source_record_ids=list(data.get("source_record_ids", []) or []),
            segment_ids=list(data.get("segment_ids", []) or []),
            source_locations=list(data.get("source_locations", []) or []),
            sample_contexts=list(data.get("sample_contexts", []) or []),
            occurrence_count=int(data.get("occurrence_count", 1) or 1),
        )


@dataclass
class ReviewedTerm:
    review_id: str
    surface_form: str
    term_type: str
    decision: str
    source_locations: List[Dict[str, Any]]
    decision_reason: str
    sample_contexts: List[str] = field(default_factory=list)
    occurrence_count: int = 1
    review_reason: str = ""

    def to_dict(self) -> Dict[str, Any]:
        data = asdict(self)
        data["source_locations"] = list(data.get("source_locations", []) or [])[:10]
        data["sample_contexts"] = list(data.get("sample_contexts", []) or [])[:3]
        return data

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "ReviewedTerm":
        return cls(
            review_id=str(data.get("review_id", "")),
            surface_form=str(data.get("surface_form", "")),
            term_type=str(data.get("term_type", "")),
            decision=str(data.get("decision", "")),
            source_locations=list(data.get("source_locations", []) or []),
            decision_reason=str(data.get("decision_reason", "")),
            sample_contexts=list(data.get("sample_contexts", []) or []),
            occurrence_count=int(data.get("occurrence_count", 1) or 1),
            review_reason=str(data.get("review_reason", "")),
        )


@dataclass
class NonTransRegexRule:
    rule_id: str
    name: str
    pattern: str
    element_type: str
    open_pattern: str = ""
    close_pattern: str = ""
    empty_pattern: str = ""
    enabled: bool = True
    order_index: int = 0
    examples: List[str] = field(default_factory=list)
    source: str = ""
    conflict_notes: str = ""

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "NonTransRegexRule":
        return cls(
            rule_id=str(data.get("rule_id", "")),
            name=str(data.get("name", "")),
            pattern=str(data.get("pattern", "")),
            element_type=str(data.get("element_type", data.get("category", "")) or ""),
            open_pattern=str(data.get("open_pattern", data.get("open", "")) or ""),
            close_pattern=str(data.get("close_pattern", data.get("close", "")) or ""),
            empty_pattern=str(data.get("empty_pattern", data.get("empty", "")) or ""),
            enabled=bool(data.get("enabled", True)),
            order_index=int(data.get("order_index", 0) or 0),
            examples=list(data.get("examples", []) or []),
            source=str(data.get("source", "")),
            conflict_notes=str(data.get("conflict_notes", "")),
        )


@dataclass
class NonTransElement:
    element_id: str
    element: str
    element_type: str
    source_record_ids: List[str] = field(default_factory=list)
    sample_contexts: List[str] = field(default_factory=list)
    occurrence_count: int = 1

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "NonTransElement":
        return cls(
            element_id=str(data.get("element_id", "")),
            element=str(data.get("element", "")),
            element_type=str(data.get("element_type", data.get("category", "")) or ""),
            source_record_ids=list(data.get("source_record_ids", []) or []),
            sample_contexts=list(data.get("sample_contexts", []) or []),
            occurrence_count=int(data.get("occurrence_count", 1) or 1),
        )


@dataclass
class NonTransRegexRow:
    row_id: str
    rule_id: str
    name: str
    regex: str
    role: str
    element_type: str
    order_index: int = 0
    examples: List[str] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "NonTransRegexRow":
        return cls(
            row_id=str(data.get("row_id", "")),
            rule_id=str(data.get("rule_id", "")),
            name=str(data.get("name", "")),
            regex=str(data.get("regex", data.get("pattern", "")) or ""),
            role=str(data.get("role", "")),
            element_type=str(data.get("element_type", data.get("category", "")) or ""),
            order_index=int(data.get("order_index", 0) or 0),
            examples=list(data.get("examples", []) or []),
        )


@dataclass
class NonTransMatch:
    placeholder: str
    text: str
    rule_id: str
    row_id: str
    role: str
    element_type: str
    start: int
    end: int

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


@dataclass
class ProtectedText:
    original_text: str
    protected_text: str
    matches: List[NonTransMatch] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        data = asdict(self)
        data["matches"] = [item.to_dict() for item in self.matches]
        return data


@dataclass
class TermRecallCleanRecord:
    source_record_id: str
    original_text: str
    nontrans_protected_text: str
    term_recall_clean_text: str
    dedupe_key: str

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "TermRecallCleanRecord":
        return cls(
            source_record_id=str(data.get("source_record_id", "")),
            original_text=str(data.get("original_text", "")),
            nontrans_protected_text=str(data.get("nontrans_protected_text", "")),
            term_recall_clean_text=str(data.get("term_recall_clean_text", "")),
            dedupe_key=str(data.get("dedupe_key", "")),
        )


@dataclass
class FailureRecord:
    task_type: str
    item_id: str
    reason: str
    guidance: str = ""
    stage: str = ""
    source_excerpt: str = ""
    attempt_count: int = 0

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "FailureRecord":
        return cls(
            task_type=str(data.get("task_type", "")),
            item_id=str(data.get("item_id", "")),
            reason=str(data.get("reason", "")),
            guidance=str(data.get("guidance", "")),
            stage=str(data.get("stage", "")),
            source_excerpt=str(data.get("source_excerpt", "")),
            attempt_count=int(data.get("attempt_count", 0) or 0),
        )


@dataclass
class RuntimeTaskState:
    cache_version: int
    task_id: str
    stage: str
    created_at: str
    updated_at: str
    input_config: Dict[str, Any]
    provider_name: str
    model_name: str
    processed_files: List[str] = field(default_factory=list)
    completed_batch_ids: List[str] = field(default_factory=list)
    source_records: List[Dict[str, Any]] = field(default_factory=list)
    segments: List[Dict[str, Any]] = field(default_factory=list)
    candidate_terms: List[Dict[str, Any]] = field(default_factory=list)
    reviewed_terms: List[Dict[str, Any]] = field(default_factory=list)
    extracted_terms: List[str] = field(default_factory=list)
    analyzed_results: List[Dict[str, Any]] = field(default_factory=list)
    approved_results: List[Dict[str, Any]] = field(default_factory=list)
    review_results: List[Dict[str, Any]] = field(default_factory=list)
    failed_items: List[FailureRecord] = field(default_factory=list)
    stats: Dict[str, Any] = field(default_factory=dict)
    output_file: str = ""
    last_error: str = ""

    def touch(self) -> None:
        self.updated_at = now_iso()

    def to_dict(self) -> Dict[str, Any]:
        extraction_mode = normalize_extraction_mode(self.input_config.get("extraction_mode", "terms"))
        stage = str(self.stage or "")
        persist_source_records = list(self.source_records)
        persist_segments = list(self.segments)
        persist_candidate_terms = list(self.candidate_terms)
        persist_reviewed_terms = list(self.reviewed_terms)

        if extraction_mode == "terms" and persist_segments:
            persist_source_records = []
        if persist_reviewed_terms:
            persist_candidate_terms = []
        if stage in {"AGGREGATING_TERMS", "EXPORTING", "COMPLETED"} and persist_reviewed_terms:
            persist_segments = []

        return {
            "cache_version": self.cache_version,
            "task_id": self.task_id,
            "stage": self.stage,
            "created_at": self.created_at,
            "updated_at": self.updated_at,
            "input_config": dict(self.input_config),
            "provider_name": self.provider_name,
            "model_name": self.model_name,
            "processed_files": list(self.processed_files),
            "completed_batch_ids": list(self.completed_batch_ids),
            "source_records": persist_source_records,
            "segments": persist_segments,
            "candidate_terms": persist_candidate_terms,
            "reviewed_terms": persist_reviewed_terms,
            "extracted_terms": list(self.extracted_terms),
            "analyzed_results": list(self.analyzed_results),
            "approved_results": list(self.approved_results),
            "review_results": list(self.review_results),
            "failed_items": [item.to_dict() for item in self.failed_items],
            "stats": dict(self.stats),
            "output_file": self.output_file,
            "last_error": self.last_error,
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "RuntimeTaskState":
        raw_completed = list(data.get("completed_batch_ids", []) or [])
        return cls(
            cache_version=int(data.get("cache_version", 1)),
            task_id=str(data.get("task_id", "")),
            stage=str(data.get("stage", "IDLE")),
            created_at=str(data.get("created_at", now_iso())),
            updated_at=str(data.get("updated_at", now_iso())),
            input_config=dict(data.get("input_config", {}) or {}),
            provider_name=str(data.get("provider_name", "")),
            model_name=str(data.get("model_name", "")),
            processed_files=list(data.get("processed_files", []) or []),
            completed_batch_ids=[str(item) for item in raw_completed],
            source_records=list(data.get("source_records", []) or []),
            segments=list(data.get("segments", []) or []),
            candidate_terms=list(data.get("candidate_terms", []) or []),
            reviewed_terms=list(data.get("reviewed_terms", []) or []),
            extracted_terms=list(data.get("extracted_terms", []) or []),
            analyzed_results=list(data.get("analyzed_results", []) or []),
            approved_results=list(data.get("approved_results", []) or []),
            review_results=list(data.get("review_results", []) or []),
            failed_items=[FailureRecord.from_dict(item) for item in list(data.get("failed_items", []) or [])],
            stats=dict(data.get("stats", {}) or {}),
            output_file=str(data.get("output_file", "")),
            last_error=str(data.get("last_error", "")),
        )


@dataclass
class FolderScanResult:
    folder_path: str
    file_type: str
    file_count: int
    headers: List[str]
    files: List[str]

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


@dataclass
class LLMRequest:
    task_id: str
    task_type: str
    prompt: str
    messages: List[Dict[str, str]] = field(default_factory=list)
    metadata: Dict[str, Any] = field(default_factory=dict)


@dataclass
class LLMResponse:
    task_id: str
    task_type: str
    content: str
    provider: str
    model: str
    latency_ms: int
    attempts: int
    success: bool
    error: str = ""
    error_type: str = ""
    retryable: bool = False
    response_metadata: Dict[str, Any] = field(default_factory=dict)


@dataclass
class SchedulerSnapshot:
    processed_count: int
    total_count: int
    current_concurrency: int
    success_count: int
    failure_count: int
    retry_count: int
