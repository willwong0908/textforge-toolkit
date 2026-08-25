"""Persistent storage helpers."""

from __future__ import annotations

import json
import os
import re
import sys
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Optional

from .constants import (
    CONFIG_VERSION,
    DEFAULT_CANDIDATE_SYSTEM_PROMPT,
    DEFAULT_CANDIDATE_USER_PROMPT,
    DEFAULT_CLASSIFICATION_SYSTEM_PROMPT,
    DEFAULT_CLASSIFICATION_USER_PROMPT,
    DEFAULT_NONTRANS_DISCOVERY_SYSTEM_PROMPT,
    DEFAULT_NONTRANS_DISCOVERY_USER_PROMPT,
    DEFAULT_NONTRANS_REGEX_SYSTEM_PROMPT,
    DEFAULT_NONTRANS_REGEX_USER_PROMPT,
    DEFAULT_NONTRANS_REORDER_SYSTEM_PROMPT,
    DEFAULT_NONTRANS_REORDER_USER_PROMPT,
    DEFAULT_REQUEST_LIMITS,
    DEFAULT_UI_PREFERENCES,
    PROVIDER_PRESETS,
    RUNTIME_CACHE_VERSION,
)
from .models import AppSettings, NonTransRegexRule, ProviderSettings, RuntimeTaskState, now_iso, sync_extraction_flags


@dataclass
class AppPaths:
    root_dir: Path
    output_dir: Path
    debug_failed_responses_dir: Path
    task_logs_dir: Path
    settings_file: Path
    runtime_cache_file: Path
    log_file: Path


def get_app_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


def get_app_paths() -> AppPaths:
    root_dir = get_app_root()
    output_dir = root_dir / "output"
    return AppPaths(
        root_dir=root_dir,
        output_dir=output_dir,
        debug_failed_responses_dir=output_dir / "debug_failed_responses",
        task_logs_dir=output_dir / "task_logs",
        settings_file=root_dir / "settings.json",
        runtime_cache_file=output_dir / "runtime_cache.json",
        log_file=output_dir / "log.txt",
    )


def ensure_directories(paths: AppPaths) -> None:
    paths.output_dir.mkdir(parents=True, exist_ok=True)
    paths.debug_failed_responses_dir.mkdir(parents=True, exist_ok=True)
    paths.task_logs_dir.mkdir(parents=True, exist_ok=True)


def atomic_write_json(path: Path, data: Dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    serialized = json.dumps(data, ensure_ascii=False, indent=2)
    temp_path = path.with_suffix(path.suffix + ".tmp")

    last_error = None
    for attempt in range(6):
        try:
            temp_path.write_text(serialized, encoding="utf-8")
            os.replace(str(temp_path), str(path))
            return
        except PermissionError as exc:
            last_error = exc
            time.sleep(0.15 * (attempt + 1))
        except OSError as exc:
            last_error = exc
            if getattr(exc, "winerror", None) != 5:
                raise
            time.sleep(0.15 * (attempt + 1))

    try:
        if temp_path.exists():
            temp_path.unlink()
    except OSError:
        pass

    try:
        path.write_text(serialized, encoding="utf-8")
        return
    except Exception:
        pass

    if last_error is not None:
        raise PermissionError(
            "无法写入缓存文件：{0}。请将程序放到非同步目录，或关闭占用该文件的同步/安全软件后重试。".format(path)
        ) from last_error
    raise PermissionError("无法写入缓存文件：{0}".format(path))


def _detect_legacy_provider(_legacy: Dict[str, Any]) -> str:
    return "DeepSeek"


def build_default_settings() -> AppSettings:
    provider_settings = {}
    for provider_name, preset in PROVIDER_PRESETS.items():
        provider_settings[provider_name] = ProviderSettings(
            api_key="",
            base_url=str(preset["base_url"]),
            model=str(preset["default_model"]),
            timeout_seconds=90,
            max_concurrency=int(preset["max_concurrency"]),
            extra_headers={},
            disable_system_proxy=True,
        )

    settings = AppSettings(
        config_version=CONFIG_VERSION,
        input_defaults={
            "folder_path": "",
            "header_name": "",
            "source_language": "中文",
            "export_review_sheet": False,
            "ascii_filter_blacklist": [
                "playername",
                "npcid",
                "skilllevel",
                "itemiconpath",
                "item_id",
                "quest_name",
                "gathering_rate_mod",
                "gathering_loot_ratio_add",
                "advanced_gathering_hit_rate",
                "advanced_rich_gathering_hit_rate",
                "top_gathering_hit_rate",
                "top_rich_gathering_hit_rate",
            ],
            "ascii_filter_whitelist": ["SSR", "SR", "UR", "SP", "Exergy"],
            "recall_scopes": [],
            "extraction_mode": "terms",
            "enable_nontrans_extraction": True,
            "enable_term_extraction": True,
            "nontrans_placeholder_format": "<{n}>",
            "numeric_normalization_enabled": True,
            "numeric_normalization_mode": "duplicate_group_only",
            "single_occurrence_approved_policy": "allow_to_library",
            "nontrans_stage_settings": {
                "chunk_char_limit": 20000,
                "enable_thinking": True,
                "reasoning_effort": "low",
                "builtin_regex_enabled": True,
                "ai_discovery_enabled": True,
                "ai_regex_generation_enabled": True,
            },
            "ascii_candidate_patterns": [
                {"name": "angle_brackets", "pattern": r"<[^<>]+>", "enabled": True, "order_index": 1},
                {"name": "brace_block", "pattern": r"\{[^{}]+\}", "enabled": True, "order_index": 2},
                {"name": "bracket_block", "pattern": r"\[[^\[\]]+\]", "enabled": True, "order_index": 3},
                {
                    "name": "dollar_var",
                    "pattern": r"\$(?:\{[^{}]+\}|[A-Za-z_][A-Za-z0-9_.\[\]-]*)",
                    "enabled": True,
                    "order_index": 4,
                },
                {
                    "name": "percent_token",
                    "pattern": r"%(?:\d+\$)?[#0\- +']*(?:\d+)?(?:\.\d+)?[A-Za-z]|%[A-Za-z_][A-Za-z0-9_]*%",
                    "enabled": True,
                    "order_index": 5,
                },
                {
                    "name": "escape_seq",
                    "pattern": r"\\(?:[nrtbfav\\'\"0]|u[0-9A-Fa-f]{4}|x[0-9A-Fa-f]{2})",
                    "enabled": True,
                    "order_index": 6,
                },
                {
                    "name": "html_entity",
                    "pattern": r"&(?:[A-Za-z][A-Za-z0-9]{1,31}|#\d{1,7}|#x[0-9A-Fa-f]{1,6});",
                    "enabled": True,
                    "order_index": 7,
                },
            ],
            "term_recall_stage_settings": {
                "single_item_char_limit": 500,
                "batch_request_char_limit": 20000,
                "enable_thinking": True,
                "reasoning_effort": "low",
            },
            "term_review_stage_settings": {
                "batch_request_char_limit": 20000,
                "max_context_chars": 220,
                "enable_thinking": True,
                "reasoning_effort": "low",
            },
            "ai_review_stage_settings": {
                "batch_request_char_limit": 20000,
                "max_items_per_request": 80,
                "enable_thinking": True,
                "workspace_enable_thinking": True,
                "reasoning_effort": "low",
                "auto_start_after_inspection": False,
                "debug_payload_logging": False,
            },
            "term_stage_settings": {
                "single_item_char_limit": 500,
                "batch_request_char_limit": 3000,
                "enable_thinking": False,
            },
        },
        provider_name="DeepSeek",
        provider_settings=provider_settings,
        prompt_templates={
            "candidate_system_prompt_template": DEFAULT_CANDIDATE_SYSTEM_PROMPT,
            "candidate_user_prompt_template": DEFAULT_CANDIDATE_USER_PROMPT,
            "classification_system_prompt_template": DEFAULT_CLASSIFICATION_SYSTEM_PROMPT,
            "classification_user_prompt_template": DEFAULT_CLASSIFICATION_USER_PROMPT,
            "nontrans_discovery_system_prompt_template": DEFAULT_NONTRANS_DISCOVERY_SYSTEM_PROMPT,
            "nontrans_discovery_user_prompt_template": DEFAULT_NONTRANS_DISCOVERY_USER_PROMPT,
            "nontrans_regex_system_prompt_template": DEFAULT_NONTRANS_REGEX_SYSTEM_PROMPT,
            "nontrans_regex_user_prompt_template": DEFAULT_NONTRANS_REGEX_USER_PROMPT,
            "nontrans_reorder_system_prompt_template": DEFAULT_NONTRANS_REORDER_SYSTEM_PROMPT,
            "nontrans_reorder_user_prompt_template": DEFAULT_NONTRANS_REORDER_USER_PROMPT,
        },
        request_limits=dict(DEFAULT_REQUEST_LIMITS),
        ui_preferences=dict(DEFAULT_UI_PREFERENCES),
    )
    sync_extraction_flags(settings.input_defaults)
    return settings


def _normalize_pending_nontrans_rule_item(data: Dict[str, Any], index: int = 0) -> Dict[str, Any]:
    direct_role = str(data.get("role", "") or "").strip()
    direct_regex = str(data.get("regex", "") or "").strip()
    if direct_role in {"open", "close", "empty"} and direct_regex:
        cache_id = str(data.get("cache_id", "") or data.get("pending_id", "")).strip() or "pending_rule_{0:04d}".format(index)
        created_at = str(data.get("created_at", "") or now_iso()).strip()
        updated_at = str(data.get("updated_at", "") or created_at).strip()
        return {
            "cache_id": cache_id,
            "rule_id": str(data.get("rule_id", "") or cache_id).strip() or cache_id,
            "name": str(data.get("name", "") or cache_id).strip() or cache_id,
            "role": direct_role,
            "regex": direct_regex,
            "element_type": str(data.get("element_type", "") or "other").strip() or "other",
            "enabled": bool(data.get("enabled", True)),
            "examples": [str(item or "").strip() for item in list(data.get("examples", []) or []) if str(item or "").strip()],
            "created_at": created_at,
            "updated_at": updated_at,
        }

    rule = NonTransRegexRule.from_dict(data)
    role = "empty"
    regex = str(rule.empty_pattern or "").strip()
    if str(rule.open_pattern or "").strip():
        role = "open"
        regex = str(rule.open_pattern or "").strip()
    elif str(rule.close_pattern or "").strip():
        role = "close"
        regex = str(rule.close_pattern or "").strip()
    elif not regex:
        regex = str(rule.pattern or "").strip()
    cache_id = str(data.get("cache_id", "") or data.get("pending_id", "")).strip() or "pending_rule_{0:04d}".format(index)
    created_at = str(data.get("created_at", "") or now_iso()).strip()
    updated_at = str(data.get("updated_at", "") or created_at).strip()
    return {
        "cache_id": cache_id,
        "rule_id": str(rule.rule_id or "").strip() or cache_id,
        "name": str(rule.name or "").strip() or cache_id,
        "role": role,
        "regex": regex,
        "element_type": str(rule.element_type or "").strip() or "other",
        "enabled": bool(data.get("enabled", rule.enabled)),
        "examples": [str(item or "").strip() for item in list(rule.examples or []) if str(item or "").strip()],
        "created_at": created_at,
        "updated_at": updated_at,
    }


def _merge_pending_nontrans_rule_entries(existing: list[Dict[str, Any]], incoming: list[Dict[str, Any]]) -> list[Dict[str, Any]]:
    merged: Dict[tuple[str, str, str], Dict[str, Any]] = {}
    ordered_keys = []
    for source in (existing, incoming):
        for index, item in enumerate(source, start=1):
            normalized = _normalize_pending_nontrans_rule_item(item, index=index)
            key = (
                str(normalized["role"]),
                str(normalized["element_type"]),
                str(normalized["regex"]),
            )
            current = merged.get(key)
            if current is None:
                merged[key] = normalized
                ordered_keys.append(key)
                continue
            seen_examples = list(current.get("examples", []) or [])
            for example in list(normalized.get("examples", []) or []):
                if example and example not in seen_examples:
                    seen_examples.append(example)
            current["examples"] = seen_examples[:3]
            if normalized.get("updated_at"):
                current["updated_at"] = normalized["updated_at"]
            if not current.get("name") and normalized.get("name"):
                current["name"] = normalized["name"]
            if not current.get("rule_id") and normalized.get("rule_id"):
                current["rule_id"] = normalized["rule_id"]
    ordered_items = [merged[key] for key in ordered_keys]
    return _prune_subsumed_pending_nontrans_rules(ordered_items)


def _rule_representative_examples(item: Dict[str, Any]) -> list[str]:
    examples = [str(example or "").strip() for example in list(item.get("examples", []) or []) if str(example or "").strip()]
    regex_text = str(item.get("regex", "") or "").strip()
    if not examples:
        return [regex_text] if regex_text else []
    if not regex_text:
        return examples[:3]
    try:
        own_compiled = re.compile(regex_text)
    except re.error:
        return examples[:3]
    matched = [example for example in examples if own_compiled.fullmatch(example)]
    if matched:
        return matched[:3]
    return [regex_text]


def _rule_examples_match_regex(item: Dict[str, Any], pattern: str) -> bool:
    examples = _rule_representative_examples(item)
    if not examples:
        return False
    try:
        compiled = re.compile(str(pattern or ""))
    except re.error:
        return False
    return all(compiled.fullmatch(example) for example in examples)


def _prune_subsumed_pending_nontrans_rules(items: list[Dict[str, Any]]) -> list[Dict[str, Any]]:
    pruned: list[Dict[str, Any]] = []
    for item in items:
        role = str(item.get("role", "") or "").strip()
        element_type = str(item.get("element_type", "") or "").strip()
        regex = str(item.get("regex", "") or "").strip()
        if not role or not element_type or not regex:
            pruned.append(item)
            continue

        replaced_existing = False
        kept_items: list[Dict[str, Any]] = []
        for existing in pruned:
            existing_role = str(existing.get("role", "") or "").strip()
            existing_type = str(existing.get("element_type", "") or "").strip()
            existing_regex = str(existing.get("regex", "") or "").strip()
            same_bucket = role == existing_role and element_type == existing_type
            if not same_bucket:
                kept_items.append(existing)
                continue

            new_covers_existing = regex != existing_regex and _rule_examples_match_regex(existing, regex)
            existing_covers_new = regex != existing_regex and _rule_examples_match_regex(item, existing_regex)

            if new_covers_existing and not existing_covers_new:
                merged_examples = [str(example or "").strip() for example in list(item.get("examples", []) or []) if str(example or "").strip()]
                for example in list(existing.get("examples", []) or []):
                    cleaned = str(example or "").strip()
                    if cleaned and cleaned not in merged_examples:
                        merged_examples.append(cleaned)
                item["examples"] = merged_examples[:3]
                replaced_existing = True
                continue

            if existing_covers_new and not new_covers_existing:
                merged_examples = [str(example or "").strip() for example in list(existing.get("examples", []) or []) if str(example or "").strip()]
                for example in list(item.get("examples", []) or []):
                    cleaned = str(example or "").strip()
                    if cleaned and cleaned not in merged_examples:
                        merged_examples.append(cleaned)
                existing["examples"] = merged_examples[:3]
                kept_items.append(existing)
                replaced_existing = True
                break

            kept_items.append(existing)

        pruned = kept_items
        if replaced_existing and any(
            role == str(existing.get("role", "") or "").strip()
            and element_type == str(existing.get("element_type", "") or "").strip()
            and regex == str(existing.get("regex", "") or "").strip()
            for existing in pruned
        ):
            continue
        pruned.append(item)
    return pruned


def load_pending_nontrans_rule_imports(settings: AppSettings) -> list[Dict[str, Any]]:
    items = list(settings.ui_preferences.get("pending_nontrans_rule_imports", []) or [])
    return _merge_pending_nontrans_rule_entries([], items)


def append_pending_nontrans_rule_imports(
    settings: AppSettings,
    rules: list[NonTransRegexRule],
) -> list[Dict[str, Any]]:
    existing = load_pending_nontrans_rule_imports(settings)
    next_index = len(existing) + 1
    incoming = []
    for offset, rule in enumerate(rules, start=0):
        incoming.append(
            _normalize_pending_nontrans_rule_item(
                {
                    **rule.to_dict(),
                    "cache_id": "pending_rule_{0:04d}".format(next_index + offset),
                    "created_at": now_iso(),
                    "updated_at": now_iso(),
                },
                index=next_index + offset,
            )
        )
    merged = _merge_pending_nontrans_rule_entries(existing, incoming)
    settings.ui_preferences["pending_nontrans_rule_imports"] = merged
    settings.ui_preferences["pending_nontrans_rule_notice_seen"] = False
    settings.ui_preferences["pending_nontrans_rule_library_seen"] = False
    return merged


def mark_pending_nontrans_rule_notice(
    settings: AppSettings,
    *,
    notice_seen: Optional[bool] = None,
    library_seen: Optional[bool] = None,
) -> None:
    if notice_seen is not None:
        settings.ui_preferences["pending_nontrans_rule_notice_seen"] = bool(notice_seen)
    if library_seen is not None:
        settings.ui_preferences["pending_nontrans_rule_library_seen"] = bool(library_seen)


def clear_pending_nontrans_rule_imports(
    settings: AppSettings,
    cache_ids: Optional[list[str]] = None,
) -> list[Dict[str, Any]]:
    existing = load_pending_nontrans_rule_imports(settings)
    if not cache_ids:
        remaining: list[Dict[str, Any]] = []
    else:
        id_set = {str(item or "").strip() for item in cache_ids if str(item or "").strip()}
        remaining = [item for item in existing if str(item.get("cache_id", "")).strip() not in id_set]
    settings.ui_preferences["pending_nontrans_rule_imports"] = remaining
    if not remaining:
        settings.ui_preferences["pending_nontrans_rule_notice_seen"] = True
        settings.ui_preferences["pending_nontrans_rule_library_seen"] = True
    return remaining


class SettingsStore:
    def __init__(self, paths: Optional[AppPaths] = None):
        self.paths = paths or get_app_paths()
        ensure_directories(self.paths)

    def load(self) -> AppSettings:
        if self.paths.settings_file.exists():
            raw = json.loads(self.paths.settings_file.read_text(encoding="utf-8"))
            settings = AppSettings.from_dict(raw)
            settings = self._merge_missing_defaults(settings)
            self.save(settings)
            return settings

        settings = build_default_settings()
        self.save(settings)
        return settings

    def save(self, settings: AppSettings) -> None:
        atomic_write_json(self.paths.settings_file, settings.to_dict())

    def migrate_legacy_settings(self, legacy: Dict[str, Any]) -> AppSettings:
        settings = build_default_settings()
        provider_name = _detect_legacy_provider(legacy)
        settings.provider_name = provider_name

        provider_config = settings.provider_settings.get(provider_name, ProviderSettings())
        provider_config.api_key = str(legacy.get("openai_api_key", provider_config.api_key))
        provider_config.base_url = str(legacy.get("openai_base_url", provider_config.base_url))
        provider_config.model = str(legacy.get("openai_model", provider_config.model))
        provider_config.max_concurrency = int(
            legacy.get("openai_concurrency", provider_config.max_concurrency) or provider_config.max_concurrency
        )
        provider_config.disable_system_proxy = bool(legacy.get("disable_system_proxy", True))
        settings.provider_settings[provider_name] = provider_config

        legacy_batch_limit = int(legacy.get("max_chars_per_request", settings.request_limits["batch_request_char_limit"]))
        settings.request_limits["single_item_char_limit"] = int(
            legacy.get("single_item_char_limit", settings.request_limits["single_item_char_limit"])
        )
        settings.request_limits["batch_request_char_limit"] = int(
            legacy.get("batch_request_char_limit", legacy_batch_limit)
        )
        settings.request_limits["manual_concurrency"] = int(
            legacy.get("openai_concurrency", settings.request_limits["manual_concurrency"])
        )
        settings.input_defaults["source_language"] = str(
            legacy.get("source_language", settings.input_defaults["source_language"])
        )
        settings.input_defaults["export_review_sheet"] = bool(
            legacy.get("enable_analysis", settings.input_defaults["export_review_sheet"])
        )
        sync_extraction_flags(settings.input_defaults)
        return settings

    def _merge_missing_defaults(self, settings: AppSettings) -> AppSettings:
        defaults = build_default_settings()
        settings.config_version = CONFIG_VERSION

        for key, value in defaults.input_defaults.items():
            settings.input_defaults.setdefault(key, value)
        sync_extraction_flags(settings.input_defaults)
        if "export_review_sheet" not in settings.input_defaults:
            settings.input_defaults["export_review_sheet"] = bool(settings.input_defaults.get("enable_analysis", False))

        legacy_batch_limit = settings.request_limits.pop("max_chars_per_request", None)
        if legacy_batch_limit is not None and "batch_request_char_limit" not in settings.request_limits:
            settings.request_limits["batch_request_char_limit"] = int(legacy_batch_limit or 3000)

        settings.request_limits.setdefault("single_item_char_limit", defaults.request_limits["single_item_char_limit"])
        settings.request_limits.setdefault("batch_request_char_limit", defaults.request_limits["batch_request_char_limit"])
        for key, value in defaults.request_limits.items():
            settings.request_limits.setdefault(key, value)

        settings.input_defaults.setdefault("nontrans_stage_settings", {})
        settings.input_defaults.setdefault("term_recall_stage_settings", {})
        settings.input_defaults.setdefault("term_review_stage_settings", {})
        settings.input_defaults.setdefault("ai_review_stage_settings", {})
        settings.input_defaults.setdefault("term_stage_settings", {})

        for key, value in defaults.input_defaults["nontrans_stage_settings"].items():
            settings.input_defaults["nontrans_stage_settings"].setdefault(key, value)

        legacy_term_stage = dict(settings.input_defaults.get("term_stage_settings", {}) or {})
        for key, value in defaults.input_defaults["term_recall_stage_settings"].items():
            settings.input_defaults["term_recall_stage_settings"].setdefault(key, legacy_term_stage.get(key, value))
        for key, value in defaults.input_defaults["term_review_stage_settings"].items():
            settings.input_defaults["term_review_stage_settings"].setdefault(key, legacy_term_stage.get(key, value))
        for key, value in defaults.input_defaults["ai_review_stage_settings"].items():
            settings.input_defaults["ai_review_stage_settings"].setdefault(key, value)
        for stage_key, limit_key in (
            ("nontrans_stage_settings", "chunk_char_limit"),
            ("term_recall_stage_settings", "batch_request_char_limit"),
            ("term_review_stage_settings", "batch_request_char_limit"),
        ):
            stage = settings.input_defaults[stage_key]
            try:
                configured_limit = int(stage.get(limit_key) or 0)
            except (TypeError, ValueError):
                configured_limit = 0
            if configured_limit <= 6000:
                stage[limit_key] = 20000
            stage["enable_thinking"] = True
            effort = str(stage.get("reasoning_effort") or "").strip().lower()
            stage["reasoning_effort"] = effort if effort in {"low", "medium", "high"} else "low"
        ai_review_stage = settings.input_defaults["ai_review_stage_settings"]
        try:
            configured_ai_review_limit = int(ai_review_stage.get("batch_request_char_limit") or 0)
        except (TypeError, ValueError):
            configured_ai_review_limit = 0
        if configured_ai_review_limit <= 6000:
            ai_review_stage["batch_request_char_limit"] = 20000
        ai_review_stage["enable_thinking"] = True
        ai_review_stage["workspace_enable_thinking"] = True
        ai_review_effort = str(ai_review_stage.get("reasoning_effort") or "").strip().lower()
        ai_review_stage["reasoning_effort"] = (
            "high" if ai_review_effort == "max" else ai_review_effort if ai_review_effort in {"low", "medium", "high"} else "low"
        )
        for key, value in defaults.input_defaults["term_stage_settings"].items():
            settings.input_defaults["term_stage_settings"].setdefault(key, value)

        ascii_patterns = list(settings.input_defaults.get("ascii_candidate_patterns", []) or [])
        settings.input_defaults["ascii_candidate_patterns"] = [
            item
            for item in ascii_patterns
            if not (
                isinstance(item, dict)
                and str(item.get("name", "")).strip() == "ascii_fragment"
                and str(item.get("pattern", "")).strip() == r"[A-Za-z_][A-Za-z0-9_./:\-]{3,}"
            )
        ]

        for key, value in defaults.prompt_templates.items():
            settings.prompt_templates.setdefault(key, value)

        current_nontrans_regex_user = str(
            settings.prompt_templates.get("nontrans_regex_user_prompt_template", "") or ""
        )
        if (
            "covered_ids" not in current_nontrans_regex_user
            or '"examples": ["{filter|Warning}", "{base|Warning}"]' in current_nontrans_regex_user
            or '"pattern": "\\\\{[A-Za-z_][A-Za-z0-9_]*\\\\|Warning\\\\}"' in current_nontrans_regex_user
        ):
            settings.prompt_templates["nontrans_regex_user_prompt_template"] = defaults.prompt_templates[
                "nontrans_regex_user_prompt_template"
            ]

        for legacy_key in (
            "candidate_prompt_template",
            "classification_prompt_template",
            "extract_prompt_template",
            "analyze_prompt_template",
        ):
            settings.prompt_templates.pop(legacy_key, None)

        for key, value in defaults.request_limits.items():
            settings.request_limits.setdefault(key, value)
        for key, value in defaults.ui_preferences.items():
            settings.ui_preferences.setdefault(key, value)

        defaults_provider_names = set(defaults.provider_settings.keys())
        settings.provider_settings = {
            name: provider
            for name, provider in settings.provider_settings.items()
            if name in defaults_provider_names
        }
        for name, provider_defaults in defaults.provider_settings.items():
            current = settings.provider_settings.get(name)
            if current is None:
                settings.provider_settings[name] = provider_defaults
                continue
            merged = ProviderSettings.from_dict(provider_defaults.to_dict())
            merged.api_key = current.api_key or merged.api_key
            merged.base_url = current.base_url or merged.base_url
            merged.model = current.model or merged.model
            merged.timeout_seconds = current.timeout_seconds or merged.timeout_seconds
            merged.max_concurrency = current.max_concurrency or merged.max_concurrency
            merged.extra_headers = current.extra_headers or merged.extra_headers
            merged.disable_system_proxy = current.disable_system_proxy
            settings.provider_settings[name] = merged

        if settings.provider_name not in settings.provider_settings:
            settings.provider_name = defaults.provider_name
        return settings


class RuntimeCacheStore:
    def __init__(self, paths: Optional[AppPaths] = None):
        self.paths = paths or get_app_paths()
        ensure_directories(self.paths)

    def load(self) -> Optional[RuntimeTaskState]:
        if not self.paths.runtime_cache_file.exists():
            return None
        try:
            raw = json.loads(self.paths.runtime_cache_file.read_text(encoding="utf-8"))
            return RuntimeTaskState.from_dict(raw)
        except Exception:
            return None

    def save(self, state: RuntimeTaskState) -> None:
        state.cache_version = RUNTIME_CACHE_VERSION
        state.touch()
        atomic_write_json(self.paths.runtime_cache_file, state.to_dict())

    def exists(self) -> bool:
        return self.paths.runtime_cache_file.exists()

    def clear(self) -> None:
        if self.paths.runtime_cache_file.exists():
            self.paths.runtime_cache_file.unlink()

    def summary(self) -> Optional[Dict[str, Any]]:
        state = self.load()
        if not state:
            return None
        return {
            "task_id": state.task_id,
            "stage": state.stage,
            "updated_at": state.updated_at,
            "folder_path": state.input_config.get("folder_path", ""),
            "header_name": state.input_config.get("header_name", ""),
            "source_language": state.input_config.get("source_language", ""),
            "provider_name": state.provider_name,
            "model_name": state.model_name,
            "stats": dict(state.stats),
            "failed_count": len(state.failed_items),
        }
