"""FastAPI WebUI entry point."""

from __future__ import annotations

import json
import os
import re
import subprocess
import sys
import tempfile
import threading
import traceback
import webbrowser
from pathlib import Path
from typing import Optional

import httpx
from fastapi import FastAPI, HTTPException, UploadFile, File, Form
from fastapi.responses import FileResponse, HTMLResponse, Response
from starlette.concurrency import run_in_threadpool
from pydantic import BaseModel, ConfigDict

if __package__:
    from .ai_review.cache_service import (
        create_batch as create_ai_review_batch,
        get_batch as get_ai_review_batch,
        get_batch_items as get_ai_review_batch_items,
        open_directory as open_ai_review_directory,
        replace_batch_items as replace_ai_review_batch_items,
        save_upload_file as save_ai_review_upload_file,
    )
    from .ai_review.config import OUTPUTS_DIR as AI_REVIEW_OUTPUTS_DIR
    from .ai_review.database import init_db as init_ai_review_db
    from .ai_review.conversation_routes import router as ai_review_conversation_router
    from .ai_review.term_base_routes import router as ai_review_term_base_router
    from .ai_review.directional_service import (
        delete_directional_template,
        get_directional_template,
        list_directional_templates,
        save_directional_template,
    )
    from .ai_review.excel_mapping_service import (
        delete_excel_mapping_preset,
        get_excel_mapping_preset,
        list_excel_mapping_presets,
        save_excel_mapping_preset,
    )
    from .ai_review.file_reader import (
        detect_file_type as detect_ai_review_file_type,
        read_excel_headers as read_ai_review_excel_headers,
        read_excel_items as read_ai_review_excel_items,
        read_excel_items_by_mapping as read_ai_review_excel_items_by_mapping,
        read_xliff_items as read_ai_review_xliff_items,
        read_xliff_language_metadata as read_ai_review_xliff_language_metadata,
    )
    from .ai_review.forbidden_service import (
        delete_forbidden_template,
        get_forbidden_template,
        list_forbidden_templates,
        save_forbidden_template,
    )
    from .ai_review.prompt_service import (
        delete_prompt_template as delete_ai_review_prompt_template,
        get_prompt_template as get_ai_review_prompt_template,
        list_prompt_templates as list_ai_review_prompt_templates,
        reset_default_prompt_template as reset_ai_review_prompt_template,
        save_prompt_template as save_ai_review_prompt_template,
    )
    from .ai_review.review_service import (
        ReviewTaskError,
        create_review_task,
        get_review_followup_messages,
        get_review_issue_results,
        get_review_logs,
        get_review_results,
        get_review_task,
        recover_interrupted_review_tasks,
        send_review_followup_message,
    )
    from .ai_review.shared_provider import (
        SharedProviderError,
        get_shared_ai_settings as get_ai_review_shared_ai_settings,
        list_models as list_ai_review_models,
        test_chat as test_ai_review_chat,
    )
    from .constants import APP_VERSION, UPDATE_ASSET_NAME_HINTS, UPDATE_RELEASE_API
    from .core import scan_folder
    from .cross_excel import merge_excel_files_by_headers, scan_cross_excel_folder, search_excel_rows
    from .diff_excel import (
        DiffRecord as DiffExcelRecord,
        apply_highlight_to_records as apply_diff_excel_highlight,
        apply_highlight_from_cache as apply_diff_excel_highlight_from_cache,
        export_cached_diff_records as export_diff_excel_cached_records,
        read_cached_diff_preview as read_diff_excel_cached_preview,
        run_compare_to_cache as run_diff_excel_compare_to_cache,
        scan_field_match_headers as scan_diff_excel_field_match_headers,
    )
    from .diff_task_service import DiffTaskService
    from .feedback import FeedbackError, feedback_status, submit_feedback
    from .models import TaskInput, normalize_extraction_mode, sync_extraction_flags
    from .open_utils import open_folder as open_path_folder
    from .open_utils import open_path as open_any_path
    from .open_utils import open_spreadsheet_cell
    from .nontrans import (
        NONTRANS_ELEMENT_TYPE_LABELS,
        NONTRANS_ROLE_LABELS,
        BUILTIN_NONTRANS_LIBRARY_FILE,
        deduplicate_nontrans_regex_rows,
        expand_nontrans_regex_rows,
        load_builtin_nontrans_rules,
        save_builtin_nontrans_rules,
        validate_nontrans_rule,
    )
    from .providers import ProviderRegistry
    from .service_layer import ExtractionTaskFacade
    from .storage import (
        append_pending_nontrans_rule_imports,
        build_default_settings,
        clear_pending_nontrans_rule_imports,
        get_app_root,
        load_pending_nontrans_rule_imports,
        mark_pending_nontrans_rule_notice,
    )
    from .telemetry import track_event
else:
    package_root = Path(__file__).resolve().parent.parent
    if str(package_root) not in sys.path:
        sys.path.insert(0, str(package_root))
    from term_extractor_app.ai_review.cache_service import (
        create_batch as create_ai_review_batch,
        get_batch as get_ai_review_batch,
        get_batch_items as get_ai_review_batch_items,
        open_directory as open_ai_review_directory,
        replace_batch_items as replace_ai_review_batch_items,
        save_upload_file as save_ai_review_upload_file,
    )
    from term_extractor_app.ai_review.config import OUTPUTS_DIR as AI_REVIEW_OUTPUTS_DIR
    from term_extractor_app.ai_review.database import init_db as init_ai_review_db
    from term_extractor_app.ai_review.conversation_routes import router as ai_review_conversation_router
    from term_extractor_app.ai_review.term_base_routes import router as ai_review_term_base_router
    from term_extractor_app.ai_review.directional_service import (
        delete_directional_template,
        get_directional_template,
        list_directional_templates,
        save_directional_template,
    )
    from term_extractor_app.ai_review.excel_mapping_service import (
        delete_excel_mapping_preset,
        get_excel_mapping_preset,
        list_excel_mapping_presets,
        save_excel_mapping_preset,
    )
    from term_extractor_app.ai_review.file_reader import (
        detect_file_type as detect_ai_review_file_type,
        read_excel_headers as read_ai_review_excel_headers,
        read_excel_items as read_ai_review_excel_items,
        read_excel_items_by_mapping as read_ai_review_excel_items_by_mapping,
        read_xliff_items as read_ai_review_xliff_items,
        read_xliff_language_metadata as read_ai_review_xliff_language_metadata,
    )
    from term_extractor_app.ai_review.forbidden_service import (
        delete_forbidden_template,
        get_forbidden_template,
        list_forbidden_templates,
        save_forbidden_template,
    )
    from term_extractor_app.ai_review.prompt_service import (
        delete_prompt_template as delete_ai_review_prompt_template,
        get_prompt_template as get_ai_review_prompt_template,
        list_prompt_templates as list_ai_review_prompt_templates,
        reset_default_prompt_template as reset_ai_review_prompt_template,
        save_prompt_template as save_ai_review_prompt_template,
    )
    from term_extractor_app.ai_review.review_service import (
        ReviewTaskError,
        create_review_task,
        get_review_followup_messages,
        get_review_issue_results,
        get_review_logs,
        get_review_results,
        get_review_task,
        recover_interrupted_review_tasks,
        send_review_followup_message,
    )
    from term_extractor_app.ai_review.shared_provider import (
        SharedProviderError,
        get_shared_ai_settings as get_ai_review_shared_ai_settings,
        list_models as list_ai_review_models,
        test_chat as test_ai_review_chat,
    )
    from term_extractor_app.constants import APP_VERSION, UPDATE_ASSET_NAME_HINTS, UPDATE_RELEASE_API
    from term_extractor_app.core import scan_folder
    from term_extractor_app.cross_excel import (
        merge_excel_files_by_headers,
        scan_cross_excel_folder,
        search_excel_rows,
    )
    from term_extractor_app.diff_excel import (
        DiffRecord as DiffExcelRecord,
        apply_highlight_to_records as apply_diff_excel_highlight,
        apply_highlight_from_cache as apply_diff_excel_highlight_from_cache,
        export_cached_diff_records as export_diff_excel_cached_records,
        read_cached_diff_preview as read_diff_excel_cached_preview,
        run_compare_to_cache as run_diff_excel_compare_to_cache,
    )
    from term_extractor_app.diff_task_service import DiffTaskService
    from term_extractor_app.feedback import FeedbackError, feedback_status, submit_feedback
    from term_extractor_app.models import TaskInput, normalize_extraction_mode, sync_extraction_flags
    from term_extractor_app.open_utils import open_folder as open_path_folder
    from term_extractor_app.open_utils import open_path as open_any_path
    from term_extractor_app.open_utils import open_spreadsheet_cell
    from term_extractor_app.nontrans import (
        NONTRANS_ELEMENT_TYPE_LABELS,
        NONTRANS_ROLE_LABELS,
        BUILTIN_NONTRANS_LIBRARY_FILE,
        deduplicate_nontrans_regex_rows,
        expand_nontrans_regex_rows,
        load_builtin_nontrans_rules,
        save_builtin_nontrans_rules,
        validate_nontrans_rule,
    )
    from term_extractor_app.providers import ProviderRegistry
    from term_extractor_app.service_layer import ExtractionTaskFacade
    from term_extractor_app.storage import (
        append_pending_nontrans_rule_imports,
        build_default_settings,
        clear_pending_nontrans_rule_imports,
        get_app_root,
        load_pending_nontrans_rule_imports,
        mark_pending_nontrans_rule_notice,
    )
    from term_extractor_app.telemetry import track_event


class StartTaskPayload(BaseModel):
    folder_path: str
    header_name: str = ""
    source_language: str = "中文"
    file_type: str = ""
    export_review_sheet: bool = False
    extraction_mode: str = "terms"
    single_item_char_limit: int = 500
    batch_request_char_limit: int = 3000
    resume: bool = False


class SettingsPayload(BaseModel):
    model_config = ConfigDict(protected_namespaces=())

    provider_name: Optional[str] = None
    model_name: Optional[str] = None
    api_key: Optional[str] = None
    base_url: Optional[str] = None
    timeout_seconds: Optional[int] = None
    disable_system_proxy: Optional[bool] = None
    extraction_mode: Optional[str] = None
    source_language: Optional[str] = None
    nontrans_chunk_char_limit: Optional[int] = None
    nontrans_placeholder_format: Optional[str] = None
    term_recall_batch_char_limit: Optional[int] = None
    term_review_batch_char_limit: Optional[int] = None
    term_review_max_context_chars: Optional[int] = None
    ai_review_batch_char_limit: Optional[int] = None
    ai_review_max_items_per_request: Optional[int] = None
    ai_review_workspace_enable_thinking: Optional[bool] = None
    ai_review_debug_payload_logging: Optional[bool] = None
    ai_review_reasoning_effort: Optional[str] = None
    nontrans_reasoning_effort: Optional[str] = None
    term_recall_reasoning_effort: Optional[str] = None
    term_review_reasoning_effort: Optional[str] = None
    nontrans_enable_thinking: Optional[bool] = None
    term_recall_enable_thinking: Optional[bool] = None
    term_review_enable_thinking: Optional[bool] = None
    ai_review_enable_thinking: Optional[bool] = None
    builtin_regex_enabled: Optional[bool] = None
    ai_discovery_enabled: Optional[bool] = None
    ai_regex_generation_enabled: Optional[bool] = None
    numeric_normalization_enabled: Optional[bool] = None


class AsciiCandidatePatternPayload(BaseModel):
    name: str = ""
    pattern: str
    enabled: bool = True
    order_index: int = 0


class AsciiCandidatePatternsPayload(BaseModel):
    patterns: list[AsciiCandidatePatternPayload]


class BuiltinNonTransRulePayload(BaseModel):
    rule_id: str = ""
    name: str = ""
    role: str = "empty"
    element_type: str = ""
    pattern: str = ""
    open_pattern: str = ""
    close_pattern: str = ""
    empty_pattern: str = ""
    enabled: bool = True
    examples: list[str] = []


class BuiltinNonTransRulesPayload(BaseModel):
    rules: list[BuiltinNonTransRulePayload]


class PendingNonTransRuleItemPayload(BaseModel):
    cache_id: str = ""
    rule_id: str = ""
    name: str = ""
    role: str = "empty"
    element_type: str = "other"
    pattern: str = ""
    enabled: bool = True
    examples: list[str] = []


class PendingNonTransRuleImportPayload(BaseModel):
    rules: list[PendingNonTransRuleItemPayload]


class PendingNonTransRuleSeenPayload(BaseModel):
    notice_seen: Optional[bool] = None
    library_seen: Optional[bool] = None


class PromptTemplatesPayload(BaseModel):
    templates: dict[str, str]


class PromptTemplateResetPayload(BaseModel):
    keys: list[str] = []


class AppUpdateStartPayload(BaseModel):
    force: bool = False


class ToolOpenPayload(BaseModel):
    tool_key: str


class CrossExcelSearchPayload(BaseModel):
    folder_path: str
    query: str
    limit: int = 300


class CrossExcelMergePayload(BaseModel):
    folder_path: str
    headers: list[str]
    apply_format: bool = True


class DiffExcelComparePayload(BaseModel):
    path_a: str
    path_b: str
    compare_mode: str = "position"
    reference_field: str = ""
    compare_fields: list[str] = []
    include_unmatched: bool = False


class DiffExcelExportPayload(BaseModel):
    cache_file: str
    query: str = ""
    output_file: str = ""


class DiffExcelHighlightPayload(BaseModel):
    cache_file: str
    query: str = ""
    target: str = "A"
    color_hex: str = "#FFD966"


class DiffExcelOpenCellPayload(BaseModel):
    file_path: str
    sheet_name: str
    cell_address: str


class AIReviewOpenFilePayload(BaseModel):
    file_path: str


class AIReviewSelectColumnsPayload(BaseModel):
    batch_id: str
    source_column: str
    target_column: str


class AIReviewExcelMappingPayload(BaseModel):
    batch_id: str
    mapping: dict


class AIReviewExcelMappingPresetSavePayload(BaseModel):
    id: str | None = None
    name: str
    mapping: dict


class AIReviewPromptTemplateSavePayload(BaseModel):
    id: str | None = None
    name: str
    system_prompt: str
    user_prompt: str
    forbidden_words_text: str = ""


class AIReviewDirectionalTemplateSavePayload(BaseModel):
    id: str | None = None
    name: str
    items: list[dict]


class AIReviewForbiddenTemplateSavePayload(BaseModel):
    id: str | None = None
    name: str
    words_text: str


class AIReviewStartPayload(BaseModel):
    batch_id: str
    prompt_template_id: str | None = None
    source_language: str = ""
    target_language: str = ""
    mode: str = "normal"
    directional_template_id: str | None = None
    enable_ai_review: bool = True
    enable_forbidden_check: bool = False
    forbidden_template_id: str | None = None


class AIReviewFollowupPayload(BaseModel):
    message: str


PROMPT_TEMPLATE_META = [
    {
        "key": "candidate_system_prompt_template",
        "label": "术语召回提示词（系统）",
        "description": "",
    },
    {
        "key": "candidate_user_prompt_template",
        "label": "术语召回提示词（用户）",
        "description": "",
    },
    {
        "key": "classification_system_prompt_template",
        "label": "术语校验提示词（系统）",
        "description": "",
    },
    {
        "key": "classification_user_prompt_template",
        "label": "术语校验提示词（用户）",
        "description": "",
    },
    {
        "key": "nontrans_discovery_system_prompt_template",
        "label": "非译发现提示词（系统）",
        "description": "",
    },
    {
        "key": "nontrans_discovery_user_prompt_template",
        "label": "非译发现提示词（用户）",
        "description": "",
    },
    {
        "key": "nontrans_regex_system_prompt_template",
        "label": "非译正则提示词（系统）",
        "description": "",
    },
    {
        "key": "nontrans_regex_user_prompt_template",
        "label": "非译正则提示词（用户）",
        "description": "",
    },
]


def allowed_prompt_template_keys() -> set[str]:
    return {str(item["key"]) for item in PROMPT_TEMPLATE_META}


def prompt_templates_response(settings):
    defaults = build_default_settings().prompt_templates
    templates = []
    for item in PROMPT_TEMPLATE_META:
        key = str(item["key"])
        value = str(settings.prompt_templates.get(key, defaults.get(key, "")) or "")
        default_value = str(defaults.get(key, "") or "")
        templates.append(
            {
                "key": key,
                "label": item["label"],
                "description": item["description"],
                "value": value,
                "is_default": value == default_value,
            }
        )
    return {"templates": templates}


def builtin_nontrans_rules_response():
    rules = load_builtin_nontrans_rules()
    rows = deduplicate_nontrans_regex_rows(expand_nontrans_regex_rows(rules))
    return {
        "rule_count": len(rules),
        "row_count": len(rows),
        "rows": [
            {
                "row_id": row.row_id,
                "rule_id": row.rule_id,
                "name": row.name,
                "regex": row.regex,
                "role": row.role,
                "role_label": NONTRANS_ROLE_LABELS.get(row.role, row.role),
                "element_type": row.element_type,
                "element_type_label": NONTRANS_ELEMENT_TYPE_LABELS.get(row.element_type, row.element_type),
                "order_index": index,
                "examples": list(row.examples or []),
            }
            for index, row in enumerate(rows, start=1)
        ],
        "rules": [
            {
                **rule.to_dict(),
                "role": (
                    "open"
                    if str(rule.open_pattern or "").strip()
                    else "close"
                    if str(rule.close_pattern or "").strip()
                    else "empty"
                ),
                "regex": (
                    str(rule.open_pattern or "").strip()
                    or str(rule.close_pattern or "").strip()
                    or str(rule.empty_pattern or "").strip()
                    or str(rule.pattern or "").strip()
                ),
            }
            for rule in rules
        ],
    }


def pending_nontrans_rules_response(settings) -> dict:
    items = load_pending_nontrans_rule_imports(settings)
    notice_seen = bool(settings.ui_preferences.get("pending_nontrans_rule_notice_seen", False))
    library_seen = bool(settings.ui_preferences.get("pending_nontrans_rule_library_seen", False))
    return {
        "count": len(items),
        "has_pending": bool(items),
        "notice_seen": notice_seen,
        "library_seen": library_seen,
        "show_notice_dot": bool(items) and not notice_seen,
        "show_library_dot": bool(items) and not library_seen,
        "show_notice_button": bool(items) and not notice_seen,
        "rules": items,
    }


def _normalize_version_parts(version_text: str) -> list[int]:
    normalized = str(version_text or "").strip().lower()
    if normalized.startswith("v"):
        normalized = normalized[1:]
    parts = []
    for chunk in normalized.split("."):
        chunk = chunk.strip()
        if not chunk:
            continue
        digits = "".join(char for char in chunk if char.isdigit())
        parts.append(int(digits or 0))
    return parts


def _is_remote_version_newer(current_version: str, remote_version: str) -> bool:
    left = _normalize_version_parts(current_version)
    right = _normalize_version_parts(remote_version)
    length = max(len(left), len(right), 1)
    left.extend([0] * (length - len(left)))
    right.extend([0] * (length - len(right)))
    return tuple(right) > tuple(left)


def _pick_update_asset(release_data: dict) -> Optional[dict]:
    assets = list(release_data.get("assets", []) or [])
    zip_assets = [
        item for item in assets
        if str(item.get("name", "") or "").lower().endswith(".zip")
        and str(item.get("browser_download_url", "") or "").strip()
    ]
    if not zip_assets:
        return None
    for hint in UPDATE_ASSET_NAME_HINTS:
        matched = next(
            (
                item for item in zip_assets
                if hint in str(item.get("name", "") or "").lower()
            ),
            None,
        )
        if matched is not None:
            return matched
    return zip_assets[0]


def _resolve_logo_path() -> Path:
    candidate_paths = [
        get_app_root() / "logo.png",
        get_app_root() / "_internal" / "logo.png",
        get_app_root().parent / "logo.png",
        Path(__file__).resolve().parent.parent / "logo.png",
    ]
    for path in candidate_paths:
        if path.exists() and path.is_file():
            return path
    raise FileNotFoundError("未找到 logo.png")


TOOL_OPEN_EVENT_MAP = {
    "home_guide": "tool_open.home_guide",
    "text_preprocess": "tool_open.text_preprocess",
    "ai_review": "tool_open.ai_review",
    "cross_excel": "tool_open.cross_excel",
    "diff_excel": "diff.open",
}


async def fetch_app_update_info() -> dict:
    current_version = APP_VERSION
    try:
        async with httpx.AsyncClient(timeout=15.0, follow_redirects=True) as client:
            response = await client.get(
                UPDATE_RELEASE_API,
                headers={
                    "Accept": "application/vnd.github+json",
                    "User-Agent": "Yeehe-Toolkit-Suite-Updater",
                },
            )
            response.raise_for_status()
            release_data = response.json()
    except httpx.HTTPStatusError as exc:
        if exc.response is not None and exc.response.status_code == 404:
            return {
                "supported": bool(getattr(sys, "frozen", False)),
                "current_version": current_version,
                "latest_version": current_version,
                "update_available": False,
                "release_notes": "",
                "published_at": "",
                "download_url": "",
                "asset_name": "",
                "message": "当前还没有可用的发布版本。",
            }
        return {
            "supported": bool(getattr(sys, "frozen", False)),
            "current_version": current_version,
            "latest_version": current_version,
            "update_available": False,
            "release_notes": "",
            "published_at": "",
            "download_url": "",
            "asset_name": "",
            "message": "检查更新失败：{0}".format(exc),
        }
    except Exception as exc:
        return {
            "supported": bool(getattr(sys, "frozen", False)),
            "current_version": current_version,
            "latest_version": current_version,
            "update_available": False,
            "release_notes": "",
            "published_at": "",
            "download_url": "",
            "asset_name": "",
            "message": "检查更新失败：{0}".format(exc),
        }

    latest_version = str(
        release_data.get("latest_version")
        or release_data.get("tag_name")
        or release_data.get("name")
        or current_version
    ).strip() or current_version
    asset = _pick_update_asset(release_data) or {}
    download_url = str(
        release_data.get("download_url")
        or asset.get("browser_download_url", "")
        or ""
    ).strip()
    asset_name = str(
        release_data.get("asset_name")
        or asset.get("name", "")
        or ""
    ).strip()
    return {
        "supported": bool(getattr(sys, "frozen", False)),
        "current_version": current_version,
        "latest_version": latest_version,
        "update_available": _is_remote_version_newer(current_version, latest_version),
        "release_notes": str(release_data.get("release_notes") or release_data.get("body", "") or "").strip(),
        "published_at": str(release_data.get("published_at", "") or "").strip(),
        "download_url": download_url,
        "asset_name": asset_name,
        "message": str(release_data.get("message", "") or "").strip(),
    }


def _build_update_powershell_script() -> str:
    return r"""
param(
  [string]$DownloadUrl,
  [string]$BundleRoot,
  [string]$LauncherName,
  [string]$ExeName,
  [int]$ParentPid
)

$ErrorActionPreference = "Stop"

function Write-Step {
  param([string]$Message)
  Write-Host ("[{0}] {1}" -f (Get-Date -Format "HH:mm:ss"), $Message) -ForegroundColor Cyan
}

function Find-BundlePayloadRoot {
  param([string]$ExtractRoot, [string]$ExeName, [string]$LauncherName)
  $roots = @($ExtractRoot)
  $dirs = Get-ChildItem -LiteralPath $ExtractRoot -Directory -Recurse -ErrorAction SilentlyContinue
  foreach ($dir in $dirs) {
    $roots += $dir.FullName
  }
  foreach ($root in $roots) {
    $launcherPath = Join-Path $root $LauncherName
    $programExe = Join-Path (Join-Path $root "program") $ExeName
    $directExe = Join-Path $root $ExeName
    if ((Test-Path -LiteralPath $launcherPath) -and (Test-Path -LiteralPath $programExe)) {
      return $root
    }
    if (Test-Path -LiteralPath $launcherPath) {
      return $root
    }
    if (Test-Path -LiteralPath $programExe) {
      return $root
    }
    if (Test-Path -LiteralPath $directExe) {
      return $root
    }
  }
  throw "未找到更新包中的启动文件。"
}

$workDir = Join-Path ([System.IO.Path]::GetTempPath()) ("yeehe_update_" + [guid]::NewGuid().ToString("N"))
$zipPath = Join-Path $workDir "update.zip"
$extractDir = Join-Path $workDir "extract"
New-Item -ItemType Directory -Path $workDir -Force | Out-Null
New-Item -ItemType Directory -Path $extractDir -Force | Out-Null

try {
  $host.UI.RawUI.WindowTitle = "译禾工具合集 更新中"
  Write-Host "==============================================" -ForegroundColor DarkGray
  Write-Host "译禾工具合集 正在更新" -ForegroundColor Yellow
  Write-Host "请不要关闭此窗口，更新完成后会自动重新打开程序。" -ForegroundColor Gray
  Write-Host "==============================================" -ForegroundColor DarkGray
  Write-Host ""
  Write-Step "开始下载更新包"
  Invoke-WebRequest -Uri $DownloadUrl -OutFile $zipPath -UseBasicParsing
  Write-Step "下载完成，正在解压"
  Expand-Archive -LiteralPath $zipPath -DestinationPath $extractDir -Force
  $payloadRoot = Find-BundlePayloadRoot -ExtractRoot $extractDir -ExeName $ExeName -LauncherName $LauncherName

  if ($ParentPid -gt 0) {
    Write-Step "正在关闭旧进程"
    try {
      Stop-Process -Id $ParentPid -Force -ErrorAction SilentlyContinue
    }
    catch {}
    for ($i = 0; $i -lt 120; $i++) {
      $proc = Get-Process -Id $ParentPid -ErrorAction SilentlyContinue
      if ($null -eq $proc) { break }
      Start-Sleep -Milliseconds 500
    }
  }

  Write-Step "正在替换程序文件"
  $programDir = Join-Path $BundleRoot "program"
  if (Test-Path -LiteralPath $programDir) {
    Remove-Item -LiteralPath $programDir -Recurse -Force -ErrorAction SilentlyContinue
  }
  Get-ChildItem -LiteralPath $payloadRoot -Force | ForEach-Object {
    Copy-Item -LiteralPath $_.FullName -Destination $BundleRoot -Recurse -Force
  }

  $launcherPath = Join-Path $BundleRoot $LauncherName
  Write-Step "更新完成，正在重新启动"
  if (Test-Path -LiteralPath $launcherPath) {
    Start-Process -FilePath "cmd.exe" -ArgumentList "/c", ('"{0}"' -f $launcherPath) -WorkingDirectory $BundleRoot | Out-Null
  }
  else {
    Start-Process -FilePath (Join-Path $BundleRoot "program\$ExeName") -WorkingDirectory (Join-Path $BundleRoot "program") | Out-Null
  }
  Write-Host ""
  Write-Host "更新完成，程序已重新启动。" -ForegroundColor Green
  Start-Sleep -Seconds 2
}
finally {
  Start-Sleep -Seconds 2
  Remove-Item -LiteralPath $workDir -Recurse -Force -ErrorAction SilentlyContinue
}
"""


def start_app_update(download_url: str) -> dict:
    if not getattr(sys, "frozen", False):
        raise RuntimeError("当前是源码运行模式，自动更新只支持发布版。")
    url = str(download_url or "").strip()
    if not url:
        raise RuntimeError("没有可用的更新下载地址。")

    app_root = Path(get_app_root()).resolve()
    exe_path = Path(sys.executable).resolve()
    bundle_root = app_root.parent if exe_path.parent.name.lower() == "program" else app_root
    launcher_name = "start_webui.bat"
    script_dir = Path(tempfile.mkdtemp(prefix="yeehe_update_"))
    script_path = script_dir / "run_update.ps1"
    # Windows PowerShell 5.x reads UTF-8 scripts reliably when BOM is present.
    # Without BOM, non-ASCII text in the updater script can be misparsed.
    script_path.write_text(_build_update_powershell_script(), encoding="utf-8-sig")

    subprocess.Popen(
        [
            "powershell",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(script_path),
            "-DownloadUrl",
            url,
            "-BundleRoot",
            str(bundle_root),
            "-LauncherName",
            launcher_name,
            "-ExeName",
            exe_path.name,
            "-ParentPid",
            str(os.getpid()),
        ],
        cwd=str(bundle_root),
        creationflags=(
            getattr(subprocess, "CREATE_NEW_CONSOLE", 0)
            | getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
        ),
    )
    return {
        "ok": True,
        "message": "更新已开始，当前页面会关闭，并弹出更新窗口显示进度。",
    }


def _save_builtin_nontrans_rules_to_library(rules_payload: list[BuiltinNonTransRulePayload]) -> dict:
    normalized_rules = []
    seen_rule_ids = set()

    for index, item in enumerate(rules_payload, start=1):
        rule_id = str(item.rule_id or "").strip() or "custom_rule_{0:03d}".format(index)
        if rule_id in seen_rule_ids:
            raise HTTPException(status_code=400, detail="规则 ID 重复：{0}".format(rule_id))
        seen_rule_ids.add(rule_id)

        examples = [str(example or "").strip() for example in list(item.examples or []) if str(example or "").strip()]
        role = str(item.role or "empty").strip() or "empty"
        regex_text = (
            str(item.pattern or "").strip()
            or str(item.open_pattern or "").strip()
            or str(item.close_pattern or "").strip()
            or str(item.empty_pattern or "").strip()
        )
        rule = item.model_dump()
        rule.update(
            {
                "rule_id": rule_id,
                "name": str(item.name or "").strip() or rule_id,
                "element_type": str(item.element_type or "").strip() or "other",
                "pattern": regex_text,
                "open_pattern": regex_text if role == "open" else "",
                "close_pattern": regex_text if role == "close" else "",
                "empty_pattern": regex_text if role == "empty" else "",
                "enabled": bool(item.enabled),
                "order_index": index,
                "examples": examples,
                "source": "builtin",
            }
        )
        model_rule = load_builtin_nontrans_rules()[0].from_dict(rule)
        issues = []
        if role not in {"open", "close", "empty"}:
            issues.append("命中方式无效")
        if not regex_text:
            issues.append("至少要填写一个正则")
        try:
            re.compile(regex_text)
        except re.error as exc:
            issues.append("正则编译失败：{0}".format(exc))
        if issues:
            raise HTTPException(
                status_code=400,
                detail="规则 {0} 校验失败：{1}".format(model_rule.name or model_rule.rule_id, "；".join(issues)),
            )
        normalized_rules.append(model_rule.to_dict())

    save_builtin_nontrans_rules([load_builtin_nontrans_rules()[0].from_dict(rule) for rule in normalized_rules])
    return builtin_nontrans_rules_response()


def _open_local_file(path_text: str) -> None:
    open_any_path(path_text)


def _open_excel_cell(path_text: str, sheet_name: str, cell_address: str) -> None:
    open_spreadsheet_cell(path_text, sheet_name, cell_address)


def _ai_review_batch_response(batch_id: str, message: str) -> dict:
    batch = get_ai_review_batch(batch_id)
    if not batch:
        raise HTTPException(status_code=404, detail="读取批次不存在")
    preview = get_ai_review_batch_items(batch_id, limit=5)
    metadata = dict(batch.get("metadata", {}) or {})
    return {
        "ok": True,
        "message": message,
        "batch": {
            "id": batch["id"],
            "filename": batch["original_filename"],
            "file_type": batch["file_type"],
            "source_column": batch["source_column"],
            "target_column": batch["target_column"],
            "item_count": batch["item_count"],
            "source_language": metadata.get("source_language", ""),
            "target_language": metadata.get("target_language", ""),
            "updated_at": batch["updated_at"],
            "metadata": metadata,
        },
        "preview": preview,
    }


def _diff_excel_record_from_dict(item: dict) -> DiffExcelRecord:
    return DiffExcelRecord(
        filename_a=str(item.get("filename_a", "") or ""),
        filename_b=str(item.get("filename_b", "") or ""),
        sheet=str(item.get("sheet", "") or ""),
        cell_address=str(item.get("cell_address", "") or ""),
        value_a=str(item.get("value_a", "") or ""),
        value_b=str(item.get("value_b", "") or ""),
        file_path_a=str(item.get("file_path_a", "") or ""),
        file_path_b=str(item.get("file_path_b", "") or ""),
    )


def _ai_review_excel_upload_response(
    *,
    batch_id: str,
    filename: str,
    metadata: dict,
    original_file_path: str = "",
) -> dict:
    metadata_with_path = {**metadata}
    if original_file_path:
        metadata_with_path["original_file_path"] = original_file_path
    return {
        "ok": True,
        "message": "Excel 文件已读取，请导入文本。",
        "file_type": "excel",
        "batch_id": batch_id,
        "filename": filename,
        "batch": {
            "id": batch_id,
            "filename": filename,
            "file_type": "excel",
            "source_column": "",
            "target_column": "",
            "item_count": 0,
            "source_language": "",
            "target_language": "",
            "updated_at": "",
            "metadata": metadata_with_path,
        },
        "preview": [],
        "headers": metadata["headers"],
        "headers_by_sheet": metadata["headers_by_sheet"],
        "columns_by_sheet": metadata["columns_by_sheet"],
        "sheet_names": metadata["sheet_names"],
        "needs_column_selection": False,
    }


def create_app(facade: Optional[ExtractionTaskFacade] = None) -> FastAPI:
    task_facade = facade or ExtractionTaskFacade()
    diff_task_service = DiffTaskService()
    app = FastAPI(title="AI Term Extractor WebUI")
    init_ai_review_db()
    recover_interrupted_review_tasks()
    app.include_router(ai_review_conversation_router)
    app.include_router(ai_review_term_base_router)

    @app.get("/", response_class=HTMLResponse)
    async def index() -> str:
        return build_index_html()

    @app.get("/assets/app.css")
    async def css() -> Response:
        return Response(APP_CSS, media_type="text/css")

    @app.get("/assets/app.js")
    async def js() -> Response:
        return Response(APP_JS, media_type="application/javascript")

    @app.get("/assets/logo.png")
    async def logo() -> FileResponse:
        try:
            logo_path = _resolve_logo_path()
        except FileNotFoundError as exc:
            raise HTTPException(status_code=404, detail=str(exc)) from exc
        return FileResponse(logo_path, media_type="image/png")

    @app.get("/api/feedback/status")
    async def get_feedback_status():
        return feedback_status()

    @app.post("/api/feedback/open-log")
    async def open_feedback_log():
        info = feedback_status()
        log_path = str(info.get("log_path", "") or "").strip()
        if not log_path:
            raise HTTPException(status_code=404, detail="未找到日志路径。")
        try:
            _open_local_file(log_path)
        except Exception as exc:
            raise HTTPException(status_code=500, detail="打开日志失败：{0}".format(exc)) from exc
        return {"ok": True, "log_path": log_path}

    @app.post("/api/feedback/submit")
    async def submit_feedback_endpoint(
        message: str = Form(...),
        screenshot: UploadFile | None = File(default=None),
    ):
        screenshot_bytes = None
        screenshot_name = ""
        screenshot_type = ""
        if screenshot is not None:
            screenshot_name = str(screenshot.filename or "").strip()
            screenshot_type = str(screenshot.content_type or "").strip()
            screenshot_bytes = await screenshot.read()
            if not screenshot_bytes:
                screenshot_bytes = None
        try:
            result = submit_feedback(
                message,
                screenshot_name=screenshot_name,
                screenshot_type=screenshot_type,
                screenshot_bytes=screenshot_bytes,
            )
        except FeedbackError as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc
        except Exception as exc:
            raise HTTPException(status_code=500, detail="反馈发送失败：{0}".format(exc)) from exc
        return result

    @app.post("/api/telemetry/tool-open")
    async def telemetry_tool_open(payload: ToolOpenPayload):
        event_name = TOOL_OPEN_EVENT_MAP.get(str(payload.tool_key or "").strip())
        if event_name:
            track_event(event_name)
        return {"ok": True}

    @app.get("/api/app-update")
    async def get_app_update():
        return await fetch_app_update_info()

    @app.post("/api/app-update/start")
    async def start_app_update_endpoint(payload: AppUpdateStartPayload):
        snapshot = task_facade.snapshot()
        if snapshot.is_running and not payload.force:
            raise HTTPException(status_code=409, detail="任务运行中，暂时不能更新。")
        update_info = await fetch_app_update_info()
        if not update_info.get("supported"):
            raise HTTPException(status_code=400, detail="当前运行方式不支持自动更新。")
        if not update_info.get("update_available"):
            raise HTTPException(status_code=400, detail="当前已经是最新版本。")
        if not str(update_info.get("download_url", "") or "").strip():
            raise HTTPException(status_code=400, detail="当前版本缺少可下载的更新包。")
        try:
            return start_app_update(str(update_info.get("download_url", "") or ""))
        except Exception as exc:
            raise HTTPException(status_code=500, detail=str(exc)) from exc

    @app.get("/api/settings")
    async def get_settings():
        settings = task_facade.load_settings()
        provider_settings = settings.provider_settings.get(settings.provider_name)
        nontrans = dict(settings.input_defaults.get("nontrans_stage_settings", {}) or {})
        recall = dict(settings.input_defaults.get("term_recall_stage_settings", {}) or {})
        review = dict(settings.input_defaults.get("term_review_stage_settings", {}) or {})
        return {
            "provider_name": settings.provider_name,
            "model_name": provider_settings.model if provider_settings else "",
            "api_key": provider_settings.api_key if provider_settings else "",
            "base_url": provider_settings.base_url if provider_settings else "",
            "timeout_seconds": provider_settings.timeout_seconds if provider_settings else 90,
            "disable_system_proxy": provider_settings.disable_system_proxy if provider_settings else True,
            "source_language": settings.input_defaults.get("source_language", "中文"),
            "extraction_mode": settings.input_defaults.get("extraction_mode", "terms"),
            "enable_nontrans_extraction": settings.input_defaults.get("enable_nontrans_extraction", True),
            "enable_term_extraction": settings.input_defaults.get("enable_term_extraction", True),
            "nontrans_placeholder_format": settings.input_defaults.get("nontrans_placeholder_format", "<{n}>"),
            "numeric_normalization_enabled": settings.input_defaults.get(
                "numeric_normalization_enabled", True
            ),
            "numeric_normalization_mode": settings.input_defaults.get(
                "numeric_normalization_mode", "duplicate_group_only"
            ),
            "nontrans_stage_settings": nontrans,
            "term_recall_stage_settings": recall,
            "term_review_stage_settings": review,
            "ai_review_stage_settings": dict(settings.input_defaults.get("ai_review_stage_settings", {}) or {}),
            "providers": list(settings.provider_settings.keys()),
            "pending_nontrans_rules": pending_nontrans_rules_response(settings),
        }

    @app.post("/api/settings")
    async def save_settings(payload: SettingsPayload):
        settings = task_facade.load_settings()
        if payload.provider_name and payload.provider_name in settings.provider_settings:
            settings.provider_name = payload.provider_name
        provider_settings = settings.provider_settings.get(settings.provider_name)
        if provider_settings:
            if payload.model_name is not None:
                provider_settings.model = payload.model_name
            if payload.api_key is not None:
                provider_settings.api_key = payload.api_key
            if payload.base_url is not None:
                provider_settings.base_url = payload.base_url
            if payload.timeout_seconds is not None:
                provider_settings.timeout_seconds = max(1, int(payload.timeout_seconds))
            if payload.disable_system_proxy is not None:
                provider_settings.disable_system_proxy = bool(payload.disable_system_proxy)
            settings.provider_settings[settings.provider_name] = provider_settings
        if payload.extraction_mode:
            settings.input_defaults["extraction_mode"] = normalize_extraction_mode(payload.extraction_mode)
            sync_extraction_flags(settings.input_defaults)
        if payload.source_language:
            settings.input_defaults["source_language"] = payload.source_language
        if payload.nontrans_placeholder_format is not None:
            settings.input_defaults["nontrans_placeholder_format"] = payload.nontrans_placeholder_format.strip() or "<{n}>"
        settings.input_defaults["single_occurrence_approved_policy"] = "allow_to_library"
        if payload.numeric_normalization_enabled is not None:
            settings.input_defaults["numeric_normalization_enabled"] = bool(
                payload.numeric_normalization_enabled
            )
            settings.input_defaults.setdefault("numeric_normalization_mode", "duplicate_group_only")

        nontrans = dict(settings.input_defaults.get("nontrans_stage_settings", {}) or {})
        if payload.nontrans_chunk_char_limit is not None:
            nontrans["chunk_char_limit"] = int(payload.nontrans_chunk_char_limit)
        if payload.nontrans_reasoning_effort is not None:
            effort = str(payload.nontrans_reasoning_effort or "low").strip().lower()
            nontrans["reasoning_effort"] = effort if effort in {"low", "medium", "high"} else "low"
        if payload.nontrans_enable_thinking is not None:
            nontrans["enable_thinking"] = bool(payload.nontrans_enable_thinking)
        if payload.builtin_regex_enabled is not None:
            nontrans["builtin_regex_enabled"] = bool(payload.builtin_regex_enabled)
        if payload.ai_discovery_enabled is not None:
            nontrans["ai_discovery_enabled"] = bool(payload.ai_discovery_enabled)
        if payload.ai_regex_generation_enabled is not None:
            nontrans["ai_regex_generation_enabled"] = bool(payload.ai_regex_generation_enabled)
        settings.input_defaults["nontrans_stage_settings"] = nontrans

        recall = dict(settings.input_defaults.get("term_recall_stage_settings", {}) or {})
        if payload.term_recall_batch_char_limit is not None:
            recall["batch_request_char_limit"] = int(payload.term_recall_batch_char_limit)
        if payload.term_recall_reasoning_effort is not None:
            effort = str(payload.term_recall_reasoning_effort or "low").strip().lower()
            recall["reasoning_effort"] = effort if effort in {"low", "medium", "high"} else "low"
        if payload.term_recall_enable_thinking is not None:
            recall["enable_thinking"] = bool(payload.term_recall_enable_thinking)
        settings.input_defaults["term_recall_stage_settings"] = recall

        review = dict(settings.input_defaults.get("term_review_stage_settings", {}) or {})
        if payload.term_review_batch_char_limit is not None:
            review["batch_request_char_limit"] = int(payload.term_review_batch_char_limit)
        if payload.term_review_max_context_chars is not None:
            review["max_context_chars"] = int(payload.term_review_max_context_chars)
        if payload.term_review_reasoning_effort is not None:
            effort = str(payload.term_review_reasoning_effort or "low").strip().lower()
            review["reasoning_effort"] = effort if effort in {"low", "medium", "high"} else "low"
        if payload.term_review_enable_thinking is not None:
            review["enable_thinking"] = bool(payload.term_review_enable_thinking)
        settings.input_defaults["term_review_stage_settings"] = review

        ai_review = dict(settings.input_defaults.get("ai_review_stage_settings", {}) or {})
        if payload.ai_review_batch_char_limit is not None:
            ai_review["batch_request_char_limit"] = int(payload.ai_review_batch_char_limit)
        if payload.ai_review_enable_thinking is not None:
            ai_review["enable_thinking"] = bool(payload.ai_review_enable_thinking)
        if payload.ai_review_max_items_per_request is not None:
            ai_review["max_items_per_request"] = max(1, int(payload.ai_review_max_items_per_request))
        if payload.ai_review_workspace_enable_thinking is not None:
            ai_review["workspace_enable_thinking"] = bool(payload.ai_review_workspace_enable_thinking)
        if payload.ai_review_debug_payload_logging is not None:
            ai_review["debug_payload_logging"] = bool(payload.ai_review_debug_payload_logging)
        if payload.ai_review_reasoning_effort is not None:
            effort = str(payload.ai_review_reasoning_effort or "low").strip().lower()
            ai_review["reasoning_effort"] = effort if effort in {"low", "medium", "high"} else "low"
        # These three stages always use model reasoning; the UI only exposes its intensity.
        nontrans["enable_thinking"] = True
        recall["enable_thinking"] = True
        review["enable_thinking"] = True
        settings.input_defaults["nontrans_stage_settings"] = nontrans
        settings.input_defaults["term_recall_stage_settings"] = recall
        settings.input_defaults["term_review_stage_settings"] = review
        ai_review["enable_thinking"] = True
        ai_review["workspace_enable_thinking"] = True
        ai_review.pop("workspace_model", None)
        settings.input_defaults["ai_review_stage_settings"] = ai_review

        task_facade.save_settings(settings)
        return {"ok": True}

    @app.post("/api/providers/models")
    async def load_provider_models(payload: SettingsPayload):
        settings = task_facade.load_settings()
        provider_name = payload.provider_name or settings.provider_name
        if provider_name not in settings.provider_settings:
            raise HTTPException(status_code=400, detail="Unknown provider.")

        provider_settings = settings.provider_settings[provider_name]
        if payload.model_name is not None:
            provider_settings.model = payload.model_name
        if payload.api_key is not None:
            provider_settings.api_key = payload.api_key
        if payload.base_url is not None:
            provider_settings.base_url = payload.base_url
        if payload.timeout_seconds is not None:
            provider_settings.timeout_seconds = max(1, int(payload.timeout_seconds))
        if payload.disable_system_proxy is not None:
            provider_settings.disable_system_proxy = bool(payload.disable_system_proxy)

        adapter = ProviderRegistry.create_adapter(provider_name, provider_settings)
        try:
            success, message, models = await adapter.list_models()
        finally:
            await adapter.close()

        if not success:
            raise HTTPException(status_code=400, detail=message or "加载模型失败")

        return {
            "provider_name": provider_name,
            "models": models,
            "message": message or "已刷新模型列表",
            "selected_model": provider_settings.model,
        }

    @app.get("/api/ascii-candidate-patterns")
    async def get_ascii_candidate_patterns():
        settings = task_facade.load_settings()
        patterns = list(settings.input_defaults.get("ascii_candidate_patterns", []) or [])
        return {"patterns": patterns}

    @app.post("/api/ascii-candidate-patterns")
    async def save_ascii_candidate_patterns(payload: AsciiCandidatePatternsPayload):
        normalized = []
        for index, item in enumerate(payload.patterns, start=1):
            pattern = item.pattern.strip()
            if not pattern:
                raise HTTPException(status_code=400, detail="Regex pattern cannot be empty.")
            try:
                re.compile(pattern)
            except re.error as exc:
                raise HTTPException(status_code=400, detail="Invalid regex: {0}".format(exc)) from exc
            name = item.name.strip() or pattern
            normalized.append(
                {
                    "name": name,
                    "pattern": pattern,
                    "enabled": bool(item.enabled),
                    "order_index": index,
                }
            )
        settings = task_facade.load_settings()
        settings.input_defaults["ascii_candidate_patterns"] = normalized
        task_facade.save_settings(settings)
        return {"ok": True, "patterns": normalized}

    @app.get("/api/nontrans-builtin-rules")
    async def get_nontrans_builtin_rules():
        return builtin_nontrans_rules_response()

    @app.post("/api/nontrans-builtin-rules")
    async def save_nontrans_builtin_rules(payload: BuiltinNonTransRulesPayload):
        response = _save_builtin_nontrans_rules_to_library(payload.rules)
        return {"ok": True, **response}

    @app.get("/api/nontrans-pending-rules")
    async def get_pending_nontrans_rules():
        settings = task_facade.load_settings()
        return pending_nontrans_rules_response(settings)

    @app.post("/api/nontrans-pending-rules/seen")
    async def save_pending_nontrans_rule_seen(payload: PendingNonTransRuleSeenPayload):
        settings = task_facade.load_settings()
        mark_pending_nontrans_rule_notice(
            settings,
            notice_seen=payload.notice_seen,
            library_seen=payload.library_seen,
        )
        task_facade.save_settings(settings)
        return {"ok": True, **pending_nontrans_rules_response(settings)}

    @app.post("/api/nontrans-pending-rules/import")
    async def import_pending_nontrans_rules(payload: PendingNonTransRuleImportPayload):
        if not payload.rules:
            raise HTTPException(status_code=400, detail="没有可导入的规则。")
        pending_rules = load_pending_nontrans_rule_imports(task_facade.load_settings())
        pending_by_cache_id = {
            str(item.get("cache_id", "")).strip(): item
            for item in pending_rules
            if str(item.get("cache_id", "")).strip()
        }
        if not pending_by_cache_id:
            raise HTTPException(status_code=400, detail="当前没有待导入的规则。")

        selected_cache_ids = []
        import_rows = []
        for index, item in enumerate(payload.rules, start=1):
            cache_id = str(item.cache_id or "").strip()
            if not cache_id or cache_id not in pending_by_cache_id:
                raise HTTPException(status_code=400, detail="存在无效的缓存规则，无法导入。")
            selected_cache_ids.append(cache_id)
            import_rows.append(
                BuiltinNonTransRulePayload(
                    rule_id=str(item.rule_id or "").strip() or cache_id,
                    name=str(item.name or "").strip() or cache_id,
                    role=str(item.role or "empty").strip() or "empty",
                    element_type=str(item.element_type or "other").strip() or "other",
                    pattern=str(item.pattern or "").strip(),
                    enabled=bool(item.enabled),
                    examples=[str(example or "").strip() for example in list(item.examples or []) if str(example or "").strip()],
                )
            )

        builtin_rules = load_builtin_nontrans_rules()
        builtin_by_key = {}
        merged_payload = []
        for rule in builtin_rules:
            role = (
                "open"
                if str(rule.open_pattern or "").strip()
                else "close"
                if str(rule.close_pattern or "").strip()
                else "empty"
            )
            regex = (
                str(rule.open_pattern or "").strip()
                or str(rule.close_pattern or "").strip()
                or str(rule.empty_pattern or "").strip()
                or str(rule.pattern or "").strip()
            )
            key = (role, str(rule.element_type or "").strip() or "other", regex)
            builtin_by_key[key] = rule
            merged_payload.append(
                BuiltinNonTransRulePayload(
                    rule_id=rule.rule_id,
                    name=rule.name,
                    role=role,
                    element_type=rule.element_type,
                    pattern=regex,
                    enabled=rule.enabled,
                    examples=list(rule.examples or []),
                )
            )

        for item in import_rows:
            key = (str(item.role), str(item.element_type), str(item.pattern))
            existing = builtin_by_key.get(key)
            if existing is None:
                merged_payload.append(item)
                continue
            merged_examples = [str(example or "").strip() for example in list(existing.examples or []) if str(example or "").strip()]
            for example in list(item.examples or []):
                cleaned = str(example or "").strip()
                if cleaned and cleaned not in merged_examples:
                    merged_examples.append(cleaned)
            for target in merged_payload:
                target_key = (str(target.role), str(target.element_type), str(target.pattern))
                if target_key == key:
                    target.examples = merged_examples[:3]
                    if not str(target.name or "").strip() and str(item.name or "").strip():
                        target.name = str(item.name)
                    break

        response = _save_builtin_nontrans_rules_to_library(merged_payload)
        settings = task_facade.load_settings()
        clear_pending_nontrans_rule_imports(settings)
        task_facade.save_settings(settings)
        if import_rows:
            track_event("feature_used.nontrans_regex_imported")
        return {
            "ok": True,
            **response,
            "pending_nontrans_rules": pending_nontrans_rules_response(settings),
        }

    @app.post("/api/nontrans-pending-rules/clear")
    async def clear_pending_nontrans_rules():
        settings = task_facade.load_settings()
        had_pending = bool(load_pending_nontrans_rule_imports(settings))
        clear_pending_nontrans_rule_imports(settings)
        task_facade.save_settings(settings)
        if had_pending:
            track_event("feature_used.nontrans_regex_discarded")
        return {"ok": True, **pending_nontrans_rules_response(settings)}

    @app.get("/api/prompt-templates")
    async def get_prompt_templates():
        settings = task_facade.load_settings()
        return prompt_templates_response(settings)

    @app.post("/api/prompt-templates")
    async def save_prompt_templates(payload: PromptTemplatesPayload):
        allowed_keys = allowed_prompt_template_keys()
        unknown_keys = sorted(set(payload.templates.keys()) - allowed_keys)
        if unknown_keys:
            raise HTTPException(
                status_code=400,
                detail="Unknown prompt template key: {0}".format(", ".join(unknown_keys)),
            )
        settings = task_facade.load_settings()
        for key, value in payload.templates.items():
            text = str(value)
            if not text.strip():
                raise HTTPException(status_code=400, detail="Prompt template cannot be empty: {0}".format(key))
            settings.prompt_templates[key] = text
        task_facade.save_settings(settings)
        return {"ok": True, **prompt_templates_response(settings)}

    @app.post("/api/prompt-templates/reset")
    async def reset_prompt_templates(payload: PromptTemplateResetPayload):
        allowed_keys = allowed_prompt_template_keys()
        keys = payload.keys or sorted(allowed_keys)
        unknown_keys = sorted(set(keys) - allowed_keys)
        if unknown_keys:
            raise HTTPException(
                status_code=400,
                detail="Unknown prompt template key: {0}".format(", ".join(unknown_keys)),
            )
        defaults = build_default_settings().prompt_templates
        settings = task_facade.load_settings()
        for key in keys:
            settings.prompt_templates[key] = defaults.get(key, "")
        task_facade.save_settings(settings)
        return {"ok": True, **prompt_templates_response(settings)}

    @app.get("/api/scan")
    async def scan(folder_path: str):
        try:
            result = scan_folder(folder_path)
        except Exception as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc
        return result.to_dict()

    @app.post("/api/dialog/select-folder")
    async def select_folder():
        try:
            folder_path = select_folder_dialog()
        except Exception as exc:
            raise HTTPException(status_code=500, detail=str(exc)) from exc
        return {"folder_path": folder_path, "cancelled": not bool(folder_path)}

    @app.get("/api/cross-excel/scan")
    async def cross_excel_scan(folder_path: str):
        try:
            return scan_cross_excel_folder(folder_path)
        except Exception as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc

    @app.post("/api/cross-excel/search")
    async def cross_excel_search(payload: CrossExcelSearchPayload):
        try:
            track_event("task_action.cross_excel_search")
            result = search_excel_rows(payload.folder_path, payload.query, payload.limit)
            return result
        except Exception as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc

    @app.post("/api/cross-excel/merge")
    async def cross_excel_merge(payload: CrossExcelMergePayload):
        try:
            track_event("task_action.cross_excel_merge")
            result = merge_excel_files_by_headers(
                payload.folder_path,
                payload.headers,
                apply_format=payload.apply_format,
            )
            return result
        except Exception as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc

    @app.post("/api/diff-excel/compare")
    async def diff_excel_compare(payload: DiffExcelComparePayload):
        track_event("diff.compare.start")
        try:
            path_a = Path(str(payload.path_a or "").strip())
            path_b = Path(str(payload.path_b or "").strip())
            if path_a.is_file() and path_b.is_file():
                track_event("diff.mode.file_to_file")
            elif path_a.is_dir() or path_b.is_dir():
                track_event("diff.mode.folder_to_folder")
            if payload.compare_mode == "field_match":
                track_event("diff.mode.field_match")
            else:
                track_event("diff.mode.position")
            def operation(progress_callback):
                try:
                    result = run_diff_excel_compare_to_cache(
                        payload.path_a,
                        payload.path_b,
                        compare_mode=payload.compare_mode,
                        reference_field=payload.reference_field,
                        compare_fields=payload.compare_fields,
                        include_unmatched=payload.include_unmatched,
                        ignore_case=False,
                        trim_whitespace=False,
                        progress_callback=progress_callback,
                    )
                    track_event("diff.compare.success")
                    return result
                except Exception:
                    track_event("diff.compare.fail")
                    raise

            return diff_task_service.start(operation).to_dict()
        except Exception as exc:
            track_event("diff.compare.fail")
            raise HTTPException(status_code=400, detail=str(exc)) from exc

    @app.get("/api/diff-excel/task/{task_id}")
    async def diff_excel_task(task_id: str):
        task = diff_task_service.get(task_id)
        if task is None:
            raise HTTPException(status_code=404, detail="未找到比对任务。")
        return task.to_dict()

    @app.get("/api/diff-excel/field-match-headers")
    async def diff_excel_field_match_headers(path_a: str, path_b: str):
        try:
            return scan_diff_excel_field_match_headers(path_a, path_b)
        except Exception as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc

    @app.post("/api/diff-excel/export")
    async def diff_excel_export(payload: DiffExcelExportPayload):
        try:
            return export_diff_excel_cached_records(
                payload.cache_file,
                payload.output_file,
                query=payload.query,
            )
        except Exception as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc

    @app.post("/api/diff-excel/highlight")
    async def diff_excel_highlight(payload: DiffExcelHighlightPayload):
        try:
            track_event("diff.highlight")
            changed_cells, workbook_count = apply_diff_excel_highlight_from_cache(
                payload.cache_file,
                str(payload.target or "A"),
                str(payload.color_hex or "#FFD966"),
                query=payload.query,
            )
            return {
                "ok": True,
                "changed_cells": changed_cells,
                "workbook_count": workbook_count,
                "target": str(payload.target or "A"),
                "color_hex": str(payload.color_hex or "#FFD966"),
            }
        except Exception as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc

    @app.get("/api/diff-excel/preview")
    async def diff_excel_preview(cache_file: str, query: str = "", limit: int = 200, offset: int = 0):
        try:
            return read_diff_excel_cached_preview(cache_file, query=query, limit=limit, offset=offset)
        except Exception as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc

    @app.post("/api/diff-excel/open-cell")
    async def diff_excel_open_cell(payload: DiffExcelOpenCellPayload):
        file_path = str(payload.file_path or "").strip()
        sheet_name = str(payload.sheet_name or "").strip()
        cell_address = str(payload.cell_address or "").strip()
        if not file_path or not sheet_name or not cell_address:
            raise HTTPException(status_code=400, detail="缺少定位所需的信息。")
        try:
            track_event("diff.jump")
            _open_excel_cell(file_path, sheet_name, cell_address)
        except Exception as exc:
            raise HTTPException(status_code=500, detail=str(exc)) from exc
        return {
            "ok": True,
            "file_path": file_path,
            "sheet_name": sheet_name,
            "cell_address": cell_address,
        }

    @app.get("/api/dialog/select-review-file")
    async def select_review_file():
        try:
            import tkinter as tk
            from tkinter import filedialog

            root = tk.Tk()
            root.withdraw()
            root.attributes("-topmost", True)
            selected = filedialog.askopenfilename(
                title="选择待审校文件",
                filetypes=[
                    ("支持的文件", "*.xlsx *.xlsm *.xlf *.xliff"),
                    ("Excel 文件", "*.xlsx *.xlsm"),
                    ("XLIFF 文件", "*.xlf *.xliff"),
                ],
            )
            root.destroy()
        except Exception as exc:
            raise HTTPException(status_code=500, detail=str(exc)) from exc
        return {"file_path": selected or "", "cancelled": not bool(selected)}

    @app.get("/api/dialog/select-review-files")
    async def select_review_files():
        try:
            import tkinter as tk
            from tkinter import filedialog

            root = tk.Tk()
            root.withdraw()
            root.attributes("-topmost", True)
            selected = filedialog.askopenfilenames(
                title="选择待审校文件",
                filetypes=[
                    ("支持的文件", "*.xlsx *.xlsm *.xlf *.xliff *.csv *.tsv *.txt *.md *.docx *.pptx *.pdf *.json *.xml"),
                    ("所有文件", "*.*"),
                ],
            )
            root.destroy()
        except Exception as exc:
            raise HTTPException(status_code=500, detail=str(exc)) from exc
        paths = [str(path) for path in selected if str(path)]
        return {"file_paths": paths, "cancelled": not bool(paths)}

    @app.post("/api/ai-review/file/open")
    async def ai_review_open_file(payload: AIReviewOpenFilePayload):
        file_path = str(payload.file_path or "").strip()
        if not file_path:
            raise HTTPException(status_code=400, detail="请先选择文件。")
        try:
            _open_local_file(file_path)
        except Exception as exc:
            raise HTTPException(status_code=500, detail=str(exc)) from exc
        return {"ok": True, "file_path": file_path}

    @app.post("/api/ai-review/file/load")
    async def ai_review_load_file(payload: AIReviewOpenFilePayload):
        file_path = Path(str(payload.file_path or "").strip())
        if not str(file_path):
            raise HTTPException(status_code=400, detail="请先选择文件。")
        if not file_path.exists() or not file_path.is_file():
            raise HTTPException(status_code=400, detail="文件不存在。")
        filename = file_path.name
        try:
            stored_path = save_ai_review_upload_file(filename, file_path.read_bytes())
            file_type = detect_ai_review_file_type(stored_path)
            if file_type == "excel":
                metadata = read_ai_review_excel_headers(stored_path)
                batch_id = create_ai_review_batch(
                    original_filename=filename,
                    stored_path=stored_path,
                    file_type=file_type,
                    status="uploaded",
                    metadata={**metadata, "original_file_path": str(file_path)},
                )
                return _ai_review_excel_upload_response(
                    batch_id=batch_id,
                    filename=filename,
                    metadata=metadata,
                    original_file_path=str(file_path),
                )

            items = read_ai_review_xliff_items(stored_path, filename)
            language_metadata = read_ai_review_xliff_language_metadata(stored_path)
            batch_id = create_ai_review_batch(
                original_filename=filename,
                stored_path=stored_path,
                file_type=file_type,
                status="uploaded",
                metadata={**language_metadata, "original_file_path": str(file_path)},
            )
            replace_ai_review_batch_items(
                batch_id=batch_id,
                items=items,
                metadata_update={"preview_ready": True, **language_metadata, "original_file_path": str(file_path)},
            )
            return _ai_review_batch_response(batch_id, "XLIFF 读取完成")
        except ValueError as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc
        except Exception as exc:
            raise HTTPException(status_code=500, detail="读取失败：{0}".format(exc)) from exc

    @app.post("/api/ai-review/file/upload")
    async def ai_review_upload_file(file: UploadFile = File(...)):
        filename = file.filename or "unknown"
        data = await file.read()
        if not data:
            raise HTTPException(status_code=400, detail="上传文件为空。")
        try:
            stored_path = save_ai_review_upload_file(filename, data)
            file_type = detect_ai_review_file_type(stored_path)
            if file_type == "excel":
                metadata = read_ai_review_excel_headers(stored_path)
                batch_id = create_ai_review_batch(
                    original_filename=filename,
                    stored_path=stored_path,
                    file_type=file_type,
                    status="uploaded",
                    metadata=metadata,
                )
                return _ai_review_excel_upload_response(
                    batch_id=batch_id,
                    filename=filename,
                    metadata=metadata,
                )

            items = read_ai_review_xliff_items(stored_path, filename)
            language_metadata = read_ai_review_xliff_language_metadata(stored_path)
            batch_id = create_ai_review_batch(
                original_filename=filename,
                stored_path=stored_path,
                file_type=file_type,
                status="uploaded",
                metadata=language_metadata,
            )
            replace_ai_review_batch_items(
                batch_id=batch_id,
                items=items,
                metadata_update={"preview_ready": True, **language_metadata},
            )
            return _ai_review_batch_response(batch_id, "XLIFF 读取完成")
        except ValueError as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc
        except Exception as exc:
            raise HTTPException(status_code=500, detail="读取失败：{0}".format(exc)) from exc

    @app.post("/api/ai-review/select-columns")
    async def ai_review_select_columns(payload: AIReviewSelectColumnsPayload):
        batch = get_ai_review_batch(payload.batch_id)
        if not batch:
            raise HTTPException(status_code=404, detail="读取批次不存在")
        if batch["file_type"] != "excel":
            raise HTTPException(status_code=400, detail="只有 Excel 文件需要选择列")
        try:
            items = read_ai_review_excel_items(
                Path(batch["stored_path"]),
                payload.source_column,
                payload.target_column,
                batch["original_filename"],
            )
            replace_ai_review_batch_items(
                batch_id=payload.batch_id,
                items=items,
                source_column=payload.source_column,
                target_column=payload.target_column,
                metadata_update={"preview_ready": True},
            )
            return _ai_review_batch_response(payload.batch_id, "Excel 读取完成")
        except ValueError as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc
        except Exception as exc:
            raise HTTPException(status_code=500, detail="读取失败：{0}".format(exc)) from exc

    @app.post("/api/ai-review/select-excel-mapping")
    async def ai_review_select_excel_mapping(payload: AIReviewExcelMappingPayload):
        batch = get_ai_review_batch(payload.batch_id)
        if not batch:
            raise HTTPException(status_code=404, detail="读取批次不存在")
        if batch["file_type"] != "excel":
            raise HTTPException(status_code=400, detail="只有 Excel 文件需要配置列映射")
        try:
            items = read_ai_review_excel_items_by_mapping(
                Path(batch["stored_path"]),
                payload.mapping,
                batch["original_filename"],
            )
            replace_ai_review_batch_items(
                batch_id=payload.batch_id,
                items=items,
                source_column=None,
                target_column=None,
                metadata_update={
                    "preview_ready": True,
                    "excel_mapping": payload.mapping,
                    "source_language": str(payload.mapping.get("source_language") or ""),
                    "target_language": str(payload.mapping.get("target_language") or ""),
                },
            )
            return _ai_review_batch_response(payload.batch_id, "Excel 列映射读取完成")
        except ValueError as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc
        except Exception as exc:
            raise HTTPException(status_code=500, detail="读取失败：{0}".format(exc)) from exc

    @app.get("/api/ai-review/ai-settings")
    async def ai_review_ai_settings():
        settings = get_ai_review_shared_ai_settings()
        return {
            "provider": settings.get("provider", "DeepSeek"),
            "has_api_key": bool(settings.get("api_key")),
            "selected_model": settings.get("selected_model", ""),
            "max_chars_per_request": settings.get("max_chars_per_request", 3000),
            "enable_thinking": bool(settings.get("enable_thinking", False)),
        }

    @app.post("/api/ai-review/ai/test")
    async def ai_review_test_model():
        settings = get_ai_review_shared_ai_settings()
        api_key = str(settings.get("api_key", "") or "")
        model = str(settings.get("selected_model", "") or "")
        if not api_key:
            raise HTTPException(status_code=400, detail="请先在模型设置中填写 API Key。")
        if not model:
            raise HTTPException(status_code=400, detail="请先在模型设置中选择模型。")
        try:
            content = test_ai_review_chat(
                api_key,
                model,
                enable_thinking=bool(settings.get("enable_thinking", False)),
            )
        except SharedProviderError as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc
        return {"ok": True, "message": "测试 OK", "model": model, "response": content}

    @app.get("/api/ai-review/excel-mapping-presets")
    async def ai_review_excel_mapping_presets():
        return {"presets": list_excel_mapping_presets()}

    @app.get("/api/ai-review/excel-mapping-presets/{preset_id}")
    async def ai_review_excel_mapping_preset(preset_id: str):
        try:
            return {"preset": get_excel_mapping_preset(preset_id)}
        except ValueError as exc:
            raise HTTPException(status_code=404, detail=str(exc)) from exc

    @app.post("/api/ai-review/excel-mapping-presets")
    async def ai_review_save_excel_mapping_preset(payload: AIReviewExcelMappingPresetSavePayload):
        try:
            preset = save_excel_mapping_preset(payload.name, payload.mapping, payload.id)
        except ValueError as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc
        return {"ok": True, "message": "Excel 映射模板已保存", "preset": preset}

    @app.delete("/api/ai-review/excel-mapping-presets/{preset_id}")
    async def ai_review_delete_excel_mapping_preset(preset_id: str):
        try:
            delete_excel_mapping_preset(preset_id)
        except ValueError as exc:
            raise HTTPException(status_code=404, detail=str(exc)) from exc
        return {"ok": True, "message": "Excel 映射模板已删除"}

    @app.get("/api/ai-review/prompt-templates")
    async def ai_review_prompt_templates():
        return {"templates": list_ai_review_prompt_templates()}

    @app.get("/api/ai-review/prompt-templates/{template_id}")
    async def ai_review_prompt_template(template_id: str):
        try:
            return {"template": get_ai_review_prompt_template(template_id)}
        except ValueError as exc:
            raise HTTPException(status_code=404, detail=str(exc)) from exc

    @app.post("/api/ai-review/prompt-templates")
    async def ai_review_save_prompt_template(payload: AIReviewPromptTemplateSavePayload):
        try:
            template = save_ai_review_prompt_template(
                template_id=payload.id,
                name=payload.name.strip() or "未命名模板",
                system_prompt=payload.system_prompt,
                user_prompt=payload.user_prompt,
                forbidden_words_text=payload.forbidden_words_text,
            )
        except ValueError as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc
        return {"ok": True, "message": "提示词模板已保存", "template": template}

    @app.post("/api/ai-review/prompt-templates/reset-default")
    async def ai_review_reset_prompt_template():
        template = reset_ai_review_prompt_template()
        return {"ok": True, "message": "默认提示词已恢复", "template": template}

    @app.delete("/api/ai-review/prompt-templates/{template_id}")
    async def ai_review_delete_prompt_template(template_id: str):
        try:
            fallback_template = delete_ai_review_prompt_template(template_id)
        except ValueError as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc
        return {"ok": True, "message": "提示词模板已删除", "fallback_template": fallback_template}

    @app.get("/api/ai-review/directional-templates")
    async def ai_review_directional_templates():
        return {"templates": list_directional_templates()}

    @app.get("/api/ai-review/directional-templates/{template_id}")
    async def ai_review_directional_template(template_id: str):
        try:
            return {"template": get_directional_template(template_id)}
        except ValueError as exc:
            raise HTTPException(status_code=404, detail=str(exc)) from exc

    @app.post("/api/ai-review/directional-templates")
    async def ai_review_save_directional_template(payload: AIReviewDirectionalTemplateSavePayload):
        try:
            template = save_directional_template(
                payload.id,
                payload.name.strip() or "未命名定向模板",
                payload.items,
            )
        except ValueError as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc
        return {"ok": True, "message": "定向审校模板已保存", "template": template}

    @app.delete("/api/ai-review/directional-templates/{template_id}")
    async def ai_review_delete_directional_template(template_id: str):
        try:
            delete_directional_template(template_id)
        except ValueError as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc
        return {"ok": True, "message": "定向审校模板已删除"}

    @app.get("/api/ai-review/forbidden-templates")
    async def ai_review_forbidden_templates():
        return {"templates": list_forbidden_templates()}

    @app.get("/api/ai-review/forbidden-templates/{template_id}")
    async def ai_review_forbidden_template(template_id: str):
        try:
            return {"template": get_forbidden_template(template_id)}
        except ValueError as exc:
            raise HTTPException(status_code=404, detail=str(exc)) from exc

    @app.post("/api/ai-review/forbidden-templates")
    async def ai_review_save_forbidden_template(payload: AIReviewForbiddenTemplateSavePayload):
        try:
            template = save_forbidden_template(
                payload.id,
                payload.name.strip() or "未命名禁用词模板",
                payload.words_text,
            )
        except ValueError as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc
        return {"ok": True, "message": "禁用词模板已保存", "template": template}

    @app.delete("/api/ai-review/forbidden-templates/{template_id}")
    async def ai_review_delete_forbidden_template(template_id: str):
        try:
            delete_forbidden_template(template_id)
        except ValueError as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc
        return {"ok": True, "message": "禁用词模板已删除"}

    @app.post("/api/ai-review/start")
    async def ai_review_start(payload: AIReviewStartPayload):
        try:
            task_id = create_review_task(
                payload.batch_id,
                payload.prompt_template_id,
                payload.source_language.strip(),
                payload.target_language.strip(),
                payload.mode,
                payload.directional_template_id,
                payload.enable_ai_review,
                payload.enable_forbidden_check,
                payload.forbidden_template_id,
            )
        except (ReviewTaskError, ValueError) as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc
        task = get_review_task(task_id)
        return {"ok": True, "message": "审校任务已启动", "task": task}

    @app.get("/api/ai-review/tasks/{task_id}")
    async def ai_review_task(task_id: str):
        task = get_review_task(task_id)
        if not task:
            raise HTTPException(status_code=404, detail="审校任务不存在")
        return {"task": task, "results": get_review_results(task_id, limit=20)}

    @app.get("/api/ai-review/tasks/{task_id}/logs")
    async def ai_review_task_logs(task_id: str, after_id: int = 0):
        if not get_review_task(task_id):
            raise HTTPException(status_code=404, detail="审校任务不存在")
        return {"logs": get_review_logs(task_id, after_id)}

    @app.get("/api/ai-review/tasks/{task_id}/issue-results")
    async def ai_review_task_issue_results(task_id: str):
        task = get_review_task(task_id)
        if not task:
            raise HTTPException(status_code=404, detail="审校任务不存在")
        return {"task": task, "results": get_review_issue_results(task_id)}

    @app.get("/api/ai-review/tasks/{task_id}/followup/{result_id}")
    async def ai_review_followup_messages(task_id: str, result_id: str):
        try:
            init_ai_review_db()
            if not get_review_task(task_id):
                raise HTTPException(status_code=404, detail="审校任务不存在")
            result = get_review_followup_messages(task_id, result_id)
            track_event("feature_used.ai_review_followup_open")
            return result
        except HTTPException:
            raise
        except ReviewTaskError as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc
        except Exception as exc:
            print("[AI_REVIEW][ERROR][{0}] 追问读取失败：{1}\n{2}".format(task_id, exc, traceback.format_exc()), flush=True)
            raise HTTPException(status_code=500, detail="追问读取失败：{0}".format(exc)) from exc

    @app.post("/api/ai-review/tasks/{task_id}/followup/{result_id}")
    async def ai_review_followup_send(task_id: str, result_id: str, payload: AIReviewFollowupPayload):
        try:
            init_ai_review_db()
            if not get_review_task(task_id):
                raise HTTPException(status_code=404, detail="审校任务不存在")
            track_event("feature_used.ai_review_followup_send")
            result = await run_in_threadpool(send_review_followup_message, task_id, result_id, payload.message)
            track_event("feature_success.ai_review_followup_send")
            return result
        except HTTPException:
            raise
        except ReviewTaskError as exc:
            track_event("feature_fail.ai_review_followup_send")
            raise HTTPException(status_code=400, detail=str(exc)) from exc
        except Exception as exc:
            track_event("feature_fail.ai_review_followup_send")
            print("[AI_REVIEW][ERROR][{0}] 追问发送失败：{1}\n{2}".format(task_id, exc, traceback.format_exc()), flush=True)
            raise HTTPException(status_code=500, detail="追问发送失败：{0}".format(exc)) from exc

    @app.post("/api/ai-review/outputs/open-folder")
    async def ai_review_open_outputs():
        try:
            open_ai_review_directory(AI_REVIEW_OUTPUTS_DIR)
        except Exception as exc:
            raise HTTPException(status_code=500, detail="打开目录失败：{0}".format(exc)) from exc
        return {"ok": True, "message": "已打开结果目录", "path": str(AI_REVIEW_OUTPUTS_DIR)}

    @app.post("/api/ai-review/outputs/open-file")
    async def ai_review_open_output_file(payload: AIReviewOpenFilePayload):
        file_path = str(payload.file_path or "").strip()
        if not file_path:
            raise HTTPException(status_code=400, detail="暂无可打开的结果文件。")
        try:
            _open_local_file(file_path)
        except Exception as exc:
            raise HTTPException(status_code=500, detail="打开文件失败：{0}".format(exc)) from exc
        return {"ok": True, "file_path": file_path}

    @app.post("/api/tasks/start")
    async def start_task(payload: StartTaskPayload):
        settings = task_facade.load_settings()
        task_input = None
        if not payload.resume:
            task_input = TaskInput(
                folder_path=payload.folder_path,
                header_name=payload.header_name,
                source_language=payload.source_language,
                single_item_char_limit=payload.single_item_char_limit,
                batch_request_char_limit=payload.batch_request_char_limit,
                file_type=payload.file_type,
                export_review_sheet=payload.export_review_sheet,
                extraction_mode=payload.extraction_mode,
            )
        try:
            task_facade.start(task_input, resume=payload.resume, settings=settings)
        except RuntimeError as exc:
            raise HTTPException(status_code=409, detail=str(exc)) from exc
        return {"ok": True}

    @app.post("/api/tasks/resume")
    async def resume_task():
        settings = task_facade.load_settings()
        try:
            task_facade.resume(settings=settings)
        except RuntimeError as exc:
            raise HTTPException(status_code=409, detail=str(exc)) from exc
        return {"ok": True}

    @app.post("/api/tasks/stop")
    async def stop_task():
        task_facade.stop()
        return {"ok": True}

    @app.post("/api/tasks/clear-cache")
    async def clear_task_cache():
        try:
            task_facade.clear_runtime_cache()
        except RuntimeError as exc:
            raise HTTPException(status_code=409, detail=str(exc)) from exc
        return {"ok": True}

    @app.get("/api/status")
    async def status():
        return task_facade.snapshot().to_dict()

    @app.get("/api/results/summary")
    async def results_summary(output_file: Optional[str] = None):
        return task_facade.result_summary(output_file).to_dict()

    @app.get("/api/results/download")
    async def download_result(output_file: str):
        path = Path(output_file)
        if not path.exists() or not path.is_file():
            raise HTTPException(status_code=404, detail="Output file does not exist.")
        return FileResponse(
            path,
            filename=path.name,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    @app.post("/api/results/open-folder")
    async def open_result_folder(output_file: str):
        path = Path(output_file)
        if not path.exists() or not path.is_file():
            raise HTTPException(status_code=404, detail="Output file does not exist.")
        try:
            open_file_location(path)
        except Exception as exc:
            raise HTTPException(status_code=500, detail=str(exc)) from exc
        return {"ok": True, "folder": str(path.parent)}

    @app.post("/api/results/open-file")
    async def open_result_file(output_file: str):
        path = Path(output_file)
        if not path.exists() or not path.is_file():
            raise HTTPException(status_code=404, detail="Output file does not exist.")
        try:
            open_file_directly(path)
        except Exception as exc:
            raise HTTPException(status_code=500, detail=str(exc)) from exc
        return {"ok": True, "file": str(path)}

    return app


app = create_app()


def settings_to_json_for_debug() -> str:
    return json.dumps(build_default_settings().to_dict(), ensure_ascii=False, indent=2)


def open_file_location(path: Path) -> None:
    open_path_folder(path.parent)


def open_file_directly(path: Path) -> None:
    open_any_path(path)


def select_folder_dialog() -> str:
    try:
        import tkinter as tk
        from tkinter import filedialog
    except Exception as exc:
        raise RuntimeError("当前环境不支持原生文件夹选择窗口。") from exc

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    try:
        selected = filedialog.askdirectory(title="选择待处理文件夹", mustexist=True)
    finally:
        root.destroy()
    return str(selected or "")


INDEX_HTML = """<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>译禾工具合集</title>
  <link rel="stylesheet" href="/assets/app.css" />
</head>
<body>
  <a class="skip-link" href="#mainContent">跳到主内容</a>
  <div class="shell">
    <aside class="sidebar">
      <div class="brand">
        <img class="brand-mark" src="/assets/logo.png" alt="译禾工具合集" />
        <div class="brand-copy">
          <span class="brand-version">v__APP_VERSION__</span>
          <strong>译禾工具合集</strong>
          <small>作者：王子京</small>
        </div>
      </div>
      <nav class="sidebar-nav">
        <button class="nav-link nav-link-top active" data-page-target="toolGuidePage">工具说明</button>
        <button class="nav-link nav-link-top" data-page-target="modelSettingsPage">模型设置</button>

        <details class="nav-accordion" data-accordion-key="text-preprocess">
          <summary class="nav-accordion-summary">
            <span class="nav-group-title">文本预处理工具</span>
          </summary>
          <div class="nav-submenu">
            <button class="nav-link nav-link-sub" data-page-target="overviewPage">总览</button>
            <button class="nav-link nav-link-sub" data-page-target="modelStageSettingsPage">模型阶段设置</button>
            <button class="nav-link nav-link-sub" data-page-target="nontransSettingsPage">非译元素设置</button>
            <button class="nav-link nav-link-sub" data-page-target="promptSettingsPage">提示词设置</button>
          </div>
        </details>

        <details class="nav-accordion" data-accordion-key="diff-tool">
          <summary class="nav-accordion-summary">
            <span class="nav-group-title">Diff 工具</span>
          </summary>
          <div class="nav-submenu">
            <button class="nav-link nav-link-sub" data-page-target="diffExcelPage">Excel差异比对</button>
          </div>
        </details>

        <details class="nav-accordion" data-accordion-key="ai-review-tool">
          <summary class="nav-accordion-summary">
            <span class="nav-group-title">AI 审校工具</span>
          </summary>
          <div class="nav-submenu">
            <button class="nav-link nav-link-sub" data-page-target="aiReviewTaskPage">审校任务</button>
            <button class="nav-link nav-link-sub" data-page-target="aiReviewSettingsPage">审校设置</button>
          </div>
        </details>

        <details class="nav-accordion" data-accordion-key="cross-excel-search">
          <summary class="nav-accordion-summary">
            <span class="nav-group-title">跨Excel搜索与合并</span>
          </summary>
          <div class="nav-submenu">
            <button class="nav-link nav-link-sub" data-page-target="crossExcelPage">搜索与合并</button>
          </div>
        </details>
      </nav>
      <div class="sidebar-actions-panel">
        <button id="feedbackEntryButton" class="feedback-entry-button" type="button">我要反馈</button>
      </div>
    </aside>

    <main id="mainContent" class="content" tabindex="-1">
      <header class="hero">
        <div>
          <div class="hero-context"><span class="eyebrow">本地工作台</span><span class="hero-context-divider"></span><span>Localization &amp; Text Tooling</span></div>
          <h1 id="heroTitle">工具说明</h1>
          <p id="heroLede" class="lede">先了解每个工具能做什么，再开始任务。</p>
          <button id="updateNoticeButton" class="notice-button update-notice-button" type="button" hidden>
            <span class="notice-dot"></span>
            <span>发现新版本</span>
          </button>
          <button id="pendingRuleNoticeButton" class="notice-button" type="button" hidden>
            <span class="notice-dot"></span>
            <span>发现新的非译规则</span>
          </button>
        </div>
      </header>

      <section id="toolGuidePage" class="page-section active">
        <div class="inline-notice" role="note"><span class="inline-notice-icon" aria-hidden="true"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><circle cx="12" cy="12" r="8.5"></circle><path d="M12 10.5v5M12 7.5h.01"></path></svg></span><span><strong>隐私优先</strong>仅统计匿名功能触发次数；不会收集文本、文件名、路径、账号或密钥。</span></div>
        <div class="tool-guide-grid">
          <button class="tool-guide-card tool-guide-card-preprocess" type="button" data-tool-guide="textPreprocess" aria-label="打开文本预处理工具说明">
            <span class="tool-guide-card-top"><span class="tool-guide-icon" aria-hidden="true"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7"><path d="M6.5 3.5h7l4 4v13h-11z"></path><path d="M13.5 3.5v4h4M8.5 12h7M8.5 15.5h7"></path></svg></span><span class="tool-guide-badge">批量处理</span></span>
            <span class="tool-guide-card-copy"><strong>文本预处理</strong><small>从本地表格中提取术语、非译元素和可复用规则。</small></span>
            <span class="tool-guide-card-footer"><span>目录级工作流</span><span class="tool-guide-arrow" aria-hidden="true">→</span></span>
          </button>
          <button class="tool-guide-card tool-guide-card-diff" type="button" data-tool-guide="diffExcel" aria-label="打开 Diff 工具说明">
            <span class="tool-guide-card-top"><span class="tool-guide-icon" aria-hidden="true"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7"><path d="M7 4v12M17 8v12M4 7l3-3 3 3M14 17l3 3 3-3"></path><path d="M9 7h5M10 17h4"></path></svg></span><span class="tool-guide-badge">Excel</span></span>
            <span class="tool-guide-card-copy"><strong>Diff 工具</strong><small>比较文件或目录差异，快速定位改动。</small></span>
            <span class="tool-guide-card-footer"><span>差异预览与导出</span><span class="tool-guide-arrow" aria-hidden="true">→</span></span>
          </button>
          <button class="tool-guide-card tool-guide-card-review" type="button" data-tool-guide="aiReview" aria-label="打开 AI 审校工具说明">
            <span class="tool-guide-card-top"><span class="tool-guide-icon" aria-hidden="true"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7"><path d="M12 3.5 13.7 8l4.8 1.6-4.2 2.6.1 5-4.4-2.7-4.2 2.7.1-5-4.2-2.6L6.3 8z"></path><path d="m17.5 17.5 1 1 2-2"></path></svg></span><span class="tool-guide-badge tool-guide-badge-ai">Agent</span></span>
            <span class="tool-guide-card-copy"><strong>AI 审校</strong><small>让 Workspace Agent 识别结构，再按目标语言独立审校。</small></span>
            <span class="tool-guide-card-footer"><span>流式会话工作台</span><span class="tool-guide-arrow" aria-hidden="true">→</span></span>
          </button>
          <button class="tool-guide-card tool-guide-card-excel" type="button" data-tool-guide="crossExcel" aria-label="打开跨 Excel 搜索与合并工具说明">
            <span class="tool-guide-card-top"><span class="tool-guide-icon" aria-hidden="true"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7"><rect x="4" y="4" width="16" height="16" rx="1.5"></rect><path d="M4 10h16M10 4v16M14 10v10"></path></svg></span><span class="tool-guide-badge">多文件</span></span>
            <span class="tool-guide-card-copy"><strong>跨 Excel</strong><small>跨文件搜索和按表头合并。</small></span>
            <span class="tool-guide-card-footer"><span>搜索与合并</span><span class="tool-guide-arrow" aria-hidden="true">→</span></span>
          </button>
        </div>
      </section>

      <section id="overviewPage" class="page-section">
        <div class="grid dashboard-grid">
          <section class="card">
            <div class="card-title">
              <h3>新建任务</h3>
              <p>选择目录和模式。</p>
            </div>
            <div class="grid two">
              <label>输入目录<span class="secret-field"><input id="folderPath" placeholder="D:\\项目\\文本表" /><button id="chooseFolderButton" class="mini-button" type="button">选择文件夹</button></span></label>
              <label>待提取列<select id="headerName"><option value="">请选择待提取列</option></select></label>
              <label>源语言<input id="sourceLanguage" value="Chinese" /></label>
              <label>运行模式<select id="extractionMode"><option value="terms">提取术语</option><option value="nontrans_only">仅提取非译元素</option></select></label>
            </div>
            <div class="actions">
              <button id="scanButton" class="secondary">扫描目录</button>
              <button id="startButton" class="primary">开始提取</button>
            </div>
            <pre id="scanResult" class="result-box">尚未扫描</pre>
          </section>

          <section class="card preprocess-progress-card">
            <div class="card-title">
              <div>
                <h3>任务进度</h3>
                <p id="preprocessProgressMessage">等待开始任务。</p>
              </div>
              <span id="preprocessProgressState" class="preprocess-status-pill">空闲</span>
            </div>
            <div class="preprocess-progress-overview">
              <div class="preprocess-progress-total">
                <span>已完成</span>
                <strong id="progressText">0 / 0</strong>
              </div>
              <div class="preprocess-progress-stage">
                <span>当前阶段</span>
                <strong id="preprocessProgressStage">未启动</strong>
              </div>
              <span id="preprocessProgressPercent" class="preprocess-progress-percent">未开始</span>
            </div>
            <div class="progress"><span id="progressBar"></span></div>
            <div class="preprocess-progress-metrics">
              <div><span>批次</span><strong id="batchText">0 / 0</strong></div>
              <div><span>成功</span><strong id="successText">0</strong></div>
              <div><span>失败</span><strong id="failureText">0</strong></div>
              <div><span>重试</span><strong id="retryText">0</strong></div>
              <div><span>并发</span><strong id="concurrencyText">0</strong></div>
              <div><span>AI 请求</span><strong id="preprocessLlmRequests">0</strong></div>
            </div>
            <section class="preprocess-events" aria-label="最近任务事件">
              <div class="preprocess-events-head"><strong>最近事件</strong><span>实时更新</span></div>
              <div id="preprocessRecentEvents" class="preprocess-events-list"><span class="preprocess-events-empty">暂无任务事件</span></div>
            </section>
            <div class="actions">
              <button id="resumeButton" class="secondary">继续任务</button>
              <button id="stopButton" class="danger">停止任务</button>
              <button id="clearCacheButton" class="danger">清空缓存</button>
            </div>
            <div id="errorPanel" class="error-panel" hidden>
              <strong>任务需要处理</strong>
              <span id="errorPanelText"></span>
            </div>
          </section>
        </div>

        <div class="grid two">
          <section class="card">
            <div class="card-title">
              <h3>最近输出</h3>
              <p></p>
            </div>
            <div class="result-file">
              <span>输出文件</span>
              <strong id="resultOutputFile">暂无输出</strong>
              <small id="resultState">等待任务完成。</small>
            </div>
            <div class="actions">
              <button id="downloadResultButton" class="secondary" disabled>打开结果文件</button>
              <button id="openResultFolderButton" class="secondary" disabled>打开输出目录</button>
            </div>
            <label>输出文件路径<input id="outputFile" readonly /></label>
            <label>最近错误<input id="lastError" readonly /></label>
          </section>

          <section class="card">
            <div class="card-title">
              <h3>本次结果</h3>
              <p></p>
            </div>
            <div class="metrics result-metrics">
              <div><span>术语库</span><strong id="resultTermCount">0</strong></div>
              <div><span>失败记录</span><strong id="resultFailureCount">0</strong></div>
              <div><span>非译元素正则</span><strong id="resultRegexCount">0</strong></div>
            </div>
          </section>
        </div>
      </section>

      <section id="modelSettingsPage" class="page-section">
        <section class="card">
          <div class="card-title">
            <h3>通用模型</h3>
            <p>统一设置当前任务使用的模型连接。</p>
          </div>
          <div class="grid two">
            <label>供应商<input value="DeepSeek" readonly /></label>
            <label>模型列表<select id="modelName"></select></label>
            <label>API Key<span class="secret-field"><input id="apiKey" type="password" autocomplete="off" placeholder="sk-..." /><button id="saveModelConnectionButton" class="mini-button" type="button">加载模型</button></span></label>
            <label>超时秒数<input id="timeoutSeconds" type="number" min="1" value="90" /></label>
            <label class="check"><input id="disableSystemProxy" type="checkbox" checked /> 禁用系统代理</label>
          </div>
          <div class="actions">
            <span id="modelConnectionHint" class="hint"></span>
          </div>
        </section>
      </section>

      <section id="modelStageSettingsPage" class="page-section">
        <div class="stage-grid">
          <section class="card compact-card">
            <div class="card-title">
              <h3>非译元素阶段</h3>
              <p></p>
            </div>
            <label>单次处理长度<input id="nontransLimit" type="number" min="200" value="5000" /></label>
            <label>思考强度（已开启）<select id="nontransReasoningEffort"><option value="low">低</option><option value="medium">中</option><option value="high">高</option></select></label>
          </section>
          <section class="card compact-card">
            <div class="card-title">
              <h3>术语召回阶段</h3>
              <p></p>
            </div>
            <label>单次处理长度<input id="recallLimit" type="number" min="200" value="5000" /></label>
            <label>思考强度（已开启）<select id="recallReasoningEffort"><option value="low">低</option><option value="medium">中</option><option value="high">高</option></select></label>
          </section>
          <section class="card compact-card">
            <div class="card-title">
              <h3>术语校验阶段</h3>
              <p></p>
            </div>
            <label>单次处理长度<input id="reviewLimit" type="number" min="200" value="5000" /></label>
            <label>上下文长度<input id="reviewContextLimit" type="number" min="50" max="2000" value="220" /></label>
            <label>思考强度（已开启）<select id="termReviewReasoningEffort"><option value="low">低</option><option value="medium">中</option><option value="high">高</option></select></label>
          </section>
        </div>

        <div class="sticky-actions">
          <button id="saveSettingsButton" class="primary">保存设置</button>
          <span id="saveHint" class="hint"></span>
        </div>
      </section>

      <section id="nontransSettingsPage" class="page-section">
        <div class="subnav">
          <button class="subnav-link active" data-subtab-group="nontrans" data-subtab-target="nontransRulePanel">宽泛检测规则</button>
          <button class="subnav-link" data-subtab-group="nontrans" data-subtab-target="nontransBuiltinPanel">内置规则库</button>
          <button class="subnav-link" data-subtab-group="nontrans" data-subtab-target="nontransProtectPanel">保护设置</button>
        </div>

        <div id="nontransRulePanel" class="subtab-panel active" data-subtab-panel-group="nontrans">
          <section class="card">
            <div class="card-title">
              <h3>宽泛检测规则</h3>
              <p>顺序可调整。</p>
            </div>
            <div class="pattern-table-wrap">
              <table class="pattern-table">
                <thead>
                  <tr><th>启用</th><th>名称</th><th>正则表达式</th><th>顺序</th><th>操作</th></tr>
                </thead>
                <tbody id="asciiPatternBody"></tbody>
              </table>
            </div>
            <div class="actions">
              <button id="addAsciiPatternButton" class="secondary">新增规则</button>
              <button id="saveAsciiPatternsButton" class="secondary">保存规则</button>
              <span id="asciiPatternHint" class="hint"></span>
            </div>
          </section>
        </div>

        <div id="nontransBuiltinPanel" class="subtab-panel" data-subtab-panel-group="nontrans">
          <section class="card">
            <div class="card-title">
              <h3>内置规则库</h3>
              <p></p>
            </div>
            <div class="summary-line">
              <span>规则家族：<strong id="builtinRuleCount">0</strong></span>
              <span>可执行正则：<strong id="builtinRuleRowCount">0</strong></span>
            </div>
            <div class="actions">
              <button id="addBuiltinRuleButton" class="secondary">新增规则</button>
              <button id="saveBuiltinRulesButton" class="secondary">保存规则库</button>
              <span id="builtinRuleHint" class="hint"></span>
            </div>
            <div class="pattern-table-wrap">
              <table class="pattern-table compact-rule-table builtin-editor-table">
                <thead>
                  <tr><th>启用</th><th>标题</th><th>正则表达式</th><th>命中方式</th><th>操作</th></tr>
                </thead>
                <tbody id="builtinRuleEditorBody"></tbody>
              </table>
            </div>
            <div class="pattern-table-wrap">
              <table class="pattern-table builtin-preview-table">
                <thead>
                  <tr><th>顺序</th><th>名称</th><th>开始/结束/空</th><th>类型</th><th>正则表达式</th><th>样例</th></tr>
                </thead>
                <tbody id="builtinRuleBody"></tbody>
              </table>
            </div>
            <div class="actions">
              <span class="hint">下方是保存后的执行预览。样例不会手动填写，规则实际命中后会自动记录，最多保留最新 3 个。</span>
            </div>
          </section>
        </div>

        <div id="nontransProtectPanel" class="subtab-panel" data-subtab-panel-group="nontrans">
          <div class="grid two">
            <section class="card">
              <div class="card-title">
                <h3>保护方式</h3>
                <p></p>
              </div>
              <label>占位符格式<input id="nontransPlaceholderFormat" value="<{n}>" /></label>
              <label class="check"><input id="numericNormalization" type="checkbox" checked /> 启用数值归一</label>
            </section>
            <section class="card">
              <div class="card-title">
                <h3>发现方式</h3>
                <p></p>
              </div>
              <label class="check"><input id="builtinRegex" type="checkbox" checked /> 启用内置规则库</label>
              <label class="check"><input id="aiDiscovery" type="checkbox" checked /> 启用 AI 发现</label>
              <label class="check"><input id="aiRegex" type="checkbox" checked /> 启用 AI 生成规则</label>
            </section>
          </div>
        </div>
      </section>

      <section id="promptSettingsPage" class="page-section">
        <div class="subnav">
          <button class="subnav-link active" data-subtab-group="prompt" data-subtab-target="promptRecallPanel">术语召回</button>
          <button class="subnav-link" data-subtab-group="prompt" data-subtab-target="promptReviewPanel">术语校验</button>
          <button class="subnav-link" data-subtab-group="prompt" data-subtab-target="promptNontransPanel">非译元素</button>
        </div>
        <div id="promptTemplateList">
          <div id="promptRecallPanel" class="subtab-panel active" data-subtab-panel-group="prompt"></div>
          <div id="promptReviewPanel" class="subtab-panel" data-subtab-panel-group="prompt"></div>
          <div id="promptNontransPanel" class="subtab-panel" data-subtab-panel-group="prompt"></div>
        </div>
        <div class="actions">
          <button id="savePromptTemplatesButton" class="secondary">保存提示词</button>
          <button id="resetPromptTemplatesButton" class="danger">恢复默认提示词</button>
          <span id="promptTemplateHint" class="hint"></span>
        </div>
      </section>

      <section id="runDetailsPage" class="page-section" data-internal-view="true" aria-hidden="true">
        <section class="card">
          <div class="stats-panel">
            <div class="stats-title">任务统计</div>
            <div class="metrics compact-metrics">
              <div><span>源文本</span><strong id="statSourceRecords">0</strong></div>
              <div><span>非译候选条目</span><strong id="statNontransCandidateRecords">0</strong></div>
              <div><span>非译元素</span><strong id="statNontransElements">0</strong></div>
              <div><span>非译正则</span><strong id="statNontransRegexRows">0</strong></div>
              <div><span>文本片段</span><strong id="statSegments">0</strong></div>
              <div><span>候选术语</span><strong id="statCandidates">0</strong></div>
              <div><span>正式术语</span><strong id="statApproved">0</strong></div>
              <div><span>可召回文本</span><strong id="statRecallableRecords">0</strong></div>
              <div><span>唯一召回文本</span><strong id="statUniqueRecallTexts">0</strong></div>
              <div><span>去重节省</span><strong id="statDedupedRecords">0</strong></div>
              <div><span>节省比例</span><strong id="statDedupeSavingsPercent">0%</strong></div>
              <div><span>召回批次</span><strong id="statRecallChunkBatches">0</strong></div>
              <div><span>保护改写</span><strong id="statProtectedChanged">0</strong></div>
              <div><span>清洗为空</span><strong id="statCleanLostMeaningful">0</strong></div>
              <div><span>数值归一</span><strong id="statNumericNormalized">0</strong></div>
            </div>
          </div>

          <div class="stats-panel">
            <div class="stats-title">AI 消耗</div>
            <div class="metrics compact-metrics">
              <div><span>AI 请求</span><strong id="statLlmRequests">0</strong></div>
              <div><span>总耗时(ms)</span><strong id="statLlmLatencyTotal">0</strong></div>
              <div><span>平均耗时(ms)</span><strong id="statLlmLatencyAvg">0</strong></div>
              <div><span>提示字符</span><strong id="statLlmPromptChars">0</strong></div>
              <div><span>Token 合计</span><strong id="statLlmTokens">0</strong></div>
            </div>
          </div>

          <pre id="logBox" class="log-box">暂无日志</pre>
        </section>
      </section>

      <section id="resultsPage" class="page-section" data-internal-view="true" aria-hidden="true">
        <div class="grid two">
          <section class="card">
            <div class="card-title">
              <h3>导出结果</h3>
              <p></p>
            </div>
            <div class="result-file">
              <span>输出文件</span>
              <strong id="resultOutputFileMirror">暂无输出</strong>
              <small id="resultStateMirror">等待任务完成。</small>
            </div>
            <div class="notice-list">
              <div>术语提取模式：输出术语库、失败记录、非译元素正则。</div>
              <div>仅提取非译元素模式：只输出非译元素正则。</div>
            </div>
          </section>
          <section class="card">
            <div class="card-title">
              <h3>数量概览</h3>
              <p></p>
            </div>
            <div class="metrics result-metrics">
              <div><span>术语库</span><strong id="resultTermCountMirror">0</strong></div>
              <div><span>失败记录</span><strong id="resultFailureCountMirror">0</strong></div>
              <div><span>非译元素正则</span><strong id="resultRegexCountMirror">0</strong></div>
            </div>
          </section>
        </div>
      </section>

      <section id="crossExcelPage" class="page-section">
        <div class="grid dashboard-grid">
          <section class="card">
            <div class="card-title">
              <h3>搜索范围</h3>
              <p>选择目录后可扫描全部 Excel 文件并加载表头。</p>
            </div>
            <div class="grid two">
              <label>输入目录
                <span class="secret-field">
                  <input id="crossExcelFolderPath" placeholder="D:\\项目\\Excel目录" />
                  <button id="chooseCrossExcelFolderButton" class="mini-button" type="button">选择文件夹</button>
                </span>
              </label>
              <div class="cross-summary-box">
                <span>文件数</span>
                <strong id="crossExcelFileCount">0</strong>
                <small id="crossExcelScanHint">请先选择目录并扫描。</small>
              </div>
            </div>
            <div class="actions">
              <button id="scanCrossExcelButton" class="secondary">扫描目录</button>
            </div>
          </section>

          <section class="card">
            <div class="card-title">
              <h3>全局搜索</h3>
              <p>输入关键词后，再预览命中的整行内容。</p>
            </div>
            <div class="grid two">
              <label>搜索内容
                <span class="search-inline-field">
                  <input id="crossExcelQuery" placeholder="输入要搜索的文字" />
                  <button id="searchCrossExcelButton" class="primary" type="button">搜索</button>
                </span>
              </label>
              <label>结果上限<input id="crossExcelLimit" type="number" min="1" max="2000" value="300" /></label>
            </div>
            <div class="actions action-row-compact">
              <span id="crossExcelSearchHint" class="hint"></span>
            </div>
            <div class="summary-line">
              <span>命中结果 <strong id="crossExcelMatchCount">0</strong></span>
              <span>扫描行数 <strong id="crossExcelScannedRows">0</strong></span>
              <span>状态 <strong id="crossExcelTruncatedLabel">未搜索</strong></span>
            </div>
          </section>
        </div>

        <div class="grid dashboard-grid">
          <section class="card">
            <div class="card-title">
              <h3>搜索预览</h3>
              <p>按整行展示。点击单元格即可复制内容。</p>
            </div>
            <div id="crossExcelSearchResults" class="cross-search-results">
              <div class="cross-empty-state">执行搜索后，这里会显示命中的行。</div>
            </div>
          </section>

          <section class="card">
            <div class="card-title">
              <h3>按表头合并</h3>
              <p>勾选需要保留的表头，再导出合并结果。</p>
            </div>
            <div class="actions">
              <button id="selectAllCrossHeadersButton" class="secondary" type="button">全选</button>
              <button id="clearCrossHeadersButton" class="secondary" type="button">清空</button>
              <button id="mergeCrossExcelButton" class="primary" type="button">合并</button>
            </div>
            <div id="crossExcelHeaderList" class="cross-header-list">
              <div class="cross-empty-state">扫描目录后会在这里显示全部表头。</div>
            </div>
            <div class="actions action-row-compact cross-merge-actions">
              <label class="check"><input id="crossExcelApplyFormat" type="checkbox" checked /> 保留单元格格式</label>
            </div>
            <div class="result-file">
              <span>合并结果</span>
              <strong id="crossExcelOutputFile">暂无输出</strong>
              <small id="crossExcelOutputHint">合并结果会保存到工具 output 目录。</small>
            </div>
            <div class="actions">
              <button id="openCrossExcelOutputFileButton" class="secondary" type="button" disabled>打开文件</button>
              <button id="openCrossExcelOutputButton" class="secondary" type="button" disabled>打开输出目录</button>
            </div>
          </section>
        </div>
      </section>

      <section id="diffExcelPage" class="page-section">
        <div class="grid dashboard-grid">
          <section class="card">
            <div class="card-title">
              <h3>比对范围</h3>
              <p>支持文件对文件，也支持目录对目录。</p>
            </div>
            <div class="grid two">
              <label>路径 A
                <span class="secret-field">
                  <input id="diffPathA" placeholder="D:\\项目\\旧版本 或旧文件.xlsx" />
                  <button id="chooseDiffPathAButton" class="mini-button" type="button">选择文件夹</button>
                  <button id="chooseDiffFileAButton" class="mini-button" type="button">选择文件</button>
                </span>
              </label>
              <label>路径 B
                <span class="secret-field">
                  <input id="diffPathB" placeholder="D:\\项目\\新版本 或新文件.xlsx" />
                  <button id="chooseDiffPathBButton" class="mini-button" type="button">选择文件夹</button>
                  <button id="chooseDiffFileBButton" class="mini-button" type="button">选择文件</button>
                </span>
              </label>
            </div>
            <div class="actions">
              <button id="startDiffExcelButton" class="primary" type="button">开始比对</button>
              <button id="clearDiffExcelButton" class="secondary" type="button">清空结果</button>
              <div class="diff-mode-actions">
                <select id="diffCompareMode" class="diff-mode-select" aria-label="比对方式">
                  <option value="position">按位置比对</option>
                  <option value="field_match">按字段匹配</option>
                </select>
                <button id="openDiffFieldSettingsButton" class="secondary icon-button" type="button" title="设置字段匹配" aria-label="设置字段匹配" hidden>⚙</button>
              </div>
              <span id="diffExcelHint" class="hint"></span>
            </div>
          </section>

          <section class="card">
            <div class="card-title">
              <h3>当前状态</h3>
              <p>比对完成后可直接导出结果或批量标记。</p>
            </div>
            <div class="metrics hero-metrics">
              <div><span>模式</span><strong id="diffModeLabel">未开始</strong></div>
              <div><span>A 文件数</span><strong id="diffFilesInA">0</strong></div>
              <div><span>B 文件数</span><strong id="diffFilesInB">0</strong></div>
              <div><span>配对数</span><strong id="diffMatchedPairs">0</strong></div>
              <div><span>差异数</span><strong id="diffTotalCount">0</strong></div>
              <div><span>预览数</span><strong id="diffVisibleCount">0</strong></div>
            </div>
            <div class="result-file">
              <span>导出文件</span>
              <strong id="diffOutputFile">暂无输出</strong>
              <small id="diffOutputHint">比对完成后可导出差异结果。</small>
            </div>
            <div class="actions">
              <button id="exportDiffExcelButton" class="secondary" type="button" disabled>导出结果</button>
              <button id="openDiffOutputFileButton" class="secondary" type="button" disabled>打开文件</button>
              <button id="openDiffOutputFolderButton" class="secondary" type="button" disabled>打开输出目录</button>
            </div>
          </section>
        </div>

        <div class="grid dashboard-grid">
          <section class="card">
            <div class="card-title">
              <h3>批量标记</h3>
              <p>把当前预览结果批量标记回原表。</p>
            </div>
            <div class="grid two">
              <label>标记到
                <select id="diffMarkTarget">
                  <option value="A">文件 A</option>
                  <option value="B">文件 B</option>
                </select>
              </label>
              <div class="color-picker-block">
                <label>颜色
                  <input id="diffHighlightColor" type="color" value="#FFD966" />
                </label>
                <div class="preset-color-row" aria-label="常用标记颜色">
                  <button type="button" class="preset-color-btn" data-color="#F44336" title="红色"></button>
                  <button type="button" class="preset-color-btn" data-color="#FF9800" title="橙色"></button>
                  <button type="button" class="preset-color-btn" data-color="#FFEB3B" title="黄色"></button>
                  <button type="button" class="preset-color-btn" data-color="#4CAF50" title="绿色"></button>
                  <button type="button" class="preset-color-btn" data-color="#2196F3" title="蓝色"></button>
                  <button type="button" class="preset-color-btn" data-color="#9C27B0" title="紫色"></button>
                </div>
              </div>
            </div>
            <div class="actions">
              <button id="highlightDiffExcelButton" class="primary" type="button" disabled>标记差异结果</button>
              <span id="diffHighlightHint" class="hint"></span>
            </div>
          </section>
        </div>

        <section class="card">
          <div class="card-title">
            <h3>差异预览</h3>
          </div>
          <div class="pattern-table-wrap diff-preview-table-wrap">
            <table class="pattern-table diff-preview-table">
              <colgroup>
                <col class="diff-preview-file-column" />
                <col class="diff-preview-file-column" />
                <col class="diff-preview-sheet-column" />
                <col class="diff-preview-location-column" />
                <col class="diff-preview-comparison-column" />
              </colgroup>
              <thead>
                <tr id="diffExcelHead">
                  <th>文件 A</th>
                  <th>文件 B</th>
                  <th>Sheet</th>
                  <th>单元格</th>
                  <th>差异对照</th>
                </tr>
              </thead>
              <tbody id="diffExcelBody">
                <tr><td colspan="5" class="empty-cell">暂无差异结果</td></tr>
              </tbody>
            </table>
          </div>
          <div class="actions action-row-compact">
            <button id="diffPreviewPreviousButton" class="secondary" type="button" disabled>上一页</button>
            <span id="diffPreviewPageLabel" class="hint">暂无结果</span>
            <button id="diffPreviewNextButton" class="secondary" type="button" disabled>下一页</button>
          </div>
        </section>
      </section>

      <div id="diffFieldSettingsOverlay" class="modal-overlay" hidden>
        <section class="modal-card diff-field-settings-modal" role="dialog" aria-modal="true" aria-labelledby="diffFieldSettingsTitle">
          <header class="modal-header">
            <div>
              <h3 id="diffFieldSettingsTitle">字段匹配设置</h3>
              <p id="diffFieldSettingsSummary">选择用于匹配数据行和比较内容的表头。</p>
            </div>
            <button id="closeDiffFieldSettingsButton" class="modal-close" type="button" aria-label="关闭">×</button>
          </header>
          <div class="diff-field-settings-grid">
            <label>参考字段
              <select id="diffReferenceField"><option value="">请选择参考字段</option></select>
            </label>
            <div class="diff-header-presence">
              <span>文件 A 表头 <strong id="diffHeadersACount">0</strong></span>
              <span>文件 B 表头 <strong id="diffHeadersBCount">0</strong></span>
              <span>共同表头 <strong id="diffHeadersCommonCount">0</strong></span>
            </div>
          </div>
          <div class="diff-compare-fields-block">
            <div class="field-block-head">
              <strong>比对字段</strong>
              <span>仅显示两侧均存在的表头</span>
            </div>
            <div id="diffCompareFieldList" class="diff-compare-field-list"></div>
          </div>
          <label class="check-line diff-unmatched-line">
            <input id="diffIncludeUnmatched" type="checkbox" />
            <span>标记未匹配行</span>
            <span class="info-tip" tabindex="0" aria-label="说明">i<span class="info-tip-content">开启后，其中一侧存在另一侧没有的参考值时也会列为差异；批量标记时会标记该侧对应的整行。关闭则会忽略该参考值。</span></span>
          </label>
          <div class="modal-footer">
            <span id="diffFieldSettingsHint" class="hint"></span>
            <div class="modal-footer-actions">
              <button id="cancelDiffFieldSettingsButton" class="secondary" type="button">取消</button>
              <button id="saveDiffFieldSettingsButton" class="primary" type="button">确定</button>
            </div>
          </div>
        </section>
      </div>

      <section id="aiReviewTaskPage" class="page-section">
        <div class="review-conversation-shell">
          <aside class="review-session-sidebar">
            <button id="newReviewConversationButton" class="primary review-new-chat" type="button">＋ 新建审校</button>
            <div id="reviewConversationList" class="review-conversation-list"></div>
          </aside>
          <div class="review-conversation-main">
            <div class="review-conversation-head">
              <div>
                <div class="review-conversation-title-row">
                  <h3 id="reviewConversationTitle">新审校</h3>
                  <input id="reviewConversationTitleInput" class="review-title-input hidden" type="text" maxlength="120" aria-label="会话名称" />
                  <button id="renameReviewConversationButton" class="review-title-edit" type="button" aria-label="重命名会话" title="重命名会话">✎</button>
                </div>
                <p id="reviewConversationStatus">添加文件或直接输入待审校文本</p>
              </div>
              <div class="actions compact-actions">
                <label class="check-line"><input id="reviewAutoStart" type="checkbox" /><span>识别后自动开始</span></label>
                <button id="deleteReviewConversationButton" class="secondary danger-text" type="button">删除会话</button>
              </div>
            </div>
            <div id="reviewConversationMessages" class="review-conversation-messages"></div>
            <div id="reviewComposer" class="review-composer">
              <div id="reviewAttachmentChips" class="review-attachment-chips"></div>
              <textarea id="reviewComposerInput" rows="3" placeholder="输入待审校文本，或拖拽多个文件到这里…"></textarea>
              <div class="review-composer-footer">
                <div class="review-composer-tools">
                  <button id="reviewPromptChip" class="composer-chip" type="button">提示词</button>
                  <button id="reviewSourceLanguageChip" class="composer-chip" type="button">源语言：自动</button>
                  <button id="reviewTargetLanguageChip" class="composer-chip" type="button">目标语言：自动</button>
                  <button id="reviewTermBaseChip" class="composer-chip review-term-base-chip" type="button" title="选择术语表">术语表</button>
                  <button id="reviewUploadChip" class="composer-chip" type="button">＋ 文件</button>
                  <input id="reviewConversationFileInput" type="file" accept=".xlsx,.xlsm,.xlf,.xliff,.csv,.tsv,.txt,.md,.docx,.pptx,.pdf,.json,.xml" multiple hidden />
                  <input id="reviewTermBaseFileInput" type="file" accept=".xlsx,.xlsm" hidden />
                </div>
                <button id="sendReviewConversationButton" class="primary review-send-button" type="button">发送</button>
              </div>
              <div id="reviewLanguagePopover" class="review-language-popover hidden" role="dialog" aria-label="选择语言">
                <div class="review-language-popover-head">
                  <button id="closeReviewLanguagePopoverButton" class="review-language-back" type="button" aria-label="收起语言列表">‹</button>
                  <strong id="reviewLanguagePopoverTitle">选择语言</strong>
                  <span id="reviewLanguagePopoverHint">单选</span>
                </div>
                <div class="review-language-search-wrap">
                  <span>⌕</span><input id="reviewLanguageSearch" type="search" placeholder="搜索语言" autocomplete="off" />
                </div>
                <div id="reviewLanguageOptions" class="review-language-options"></div>
              </div>
              <div id="reviewTermBasePopover" class="review-term-base-popover hidden" role="menu" aria-label="选择术语表">
                <button id="uploadReviewTermBaseButton" class="review-term-base-action" type="button" role="menuitem">
                  <span class="review-term-base-action-icon">＋</span><span><strong>上传术语表</strong><small>支持 .xlsx、.xlsm；自动识别语种列</small></span>
                </button>
                <button id="clearReviewTermBaseButton" class="review-term-base-action" type="button" role="menuitem">
                  <span class="review-term-base-action-icon">∅</span><span><strong>不使用术语表</strong><small>当前会话不进行术语匹配</small></span>
                </button>
                <div id="reviewTermBaseOptions" class="review-term-base-options"></div>
              </div>
              <span id="reviewConversationHint" class="hint"></span>
            </div>
          </div>
        </div>

        <dialog id="reviewAttachmentMappingDialog" class="modal">
          <form method="dialog" class="modal-card review-attachment-mapping-card">
            <div class="review-mapping-file-head">
              <div class="review-mapping-file-icon" aria-hidden="true">XLSX</div>
              <div class="review-mapping-file-copy">
                <span>文件导入设置</span>
                <h2 id="reviewAttachmentMappingTitle">文件导入方式</h2>
              </div>
              <button id="closeReviewAttachmentMappingButton" class="review-mapping-close" value="cancel" type="submit" aria-label="关闭">×</button>
            </div>
            <p class="review-mapping-lead">选择识别方式。Excel 可读取实际工作表和表头后手动调整映射。</p>
            <div class="review-import-mode-grid">
              <label class="review-import-mode-option">
                <input type="radio" name="reviewAttachmentMappingMode" value="ai" checked />
                <span class="review-import-mode-icon review-import-mode-ai" aria-hidden="true">✦</span>
                <span><strong>AI 识别 <em>默认</em></strong><small>由 Workspace Agent 判断原文、译文和同一行上下文列</small></span>
                <i aria-hidden="true">✓</i>
              </label>
              <label class="review-import-mode-option">
                <input type="radio" name="reviewAttachmentMappingMode" value="preset" />
                <span class="review-import-mode-icon" aria-hidden="true">▦</span>
                <span><strong>手动映射</strong><small>读取工作表结构，选择并编辑原文、译文和信息列</small></span>
                <i aria-hidden="true">✓</i>
              </label>
            </div>
            <div id="reviewAttachmentPresetField" class="review-mapping-preset-panel hidden">
              <div class="review-mapping-field-label"><label for="reviewAttachmentPresetSelect">映射模板</label><span>可选</span></div>
              <div class="review-mapping-preset-row">
                <select id="reviewAttachmentPresetSelect"></select>
                <button id="editReviewAttachmentMappingButton" class="secondary" type="button">读取结构并编辑</button>
              </div>
              <small>编辑器会显示当前文件的工作表、表头和列位置；现有模板也可以直接修改并更新。</small>
            </div>
            <p id="reviewAttachmentMappingHint" class="hint"></p>
            <div class="modal-footer">
              <button class="secondary" value="cancel" type="submit">取消</button>
              <button id="saveReviewAttachmentMappingButton" class="primary" type="button">确认</button>
            </div>
          </form>
        </dialog>

        <section id="reviewConversationResultCard" class="card review-conversation-results hidden">
          <div class="card-title">
            <div><h3>审校结果预览</h3><p>每个目标语言独立审校并生成结果文件。</p></div>
            <div class="actions compact-actions">
              <div id="reviewTargetTabs" class="review-target-tabs"></div>
            </div>
          </div>
          <div class="review-inline-progress">
            <span id="reviewProgress">尚未开始</span>
            <strong id="reviewProgressCount">0 / 0</strong>
            <div class="progress"><span id="reviewProgressBar"></span></div>
            <span id="reviewFailedCount" class="hidden">0</span>
            <span id="reviewRequestedCount" class="hidden">0</span>
            <span id="reviewProgressPercent" class="hidden">未开始</span>
          </div>
          <section id="reviewRequestQueue" class="review-request-queue hidden" aria-label="AI 请求状态">
            <div class="review-request-queue-head"><strong>AI 请求状态</strong><span id="reviewRequestQueueSummary">等待创建请求</span></div>
            <div id="reviewRequestColumns" class="review-request-columns"></div>
          </section>
          <div id="outputPanel" class="result-file hidden"><span>结果文件</span><strong id="outputPath">暂无输出</strong></div>
          <div class="actions">
            <button id="openOutputDirButton" class="secondary" type="button">打开输出目录</button>
            <button id="openOutputFileButton" class="secondary" type="button" disabled>打开结果文件</button>
          </div>
          <div class="review-result-preview-head">
            <span>预览</span>
            <button id="openReviewDetailButton" class="secondary" type="button" disabled>详情</button>
          </div>
          <div class="pattern-table-wrap review-result-table-wrap">
            <table class="pattern-table review-result-table">
              <thead id="reviewResultHead"><tr><th>原文</th><th>译文</th><th>是否有问题</th><th>问题类型</th><th>问题说明</th><th>修改建议</th></tr></thead>
              <tbody id="reviewResultBody"><tr><td colspan="6" class="empty-cell">暂无审校结果</td></tr></tbody>
            </table>
          </div>
        </section>

        <div class="hidden" aria-hidden="true">
          <div id="reviewFilePath"></div><button id="chooseReviewFileButton" type="button"></button>
          <input id="reviewFileInput" type="file" /><button id="openExcelMappingButton" type="button"></button>
          <div id="reviewBatchCount"></div><div id="reviewFileHint"></div><div id="excelMappingSummary"></div>
          <table><tbody id="previewBody"></tbody></table><input id="sourceLanguageInput" /><input id="targetLanguageInput" />
          <button id="startReviewButton" type="button"></button><button id="openReviewSettingsButton" type="button"></button>
          <button id="openReviewForbiddenButton" type="button"></button><span id="reviewTaskHint"></span><ol id="reviewLogList"></ol>
        </div>

      </section>

      <section id="aiReviewSettingsPage" class="page-section">
        <div class="grid one">
          <section class="card">
            <div class="card-title">
              <h3>参数设置</h3>
              <p>Workspace Agent 与 Task Workflow 统一使用“模型设置”中的当前模型和 API 配置。</p>
            </div>
            <div class="grid two">
              <label>思考强度<select id="reviewReasoningEffort"><option value="low">低</option><option value="medium">中</option><option value="high">高</option></select></label>
              <label class="check-line"><input id="reviewDebugPayloadLogging" type="checkbox" /><span>记录调试请求内容</span></label>
              <label>单包正文预算<input id="reviewAiLimit" type="number" min="200" value="1500" /></label>
              <label>单包条目上限<input id="reviewMaxItems" type="number" min="1" max="500" value="20" /></label>
            </div>
            <div class="actions">
              <button id="saveReviewAgentSettingsButton" class="primary" type="button">保存设置</button>
            </div>
            <span id="reviewSettingsHint" class="hint"></span>
            <div class="hidden" aria-hidden="true">
              <input id="enableAiReview" type="checkbox" checked /><input id="enableDirectionalReview" type="checkbox" />
              <span id="directionalReviewLine"></span><select id="promptTemplateSelect"></select>
              <span id="directionalTemplatePanel"><select id="directionalTemplateSelect"></select></span>
              <button id="editPromptButton" type="button"></button><button id="editDirectionalButton" type="button"></button>
            </div>
          </section>
        </div>
      </section>

      <section id="aiReviewForbiddenPage" class="page-section hidden" aria-hidden="true">
        <input id="enableForbiddenCheck" type="checkbox" /><span id="forbiddenTemplatePanel"><select id="forbiddenTemplateSelect"></select></span>
        <button id="editForbiddenButton" type="button"></button><span id="reviewForbiddenHint"></span>
      </section>

    </main>
  </div>

  <div id="pendingRuleOverlay" class="modal-overlay" hidden>
    <div class="modal-card notice-modal">
      <div class="modal-header">
        <div>
          <h3>发现新的非译规则</h3>
          <p>这些规则来自最近的任务结果。确认后会写入内置规则库。</p>
        </div>
        <button id="closePendingRuleModalButton" class="modal-close" type="button" aria-label="关闭">×</button>
      </div>
      <div class="modal-actions">
        <button id="pendingRuleSelectAllButton" class="secondary" type="button">全选</button>
        <button id="pendingRuleClearSelectionButton" class="secondary" type="button">取消全选</button>
      </div>
      <div class="pattern-table-wrap">
        <table class="pattern-table compact-rule-table pending-rule-table">
          <thead>
            <tr><th>导入</th><th>标题</th><th>正则表达式</th><th>命中方式</th><th>样例</th></tr>
          </thead>
          <tbody id="pendingRuleEditorBody"></tbody>
        </table>
      </div>
      <div class="modal-footer">
        <span id="pendingRuleHint" class="hint"></span>
        <button id="confirmPendingRuleImportButton" class="primary" type="button">导入到内置规则库</button>
      </div>
    </div>
  </div>
  <div id="reviewDetailOverlay" class="modal-overlay" hidden>
    <div class="modal-card review-detail-modal">
      <div class="modal-header">
        <div>
          <h3>问题详情</h3>
          <p id="reviewDetailSummary">暂无问题条目</p>
        </div>
        <button id="closeReviewDetailButton" class="modal-close" type="button" aria-label="关闭">×</button>
      </div>
      <div id="reviewDetailGroups" class="review-detail-wrap">
        <div class="empty-cell">暂无问题条目</div>
      </div>
    </div>
  </div>
  <div id="reviewFollowupOverlay" class="modal-overlay" hidden>
    <div class="modal-card review-followup-modal">
      <div class="modal-header">
        <div>
          <h3>追问审校结果</h3>
          <p id="reviewFollowupSummary">围绕当前问题继续提问</p>
        </div>
        <button id="closeReviewFollowupButton" class="modal-close" type="button" aria-label="关闭">×</button>
      </div>
      <div class="review-followup-layout">
        <aside id="reviewFollowupContext" class="review-followup-context"></aside>
        <section class="review-followup-chat">
          <div id="reviewFollowupMessages" class="review-followup-messages"></div>
          <div class="review-followup-input">
            <textarea id="reviewFollowupInput" placeholder="输入你的追问，例如：这个建议为什么这样改？有没有更自然的译法？"></textarea>
            <button id="sendReviewFollowupButton" class="primary" type="button">发送</button>
          </div>
          <span id="reviewFollowupHint" class="hint"></span>
        </section>
      </div>
    </div>
  </div>

  <dialog id="promptDialog" class="dialog">
    <form method="dialog" class="dialog-body">
      <div class="dialog-head">
        <div>
          <h2>提示词模板</h2>
          <p>用户提示词必须包含 {text}，工具会把待审校 JSON 条目填入这里。</p>
        </div>
        <button id="closePromptDialogButton" class="icon-button" type="button" aria-label="关闭">×</button>
      </div>
      <input id="promptTemplateId" type="hidden" />
      <div class="field">
        <label for="promptDialogTemplateSelect">当前模板</label>
        <select id="promptDialogTemplateSelect"></select>
      </div>
      <div class="field">
        <label for="promptNameInput">模板名</label>
        <input id="promptNameInput" type="text" />
      </div>
      <div class="field">
        <label for="systemPromptInput">系统提示词</label>
        <textarea id="systemPromptInput" rows="6"></textarea>
      </div>
      <div class="field">
        <label for="userPromptInput">用户提示词</label>
        <textarea id="userPromptInput" rows="10"></textarea>
      </div>
      <div class="field">
        <label for="promptForbiddenWordsInput">禁用词（可选，每行一个）</label>
        <textarea id="promptForbiddenWordsInput" rows="5" placeholder="命中内容会由本地规则检查，不交给模型判断"></textarea>
      </div>
      <div class="dialog-actions">
        <button id="savePromptButton" type="button">保存模板</button>
        <button id="newPromptButton" class="secondary" type="button">新建模板</button>
        <button id="resetPromptButton" class="secondary" type="button">恢复默认</button>
        <button id="deletePromptButton" class="danger" type="button">删除模板</button>
        <button id="cancelPromptDialogButton" class="secondary" type="button">取消</button>
      </div>
    </form>
  </dialog>

  <dialog id="toolGuideDialog" class="dialog tool-guide-dialog">
    <div class="dialog-card">
      <header class="dialog-header">
        <div>
          <span class="dialog-kicker">工具说明</span>
          <h3 id="toolGuideDialogTitle">工具说明</h3>
        </div>
        <button id="closeToolGuideDialogButton" class="icon-button" type="button" aria-label="关闭">×</button>
      </header>
      <article id="toolGuideDialogBody" class="markdown-guide"></article>
    </div>
  </dialog>

  <dialog id="directionalDialog" class="dialog">
    <form method="dialog" class="dialog-body">
      <div class="dialog-head">
        <div>
          <h2>定向审校模板</h2>
          <p>每个启用的审校类型会成为结果 Excel 中的一列。</p>
        </div>
        <button id="closeDirectionalDialogButton" class="icon-button" type="button" aria-label="关闭">×</button>
      </div>
      <input id="directionalTemplateId" type="hidden" />
      <div class="field">
        <label for="directionalNameInput">模板名</label>
        <input id="directionalNameInput" type="text" />
      </div>
      <div id="directionalItems" class="directional-items"></div>
      <div class="actions">
        <button id="addDirectionalItemButton" class="secondary" type="button">新建选项</button>
      </div>
      <div class="dialog-actions">
        <button id="saveDirectionalButton" type="button">保存模板</button>
        <button id="newDirectionalButton" class="secondary" type="button">新建模板</button>
        <button id="cancelDirectionalDialogButton" class="secondary" type="button">取消</button>
      </div>
    </form>
  </dialog>

  <dialog id="forbiddenDialog" class="dialog">
    <form method="dialog" class="dialog-body">
      <div class="dialog-head">
        <div>
          <h2>禁用词模板</h2>
          <p>一行一个禁用词；检查译文，不区分大小写，包含即命中。</p>
        </div>
        <button id="closeForbiddenDialogButton" class="icon-button" type="button" aria-label="关闭">×</button>
      </div>
      <input id="forbiddenTemplateId" type="hidden" />
      <div class="field">
        <label for="forbiddenNameInput">模板名</label>
        <input id="forbiddenNameInput" type="text" />
      </div>
      <div class="field">
        <label for="forbiddenWordsInput">禁用词列表</label>
        <textarea id="forbiddenWordsInput" rows="12"></textarea>
      </div>
      <div class="dialog-actions">
        <button id="saveForbiddenButton" type="button">保存模板</button>
        <button id="newForbiddenButton" class="secondary" type="button">新建模板</button>
        <button id="cancelForbiddenDialogButton" class="secondary" type="button">取消</button>
      </div>
    </form>
  </dialog>
  <dialog id="excelMappingDialog" class="dialog wide-dialog">
    <form method="dialog" class="dialog-body">
      <div class="dialog-head">
        <div>
          <h2 id="excelMappingDialogTitle">Excel 映射</h2>
          <p id="excelMappingDialogSubtitle">按 sheet 选择原文列、译文列和信息列。</p>
        </div>
        <button id="closeExcelMappingDialogButton" class="icon-button" type="button" aria-label="关闭">×</button>
      </div>
      <div class="grid two">
        <div class="field">
          <label for="mappingSourceLanguageInput">原文语种</label>
          <input id="mappingSourceLanguageInput" type="text" />
        </div>
        <div class="field">
          <label for="mappingTargetLanguageInput">译文语种</label>
          <input id="mappingTargetLanguageInput" type="text" />
        </div>
      </div>
      <div class="mapping-template-bar">
        <label for="excelMappingPresetSelect">映射模板</label>
        <select id="excelMappingPresetSelect" title="选择后立即套用"></select>
        <button id="saveExcelMappingPresetButton" class="secondary" type="button">保存当前映射</button>
        <button id="deleteExcelMappingPresetButton" class="icon-button danger mapping-template-delete" type="button" aria-label="删除模板" title="删除当前模板" disabled>&#128465;</button>
      </div>
      <p id="excelMappingPresetHint" class="mapping-template-hint" aria-live="polite"></p>
      <div id="excelSheetTabs" class="sheet-tabs"></div>
      <div id="excelMappingColumns" class="excel-mapping-columns"></div>
      <div class="dialog-actions">
        <button id="applyExcelMappingButton" type="button">确认读取</button>
        <button id="cancelExcelMappingDialogButton" class="secondary" type="button">取消</button>
      </div>
    </form>
  </dialog>
  <dialog id="excelMappingPresetDialog" class="dialog compact-dialog">
    <form method="dialog" class="dialog-body">
      <div class="dialog-head">
        <div>
          <h2>保存映射模板</h2>
        </div>
        <button id="closeExcelMappingPresetDialogButton" class="icon-button" type="button" aria-label="关闭">×</button>
      </div>
      <div class="field">
        <label for="excelMappingPresetNameInput">模板名称</label>
        <input id="excelMappingPresetNameInput" type="text" maxlength="80" autocomplete="off" />
      </div>
      <p id="excelMappingPresetSaveHint" class="mapping-template-hint" aria-live="polite"></p>
      <div class="dialog-actions">
        <button id="confirmExcelMappingPresetButton" type="button">保存</button>
        <button id="cancelExcelMappingPresetDialogButton" class="secondary" type="button">取消</button>
      </div>
    </form>
  </dialog>
  <div id="appUpdateOverlay" class="modal-overlay" hidden>
    <div class="modal-card notice-modal update-modal">
      <div class="modal-header">
        <div>
          <h3>发现新版本</h3>
          <p id="appUpdateSummary">当前版本与最新版本不一致。</p>
        </div>
        <button id="closeAppUpdateModalButton" class="modal-close" type="button" aria-label="关闭">×</button>
      </div>
      <div class="summary-line">
        <span>当前版本 <strong id="appUpdateCurrentVersion">-</strong></span>
        <span>最新版本 <strong id="appUpdateLatestVersion">-</strong></span>
      </div>
      <div class="pattern-table-wrap">
        <div id="appUpdateReleaseNotes" class="update-release-notes">暂无更新日志</div>
      </div>
      <div class="modal-footer">
        <span id="appUpdateHint" class="hint"></span>
        <div class="modal-footer-actions">
          <button id="cancelAppUpdateButton" class="secondary" type="button">取消</button>
          <button id="confirmAppUpdateButton" class="primary" type="button">立即更新</button>
        </div>
      </div>
    </div>
  </div>
  <div id="feedbackOverlay" class="modal-overlay" hidden>
    <div class="modal-card feedback-modal">
      <div class="modal-header">
        <div>
          <h3>我要反馈</h3>
          <p>问题为必填。当前日志会自动附带，截图可按需补充。</p>
        </div>
        <button id="closeFeedbackModalButton" class="modal-close" type="button" aria-label="关闭">×</button>
      </div>
      <div class="grid one">
        <label>问题<textarea id="feedbackMessageInput" rows="7" placeholder="请尽量写清楚出现了什么问题、在哪一步出现。"></textarea></label>
        <div class="feedback-attachment-card">
          <div class="feedback-attachment-head">
            <strong>截图</strong>
            <button id="chooseFeedbackScreenshotButton" class="secondary" type="button">选择截图</button>
          </div>
          <button id="feedbackScreenshotDropzone" class="feedback-dropzone" type="button">
            <span class="feedback-dropzone-title">点击选择截图 / 直接拖入这里</span>
            <span class="feedback-dropzone-subtitle">弹窗打开时也可以直接粘贴截图</span>
          </button>
          <small id="feedbackScreenshotName">未选择截图</small>
          <input id="feedbackScreenshotInput" type="file" accept="image/*" hidden />
        </div>
        <div class="feedback-log-card">
          <div>
            <strong>日志</strong>
            <small id="feedbackLogHint">会自动附带当前日志。</small>
          </div>
          <div class="actions feedback-log-actions">
            <button id="openFeedbackLogButton" class="secondary" type="button">打开日志</button>
          </div>
          <code id="feedbackLogPath">output/log.txt</code>
        </div>
      </div>
      <div class="modal-footer">
        <span id="feedbackSubmitHint" class="hint"></span>
        <div class="modal-footer-actions">
          <button id="cancelFeedbackButton" class="secondary" type="button">取消</button>
          <button id="submitFeedbackButton" class="primary" type="button">提交反馈</button>
        </div>
      </div>
    </div>
  </div>
  <script src="/assets/app.js"></script>
</body>
</html>
"""


def build_index_html() -> str:
    return INDEX_HTML.replace("__APP_VERSION__", APP_VERSION)


APP_CSS = r"""
:root {
  --bg: #eef3f8;
  --panel: #ffffff;
  --ink: #162033;
  --muted: #66748a;
  --line: #dbe4ef;
  --primary: #1f6f68;
  --primary-strong: #16564f;
  --danger: #b4443f;
  --shadow: 0 18px 45px rgba(22, 32, 51, 0.10);
}
* { box-sizing: border-box; }
body {
  margin: 0;
  background:
    radial-gradient(circle at top left, rgba(31,111,104,.18), transparent 34rem),
    linear-gradient(135deg, #f8fbfd 0%, var(--bg) 55%, #e7eef6 100%);
  color: var(--ink);
  font-family: "Microsoft YaHei UI", "Segoe UI", sans-serif;
  font-size: 15px;
}
.hidden { display: none !important; }
.shell { display: grid; grid-template-columns: 260px 1fr; min-height: 100vh; }
.sidebar {
  padding: 28px 22px;
  background: rgba(255,255,255,.74);
  border-right: 1px solid var(--line);
  backdrop-filter: blur(18px);
}
.brand { display: flex; gap: 12px; align-items: center; margin-bottom: 36px; }
.brand-mark { width: 38px; height: 38px; border-radius: 13px; object-fit: cover; display: block; box-shadow: var(--shadow); background: #ffffff; }
.brand-copy { min-width: 0; display: grid; gap: 2px; }
.brand-version {
  display: inline-flex;
  width: fit-content;
  padding: 2px 8px;
  border-radius: 999px;
  background: #e7f1ef;
  color: var(--primary-strong);
  font-size: 11px;
  font-weight: 800;
  letter-spacing: 0.03em;
}
.brand strong { display:block; font-size: 18px; line-height: 1.2; }
.brand small { color: var(--muted); line-height: 1.3; }
.sidebar-nav { display: grid; gap: 12px; margin-bottom: 22px; }
.nav-group { display: grid; gap: 8px; }
.nav-group-title {
  padding: 0;
  color: #24364d;
  font-size: 14px;
  font-weight: 700;
  line-height: 1.25;
  letter-spacing: 0.02em;
}
.nav-accordion {
  border: 1px solid rgba(209, 220, 234, 0.9);
  border-radius: 16px;
  background: rgba(255,255,255,.64);
  overflow: hidden;
  box-shadow: 0 8px 22px rgba(22, 32, 51, 0.04);
}
.nav-accordion[open] {
  background: rgba(255,255,255,.9);
  box-shadow: 0 14px 28px rgba(22, 32, 51, 0.08);
}
.nav-accordion-summary {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  padding: 12px 14px;
  cursor: pointer;
  list-style: none;
}
.nav-accordion-summary::-webkit-details-marker { display: none; }
.nav-accordion-summary::after {
  content: "+";
  width: 24px;
  height: 24px;
  display: inline-grid;
  place-items: center;
  border-radius: 999px;
  background: #f1f5f9;
  color: #64748b;
  font-size: 15px;
  font-weight: 700;
  line-height: 1;
}
.nav-accordion[open] .nav-accordion-summary::after { content: "-"; }
.nav-submenu {
  display: grid;
  gap: 6px;
  padding: 0 10px 10px;
}
.nav-accordion:not([open]) > .nav-submenu {
  display: none;
}
.nav-link {
  width: 100%;
  justify-content: flex-start;
  min-height: 40px;
  background: transparent;
  color: #526277;
  border: 0;
  text-align: left;
  padding: 0 14px;
  border-radius: 12px;
  font-weight: 600;
  transition: background .18s ease, color .18s ease, transform .18s ease;
}
.nav-link-top {
  min-height: 46px;
  padding: 0 14px;
  font-size: 15px;
  font-weight: 700;
  color: #203248;
  background: rgba(255,255,255,.72);
  border: 1px solid rgba(209, 220, 234, 0.95);
  box-shadow: 0 8px 20px rgba(22, 32, 51, 0.05);
}
.nav-link-sub {
  padding-left: 14px;
  font-size: 14px;
  min-height: 38px;
}
.nav-link-placeholder {
  opacity: 0.58;
  border: 1px dashed #cbd8e6;
  background: rgba(237,243,248,.45);
}
.nav-link.active, .nav-link:hover {
  background: linear-gradient(180deg, #ecf7f5 0%, #e2f0ed 100%);
  color: var(--primary-strong);
}
.nav-link:hover { transform: translateY(-1px); }
.sidebar-status {
  display: grid;
  gap: 12px;
  padding: 14px;
  background: rgba(255,255,255,.78);
  border: 1px solid var(--line);
  border-radius: 18px;
}
.sidebar-actions-panel {
  margin-top: 14px;
}
.feedback-entry-button {
  width: 100%;
  min-height: 44px;
  border: 1px solid rgba(31, 111, 104, 0.24);
  border-radius: 14px;
  background: linear-gradient(135deg, #f7fbfb 0%, #e6f3f1 100%);
  color: var(--primary-strong);
  font-size: 15px;
  font-weight: 800;
  cursor: pointer;
  transition: transform .18s ease, box-shadow .18s ease, border-color .18s ease;
}
.feedback-entry-button:hover {
  transform: translateY(-1px);
  border-color: rgba(31, 111, 104, 0.42);
  box-shadow: 0 12px 26px rgba(31, 111, 104, 0.12);
}
.status-block {
  display: grid;
  gap: 6px;
  padding-bottom: 10px;
  border-bottom: 1px solid rgba(219,228,239,.9);
}
.status-block:last-of-type {
  padding-bottom: 0;
  border-bottom: 0;
}
.status-caption {
  color: var(--muted);
  font-size: 12px;
  font-weight: 800;
  letter-spacing: 0.06em;
  text-transform: uppercase;
}
.service-tip { color: var(--muted); line-height: 1.5; }
.service-tip code { font-family: "Cascadia Mono", "Consolas", monospace; font-size: 12px; }
.content { padding: 28px; max-width: 1280px; width: 100%; }
.hero {
  display: grid;
  grid-template-columns: minmax(0, 1fr) 260px;
  gap: 20px;
  align-items: stretch;
  margin-bottom: 20px;
}
.hero h1 { margin: 0; font-size: clamp(28px, 3vw, 44px); line-height: 1.12; letter-spacing: -0.04em; max-width: 840px; }
.eyebrow { margin: 0 0 10px; color: var(--primary); font-weight: 800; letter-spacing: .08em; text-transform: uppercase; }
.lede { color: var(--muted); font-size: 16px; max-width: 760px; }
.hero-card, .card, .advanced-card {
  background: rgba(255,255,255,.88);
  border: 1px solid rgba(219,228,239,.9);
  border-radius: 24px;
  box-shadow: var(--shadow);
}
.hero-card { padding: 22px; display: grid; align-content: center; gap: 10px; }
.hero-card strong { font-size: 26px; }
.hero-card-label { color: var(--muted); font-weight: 700; }
.hero-card-meta {
  color: #607287;
  font-size: 13px;
  line-height: 1.55;
  padding-top: 2px;
}
.pill { width: max-content; padding: 7px 11px; border-radius: 999px; background: #e7f1ef; color: var(--primary-strong); font-weight: 800; }
.pill.running { background: #fff2cc; color: #865d10; }
.pill.failed { background: #fde5e2; color: var(--danger); }
.card { padding: 22px; margin-bottom: 18px; }
.advanced-card { padding: 0; margin-bottom: 18px; overflow: hidden; }
.advanced-card summary {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 16px;
  padding: 18px 22px;
  cursor: pointer;
  list-style: none;
}
.advanced-card summary::-webkit-details-marker { display: none; }
.advanced-card summary::after {
  content: "灞曞紑";
  min-height: 34px;
  padding: 7px 12px;
  border-radius: 999px;
  background: #e6edf5;
  color: #25364c;
  font-weight: 800;
}
.advanced-card[open] summary::after { content: "鏀惰捣"; }
.advanced-card summary strong { display: block; font-size: 20px; }
.advanced-card summary small { display: block; margin-top: 4px; color: var(--muted); }
.advanced-shortcuts { display: flex; gap: 12px; flex-wrap: wrap; padding: 0 22px 22px; }
.shortcut-button { min-height: 40px; }
.page-section { display: none; }
.page-section.active { display: block; }
.inline-notice {
  margin-bottom: 14px;
  padding: 12px 14px;
  border-radius: 14px;
  border: 1px solid rgba(205, 220, 215, 0.95);
  background: rgba(245, 249, 247, 0.92);
  color: #506272;
  font-size: 13px;
  line-height: 1.6;
}
.tool-guide-grid {
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(180px, 1fr));
  gap: 12px;
  margin-bottom: 18px;
}
.tool-guide-card {
  min-height: 86px;
  justify-content: center;
  align-items: center;
  padding: 16px;
  border-radius: 18px;
  background: linear-gradient(145deg, rgba(255,255,255,.94), rgba(231,241,238,.88));
  border: 1px solid rgba(204, 220, 216, .95);
  box-shadow: 0 14px 34px rgba(22, 32, 51, 0.08);
  color: #16263c;
  font-size: 17px;
  font-weight: 900;
  text-align: center;
  letter-spacing: -0.02em;
}
.tool-guide-card:hover,
.tool-guide-card:focus-visible {
  background: linear-gradient(145deg, #e7f5f1, #d7eae4);
  color: var(--primary-strong);
  transform: translateY(-2px);
}
.tool-guide-dialog {
  width: min(760px, calc(100vw - 40px));
  max-height: calc(100vh - 56px);
}
.tool-guide-dialog .dialog-card {
  display: grid;
  gap: 18px;
  max-height: calc(100vh - 56px);
  overflow: auto;
  padding: 24px;
  border: 1px solid var(--line);
  border-radius: 22px;
  background: #ffffff;
  box-shadow: 0 28px 80px rgba(15, 23, 42, 0.24);
}
.tool-guide-dialog .dialog-header {
  display: flex;
  justify-content: space-between;
  align-items: start;
  gap: 18px;
  padding-bottom: 14px;
  border-bottom: 1px solid #e4edf5;
}
.tool-guide-dialog .dialog-header h3 {
  margin: 0;
  font-size: 28px;
  letter-spacing: -0.03em;
}
.dialog-kicker {
  display: block;
  margin-bottom: 6px;
  color: var(--primary);
  font-size: 12px;
  font-weight: 900;
  letter-spacing: .1em;
  text-transform: uppercase;
}
.markdown-guide {
  display: grid;
  gap: 18px;
  color: #24364d;
}
.markdown-guide h4 {
  margin: 0 0 8px;
  font-size: 18px;
}
.markdown-guide p {
  margin: 0;
  color: var(--muted);
  line-height: 1.75;
}
.markdown-guide ul,
.markdown-guide ol {
  margin: 0;
  padding-left: 22px;
  color: #334b62;
  line-height: 1.75;
}
.markdown-guide li + li { margin-top: 4px; }
.markdown-guide code {
  padding: 2px 6px;
  border-radius: 8px;
  background: #eef5f4;
  color: var(--primary-strong);
  font-family: "Cascadia Mono", "Consolas", monospace;
  font-size: 12px;
}
.section-header { margin: 0 0 16px; }
.section-header h2 { margin: 0; font-size: 26px; }
.section-header p { margin: 6px 0 0; color: var(--muted); }
.subnav { display: flex; gap: 10px; flex-wrap: wrap; margin-bottom: 16px; }
.subnav-link {
  min-height: 36px;
  padding: 0 14px;
  background: #f3f6fa;
  color: #617288;
  border-radius: 999px;
  font-size: 13px;
  font-weight: 700;
  border: 1px solid #dbe4ef;
}
.subnav-link.active {
  background: #ffffff;
  color: var(--primary-strong);
  border-color: #bfddd7;
  box-shadow: 0 8px 18px rgba(31, 111, 104, 0.08);
}
.link-with-dot,
.section-title-with-dot {
  display: inline-flex;
  align-items: center;
  gap: 8px;
}
.badge-dot {
  width: 9px;
  height: 9px;
  border-radius: 999px;
  background: #e14b4b;
  box-shadow: 0 0 0 3px rgba(225, 75, 75, 0.14);
}
.subtab-panel { display: none; }
.subtab-panel.active { display: block; }
.dashboard-grid { grid-template-columns: 1.2fr .8fr; align-items: start; }
.card-title { display: flex; align-items: end; justify-content: space-between; gap: 20px; margin-bottom: 18px; }
.card-title h2 { margin: 0; font-size: 22px; }
.card-title h3 { margin: 0; font-size: 20px; }
.card-title p { margin: 0; color: var(--muted); max-width: 620px; }
.card-title p:empty, .section-header p:empty { display: none; }
.grid.two { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 14px; }
.stage-grid { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 14px; }
.compact-card { margin-bottom: 0; }
.stage-panel { background: #f7fafc; border: 1px solid var(--line); border-radius: 18px; padding: 16px; display: grid; gap: 12px; }
.stage-panel h3 { margin: 0; font-size: 17px; }
label { display: grid; gap: 7px; color: var(--muted); font-weight: 700; }
input, select {
  width: 100%;
  min-height: 42px;
  border: 1px solid #cbd8e6;
  border-radius: 12px;
  padding: 9px 11px;
  background: #fff;
  color: var(--ink);
  font: inherit;
}
.search-inline-field {
  display: grid;
  grid-template-columns: minmax(0, 1fr) auto;
  gap: 10px;
  align-items: center;
}
.search-inline-field input {
  min-width: 0;
}
.secret-field { display: grid; grid-template-columns: minmax(0, 1fr) auto; gap: 8px; align-items: center; }
.secret-field input { min-width: 0; }
.field-label {
  display: block;
  margin-bottom: 7px;
  color: var(--muted);
  font-size: 13px;
  font-weight: 800;
}
.file-card-row {
  display: grid;
  grid-template-columns: minmax(0, 1fr) auto;
  gap: 10px;
  align-items: stretch;
}
.file-card {
  min-height: 42px;
  display: flex;
  align-items: center;
  min-width: 0;
  padding: 10px 13px;
  border: 1px solid var(--line);
  border-radius: 14px;
  background: #f8fbfd;
  color: var(--ink);
  font-weight: 700;
  line-height: 1.35;
  overflow-wrap: anywhere;
}
.file-card.empty {
  color: var(--muted);
  font-weight: 600;
}
.mini-button { min-height: 42px; padding: 0 14px; background: #e6edf5; color: #25364c; border: 1px solid #cbd8e6; }
.check { display: flex; align-items: center; gap: 9px; color: var(--ink); }
.check input[type="checkbox"],
.check-line input[type="checkbox"],
.mapping-source,
.directional-item input[type="checkbox"],
.pattern-table input[type="checkbox"],
.scope-item input[type="checkbox"],
#crossExcelHeaderList input[type="checkbox"] {
  width: 16px;
  height: 16px;
  min-width: 16px;
  min-height: 16px;
  margin: 0;
  accent-color: var(--primary);
}
.check-line {
  display: inline-flex;
  align-items: center;
  gap: 8px;
  min-height: 36px;
  padding: 8px 12px;
  border: 1px solid var(--line);
  border-radius: 999px;
  background: #f7fafc;
  color: var(--ink);
  font-weight: 700;
}
.actions { display: flex; align-items: center; gap: 12px; flex-wrap: wrap; margin-top: 16px; }
.action-row-compact {
  margin-top: 12px;
  align-items: center;
}
.action-row-compact .hint {
  min-height: 24px;
  display: inline-flex;
  align-items: center;
}
.sticky-actions {
  position: sticky;
  bottom: 18px;
  display: flex;
  justify-content: flex-end;
  align-items: center;
  gap: 12px;
  margin-top: 16px;
  padding: 14px 18px;
  border: 1px solid var(--line);
  border-radius: 18px;
  background: rgba(255,255,255,.92);
  box-shadow: var(--shadow);
}
button {
  min-height: 40px;
  border: 0;
  border-radius: 12px;
  padding: 0 16px;
  font: inherit;
  font-weight: 700;
  cursor: pointer;
  transition: transform .18s ease, box-shadow .18s ease, background .18s ease, color .18s ease;
}
button.primary {
  background: linear-gradient(180deg, #2d8a80 0%, #1f6f68 100%);
  color: white;
  box-shadow: 0 10px 20px rgba(31, 111, 104, 0.18);
}
button.secondary {
  background: #f5f8fb;
  color: #304257;
  border: 1px solid #d7e1ec;
}
button.danger { background: #fff1ef; color: var(--danger); border: 1px solid #f2d1cc; }
button:hover:not(:disabled) {
  transform: translateY(-1px);
}
button.primary:hover:not(:disabled) {
  box-shadow: 0 14px 24px rgba(31, 111, 104, 0.22);
}
button:disabled { opacity: .58; cursor: not-allowed; }
.error-panel {
  display: grid;
  gap: 4px;
  margin: 0 0 14px;
  padding: 13px 14px;
  border: 1px solid #f2b8b5;
  border-radius: 14px;
  background: #fff1ef;
  color: var(--danger);
}
.error-panel[hidden] { display: none; }
.pattern-table-wrap { overflow-x: auto; border: 1px solid var(--line); border-radius: 16px; }
.pattern-table { width: 100%; border-collapse: collapse; background: #fff; }
.pattern-table th, .pattern-table td { padding: 10px; border-bottom: 1px solid var(--line); text-align: left; vertical-align: middle; }
.pattern-table th { color: var(--muted); font-size: 13px; background: #f7fafc; }
.pattern-table tr:last-child td { border-bottom: 0; }
.pattern-table input[type="text"] { min-height: 36px; }
.pattern-table input[type="checkbox"] { width: 18px; min-height: 18px; }
.diff-preview-table-wrap {
  max-width: 100%;
  overscroll-behavior-inline: contain;
}
.diff-preview-table {
  width: 100%;
  min-width: 1240px;
  table-layout: fixed;
}
.diff-preview-file-column { width: 150px; }
.diff-preview-sheet-column { width: 130px; }
.diff-preview-location-column { width: 230px; }
.diff-preview-comparison-column { width: 620px; }
.diff-preview-table th,
.diff-preview-table td {
  vertical-align: top;
  white-space: pre-wrap;
  overflow-wrap: anywhere;
}
.diff-preview-table td:nth-child(-n + 4) {
  color: #50657a;
  line-height: 1.55;
}
.diff-preview-table td:last-child {
  padding: 12px;
  overflow: visible;
}
.review-input-table-wrap,
.review-result-table-wrap {
  max-width: 100%;
  overscroll-behavior-inline: contain;
}
.review-input-table {
  width: 100%;
  min-width: 1260px;
  table-layout: fixed;
}
.review-input-file-column { width: 170px; }
.review-input-location-column { width: 150px; }
.review-input-row-column { width: 78px; }
.review-input-text-column { width: 350px; }
.review-input-hint-column { width: 150px; }
.review-input-table th,
.review-input-table td,
.review-result-table th,
.review-result-table td {
  vertical-align: top;
  white-space: pre-wrap;
  overflow-wrap: anywhere;
  line-height: 1.55;
}
.review-input-table td:nth-child(-n + 3),
.review-input-table td:last-child {
  color: #50657a;
}
.review-result-table {
  width: 100%;
  min-width: 0;
  table-layout: fixed;
}
.review-result-table-wrap { overflow-x: hidden; }
.review-result-preview-head {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  margin-top: 16px;
  margin-bottom: 8px;
  color: var(--muted);
  font-size: 13px;
  font-weight: 700;
}
.review-result-table th,
.review-result-table td {
  min-width: 0;
  overflow-wrap: anywhere;
  word-break: break-word;
}
.review-result-table th:nth-child(3),
.review-result-table td:nth-child(3) {
  text-align: center;
}
.compact-rule-table th, .compact-rule-table td { padding: 8px; }
.compact-rule-table input[type="text"],
.compact-rule-table select {
  min-height: 34px;
  padding: 6px 9px;
  border-radius: 10px;
  font-size: 14px;
}
.compact-rule-table { table-layout: fixed; }
.builtin-editor-table th:nth-child(1),
.builtin-editor-table td:nth-child(1),
.pending-rule-table th:nth-child(1),
.pending-rule-table td:nth-child(1) { width: 54px; }
.builtin-editor-table th:nth-child(2),
.builtin-editor-table td:nth-child(2),
.pending-rule-table th:nth-child(2),
.pending-rule-table td:nth-child(2) { width: 120px; }
.builtin-editor-table th:nth-child(4),
.builtin-editor-table td:nth-child(4),
.pending-rule-table th:nth-child(4),
.pending-rule-table td:nth-child(4) { width: 92px; }
.pending-rule-table th:nth-child(5),
.pending-rule-table td:nth-child(5) { width: 360px; }
.modal-card.review-detail-modal {
  width: calc(100vw - 28px);
  height: calc(100vh - 28px);
  max-height: calc(100vh - 28px);
  display: grid;
  grid-template-rows: auto minmax(0, 1fr);
}
.review-detail-wrap {
  min-height: 0;
  max-height: none;
  overflow: auto;
  padding: 0 2px 2px;
}
.review-detail-table {
  width: 100%;
  min-width: 0;
  table-layout: fixed;
}
.review-detail-table th:nth-child(1),
.review-detail-table td:nth-child(1) {
  width: 20%;
}
.review-detail-table th:nth-child(2),
.review-detail-table td:nth-child(2) {
  width: 20%;
}
.review-detail-table th:nth-child(3),
.review-detail-table td:nth-child(3) {
  width: 20%;
}
.review-detail-table th:nth-child(4),
.review-detail-table td:nth-child(4) {
  width: 20%;
}
.review-detail-table th:nth-child(5),
.review-detail-table td:nth-child(5) {
  width: 20%;
}
.review-detail-table td {
  vertical-align: top;
  white-space: normal;
  overflow-wrap: anywhere;
  line-height: 1.6;
}
.review-detail-group + .review-detail-group {
  margin-top: 24px;
  padding-top: 20px;
  border-top: 1px solid var(--line);
}
.review-detail-group-head {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  margin: 0 0 9px;
}
.review-detail-group-head h4 {
  margin: 0;
  color: var(--text);
  font-size: 14px;
  line-height: 1.35;
  overflow-wrap: anywhere;
}
.review-detail-group-head span {
  flex: 0 0 auto;
  color: var(--muted);
  font-size: 12px;
}
.review-detail-group-table-wrap {
  overflow-x: hidden;
  border: 1px solid var(--line);
  border-radius: var(--radius-sm);
}
.review-detail-issue-cell {
  position: relative;
  padding-bottom: 48px;
}
.review-diff-cell .diff-inline-side {
  min-height: 0;
  padding: 10px;
  border-radius: 10px;
}
.review-diff-cell .diff-inline-label {
  margin-bottom: 6px;
}
.review-followup-inline-button {
  position: absolute;
  right: 12px;
  bottom: 10px;
  min-height: 32px;
  padding: 0 12px;
  border-radius: 999px;
  font-size: 13px;
}
.review-followup-modal {
  width: min(1180px, calc(100vw - 36px));
  height: min(820px, calc(100vh - 36px));
  display: grid;
  grid-template-rows: auto minmax(0, 1fr);
}
.review-followup-layout {
  min-height: 0;
  display: grid;
  grid-template-columns: minmax(260px, 360px) minmax(0, 1fr);
  gap: 16px;
}
.review-followup-context {
  min-height: 0;
  overflow: auto;
  padding: 14px;
  border: 1px solid var(--line);
  border-radius: 18px;
  background: #f7fafc;
}
.followup-context-card {
  display: grid;
  gap: 12px;
}
.followup-context-section {
  display: grid;
  gap: 5px;
}
.followup-context-section span {
  color: var(--muted);
  font-size: 12px;
  font-weight: 800;
}
.followup-context-section p {
  margin: 0;
  max-height: 138px;
  overflow: auto;
  padding-right: 4px;
  color: #24364d;
  line-height: 1.6;
  white-space: pre-wrap;
  overflow-wrap: anywhere;
}
.review-followup-chat {
  min-height: 0;
  display: grid;
  grid-template-rows: minmax(0, 1fr) auto auto;
  gap: 12px;
}
.review-followup-messages {
  min-height: 0;
  overflow: auto;
  display: grid;
  align-content: start;
  gap: 10px;
  padding: 14px;
  border: 1px solid var(--line);
  border-radius: 18px;
  background: #fbfdff;
}
.followup-message {
  width: min(78%, 720px);
  padding: 11px 13px;
  border-radius: 16px;
  line-height: 1.65;
  white-space: pre-wrap;
  overflow-wrap: anywhere;
}
.followup-message.user {
  justify-self: end;
  background: #e7f1ef;
  color: #123f3a;
}
.followup-message.assistant {
  justify-self: start;
  background: #f1f5f9;
  color: #24364d;
}
.followup-message-empty {
  align-self: center;
  justify-self: center;
  color: var(--muted);
  padding: 20px;
}
.review-followup-input {
  display: grid;
  grid-template-columns: minmax(0, 1fr) auto;
  gap: 10px;
  align-items: end;
}
.review-followup-input textarea {
  min-height: 82px;
  max-height: 180px;
  resize: vertical;
  border: 1px solid #cbd8e6;
  border-radius: 14px;
  padding: 11px 12px;
  font: inherit;
  line-height: 1.55;
}
@media (max-width: 860px) {
  .review-followup-layout {
    grid-template-columns: 1fr;
  }
  .review-followup-context {
    max-height: 260px;
  }
}
.builtin-preview-table { table-layout: fixed; }
.builtin-preview-table th:nth-child(1),
.builtin-preview-table td:nth-child(1) { width: 68px; }
.builtin-preview-table th:nth-child(2),
.builtin-preview-table td:nth-child(2) { width: 150px; }
.builtin-preview-table th:nth-child(3),
.builtin-preview-table td:nth-child(3) { width: 120px; }
.builtin-preview-table th:nth-child(4),
.builtin-preview-table td:nth-child(4) { width: 120px; }
.builtin-preview-table th:nth-child(6),
.builtin-preview-table td:nth-child(6) { width: 360px; }
.compact-rule-table .pattern-actions button {
  min-height: 32px;
  padding: 0 10px;
}
.empty-cell { color: var(--muted); text-align: center; padding: 22px !important; }
.pattern-actions { display: flex; gap: 6px; flex-wrap: wrap; }
.pattern-actions button { min-height: 34px; padding: 0 10px; }
.color-picker-block {
  display: grid;
  gap: 10px;
}
.preset-color-row {
  display: flex;
  flex-wrap: wrap;
  gap: 8px;
  align-items: center;
}
.preset-color-btn {
  width: 22px;
  height: 22px;
  padding: 0;
  border: 2px solid #111827;
  border-radius: 5px;
  background: #ffffff;
  box-shadow: 0 1px 2px rgba(17, 24, 39, 0.12);
  cursor: pointer;
  position: relative;
  transition: transform .15s ease, box-shadow .15s ease, border-color .15s ease;
}
.preset-color-btn::after {
  content: "";
  position: absolute;
  inset: 2px;
  border-radius: 3px;
  border: 1px solid rgba(255, 255, 255, 0.45);
  pointer-events: none;
}
.preset-color-btn.active {
  border-color: #111827;
  box-shadow: 0 0 0 2px rgba(17, 24, 39, 0.18), 0 1px 2px rgba(17, 24, 39, 0.14);
  transform: translateY(-1px);
}
.preset-color-btn:hover {
  transform: translateY(-1px);
}
.diff-mode-actions { display: inline-flex; align-items: center; gap: 8px; }
.diff-mode-select { width: auto; min-width: 136px; min-height: 40px; padding: 0 34px 0 12px; font-weight: 800; }
.icon-button { width: 40px; min-width: 40px; min-height: 40px; padding: 0; font-size: 18px; line-height: 1; }
.diff-field-settings-modal { width: min(760px, 100%); }
.diff-field-settings-grid { display: grid; gap: 12px; margin-bottom: 18px; }
.diff-header-presence { display: grid; grid-template-columns: 1fr; gap: 6px; }
.diff-header-presence span { display: flex; align-items: center; justify-content: space-between; gap: 12px; padding: 8px 10px; border: 1px solid var(--line); border-radius: 8px; background: #f8fbfd; color: var(--muted); font-size: 13px; }
.diff-header-presence strong { color: var(--ink); font-size: 14px; }
.diff-compare-fields-block { display: grid; gap: 8px; margin-bottom: 16px; }
.field-block-head { display: flex; align-items: baseline; justify-content: space-between; gap: 12px; }
.field-block-head span { color: var(--muted); font-size: 13px; }
.diff-compare-field-list { display: grid; grid-template-columns: 1fr; gap: 0; max-height: 320px; overflow: auto; border: 1px solid var(--line); border-radius: 8px; background: #fff; }
.diff-compare-field-item { display: flex; align-items: flex-start; gap: 10px; min-width: 0; height: auto; min-height: 40px; padding: 10px 12px; border: 0; border-bottom: 1px solid #e4ebf2; border-radius: 0; background: #fff; cursor: pointer; }
.diff-compare-field-item:last-child { border-bottom: 0; }
.diff-compare-field-item:hover { background: #f7fafc; }
.diff-compare-field-item input[type="checkbox"] { width: 16px; height: 16px; min-width: 16px; min-height: 16px; margin: 2px 0 0; accent-color: var(--primary); }
.diff-compare-field-item span { flex: 1; min-width: 0; color: var(--ink); line-height: 1.45; white-space: normal; overflow-wrap: anywhere; word-break: break-all; }
.diff-unmatched-line { display: flex; width: fit-content; min-height: 32px; margin: 0; padding: 4px 0; border: 0; border-radius: 0; background: transparent; }
.info-tip { position: relative; display: inline-grid; place-items: center; width: 17px; height: 17px; border: 1px solid #7d90a5; border-radius: 50%; color: #50657a; font-size: 12px; font-weight: 800; cursor: help; }
.info-tip-content { position: absolute; z-index: 10; left: 50%; bottom: calc(100% + 9px); width: min(330px, 70vw); padding: 10px 12px; border-radius: 10px; background: #203147; color: #fff; font-size: 13px; line-height: 1.55; font-weight: 500; box-shadow: 0 10px 28px rgba(15,23,42,.22); opacity: 0; pointer-events: none; transform: translate(-50%, 4px); transition: opacity .14s ease, transform .14s ease; }
.info-tip:hover .info-tip-content, .info-tip:focus .info-tip-content { opacity: 1; transform: translate(-50%, 0); }
@media (max-width: 680px) { .diff-header-presence { padding-bottom: 0; } .field-block-head { align-items: flex-start; flex-direction: column; gap: 3px; } .diff-mode-actions { width: 100%; } .diff-mode-select { flex: 1; } }
.diff-inline-card {
  display: grid;
  grid-template-columns: repeat(2, minmax(280px, 1fr));
  gap: 10px;
  min-width: 0;
}
.diff-inline-card.diff-inline-single { grid-template-columns: minmax(0, 1fr); }
.diff-inline-side {
  min-width: 0;
  border-radius: 12px;
  padding: 10px 12px;
  border: 1px solid var(--line);
  cursor: pointer;
  transition: transform .15s ease, box-shadow .15s ease, border-color .15s ease;
}
.diff-inline-side:hover {
  transform: translateY(-1px);
  box-shadow: 0 8px 18px rgba(17, 24, 39, 0.08);
}
.diff-inline-side:active {
  transform: translateY(0);
}
.diff-inline-delete {
  background: #fff4f2;
  border-color: #f3c7c0;
}
.diff-inline-add {
  background: #eefaf4;
  border-color: #bfe4cb;
}
.diff-inline-label {
  margin-bottom: 6px;
  font-size: 12px;
  font-weight: 800;
  letter-spacing: 0.04em;
}
.diff-inline-delete .diff-inline-label { color: #b4443f; }
.diff-inline-add .diff-inline-label { color: #1f6f68; }
.diff-inline-text {
  color: #24364d;
  line-height: 1.7;
  overflow-wrap: anywhere;
  word-break: normal;
  white-space: pre-wrap;
}
.diff-inline-empty {
  color: #9aa7b6;
}
.diff-plain-token {
  color: inherit;
}
.diff-token {
  padding: 1px 2px;
  border-radius: 4px;
  font-weight: 700;
}
.diff-token-delete {
  background: rgba(244, 67, 54, 0.16);
  color: #9f2e29;
}
.diff-token-add {
  background: rgba(76, 175, 80, 0.18);
  color: #236a2d;
}
.summary-line { display: flex; gap: 14px; flex-wrap: wrap; margin-bottom: 14px; color: var(--muted); font-weight: 700; }
.summary-line span {
  padding: 8px 12px;
  border: 1px solid #dbe4ef;
  border-radius: 999px;
  background: #f8fbfd;
  color: #607287;
  font-size: 13px;
}
.summary-line strong { color: var(--ink); }
.update-leaving-screen {
  min-height: 100vh;
  display: grid;
  place-items: center;
  padding: 32px;
}
.update-leaving-card {
  width: min(460px, 100%);
  padding: 28px;
  border-radius: 22px;
  border: 1px solid var(--line);
  background: rgba(255,255,255,.92);
  box-shadow: var(--shadow);
  text-align: center;
}
.update-leaving-card h2 { margin: 0 0 10px; font-size: 26px; }
.update-leaving-card p { margin: 0; color: var(--muted); line-height: 1.7; }
.regex-cell { font: 13px/1.45 "Cascadia Mono", "Consolas", monospace; overflow-wrap: anywhere; color: var(--ink); }
.examples-cell { min-width: 0; }
.examples-scroll {
  max-height: 88px;
  overflow-y: auto;
  padding-right: 4px;
  line-height: 1.45;
  white-space: normal;
  overflow-wrap: anywhere;
}
.examples-item + .examples-item { margin-top: 6px; }
.scope-toolbar { display: flex; gap: 10px; flex-wrap: wrap; margin-bottom: 12px; }
.scope-list { display: grid; gap: 12px; }
.scope-item { display: grid; grid-template-columns: auto minmax(0, 1fr) auto; gap: 10px 12px; align-items: start; padding: 14px; border: 1px solid var(--line); border-radius: 16px; background: #f7fafc; }
.scope-toggle { display: flex; align-items: center; justify-content: center; min-height: 42px; }
.scope-item input[type="checkbox"] { width: 18px; min-height: 18px; margin: 0; }
.scope-fields { display: grid; gap: 8px; min-width: 0; }
.scope-fields textarea { min-height: 46px; resize: vertical; border: 1px solid #cbd8e6; border-radius: 12px; padding: 9px 11px; font: inherit; }
.scope-delete { min-height: 42px; padding: 0 14px; background: #fde5e2; color: var(--danger); }
.scope-empty { padding: 18px; border: 1px dashed #cbd8e6; border-radius: 16px; background: #f7fafc; color: var(--muted); text-align: center; }
.prompt-grid { display: grid; grid-template-columns: repeat(2, minmax(0,1fr)); gap: 14px; }
.prompt-card { display: grid; gap: 9px; padding: 14px; border: 1px solid var(--line); border-radius: 16px; background: #f7fafc; }
.prompt-card header { display: flex; justify-content: space-between; gap: 12px; align-items: start; }
.prompt-card strong { color: var(--ink); }
.prompt-card small { color: var(--muted); line-height: 1.45; }
.prompt-card-actions {
  display: inline-flex;
  align-items: center;
  gap: 8px;
  flex-wrap: wrap;
  justify-content: flex-end;
}
.prompt-reset-button {
  min-height: 30px;
  padding: 0 10px;
  border-radius: 10px;
}
.prompt-card textarea {
  min-height: 220px;
  resize: vertical;
  border: 1px solid #cbd8e6;
  border-radius: 12px;
  padding: 11px;
  font: 13px/1.55 "Cascadia Mono", "Consolas", monospace;
  color: var(--ink);
}
.prompt-badge { white-space: nowrap; border-radius: 999px; padding: 4px 9px; background: #e6edf5; color: var(--muted); font-size: 12px; font-weight: 800; }
.prompt-badge.modified { background: #fff0cd; color: #8a5d00; }
.result-box, .log-box {
  margin: 16px 0 0;
  padding: 14px;
  border-radius: 14px;
  background: #111827;
  color: #d7e4f5;
  overflow: auto;
  font-size: 13px;
}
.cross-summary-box {
  display: grid;
  gap: 6px;
  align-content: center;
  min-height: 100%;
  padding: 14px;
  border: 1px solid var(--line);
  border-radius: 16px;
  background: #f7fafc;
}
.cross-summary-box span,
.cross-summary-box small { color: var(--muted); }
.cross-summary-box strong { font-size: 24px; }
.cross-header-list {
  display: grid;
  gap: 10px;
  max-height: 320px;
  overflow: auto;
  margin-top: 14px;
  padding: 4px;
}
.cross-header-item {
  display: flex;
  align-items: center;
  gap: 10px;
  padding: 11px 14px;
  border: 1px solid #dbe4ef;
  border-radius: 14px;
  background: linear-gradient(180deg, #fbfdff 0%, #f3f7fb 100%);
}
.cross-header-item input {
  width: 18px;
  min-height: 18px;
  margin: 0;
}
.cross-header-item span {
  min-width: 0;
  color: #223449;
  font-weight: 600;
  overflow-wrap: anywhere;
}
.cross-merge-actions {
  justify-content: space-between;
  gap: 16px;
}
.cross-search-results {
  display: grid;
  gap: 14px;
  max-height: 680px;
  overflow: auto;
  padding-right: 4px;
}
.cross-empty-state {
  padding: 18px;
  border: 1px dashed #cbd8e6;
  border-radius: 16px;
  background: #f7fafc;
  color: var(--muted);
  text-align: center;
}
.cross-result-card {
  border: 1px solid var(--line);
  border-radius: 18px;
  background: #fff;
  overflow: hidden;
}
.cross-result-meta {
  display: flex;
  gap: 10px;
  flex-wrap: wrap;
  padding: 14px 16px;
  border-bottom: 1px solid var(--line);
  background: #f7fafc;
}
.cross-result-meta span {
  padding: 7px 11px;
  border-radius: 999px;
  background: #ffffff;
  border: 1px solid var(--line);
  color: var(--muted);
  font-size: 13px;
  font-weight: 700;
}
.cross-row-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(160px, 1fr));
  gap: 10px;
  padding: 16px;
}
.cross-cell {
  min-height: 60px;
  padding: 12px 12px 10px;
  border: 1px solid var(--line);
  border-radius: 14px;
  background: #fff;
  text-align: left;
  transition: border-color .18s ease, box-shadow .18s ease, background .18s ease;
}
.cross-cell:hover {
  border-color: #b8cbe0;
  box-shadow: 0 8px 18px rgba(22, 32, 51, 0.08);
}
.cross-cell.matched {
  background: #fff8e8;
  border-color: #efd39a;
}
.cross-cell-index {
  display: block;
  margin-bottom: 7px;
  color: var(--muted);
  font-size: 12px;
  font-weight: 800;
}
.cross-cell-text {
  display: block;
  width: 100%;
  white-space: pre-wrap;
  overflow-wrap: anywhere;
  word-break: break-word;
  line-height: 1.55;
  color: var(--ink);
}
.metrics { display: grid; grid-template-columns: repeat(6, minmax(0,1fr)); gap: 12px; margin-bottom: 14px; }
.metrics div { padding: 14px; background: #f7fafc; border: 1px solid var(--line); border-radius: 16px; }
.metrics span { display:block; color: var(--muted); font-size: 13px; }
.metrics strong { font-size: 22px; }
.stats-panel { margin: 0 0 14px; padding: 14px; border: 1px solid var(--line); border-radius: 18px; background: rgba(247,250,252,.72); }
.stats-title { margin-bottom: 10px; color: var(--muted); font-weight: 800; }
.compact-metrics { grid-template-columns: repeat(4, minmax(0,1fr)); margin-bottom: 0; }
.compact-metrics div { background: #fff; }
.compact-metrics strong { font-size: 18px; }
.result-file {
  display: grid;
  gap: 5px;
  padding: 14px;
  margin-bottom: 14px;
  border: 1px solid var(--line);
  border-radius: 16px;
  background: #f7fafc;
}
.result-file span, .result-file small { color: var(--muted); }
.result-file strong { overflow-wrap: anywhere; }
.result-metrics { grid-template-columns: repeat(4, minmax(0,1fr)); }
.hero-metrics { margin-bottom: 10px; }
.notice-list { display: grid; gap: 10px; color: var(--ink); }
.notice-list div { padding: 12px 14px; border: 1px solid var(--line); border-radius: 14px; background: #f7fafc; }
.dialog {
  width: min(760px, calc(100vw - 36px));
  max-height: calc(100vh - 48px);
  border: 0;
  border-radius: 18px;
  padding: 0;
  background: transparent;
}
.dialog::backdrop { background: rgba(15, 23, 42, 0.38); }
.wide-dialog { width: min(1120px, calc(100vw - 36px)); }
.compact-dialog { width: min(460px, calc(100vw - 36px)); }
.dialog-body {
  display: grid;
  gap: 16px;
  max-height: calc(100vh - 48px);
  overflow: auto;
  padding: 22px;
  border: 1px solid var(--line);
  border-radius: 18px;
  background: #ffffff;
  box-shadow: var(--shadow);
}
.dialog-head {
  display: flex;
  justify-content: space-between;
  align-items: start;
  gap: 16px;
}
.dialog-head h2 { margin: 0 0 6px; font-size: 22px; }
.dialog-head p { margin: 0; color: var(--muted); line-height: 1.55; }
.dialog-actions {
  display: flex;
  gap: 10px;
  justify-content: flex-end;
  flex-wrap: wrap;
  padding-top: 4px;
}
.icon-button {
  width: 36px;
  min-height: 36px;
  padding: 0;
  border-radius: 999px;
  font-size: 20px;
  line-height: 1;
}
.field { display: grid; gap: 7px; min-width: 0; }
.field label { margin: 0; }
.field textarea {
  width: 100%;
  min-height: 140px;
  resize: vertical;
  border: 1px solid #cbd8e6;
  border-radius: 12px;
  padding: 11px 12px;
  font: inherit;
}
.directional-items {
  display: grid;
  gap: 10px;
  max-height: 360px;
  overflow: auto;
}
.directional-item {
  display: grid;
  grid-template-columns: auto minmax(0, 1fr);
  gap: 10px;
  align-items: center;
  padding: 10px 12px;
  border: 1px solid var(--line);
  border-radius: 12px;
  background: #f7fafc;
}
.directional-item input[type="checkbox"] { margin: 0; }
.sheet-tabs {
  display: flex;
  gap: 8px;
  overflow-x: auto;
  padding-bottom: 2px;
}
.sheet-tab {
  min-height: 36px;
  padding: 0 13px;
  border: 1px solid var(--line);
  border-radius: 999px;
  background: #f7fafc;
  color: var(--muted);
  white-space: nowrap;
}
.sheet-tab.active {
  border-color: rgba(31,111,104,.35);
  background: #e8f2ef;
  color: var(--primary-strong);
}
.mapping-template-bar {
  display: grid;
  grid-template-columns: auto minmax(180px, 1fr) auto auto;
  align-items: center;
  gap: 10px;
  padding: 10px 12px;
  border: 1px solid var(--line);
  border-radius: 12px;
  background: #f8fbfd;
}
.mapping-template-bar label { color: var(--muted); font-weight: 700; white-space: nowrap; }
.mapping-template-bar select { min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.mapping-template-bar button { min-height: 36px; white-space: nowrap; }
.mapping-template-delete { width: 36px; border-radius: 10px; font-size: 17px; }
.mapping-template-hint { min-height: 0; margin: -7px 0 0; color: var(--danger); font-size: 13px; line-height: 1.5; }
.mapping-template-hint:empty { display: none; }
.excel-mapping-columns {
  display: grid;
  gap: 10px;
  max-height: min(520px, 52vh);
  overflow: auto;
  padding-right: 4px;
}
.mapping-row {
  display: grid;
  grid-template-columns: 58px minmax(160px, 1fr) 86px minmax(150px, .8fr) minmax(150px, .8fr) minmax(130px, .7fr);
  gap: 10px;
  align-items: center;
  padding: 10px;
  border: 1px solid var(--line);
  border-radius: 12px;
  background: #f8fbfd;
}
.mapping-col-id {
  color: var(--primary-strong);
  font-weight: 800;
}
.mapping-header {
  min-width: 0;
  overflow-wrap: anywhere;
  color: var(--ink);
  font-weight: 700;
}
.mapping-row select,
.mapping-row input[type="text"] {
  width: 100%;
  min-height: 36px;
  padding: 6px 9px;
  border-radius: 10px;
  font-size: 14px;
}
.check-line.disabled-line { opacity: .55; }
.review-log-list {
  display: grid;
  gap: 8px;
  max-height: 260px;
  overflow: auto;
  margin: 0;
  padding-left: 20px;
}
.review-log-list li {
  white-space: pre-wrap;
  overflow-wrap: anywhere;
  line-height: 1.5;
  color: var(--muted);
}
.review-log-list li.error { color: var(--danger); }
.review-log-list li.debug { color: #496579; }
.progress { height: 12px; border-radius: 999px; background: #dfe8f2; overflow: hidden; margin-bottom: 14px; }
.progress span { display:block; height: 100%; width: 0; background: linear-gradient(90deg, var(--primary), #e7b95e); transition: width .25s ease; }
.hint { color: var(--primary-strong); font-weight: 700; }
.hero-side {
  display: grid;
  gap: 12px;
  justify-items: end;
}
.notice-button {
  display: inline-flex;
  align-items: center;
  gap: 10px;
  min-height: 42px;
  padding: 0 16px;
  border: 1px solid #f5d2a4;
  border-radius: 999px;
  background: #fff6e7;
  color: #8a5d00;
  box-shadow: 0 10px 26px rgba(138, 93, 0, 0.12);
}
.update-notice-button {
  border-color: #c7d6ea;
  background: #eef5ff;
  color: #194b7a;
  box-shadow: 0 10px 26px rgba(25, 75, 122, 0.10);
}
.notice-button[hidden] { display: none !important; }
.notice-dot {
  width: 10px;
  height: 10px;
  border-radius: 999px;
  background: #e14b4b;
}
.modal-overlay {
  position: fixed;
  inset: 0;
  display: grid;
  place-items: center;
  padding: 24px;
  background: rgba(15, 23, 42, 0.36);
  z-index: 1000;
}
.modal-overlay[hidden] { display: none; }
.modal-card {
  width: min(1120px, 100%);
  max-height: calc(100vh - 48px);
  overflow: auto;
  padding: 22px;
  border-radius: 22px;
  background: #fff;
  box-shadow: 0 30px 80px rgba(15, 23, 42, 0.24);
}
.modal-header {
  display: flex;
  align-items: start;
  justify-content: space-between;
  gap: 16px;
  margin-bottom: 18px;
}
.modal-header h3 {
  margin: 0;
  font-size: 24px;
}
.modal-header p {
  margin: 8px 0 0;
  color: var(--muted);
}
.modal-close {
  min-height: 38px;
  min-width: 38px;
  padding: 0;
  border-radius: 999px;
  background: #edf3f8;
  color: #25364c;
}
.modal-actions {
  display: flex;
  gap: 10px;
  flex-wrap: wrap;
  margin-bottom: 14px;
}
.modal-footer {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  flex-wrap: wrap;
  margin-top: 16px;
}
.modal-footer-actions {
  display: inline-flex;
  align-items: center;
  gap: 10px;
}
.update-release-notes {
  min-height: 180px;
  max-height: 360px;
  overflow-y: auto;
  padding: 14px;
  border: 1px solid var(--line);
  border-radius: 14px;
  background: #f7fafc;
  color: var(--ink);
  line-height: 1.6;
  white-space: pre-wrap;
  overflow-wrap: anywhere;
}
.update-modal .summary-line {
  margin-bottom: 16px;
}
.feedback-modal {
  width: min(680px, calc(100vw - 36px));
}
.feedback-attachment-card,
.feedback-log-card {
  display: grid;
  gap: 10px;
  padding: 14px 16px;
  border: 1px solid var(--line);
  border-radius: 16px;
  background: #f9fbfd;
}
.feedback-attachment-head {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
}
.feedback-dropzone {
  display: grid;
  gap: 6px;
  width: 100%;
  padding: 18px 16px;
  border: 1px dashed rgba(31, 111, 104, 0.28);
  border-radius: 14px;
  background: linear-gradient(180deg, #fbfefd 0%, #f2f8f7 100%);
  color: var(--ink);
  text-align: left;
  cursor: pointer;
  transition: border-color .18s ease, background .18s ease, transform .18s ease;
}
.feedback-dropzone:hover,
.feedback-dropzone:focus-visible {
  border-color: rgba(31, 111, 104, 0.52);
  background: linear-gradient(180deg, #f8fdfc 0%, #eaf6f3 100%);
  transform: translateY(-1px);
  outline: none;
}
.feedback-dropzone.drag-over {
  border-color: rgba(31, 111, 104, 0.72);
  background: linear-gradient(180deg, #ecfaf6 0%, #dff4ee 100%);
}
.feedback-dropzone-title {
  font-weight: 800;
  color: var(--primary-strong);
}
.feedback-dropzone-subtitle {
  color: var(--muted);
  font-size: 13px;
}
.feedback-log-card code {
  display: block;
  padding: 10px 12px;
  border-radius: 12px;
  background: #eef3f8;
  color: #43546a;
  font-size: 12px;
  white-space: pre-wrap;
  word-break: break-all;
}
.feedback-log-actions {
  justify-content: flex-start;
}
.review-conversation-shell {
  display: grid;
  grid-template-columns: 236px minmax(0, 1fr);
  height: clamp(620px, calc(100vh - 190px), 820px);
  min-height: 620px;
  border: 1px solid var(--line);
  border-radius: 18px;
  overflow: hidden;
  background: var(--panel);
}
.review-session-sidebar { min-height: 0; padding: 14px; border-right: 1px solid var(--line); background: rgba(20, 29, 48, 0.025); overflow: hidden; }
.review-new-chat { width: 100%; }
.review-conversation-list { display: grid; gap: 7px; margin-top: 14px; max-height: calc(100% - 56px); overflow-y: auto; }
.review-session-row { display: grid; grid-template-columns: minmax(0, 1fr) 32px; align-items: center; gap: 3px; padding: 3px; border-radius: 12px; transition: background .16s ease, box-shadow .16s ease; }
.review-session-row:hover, .review-session-row.active { background: rgba(48, 111, 214, 0.09); }
.review-session-row.active { box-shadow: inset 3px 0 0 var(--primary); }
.review-session-item { min-width: 0; width: 100%; text-align: left; padding: 8px 9px; border: 0; border-radius: 9px; background: transparent; color: var(--text); }
.review-session-item:hover { background: transparent; }
.review-session-item strong, .review-session-item span { display: block; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.review-session-item span { margin-top: 3px; color: var(--muted); font-size: 12px; }
.review-session-rename { width: 30px; height: 30px; min-height: 30px; padding: 0; border: 0; border-radius: 8px; background: transparent; color: #718096; opacity: .72; }
.review-session-row:hover .review-session-rename, .review-session-row.active .review-session-rename, .review-session-rename:focus-visible { opacity: 1; }
.review-session-rename:hover { color: var(--primary); background: rgba(31, 111, 104, .10); }
.review-session-inline-input { min-width: 0; width: 100%; height: 34px; padding: 5px 8px; border-color: rgba(31, 111, 104, .42); font-weight: 700; }
.review-conversation-main { display: grid; grid-template-rows: auto minmax(0, 1fr) auto; min-width: 0; min-height: 0; overflow: hidden; }
.review-conversation-head { display: flex; align-items: center; justify-content: space-between; gap: 16px; padding: 17px 20px; border-bottom: 1px solid var(--line); }
.review-conversation-head h3, .review-conversation-head p { margin: 0; }
.review-conversation-head p { margin-top: 4px; color: var(--muted); font-size: 13px; }
.review-conversation-title-row { display: flex; align-items: center; gap: 7px; }
.review-title-input { width: min(360px, 52vw); padding: 5px 8px; font-size: 17px; font-weight: 750; }
.review-title-edit { width: 28px; height: 28px; padding: 0; border: 0; border-radius: 50%; background: transparent; color: var(--muted); }
.review-title-edit:hover { color: var(--primary); background: rgba(48, 111, 214, 0.08); }
.review-conversation-messages { min-height: 0; padding: 24px max(24px, 8%); overflow-y: auto; overscroll-behavior: contain; display: flex; flex-direction: column; gap: 18px; scrollbar-gutter: stable; }
.review-chat-message { max-width: min(820px, 88%); white-space: pre-wrap; line-height: 1.65; }
.review-chat-message.user { align-self: flex-end; padding: 11px 15px; border-radius: 16px 16px 4px 16px; background: rgba(48, 111, 214, 0.12); }
.review-chat-message.assistant { align-self: flex-start; }
.review-chat-message.workspace-report { width: min(900px, 96%); padding: 16px 18px; border: 1px solid rgba(31, 111, 104, 0.18); border-radius: 16px; background: linear-gradient(180deg, #fbfefd 0%, #f5faf9 100%); box-shadow: 0 6px 20px rgba(23, 55, 70, 0.04); font-size: 13px; }
.review-workspace-preview { display: grid; gap: 12px; white-space: normal; }
.review-workspace-preview-head, .review-workspace-file-head, .review-workspace-mapping-head { display: flex; align-items: center; justify-content: space-between; gap: 12px; }
.review-workspace-preview-head > div, .review-workspace-file-head > div { min-width: 0; }
.review-workspace-preview-head strong, .review-workspace-file-head strong { display: block; color: var(--text); font-size: 14px; line-height: 1.35; }
.review-workspace-preview-head span, .review-workspace-file-head span, .review-workspace-mapping-head > span:last-child { color: var(--muted); font-size: 11px; }
.review-workspace-total, .review-workspace-language { flex: 0 0 auto; padding: 3px 8px; border: 1px solid rgba(31, 111, 104, .24); border-radius: 999px; background: rgba(31, 111, 104, .07); color: #17685e !important; font-size: 11px !important; font-weight: 750; }
.review-workspace-file-card { display: grid; gap: 10px; padding: 13px; border: 1px solid rgba(31, 111, 104, .18); border-radius: 11px; background: rgba(255,255,255,.52); }
.review-workspace-file-head { justify-content: flex-start; }
.review-workspace-file-icon { display: grid; flex: 0 0 auto; place-items: center; width: 31px; height: 31px; border: 1px solid rgba(48, 111, 214, .22); border-radius: 7px; background: rgba(48, 111, 214, .08); color: #2768ad; font-size: 9px; font-weight: 850; letter-spacing: .25px; }
.review-workspace-mapping { display: grid; gap: 9px; padding: 11px; border-top: 1px solid rgba(31, 111, 104, .13); }
.review-workspace-mapping-head { min-height: 21px; }
.review-workspace-flow { display: grid; grid-template-columns: minmax(0, 1fr) 22px minmax(0, 1fr); align-items: stretch; gap: 7px; }
.review-workspace-endpoint { min-width: 0; padding: 8px 9px; border: 1px solid #dbe7e3; border-radius: 8px; background: #f8fbfa; }
.review-workspace-endpoint:last-child { border-color: #d7e2f0; background: #f7faff; }
.review-workspace-endpoint.is-empty { opacity: .72; }
.review-workspace-endpoint span { display: block; margin-bottom: 3px; color: var(--muted); font-size: 10px; font-weight: 750; }
.review-workspace-endpoint strong { display: -webkit-box; overflow: hidden; color: var(--text); font-size: 11px; line-height: 1.42; overflow-wrap: anywhere; -webkit-box-orient: vertical; -webkit-line-clamp: 2; }
.review-workspace-arrow { display: grid; place-items: center; color: #6e91c8; font-size: 16px; }
.review-workspace-reference { padding-left: 1px; color: #5d7186; font-size: 11px; }
.review-workspace-samples { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 7px; }
.review-workspace-sample { min-width: 0; padding: 8px; border: 1px solid #e0e8ed; border-radius: 8px; background: rgba(249,251,252,.82); }
.review-workspace-sample > span { display: block; overflow: hidden; margin-bottom: 6px; color: #78879a; font-size: 10px; font-variant-numeric: tabular-nums; text-overflow: ellipsis; white-space: nowrap; }
.review-workspace-sample-values { display: grid; gap: 6px; }
.review-workspace-sample-source, .review-workspace-sample-target { display: -webkit-box; overflow: hidden; padding-left: 7px; border-left: 2px solid #adc8bc; color: #42556b; font-size: 11px; line-height: 1.45; overflow-wrap: anywhere; -webkit-box-orient: vertical; -webkit-line-clamp: 3; }
.review-workspace-sample-target { border-left-color: #9ab8e5; color: #284c78; }
.review-workspace-empty { padding: 9px; color: var(--muted); font-size: 12px; }
.review-workspace-notes { border-top: 1px solid rgba(31, 111, 104, .14); color: var(--muted); }
.review-workspace-notes summary { padding-top: 9px; cursor: pointer; font-size: 11px; font-weight: 750; }
.review-workspace-note-copy { display: grid; gap: 5px; padding-top: 7px; }
.review-workspace-note-copy p { margin: 0; color: var(--muted); font-size: 11px; line-height: 1.55; }
.review-chat-message.error { color: var(--danger); }
.review-chat-message-meta { color: var(--muted); font-size: 11px; margin-top: 4px; }
.review-decision-actions { display: grid; gap: 9px; margin-top: 12px; }
.review-decision-confirm-row { display: flex; }
.review-decision-other-row { display: flex; align-items: center; gap: 8px; }
.review-decision-other-row input { width: min(420px, 100%); min-width: 220px; }
.review-decision-other-row button { flex: 0 0 auto; }
.review-composer { position: relative; margin: 0 max(24px, 8%) 22px; padding: 12px 14px; border: 1px solid var(--line-strong); border-radius: 18px; background: var(--panel); box-shadow: 0 10px 34px rgba(20, 30, 50, 0.08); }
.review-composer.dragging { border-color: var(--primary); background: rgba(48, 111, 214, 0.04); }
.review-composer textarea { width: 100%; border: 0; resize: vertical; min-height: 64px; box-shadow: none; background: transparent; }
.review-composer textarea:focus { outline: 0; }
.review-composer-footer { display: flex; justify-content: space-between; gap: 12px; align-items: flex-end; }
.review-composer-tools, .review-attachment-chips, .review-target-tabs { display: flex; flex-wrap: wrap; gap: 7px; }
.composer-chip, .review-attachment-chip, .review-target-tab { padding: 7px 10px; border-radius: 999px; border: 1px solid var(--line); background: var(--panel); color: var(--text); font-size: 12px; }
.composer-chip:hover { border-color: rgba(31, 111, 104, 0.42); background: #f7fbfa; }
.review-attachment-chip { display: inline-flex; align-items: center; gap: 6px; padding: 0; overflow: hidden; background: rgba(30, 148, 96, 0.08); }
.review-attachment-open { padding: 7px 3px 7px 10px; border: 0; background: transparent; color: inherit; font-size: inherit; }
.review-attachment-remove { width: 26px; height: 26px; padding: 0; margin-right: 3px; border: 0; border-radius: 50%; background: transparent; color: var(--muted); }
.review-attachment-remove:hover { color: var(--danger); background: rgba(190, 44, 44, 0.09); }
.review-message-files { display: flex; flex-wrap: wrap; gap: 7px; margin-top: 9px; }
.review-message-file { display: inline-flex; align-items: center; gap: 7px; padding: 8px 11px; border: 1px solid rgba(48, 111, 214, 0.2); border-radius: 10px; background: rgba(255,255,255,0.75); color: #245aa5; cursor: default; }
.review-stream-output { max-height: 190px; margin-top: 8px; padding: 10px 12px; overflow: auto; border-radius: 10px; background: #f4f7fa; color: #536174; font: 12px/1.55 ui-monospace, SFMono-Regular, Consolas, monospace; white-space: pre-wrap; word-break: break-word; }
.review-thinking-details { margin-top: 10px; border: 1px solid #dfe6ec; border-radius: 11px; background: #f7f9fb; overflow: hidden; }
.review-thinking-details summary { padding: 9px 11px; color: #526175; font-size: 12px; font-weight: 750; cursor: pointer; user-select: none; }
.review-thinking-details[open] summary { border-bottom: 1px solid #e4e9ee; }
.review-thinking-details .review-stream-output { max-height: 260px; margin: 0; border-radius: 0; background: #f7f9fb; }
dialog.modal { width: auto; max-width: none; border: 0; padding: 0; background: transparent; overflow: visible; }
dialog.modal::backdrop { background: rgba(20, 31, 48, .38); backdrop-filter: blur(2px); }
.review-attachment-mapping-card { width: min(620px, calc(100vw - 30px)); padding: 0; overflow: hidden; border: 1px solid rgba(207, 218, 229, .9); border-radius: 22px; }
.review-mapping-file-head { display: grid; grid-template-columns: 48px minmax(0, 1fr) 36px; align-items: center; gap: 13px; padding: 22px 24px 13px; }
.review-mapping-file-icon { display: grid; place-items: center; width: 48px; height: 48px; border-radius: 14px; background: linear-gradient(145deg, #e4f4ee, #f3faf7); color: #167052; font-size: 10px; font-weight: 850; letter-spacing: .3px; box-shadow: inset 0 0 0 1px rgba(31, 111, 104, .12); }
.review-mapping-file-copy { min-width: 0; }
.review-mapping-file-copy > span { display: block; margin-bottom: 3px; color: var(--muted); font-size: 12px; font-weight: 700; }
.review-mapping-file-copy h2 { margin: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; color: #172235; font-size: 20px; line-height: 1.3; }
.review-mapping-close { width: 34px; height: 34px; min-height: 34px; padding: 0; border: 0; border-radius: 50%; background: #f2f5f7; color: #536174; font-size: 20px; }
.review-mapping-close:hover { background: #e8edf1; color: #172235; }
.review-mapping-lead { margin: 0; padding: 0 24px 17px; color: #66758a; font-size: 13px; line-height: 1.55; }
.review-import-mode-grid { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 10px; padding: 0 24px; }
.review-import-mode-option { position: relative; display: grid; grid-template-columns: 38px minmax(0, 1fr) 20px; align-items: center; gap: 10px; min-height: 92px; padding: 14px; border: 1px solid #dbe3eb; border-radius: 15px; background: #fff; cursor: pointer; transition: border-color .16s ease, background .16s ease, box-shadow .16s ease, transform .16s ease; }
.review-import-mode-option:hover { border-color: #a9c7bd; background: #fbfefd; transform: translateY(-1px); }
.review-import-mode-option.selected { border-color: rgba(31, 111, 104, .62); background: #f2f9f6; box-shadow: 0 0 0 3px rgba(31, 111, 104, .08); }
.review-import-mode-option.disabled { cursor: not-allowed; opacity: .48; transform: none; }
.review-import-mode-option input { position: absolute !important; width: 1px !important; height: 1px !important; opacity: 0; pointer-events: none; }
.review-import-mode-option > span:not(.review-import-mode-icon) { display: grid; gap: 5px; min-width: 0; }
.review-import-mode-option strong { color: #243247; font-size: 14px; }
.review-import-mode-option strong em { display: inline-block; margin-left: 5px; padding: 2px 6px; border-radius: 999px; background: #dcefe8; color: #1f6f68; font-size: 10px; font-style: normal; vertical-align: 1px; }
.review-import-mode-option small { color: #718096; font-size: 12px; line-height: 1.45; }
.review-import-mode-option > i { display: grid; place-items: center; width: 19px; height: 19px; border-radius: 50%; background: var(--primary); color: #fff; font-size: 11px; font-style: normal; opacity: 0; transform: scale(.78); transition: opacity .16s ease, transform .16s ease; }
.review-import-mode-option.selected > i { opacity: 1; transform: scale(1); }
.review-import-mode-icon { display: grid; place-items: center; width: 38px; height: 38px; border-radius: 11px; background: #eef3f7; color: #52647a; font-size: 19px; }
.review-import-mode-option.selected .review-import-mode-icon { background: #dcefe8; color: #1f6f68; }
.review-import-mode-ai { font-size: 18px; }
.review-mapping-preset-panel { display: grid; gap: 9px; margin: 14px 24px 0; padding: 14px; border: 1px solid #dde6ed; border-radius: 14px; background: #f8fafb; }
.review-mapping-field-label { display: flex; align-items: center; justify-content: space-between; gap: 10px; }
.review-mapping-field-label label { margin: 0; color: #34445a; font-size: 13px; font-weight: 750; }
.review-mapping-field-label span { color: #8b97a7; font-size: 11px; }
.review-mapping-preset-row { display: grid; grid-template-columns: minmax(0, 1fr) auto; gap: 9px; }
.review-mapping-preset-row select { min-width: 0; }
.review-mapping-preset-row button { min-height: 40px; white-space: nowrap; }
.review-mapping-preset-panel > small { color: #7a8797; line-height: 1.5; }
.review-attachment-mapping-card > .hint { min-height: 0; margin: 10px 24px 0; color: var(--danger); font-size: 12px; }
.review-attachment-mapping-card > .hint:empty { display: none; }
.review-attachment-mapping-card > .modal-footer { margin-top: 18px; padding: 15px 24px 20px; border-top: 1px solid #e7edf1; background: #fbfcfd; }
.review-target-tab.active { color: white; background: var(--primary); border-color: var(--primary); }
.review-send-button { min-width: 68px; border-radius: 999px; }
.review-conversation-results { margin-top: 18px; }
.review-inline-progress { display: grid; grid-template-columns: auto auto minmax(120px, 1fr); align-items: center; gap: 12px; margin-bottom: 14px; }
.review-language-popover { position: absolute; z-index: 30; bottom: 56px; left: 14px; width: min(650px, calc(100% - 28px)); overflow: hidden; border: 1px solid #d9e1ea; border-radius: 14px; background: #fff; box-shadow: 0 18px 50px rgba(29, 43, 68, 0.18); }
.review-language-popover-head { display: grid; grid-template-columns: 32px 1fr auto; align-items: center; gap: 8px; padding: 11px 14px 8px; }
.review-language-popover-head strong { font-size: 14px; }
.review-language-popover-head span { color: var(--muted); font-size: 12px; }
.review-language-back { width: 30px; height: 30px; padding: 0; border: 0; border-radius: 50%; background: transparent; color: #56647a; font-size: 25px; line-height: 1; }
.review-language-back:hover { background: #f0f3f7; }
.review-language-search-wrap { display: flex; align-items: center; gap: 8px; padding: 0 15px 10px; border-bottom: 1px solid #e4e9ef; color: #6c788a; }
.review-language-search-wrap input { width: 100%; padding: 7px 0; border: 0; border-radius: 0; box-shadow: none; background: transparent; }
.review-language-search-wrap input:focus { outline: 0; }
.review-language-options { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 2px 14px; max-height: 330px; overflow: auto; padding: 9px; }
.review-language-option { position: relative; display: flex; align-items: center; gap: 8px; min-height: 36px; padding: 7px 10px; border: 0; border-radius: 4px; color: #26354a; cursor: pointer; user-select: none; }
.review-language-option:hover { background: #f3f6fa; }
.review-language-option.selected { color: #1769d2; background: #e8f0fe; }
.review-language-option input { position: absolute; opacity: 0; pointer-events: none; }
.review-language-check { width: 14px; color: #1769d2; font-weight: 800; visibility: hidden; }
.review-language-option.selected .review-language-check { visibility: visible; }
.review-term-base-chip { max-width: 220px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.review-term-base-popover { position: absolute; z-index: 62; bottom: 56px; left: 14px; width: min(410px, calc(100% - 28px)); max-height: 390px; overflow: hidden auto; padding: 7px; border: 1px solid #d9e1ea; border-radius: 14px; background: #fff; box-shadow: 0 18px 50px rgba(29,43,68,.18); }
.review-term-base-action, .review-term-base-option { display: grid; grid-template-columns: 30px minmax(0,1fr) auto; align-items: center; gap: 9px; width: 100%; min-height: 48px; padding: 8px 10px; border: 0; border-radius: 8px; background: transparent; color: #26354a; text-align: left; }
.review-term-base-action:hover, .review-term-base-option:hover { background: #f3f6fa; }
.review-term-base-action-icon { display: grid; place-items: center; width: 28px; height: 28px; border-radius: 7px; background: #e8f0fe; color: #1769d2; font-weight: 800; }
.review-term-base-action > span:nth-child(2), .review-term-base-option-copy { display: grid; min-width: 0; gap: 2px; }
.review-term-base-action strong, .review-term-base-option strong { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; font-size: 12px; }
.review-term-base-action small, .review-term-base-option small { overflow: hidden; color: #6c788a; text-overflow: ellipsis; white-space: nowrap; font-size: 10px; }
.review-term-base-options { margin-top: 5px; padding-top: 5px; border-top: 1px solid #e4e9ef; }
.review-term-base-option.selected { background: #e8f0fe; color: #1769d2; }
.review-term-base-option-check { color: #1f6f68; font-weight: 850; }
.review-term-base-delete { width: 28px; height: 28px; min-height: 28px; padding: 0; border: 0; border-radius: 6px; background: transparent; color: #8b97a7; }
.review-term-base-delete:hover { background: rgba(190,44,44,.09); color: var(--danger); }
.danger-text { color: var(--danger); }
.compact-actions { align-items: center; }
@media (max-width: 980px) {
  .shell { grid-template-columns: 1fr; }
  .sidebar { position: static; }
  .hero, .grid.two, .stage-grid, .metrics, .prompt-grid, .dashboard-grid, .scope-list, .compact-metrics, .result-metrics, .cross-row-grid { grid-template-columns: 1fr; }
  .sticky-actions { position: static; }
  .content { padding: 18px; }
  .hero-side { justify-items: stretch; }
  .modal-card { padding: 18px; }
  .modal-footer { align-items: stretch; }
  .mapping-template-bar { grid-template-columns: 1fr auto; }
  .mapping-template-bar label { grid-column: 1 / -1; }
  .mapping-template-bar .secondary { grid-column: 1 / 2; }
  .review-conversation-shell { grid-template-columns: 1fr; }
  .review-conversation-shell { height: 760px; min-height: 0; grid-template-rows: auto minmax(0, 1fr); }
  .review-session-sidebar { border-right: 0; border-bottom: 1px solid var(--line); }
  .review-import-mode-grid { grid-template-columns: 1fr; }
  .review-mapping-preset-row { grid-template-columns: 1fr; }
  .review-conversation-list { display: flex; overflow-x: auto; }
  .review-session-item { min-width: 180px; }
  .review-language-popover { width: calc(100% - 20px); left: 10px !important; }
  .review-language-options { grid-template-columns: 1fr 1fr; }
  .review-decision-other-row { align-items: stretch; flex-direction: column; }
  .review-decision-other-row input { min-width: 0; width: 100%; }
}

/* UI/UX Pro Max design system — local B2B productivity workspace */
:root {
  --bg: #f8fafc;
  --panel: #ffffff;
  --panel-subtle: #f8fafc;
  --ink: #0f172a;
  --muted: #526174;
  --line: #dbe3ec;
  --line-strong: #c6d2df;
  --primary: #0369a1;
  --primary-strong: #075985;
  --primary-soft: #eaf4fb;
  --danger: #b42318;
  --success: #147a5f;
  --shadow: none;
  --radius-sm: 6px;
  --radius-md: 8px;
  --radius-lg: 12px;
}
html { background: var(--bg); }
body {
  min-width: 320px;
  background: var(--bg);
  color: var(--ink);
  font-family: Inter, "Microsoft YaHei UI", "Segoe UI", Arial, sans-serif;
  font-size: 14px;
  line-height: 1.5;
}
button, summary, [role="button"] { cursor: pointer; }
button:hover:not(:disabled), .nav-link:hover, .tool-guide-card:hover { transform: none; }
:where(button, input, select, textarea, summary, a):focus-visible {
  outline: 2px solid #0ea5e9;
  outline-offset: 2px;
  box-shadow: 0 0 0 4px rgba(14, 165, 233, .15);
}
.skip-link {
  position: fixed;
  z-index: 100;
  top: 12px;
  left: 12px;
  padding: 8px 12px;
  border: 1px solid #0ea5e9;
  border-radius: var(--radius-sm);
  background: #fff;
  color: #075985;
  font-weight: 750;
  transform: translateY(-160%);
  transition: transform .16s ease-out;
}
.skip-link:focus { transform: translateY(0); }
.shell { grid-template-columns: 236px minmax(0, 1fr); min-height: 100vh; }
.sidebar {
  position: sticky;
  top: 0;
  display: flex;
  min-height: 100vh;
  max-height: 100vh;
  flex-direction: column;
  padding: 20px 14px 16px;
  overflow-y: auto;
  background: #fff;
  border-right: 1px solid var(--line);
  backdrop-filter: none;
}
.brand { gap: 10px; margin: 0 6px 24px; }
.brand-mark { width: 34px; height: 34px; border: 1px solid var(--line); border-radius: 8px; box-shadow: none; }
.brand-copy { gap: 1px; }
.brand-version { order: 3; padding: 0; background: transparent; color: #64748b; font-size: 10px; font-weight: 650; letter-spacing: .04em; }
.brand strong { font-size: 15px; letter-spacing: -.02em; }
.brand small { font-size: 11px; color: #718096; }
.sidebar-nav { gap: 4px; margin: 0; }
.nav-link, .nav-link-top, .nav-link-sub {
  min-height: 34px;
  padding: 0 10px;
  border: 1px solid transparent;
  border-radius: var(--radius-sm);
  background: transparent;
  box-shadow: none;
  color: #475569;
  font-size: 13px;
  font-weight: 650;
  transition: background .16s ease, border-color .16s ease, color .16s ease;
}
.nav-link-top { min-height: 38px; color: #1e293b; font-weight: 720; }
.nav-link-sub { min-height: 32px; padding-left: 28px; font-size: 12px; }
.nav-link.active, .nav-link:hover, .nav-link-top:hover {
  background: var(--primary-soft);
  border-color: #c8e5f7;
  color: var(--primary-strong);
}
.nav-accordion { border: 0; border-radius: 0; background: transparent; box-shadow: none; overflow: visible; }
.nav-accordion[open] { background: transparent; box-shadow: none; }
.nav-accordion-summary { min-height: 38px; padding: 0 10px; border: 1px solid transparent; border-radius: var(--radius-sm); }
.nav-accordion-summary:hover { background: #f5f8fb; border-color: #e5ebf1; }
.nav-accordion-summary::after { width: 18px; height: 18px; border-radius: 4px; background: transparent; color: #64748b; font-size: 16px; font-weight: 500; }
.nav-group-title { color: #1e293b; font-size: 13px; font-weight: 720; letter-spacing: 0; }
.nav-submenu { gap: 2px; padding: 2px 0 4px; }
.sidebar-status {
  gap: 10px;
  margin: auto 0 0;
  padding: 12px;
  background: #f8fafc;
  border: 1px solid var(--line);
  border-radius: var(--radius-md);
}
.sidebar-status-heading { display: flex; align-items: center; justify-content: space-between; color: #475569; font-size: 11px; font-weight: 760; letter-spacing: .06em; text-transform: uppercase; }
.sidebar-local-badge { padding: 2px 5px; border: 1px solid #cfe5dc; border-radius: 4px; background: #f3fbf7; color: var(--success); font-size: 10px; letter-spacing: 0; }
.status-block { gap: 4px; padding-bottom: 9px; border-bottom-color: #e3eaf1; }
.status-caption { color: #718096; font-size: 10px; font-weight: 760; letter-spacing: .06em; }
.status-block strong { color: #1e293b; font-size: 12px; }
.status-line { display: flex; align-items: center; justify-content: space-between; gap: 8px; }
.status-line strong { font-size: 12px; }
.pill { padding: 2px 7px; border: 1px solid #cfe5dc; border-radius: 999px; background: #f3fbf7; color: var(--success); font-size: 11px; font-weight: 720; }
.pill.running { border-color: #f4d698; background: #fff9e8; color: #9a6700; }
.pill.failed { border-color: #f0c4bf; background: #fff5f4; color: var(--danger); }
.service-tip { display: flex; align-items: center; justify-content: space-between; gap: 8px; color: #718096; font-size: 11px; }
.service-tip code { padding: 2px 4px; border: 1px solid #e3eaf1; border-radius: 4px; background: #fff; color: #475569; font-size: 10px; }
.sidebar-actions-panel { margin-top: 8px; }
.feedback-entry-button { min-height: 34px; border: 1px solid var(--line); border-radius: var(--radius-sm); background: #fff; color: #475569; font-size: 12px; font-weight: 680; box-shadow: none; }
.feedback-entry-button:hover { border-color: #9dcde9; background: var(--primary-soft); color: var(--primary-strong); box-shadow: none; }
.content { max-width: 1480px; padding: 32px 40px 56px; }
.hero { display: block; margin: 0 0 24px; padding: 0 0 20px; border-bottom: 1px solid var(--line); }
.hero-context { display: flex; align-items: center; gap: 8px; margin-bottom: 10px; color: #718096; font-size: 11px; font-weight: 650; letter-spacing: .025em; }
.eyebrow { margin: 0; color: var(--primary-strong); font-size: 11px; font-weight: 760; letter-spacing: .08em; }
.hero-context-divider { width: 3px; height: 3px; border-radius: 50%; background: #94a3b8; }
.hero h1 { max-width: none; font-size: clamp(26px, 2.4vw, 34px); letter-spacing: -.035em; }
.lede { margin: 7px 0 0; max-width: 760px; color: var(--muted); font-size: 14px; line-height: 1.65; }
.inline-notice { display: flex; align-items: flex-start; gap: 10px; margin-bottom: 16px; padding: 10px 12px; border: 1px solid #cfe2f0; border-radius: var(--radius-md); background: #f4f9fd; color: #526174; font-size: 12px; line-height: 1.55; }
.inline-notice strong { color: #24435a; font-weight: 750; }
.inline-notice-icon { display: grid; flex: 0 0 auto; width: 18px; height: 18px; place-items: center; color: var(--primary); }
.inline-notice-icon svg { width: 16px; height: 16px; }
.tool-guide-grid { grid-template-columns: repeat(6, minmax(0, 1fr)); gap: 12px; margin-bottom: 24px; }
.tool-guide-card {
  display: grid;
  min-height: 198px;
  align-content: space-between;
  gap: 18px;
  padding: 16px;
  border: 1px solid var(--line);
  border-radius: var(--radius-md);
  background: var(--panel);
  box-shadow: none;
  color: var(--ink);
  text-align: left;
  transition: border-color .16s ease, background .16s ease;
}
.tool-guide-card:nth-child(1), .tool-guide-card:nth-child(2) { grid-column: span 3; }
.tool-guide-card:nth-child(3) { grid-column: span 4; }
.tool-guide-card:nth-child(4) { grid-column: span 2; }
.tool-guide-card:hover, .tool-guide-card:focus-visible { border-color: #8ec7e8; background: #fbfdff; color: var(--ink); }
.tool-guide-card-top, .tool-guide-card-footer { display: flex; align-items: center; justify-content: space-between; gap: 8px; }
.tool-guide-icon { display: grid; width: 34px; height: 34px; place-items: center; border: 1px solid #cfe2f0; border-radius: var(--radius-sm); background: #f4f9fd; color: var(--primary-strong); }
.tool-guide-icon svg { width: 19px; height: 19px; }
.tool-guide-badge { padding: 3px 6px; border: 1px solid #e1e8ef; border-radius: 4px; color: #64748b; font-size: 10px; font-weight: 720; letter-spacing: .04em; text-transform: uppercase; }
.tool-guide-badge-ai { border-color: #d7d8f6; background: #f7f7ff; color: #5854a8; }
.tool-guide-card-copy { display: grid; gap: 5px; }
.tool-guide-card-copy strong { color: #172033; font-size: 17px; font-weight: 760; letter-spacing: -.02em; }
.tool-guide-card-copy small { max-width: 420px; color: #64748b; font-size: 12px; font-weight: 500; line-height: 1.55; }
.tool-guide-card-footer { padding-top: 10px; border-top: 1px solid #edf1f5; color: #718096; font-size: 11px; font-weight: 650; }
.tool-guide-arrow { color: var(--primary); font-size: 16px; line-height: 1; }
.tool-guide-card-review { border-color: #c9daed; background: #fbfdff; }
.tool-guide-card-review .tool-guide-icon { border-color: #d7d8f6; background: #f7f7ff; color: #5854a8; }
.card, .advanced-card, .hero-card { border: 1px solid var(--line); border-radius: var(--radius-md); background: var(--panel); box-shadow: none; }
.card { padding: 20px; margin-bottom: 16px; }
.advanced-card { margin-bottom: 16px; }
.advanced-card summary { padding: 16px 20px; }
.advanced-card summary::after { min-height: 28px; padding: 4px 8px; border: 1px solid var(--line); border-radius: 5px; background: #f8fafc; color: #475569; font-size: 11px; }
.stage-panel, .review-followup-context, .review-followup-messages { border-radius: var(--radius-md); background: #f8fafc; }
.section-header { margin-bottom: 14px; }
.section-header h2 { font-size: 22px; letter-spacing: -.025em; }
.card-title { margin-bottom: 16px; }
.card-title h2 { font-size: 20px; }
.card-title h3 { font-size: 17px; }
.card-title p, .section-header p { font-size: 13px; line-height: 1.55; }
.grid.two, .stage-grid { gap: 12px; }
label { gap: 6px; color: #526174; font-size: 12px; font-weight: 680; }
input, select, textarea { border-color: var(--line-strong); border-radius: var(--radius-sm); box-shadow: none; }
input, select { min-height: 38px; padding: 8px 10px; }
textarea { padding: 9px 10px; }
input:hover, select:hover, textarea:hover { border-color: #9bb9d0; }
.actions { gap: 8px; margin-top: 14px; }
button { min-height: 36px; border-radius: var(--radius-sm); padding: 0 12px; font-size: 13px; font-weight: 680; box-shadow: none; transition: border-color .16s ease, background .16s ease, color .16s ease; }
button.primary { border: 1px solid #075985; background: var(--primary); color: #fff; box-shadow: none; }
button.primary:hover:not(:disabled) { border-color: #075985; background: #075985; box-shadow: none; }
button.secondary, .mini-button { border: 1px solid var(--line-strong); border-radius: var(--radius-sm); background: #fff; color: #334155; }
button.secondary:hover:not(:disabled), .mini-button:hover:not(:disabled) { border-color: #9dcde9; background: var(--primary-soft); color: var(--primary-strong); }
button.danger { border-color: #f0c4bf; border-radius: var(--radius-sm); background: #fff5f4; }
.check-line { min-height: 32px; padding: 6px 9px; border-radius: var(--radius-sm); background: #fff; }
.subnav { gap: 6px; margin-bottom: 14px; }
.subnav-link { min-height: 32px; border-radius: var(--radius-sm); background: #fff; font-size: 12px; }
.subnav-link.active { border-color: #b9dbee; background: var(--primary-soft); box-shadow: none; }
.sticky-actions { padding: 10px 12px; border-radius: var(--radius-md); background: rgba(255,255,255,.96); box-shadow: none; }
.pattern-table-wrap, .result-box, .review-result-table-wrap, .review-input-table-wrap { border-radius: var(--radius-md); }
.pattern-table th, .pattern-table td { padding: 9px 10px; }
.pattern-table th { background: #f8fafc; color: #526174; font-size: 11px; }
dialog.modal::backdrop { background: rgba(15, 23, 42, .38); backdrop-filter: none; }
.modal-card, .tool-guide-dialog .dialog-card { border-radius: var(--radius-lg); box-shadow: 0 16px 40px rgba(15, 23, 42, .16); }
.tool-guide-dialog .dialog-card { padding: 20px; }
.tool-guide-dialog .dialog-header h3 { font-size: 24px; }
.review-language-popover { border-radius: var(--radius-md); box-shadow: 0 16px 36px rgba(15, 23, 42, .14); }
@media (max-width: 980px) {
  .shell { grid-template-columns: 1fr; }
  .sidebar { position: static; max-height: none; min-height: 0; padding: 14px; }
  .brand { margin-bottom: 14px; }
  .sidebar-nav { grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 4px; }
  .nav-accordion { grid-column: span 2; }
  .sidebar-status { margin-top: 14px; }
  .content { padding: 24px 20px 40px; }
  .tool-guide-grid { grid-template-columns: repeat(2, minmax(0, 1fr)); }
  .tool-guide-card:nth-child(n) { grid-column: span 1; min-height: 176px; }
}
@media (max-width: 560px) {
  .content { padding: 20px 14px 32px; }
  .hero { margin-bottom: 18px; padding-bottom: 16px; }
  .hero h1 { font-size: 25px; }
  .sidebar-nav { grid-template-columns: 1fr; }
  .nav-accordion { grid-column: span 1; }
  .tool-guide-grid { grid-template-columns: 1fr; }
  .tool-guide-card:nth-child(n) { min-height: 152px; }
  .inline-notice { font-size: 11px; }
}
@media (prefers-reduced-motion: reduce) {
  *, *::before, *::after { scroll-behavior: auto !important; transition-duration: .01ms !important; animation-duration: .01ms !important; animation-iteration-count: 1 !important; }
}

/* Nocturne workspace system: restrained glass only on interaction layers. */
:root {
  --bg: #0a1020;
  --panel: #101a2d;
  --panel-subtle: #0d1627;
  --ink: #edf3ff;
  --muted: #9aa9c0;
  --line: #263752;
  --line-strong: #3a5075;
  --primary: #5d8dff;
  --primary-strong: #8eb1ff;
  --primary-soft: rgba(93, 141, 255, .14);
  --danger: #ff8f95;
  --success: #58d7c2;
  --glass: rgba(17, 29, 50, .72);
  --glass-strong: rgba(26, 43, 72, .78);
  --glass-border: rgba(132, 162, 214, .34);
  --radius-sm: 5px;
  --radius-md: 8px;
  --radius-lg: 12px;
}
html { min-width: 320px; overflow-x: hidden; background: var(--bg); color-scheme: dark; }
body {
  min-width: 320px;
  overflow-x: hidden;
  background:
    radial-gradient(circle at 88% -8%, rgba(65, 103, 192, .20), transparent 31rem),
    radial-gradient(circle at 5% 110%, rgba(45, 135, 165, .10), transparent 28rem),
    linear-gradient(rgba(255,255,255,.021) 1px, transparent 1px),
    linear-gradient(90deg, rgba(255,255,255,.021) 1px, transparent 1px),
    var(--bg);
  background-size: auto, auto, 32px 32px, 32px 32px, auto;
  color: var(--ink);
  font-family: Inter, "Microsoft YaHei UI", "Segoe UI", Arial, sans-serif;
  font-size: 14px;
  line-height: 1.5;
}
::selection { background: rgba(110, 153, 255, .38); color: #fff; }
* { min-width: 0; }
*::-webkit-scrollbar { width: 10px; height: 10px; }
*::-webkit-scrollbar-track { background: transparent; }
*::-webkit-scrollbar-thumb { border: 3px solid transparent; border-radius: 999px; background: #405373; background-clip: padding-box; }
*::-webkit-scrollbar-thumb:hover { background: #5a7099; background-clip: padding-box; }
:where(button, input, select, textarea, summary, a):focus-visible {
  outline: 2px solid #91b3ff;
  outline-offset: 2px;
  box-shadow: 0 0 0 4px rgba(93, 141, 255, .22);
}
.shell { grid-template-columns: 244px minmax(0, 1fr); min-height: 100dvh; }
.sidebar {
  position: sticky;
  top: 0;
  min-height: 100dvh;
  max-height: 100dvh;
  padding: 22px 14px 16px;
  overflow-y: auto;
  background: rgba(8, 14, 27, .90);
  border-right: 1px solid rgba(100, 130, 179, .30);
  box-shadow: inset -1px 0 rgba(255,255,255,.025);
  backdrop-filter: blur(18px) saturate(120%);
}
.brand { gap: 10px; margin: 0 8px 28px; }
.brand-mark { width: 34px; height: 34px; border: 1px solid rgba(141,177,255,.55); border-radius: 7px; background: #132342; box-shadow: inset 0 1px 0 rgba(255,255,255,.18), 0 8px 22px rgba(0,0,0,.18); }
.brand-copy { gap: 1px; }
.brand-version { order: 3; padding: 0; color: #7083a2; background: transparent; font-size: 10px; font-weight: 650; letter-spacing: .08em; }
.brand strong { color: #f3f6ff; font-size: 15px; letter-spacing: -.02em; }
.brand small { color: #8393aa; font-size: 11px; }
.sidebar-nav { gap: 5px; margin: 0; }
.nav-link, .nav-link-top, .nav-link-sub {
  min-height: 38px;
  padding: 0 10px;
  border: 1px solid transparent;
  border-radius: 5px;
  background: transparent;
  color: #aebbd0;
  font-size: 13px;
  font-weight: 650;
  transition: color .18s ease, background .18s ease, border-color .18s ease, box-shadow .18s ease;
}
.nav-link-top { min-height: 40px; color: #dce6f7; font-weight: 720; }
.nav-link-sub { min-height: 34px; padding-left: 28px; color: #9baac0; font-size: 12px; }
.nav-link:hover, .nav-link-top:hover, .nav-accordion-summary:hover {
  border-color: rgba(132, 162, 214, .20);
  background: rgba(82, 110, 164, .13);
  color: #f2f6ff;
}
.nav-link.active {
  border-color: rgba(130, 165, 236, .42);
  background: linear-gradient(115deg, rgba(93,141,255,.23), rgba(88,215,194,.06));
  color: #fff;
  box-shadow: inset 0 1px 0 rgba(255,255,255,.13), 0 8px 20px rgba(0,0,0,.14);
  backdrop-filter: blur(10px) saturate(130%);
}
.nav-accordion, .nav-accordion[open] { overflow: visible; border: 0; border-radius: 0; background: transparent; box-shadow: none; }
.nav-accordion-summary { min-height: 40px; padding: 0 10px; border: 1px solid transparent; border-radius: 5px; color: #dce6f7; transition: color .18s ease, background .18s ease, border-color .18s ease; }
.nav-accordion-summary::after { width: 18px; height: 18px; border: 0; border-radius: 4px; background: transparent; color: #8496b3; font-size: 17px; font-weight: 500; }
.nav-group-title { color: inherit; font-size: 13px; font-weight: 720; }
.nav-submenu { gap: 2px; padding: 3px 0 5px; }
.sidebar-status {
  gap: 10px;
  margin: auto 0 0;
  padding: 12px;
  border: 1px solid rgba(113, 145, 196, .28);
  border-radius: var(--radius-md);
  background: rgba(17, 29, 50, .62);
  box-shadow: inset 0 1px 0 rgba(255,255,255,.05);
  backdrop-filter: blur(12px) saturate(120%);
}
.sidebar-status-heading { color: #9eb0ca; font-size: 10px; font-weight: 760; letter-spacing: .09em; }
.sidebar-local-badge { border-color: rgba(88, 215, 194, .36); border-radius: 3px; background: rgba(88, 215, 194, .09); color: #83e4d2; }
.status-block { gap: 4px; border-bottom-color: rgba(105, 135, 184, .22); }
.status-caption { color: #7f91ad; }
.status-block strong { color: #e5ecfa; }
.service-tip { color: #8394ad; }
.service-tip code { border-color: rgba(113,145,196,.26); background: rgba(4, 10, 21, .42); color: #a9b9d0; }
.pill { border-color: rgba(88,215,194,.34); background: rgba(88,215,194,.10); color: #8be8d5; }
.pill.running { border-color: rgba(244,183,64,.38); background: rgba(244,183,64,.12); color: #ffd577; }
.pill.failed { border-color: rgba(255,143,149,.38); background: rgba(255,143,149,.12); color: #ffadb1; }
.feedback-entry-button { min-height: 38px; border: 1px solid rgba(123, 155, 212, .34); border-radius: 5px; background: rgba(21, 35, 59, .68); color: #c5d3e8; box-shadow: inset 0 1px 0 rgba(255,255,255,.06); backdrop-filter: blur(12px); }
.feedback-entry-button:hover { border-color: rgba(159,190,249,.62); background: rgba(73, 104, 164, .23); color: #fff; box-shadow: inset 0 1px 0 rgba(255,255,255,.12); }
.content { width: min(1480px, 100%); max-width: none; padding: 36px 44px 56px; }
.hero { margin: 0 0 24px; padding: 0 0 22px; border-bottom: 1px solid rgba(104, 137, 188, .32); }
.hero-context { color: #8194b3; font-size: 11px; letter-spacing: .08em; }
.eyebrow { color: #83e4d2; font-size: 11px; font-weight: 760; }
.hero-context-divider { background: #627694; }
.hero h1 { color: #f2f6ff; font-size: clamp(27px, 2.4vw, 35px); letter-spacing: -.045em; }
.lede { color: #a5b4c9; font-size: 14px; }
.inline-notice {
  border-color: rgba(107, 140, 194, .35);
  border-radius: var(--radius-md);
  background: rgba(18, 33, 58, .66);
  color: #b5c3d7;
  box-shadow: inset 0 1px 0 rgba(255,255,255,.05);
  backdrop-filter: blur(14px) saturate(120%);
}
.inline-notice strong { color: #eff5ff; }
.inline-notice-icon { color: #83e4d2; }
.tool-guide-grid { grid-template-columns: repeat(6, minmax(0, 1fr)); gap: 12px; }
.tool-guide-card {
  min-height: 196px;
  border-color: rgba(100, 132, 183, .38);
  border-radius: var(--radius-md);
  background: linear-gradient(135deg, rgba(24, 40, 68, .80), rgba(13, 23, 41, .80));
  box-shadow: inset 0 1px 0 rgba(255,255,255,.06), 0 16px 38px rgba(0,0,0,.13);
  color: #eff4ff;
  backdrop-filter: blur(16px) saturate(115%);
  transition: border-color .18s ease, background .18s ease, box-shadow .18s ease;
}
.tool-guide-card:hover, .tool-guide-card:focus-visible { border-color: rgba(147,181,245,.70); background: linear-gradient(135deg, rgba(31, 54, 91, .88), rgba(15, 29, 51, .88)); color: #fff; box-shadow: inset 0 1px 0 rgba(255,255,255,.10), 0 18px 40px rgba(0,0,0,.20); }
.tool-guide-icon { border-color: rgba(131,169,235,.45); border-radius: 5px; background: rgba(93,141,255,.12); color: #a8c2ff; }
.tool-guide-badge { border-color: rgba(116,150,205,.32); border-radius: 3px; color: #96a9c5; }
.tool-guide-badge-ai { border-color: rgba(130,150,255,.36); background: rgba(116,125,255,.10); color: #bdc5ff; }
.tool-guide-card-copy strong { color: #f2f6ff; }
.tool-guide-card-copy small { color: #a9b8ce; }
.tool-guide-card-footer { border-top-color: rgba(112,145,196,.22); color: #8498b6; }
.tool-guide-arrow { color: #83e4d2; }
.tool-guide-card-review { border-color: rgba(126,155,233,.50); background: linear-gradient(135deg, rgba(34, 48, 92, .86), rgba(15, 26, 53, .84)); }
.tool-guide-card-review .tool-guide-icon { border-color: rgba(150,170,255,.45); background: rgba(121,132,255,.14); color: #c0c9ff; }
.card, .advanced-card, .hero-card {
  border-color: rgba(100, 132, 183, .36);
  border-radius: var(--radius-md);
  background: rgba(16, 27, 47, .80);
  box-shadow: inset 0 1px 0 rgba(255,255,255,.05), 0 16px 38px rgba(0,0,0,.12);
  backdrop-filter: blur(16px) saturate(115%);
}
.card { padding: 20px; margin-bottom: 16px; }
.advanced-card summary { padding: 16px 20px; }
.advanced-card summary::after { border-color: rgba(119,151,205,.34); border-radius: 4px; background: rgba(82,110,164,.16); color: #b8c8e0; }
.advanced-card summary strong, .section-header h2, .card-title h2, .card-title h3 { color: #eef4ff; }
.advanced-card summary small, .card-title p, .section-header p { color: #9eafc6; }
.stage-panel, .review-followup-context, .review-followup-messages { border: 1px solid rgba(100,132,183,.30); border-radius: var(--radius-md); background: rgba(7,15,30,.36); }
label, .field-label { color: #aebdd1; }
input, select, textarea {
  border-color: rgba(114, 148, 202, .48);
  border-radius: var(--radius-sm);
  background: rgba(7, 15, 30, .48);
  color: #edf3ff;
  box-shadow: inset 0 1px 0 rgba(255,255,255,.025);
  backdrop-filter: blur(10px);
}
input::placeholder, textarea::placeholder { color: #7385a1; }
input:hover, select:hover, textarea:hover { border-color: rgba(155, 187, 242, .65); }
input:focus, select:focus, textarea:focus { border-color: #8eb1ff; }
select option { background: #101a2d; color: #edf3ff; }
.file-card, .check-line { border-color: rgba(114,148,202,.40); border-radius: var(--radius-sm); background: rgba(9,18,34,.42); color: #d9e4f4; }
.file-card.empty { color: #8294ae; }
.check-line { min-height: 36px; }
.actions, .modal-actions, .modal-footer, .modal-footer-actions, .sticky-actions, .review-composer-footer { row-gap: 8px; }
button { min-height: 38px; border-radius: 5px; padding: 0 13px; color: #dce8f9; transition: transform .18s ease, border-color .18s ease, color .18s ease, background .18s ease, box-shadow .18s ease; }
button.primary {
  position: relative;
  isolation: isolate;
  overflow: hidden;
  min-height: 42px;
  border: 1px solid rgba(164,193,255,.60);
  background: linear-gradient(115deg, rgba(122,164,255,.34), rgba(90,215,194,.15));
  color: #f2f7ff;
  box-shadow: inset 0 1px 0 rgba(255,255,255,.34), inset 0 -1px 0 rgba(4,13,30,.30), 0 8px 22px rgba(43,92,194,.20);
  backdrop-filter: blur(16px) saturate(145%);
}
button.primary::before { content: ""; position: absolute; z-index: -1; inset: 1px 48% 1px 1px; background: linear-gradient(90deg, rgba(255,255,255,.18), transparent); }
button.primary:hover:not(:disabled) { transform: translateY(-1px); border-color: rgba(198,218,255,.92); background: linear-gradient(115deg, rgba(127,169,255,.48), rgba(90,215,194,.25)); box-shadow: inset 0 1px 0 rgba(255,255,255,.45), 0 12px 28px rgba(43,92,194,.30); }
button.primary:active:not(:disabled) { transform: translateY(0); box-shadow: inset 0 1px 5px rgba(3,13,33,.42); }
button.secondary, .mini-button {
  border-color: rgba(119,151,205,.40);
  border-radius: 5px;
  background: rgba(20, 34, 57, .60);
  color: #cbd8ea;
  box-shadow: inset 0 1px 0 rgba(255,255,255,.06);
  backdrop-filter: blur(12px);
}
button.secondary:hover:not(:disabled), .mini-button:hover:not(:disabled) { border-color: rgba(161,193,250,.66); background: rgba(70,98,151,.22); color: #fff; }
button.danger { border-color: rgba(255,143,149,.38); border-radius: 5px; background: rgba(147,43,55,.16); color: #ffb2b7; }
button:disabled { opacity: .48; }
.subnav-link { border-color: rgba(114,148,202,.34); border-radius: 5px; background: rgba(16,27,47,.62); color: #a8b8cf; }
.subnav-link.active { border-color: rgba(140,178,246,.52); background: rgba(93,141,255,.14); color: #e8f1ff; }
.sticky-actions { bottom: 14px; border-color: rgba(110,143,194,.40); border-radius: var(--radius-md); background: rgba(12,22,39,.82); box-shadow: inset 0 1px 0 rgba(255,255,255,.06), 0 16px 34px rgba(0,0,0,.18); backdrop-filter: blur(16px); }
.pattern-table-wrap, .result-box, .review-result-table-wrap, .review-input-table-wrap { border-color: rgba(100,132,183,.36); border-radius: var(--radius-md); background: rgba(7,15,30,.38); }
.pattern-table, .result-box { background: transparent; color: #dce7f8; }
.pattern-table th { background: rgba(78,104,151,.16); color: #9eb0c9; }
.pattern-table th, .pattern-table td { border-bottom-color: rgba(100,132,183,.26); }
.pattern-table tbody tr:hover { background: rgba(93,141,255,.07); }
.error-panel { border-color: rgba(255,143,149,.38); border-radius: var(--radius-md); background: rgba(137,42,53,.17); color: #ffbbc0; }
.notice-button { border-color: rgba(244,183,64,.38); border-radius: 5px; background: rgba(244,183,64,.12); color: #ffdc88; box-shadow: inset 0 1px 0 rgba(255,255,255,.05); }
.update-notice-button { border-color: rgba(134,174,246,.45); background: rgba(93,141,255,.14); color: #b8ceff; }
.modal-overlay { z-index: 1200; padding: max(16px, env(safe-area-inset-top)) max(16px, env(safe-area-inset-right)) max(16px, env(safe-area-inset-bottom)) max(16px, env(safe-area-inset-left)); background: rgba(3,8,17,.72); backdrop-filter: blur(7px); }
.modal-card, .tool-guide-dialog .dialog-card {
  width: min(1120px, calc(100vw - 32px));
  max-height: min(820px, calc(100dvh - 32px));
  border: 1px solid rgba(122,154,208,.46);
  border-radius: var(--radius-lg);
  background: linear-gradient(135deg, rgba(25,42,71,.94), rgba(12,22,40,.94));
  color: #eaf1ff;
  box-shadow: inset 0 1px 0 rgba(255,255,255,.09), 0 28px 80px rgba(0,0,0,.46);
  backdrop-filter: blur(22px) saturate(125%);
}
.modal-header { border-bottom: 1px solid rgba(112,145,196,.24); padding-bottom: 14px; }
.modal-header h3, .dialog-head h2, .tool-guide-dialog .dialog-header h3 { color: #f2f6ff; }
.modal-header p, .dialog-head p { color: #a5b5ca; }
.modal-close, .review-mapping-close { border: 1px solid rgba(116,149,202,.30); border-radius: 5px; background: rgba(89,116,167,.15); color: #b7c8e2; }
.modal-close:hover, .review-mapping-close:hover { background: rgba(89,116,167,.28); color: #fff; }
.skip-link { border-color: rgba(130,165,236,.72); border-radius: 5px; background: rgba(13,27,50,.97); color: #dbe9ff; box-shadow: 0 10px 30px rgba(0,0,0,.38); }
.skip-link:focus { box-shadow: 0 0 0 3px rgba(110,158,255,.24), 0 10px 30px rgba(0,0,0,.38); }
.dialog-body {
  border: 1px solid rgba(122,154,208,.46);
  border-radius: var(--radius-lg);
  background: linear-gradient(145deg, rgba(25,42,71,.96), rgba(12,22,40,.97));
  color: #eaf1ff;
  box-shadow: inset 0 1px 0 rgba(255,255,255,.09), 0 28px 80px rgba(0,0,0,.46);
  backdrop-filter: blur(22px) saturate(125%);
}
.dialog-body .dialog-head { padding-bottom: 14px; border-bottom: 1px solid rgba(112,145,196,.24); }
.dialog-body .dialog-actions { margin-top: 2px; padding-top: 14px; border-top: 1px solid rgba(112,145,196,.20); }
.dialog-body .icon-button {
  border: 1px solid rgba(116,149,202,.30);
  border-radius: 5px;
  background: rgba(89,116,167,.15);
  color: #b7c8e2;
}
.dialog-body .icon-button:hover { background: rgba(89,116,167,.28); color: #fff; }
.dialog-body .field label, .dialog-body .mapping-template-hint { color: #aebdd1; }
.dialog-body .hint { color: #9eb0c8; }
dialog.modal, .dialog { z-index: 1300; }
dialog.modal::backdrop, .dialog::backdrop { background: rgba(3,8,17,.72); backdrop-filter: blur(7px); }
.update-release-notes, .feedback-attachment-card, .feedback-log-card { border-color: rgba(100,132,183,.30); border-radius: var(--radius-md); background: rgba(7,15,30,.38); color: #dbe7f8; }
.feedback-dropzone { border-color: rgba(130,164,224,.42); border-radius: var(--radius-md); background: rgba(18,34,59,.52); color: #e8f1ff; }
.feedback-dropzone:hover, .feedback-dropzone:focus-visible { border-color: rgba(163,195,252,.70); background: rgba(44,70,113,.52); }
.feedback-dropzone-title { color: #a9c3ff; }
.feedback-log-card code { background: rgba(3,10,21,.56); color: #b9c9df; }
.review-conversation-shell { height: clamp(580px, calc(100dvh - 188px), 820px); min-height: 580px; border-color: rgba(109,142,193,.42); border-radius: var(--radius-lg); background: rgba(11,21,38,.74); box-shadow: inset 0 1px 0 rgba(255,255,255,.06), 0 18px 42px rgba(0,0,0,.16); }
.review-session-sidebar { border-right-color: rgba(104,137,188,.28); background: rgba(5,13,27,.42); }
.review-session-row { border-radius: 5px; }
.review-session-row:hover, .review-session-row.active { background: rgba(93,141,255,.12); }
.review-session-row.active { box-shadow: inset 2px 0 0 #83e4d2; }
.review-session-item { color: #dce7f8; }
.review-session-item span { color: #899bb6; }
.review-session-rename { color: #94a7c4; }
.review-session-rename:hover { background: rgba(93,141,255,.16); color: #c9daff; }
.review-conversation-head { border-bottom-color: rgba(104,137,188,.28); background: rgba(9,18,33,.30); }
.review-conversation-head h3 { color: #eff5ff; }
.review-conversation-head p { color: #95a7c1; }
.review-conversation-messages { padding: 24px clamp(18px, 7%, 84px); gap: 16px; }
.review-chat-message.user { border: 1px solid rgba(131,166,235,.30); border-radius: 12px 12px 3px 12px; background: rgba(93,141,255,.15); color: #e9f1ff; }
.review-chat-message.workspace-report { border-color: rgba(88,215,194,.26); border-radius: var(--radius-md); background: rgba(18,50,60,.44); color: #dcece9; box-shadow: inset 0 1px 0 rgba(255,255,255,.05); }
.review-workspace-preview-head strong, .review-workspace-file-head strong { color: #eaf4f5; }
.review-workspace-preview-head span, .review-workspace-file-head span, .review-workspace-mapping-head > span:last-child { color: #8ea8ba; }
.review-workspace-total, .review-workspace-language { border-color: rgba(88,215,194,.30); background: rgba(88,215,194,.10); color: #9feadd !important; }
.review-workspace-file-card { border-color: rgba(88,215,194,.20); border-radius: 8px; background: rgba(6,20,34,.34); }
.review-workspace-file-icon { border-color: rgba(130,165,236,.32); border-radius: 5px; background: rgba(93,141,255,.13); color: #bcd2ff; }
.review-workspace-mapping { border-top-color: rgba(100,173,179,.18); }
.review-workspace-endpoint { border-color: rgba(88,215,194,.20); border-radius: 6px; background: rgba(27,73,76,.24); }
.review-workspace-endpoint:last-child { border-color: rgba(130,165,236,.25); background: rgba(34,53,88,.28); }
.review-workspace-endpoint span, .review-workspace-reference { color: #94adb9; }
.review-workspace-endpoint strong { color: #dcece9; }
.review-workspace-arrow { color: #91b4f1; }
.review-workspace-sample { border-color: rgba(112,145,196,.24); border-radius: 6px; background: rgba(4,14,27,.28); }
.review-workspace-sample > span { color: #829bb3; }
.review-workspace-sample-source { border-left-color: #5cbcae; color: #bfdbd7; }
.review-workspace-sample-target { border-left-color: #7da9ee; color: #cbdbfb; }
.review-workspace-empty { color: #94aabc; }
.review-workspace-notes { border-top-color: rgba(100,173,179,.18); color: #97adba; }
.review-workspace-note-copy p { color: #9fb5c3; }
.review-chat-message-meta { color: #8397b4; }
.review-composer { margin: 0 clamp(18px, 7%, 84px) 20px; border-color: rgba(127,161,219,.50); border-radius: var(--radius-md); background: rgba(15,29,51,.76); box-shadow: inset 0 1px 0 rgba(255,255,255,.08), 0 14px 32px rgba(0,0,0,.20); backdrop-filter: blur(16px) saturate(125%); }
.review-composer.dragging { border-color: #83e4d2; background: rgba(25,78,82,.48); }
.review-composer textarea { color: #edf3ff; }
.composer-chip, .review-attachment-chip, .review-target-tab { border-color: rgba(119,151,205,.38); border-radius: 5px; background: rgba(9,18,34,.46); color: #c9d6e9; }
.composer-chip:hover, .review-target-tab:hover { border-color: rgba(161,193,250,.62); background: rgba(75,103,160,.20); color: #fff; }
.review-attachment-chip { background: rgba(88,215,194,.10); color: #a5e9dc; }
.review-message-file { border-color: rgba(130,165,236,.34); border-radius: 5px; background: rgba(93,141,255,.10); color: #bad0ff; }
.review-stream-output, .review-thinking-details { border-color: rgba(100,132,183,.32); border-radius: 5px; background: rgba(3,10,21,.44); color: #b8c8df; }
.review-thinking-details summary { color: #aebfd7; }
.review-thinking-details[open] summary { border-bottom-color: rgba(100,132,183,.26); }
.review-thinking-details .review-stream-output { background: rgba(3,10,21,.44); }
.review-attachment-mapping-card { width: min(680px, calc(100vw - 30px)); border-color: rgba(122,154,208,.46); border-radius: var(--radius-lg); }
.review-mapping-file-head { border-bottom: 1px solid rgba(104,137,188,.23); }
.review-mapping-file-icon { border-radius: 7px; background: rgba(93,141,255,.15); color: #b5cbff; box-shadow: inset 0 0 0 1px rgba(150,180,245,.24); }
.review-mapping-file-copy > span, .review-mapping-lead, .review-mapping-field-label span, .review-mapping-preset-panel > small { color: #9eb0c8; }
.review-mapping-file-copy h2, .review-mapping-field-label label, .review-import-mode-option strong { color: #edf4ff; }
.review-import-mode-option { border-color: rgba(112,145,196,.36); border-radius: var(--radius-md); background: rgba(7,15,30,.42); }
.review-import-mode-option:hover { border-color: rgba(154,187,245,.62); background: rgba(33,56,92,.42); }
.review-import-mode-option.selected { border-color: rgba(131,168,242,.72); background: rgba(93,141,255,.14); box-shadow: inset 0 0 0 2px rgba(93,141,255,.10); }
.review-import-mode-option small { color: #9aabc2; }
.review-import-mode-option strong em { background: rgba(88,215,194,.13); color: #96eadb; }
.review-import-mode-option > i { background: #83e4d2; color: #08201f; }
.review-import-mode-icon { border-radius: 5px; background: rgba(83,112,164,.20); color: #aec2e2; }
.review-import-mode-option.selected .review-import-mode-icon { background: rgba(93,141,255,.20); color: #c6d6ff; }
.review-mapping-preset-panel { border-color: rgba(112,145,196,.30); border-radius: var(--radius-md); background: rgba(7,15,30,.34); }
.sheet-tab { border-color: rgba(112,145,196,.36); border-radius: 5px; background: rgba(9,18,34,.44); color: #9eb2ce; }
.sheet-tab:hover { border-color: rgba(154,187,245,.62); background: rgba(50,76,121,.28); color: #e7efff; }
.sheet-tab.active { border-color: rgba(131,168,242,.70); background: rgba(93,141,255,.16); color: #dbe8ff; }
.mapping-template-bar { border-color: rgba(112,145,196,.34); border-radius: 8px; background: rgba(9,18,34,.44); }
.mapping-row { border-color: rgba(112,145,196,.36); border-radius: 8px; background: rgba(9,18,34,.44); }
.mapping-col-id { color: #aabfff; }
.mapping-header { color: #e8f1ff; }
.mapping-row select, .mapping-row input[type="text"] { border-radius: 5px; background: rgba(4,13,28,.74); color: #e3ecfb; }
.mapping-row select:disabled, .mapping-row input[type="text"]:disabled { border-color: rgba(91,112,147,.28); background: rgba(5,12,24,.42); color: #71839b; opacity: 1; }
.review-attachment-mapping-card > .modal-footer { border-top-color: rgba(104,137,188,.23); background: rgba(8,16,30,.32); }
.review-target-tab.active { border-color: rgba(164,193,255,.58); background: linear-gradient(115deg, rgba(122,164,255,.34), rgba(90,215,194,.15)); color: #f2f7ff; box-shadow: inset 0 1px 0 rgba(255,255,255,.24); }
.review-send-button { min-width: 76px; border-radius: 5px; }
.review-language-popover { z-index: 60; border-color: rgba(122,154,208,.46); border-radius: var(--radius-md); background: rgba(14,26,47,.96); color: #edf3ff; box-shadow: inset 0 1px 0 rgba(255,255,255,.08), 0 20px 48px rgba(0,0,0,.44); backdrop-filter: blur(20px); }
.review-language-popover-head span, .review-language-search-wrap { color: #9cafc7; border-bottom-color: rgba(104,137,188,.26); }
.review-language-option { color: #c8d5e8; border-radius: 4px; }
.review-language-option:hover { background: rgba(93,141,255,.14); }
.review-language-option.selected { background: rgba(93,141,255,.20); color: #eaf1ff; }
.review-language-check { color: #83e4d2; }
.review-term-base-popover { border-color: rgba(122,154,208,.46); border-radius: var(--radius-md); background: rgba(14,26,47,.98); color: #edf3ff; box-shadow: inset 0 1px 0 rgba(255,255,255,.08), 0 20px 48px rgba(0,0,0,.44); backdrop-filter: blur(20px); }
.review-term-base-action, .review-term-base-option { color: #c8d5e8; }
.review-term-base-action:hover, .review-term-base-option:hover { background: rgba(93,141,255,.14); }
.review-term-base-action-icon { background: rgba(93,141,255,.18); color: #bcd0ff; }
.review-term-base-action small, .review-term-base-option small { color: #93a7c1; }
.review-term-base-options { border-top-color: rgba(104,137,188,.26); }
.review-term-base-option.selected { background: rgba(93,141,255,.20); color: #eaf1ff; }
.review-term-base-option-check { color: #83e4d2; }
.review-term-base-delete { color: #8296b0; }
@media (max-width: 1120px) {
  .content { padding: 30px 28px 46px; }
  .tool-guide-grid { grid-template-columns: repeat(2, minmax(0, 1fr)); }
  .tool-guide-card:nth-child(n) { grid-column: span 1; min-height: 178px; }
}
@media (max-width: 980px) {
  .shell { grid-template-columns: 1fr; }
  .sidebar { position: static; min-height: 0; max-height: none; padding: 14px; }
  .brand { margin-bottom: 14px; }
  .sidebar-nav { grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 5px; }
  .nav-accordion { grid-column: span 2; }
  .sidebar-status { margin-top: 14px; }
  .content { padding: 24px 20px 40px; }
  .review-conversation-shell { height: min(760px, calc(100dvh - 120px)); min-height: 560px; }
}
@media (max-width: 700px) {
  .content { padding: 20px 14px 34px; }
  .hero { margin-bottom: 18px; padding-bottom: 17px; }
  .hero h1 { font-size: 26px; }
  .card { padding: 16px; }
  .modal-card, .tool-guide-dialog .dialog-card { width: calc(100vw - 24px); max-height: calc(100dvh - 24px); padding: 16px; }
  .modal-header { align-items: flex-start; }
  .review-conversation-shell { height: min(720px, calc(100dvh - 96px)); min-height: 520px; border-radius: 8px; }
  .review-conversation-head { align-items: flex-start; padding: 14px; }
  .review-conversation-head > :last-child { flex-wrap: wrap; justify-content: flex-end; }
  .review-conversation-messages { padding: 16px; }
  .review-composer { margin: 0 12px 12px; padding: 10px; }
  .review-composer-footer { align-items: stretch; flex-direction: column; }
  .review-composer-tools { width: 100%; }
  .review-send-button { align-self: flex-end; }
  .review-workspace-samples { grid-template-columns: 1fr; }
  .review-inline-progress { grid-template-columns: auto 1fr; }
  .review-inline-progress > :last-child { grid-column: 1 / -1; }
  .review-language-popover { bottom: 48px; left: 8px !important; width: calc(100% - 16px); }
  .review-language-options { grid-template-columns: 1fr 1fr; }
}
@media (max-width: 520px) {
  .sidebar-nav { grid-template-columns: 1fr; }
  .nav-accordion { grid-column: span 1; }
  .tool-guide-grid { grid-template-columns: 1fr; }
  .tool-guide-card:nth-child(n) { min-height: 152px; }
  .grid.two, .stage-grid, .dashboard-grid { grid-template-columns: 1fr; }
  .card-title { align-items: flex-start; flex-direction: column; gap: 6px; }
  .actions > button, .modal-footer-actions > button { flex: 1 1 auto; }
  .review-conversation-shell { min-height: 500px; }
  .review-session-sidebar { padding: 10px; }
  .review-conversation-list { gap: 4px; }
  .review-session-item { padding: 7px; }
  .review-decision-other-row { align-items: stretch; flex-direction: column; }
  .review-decision-other-row input { width: 100%; min-width: 0; }
  .review-decision-other-row button { align-self: flex-end; }
  .review-workspace-flow { grid-template-columns: 1fr; gap: 5px; }
  .review-workspace-arrow { height: 16px; transform: rotate(90deg); }
  .review-language-options { grid-template-columns: 1fr; max-height: 280px; }
}

/* Palette correction: every information surface stays in the graphite family. */
:root {
  --bg: #0b1220;
  --panel: #111c2e;
  --panel-subtle: #0e1828;
  --ink: #eef3fb;
  --muted: #9aa9bd;
  --line: #2a3b57;
  --line-strong: #405675;
  --primary: #8daeff;
  --primary-strong: #b8ccff;
  --primary-soft: rgba(141, 174, 255, .13);
  --success: #7ad8cd;
}
body {
  background:
    radial-gradient(circle at 88% -12%, rgba(79, 111, 181, .14), transparent 34rem),
    radial-gradient(circle at 0 100%, rgba(56, 107, 136, .08), transparent 30rem),
    linear-gradient(rgba(255,255,255,.015) 1px, transparent 1px),
    linear-gradient(90deg, rgba(255,255,255,.015) 1px, transparent 1px),
    var(--bg);
  background-size: auto, auto, 32px 32px, 32px 32px, auto;
}
.card, .advanced-card, .hero-card {
  background: linear-gradient(145deg, rgba(19, 31, 52, .92), rgba(13, 23, 39, .88));
  border-color: rgba(102, 130, 175, .38);
  box-shadow: inset 0 1px 0 rgba(255,255,255,.045), 0 12px 30px rgba(0,0,0,.10);
}
.metrics { gap: 10px; }
.metrics div {
  min-height: 82px;
  padding: 13px 14px;
  border-color: rgba(104, 133, 181, .35);
  border-radius: 8px;
  background: linear-gradient(145deg, rgba(30, 45, 71, .82), rgba(17, 29, 49, .82));
  box-shadow: inset 0 1px 0 rgba(255,255,255,.045);
}
.metrics span { color: #9fb0c8; font-size: 12px; }
.metrics strong { color: #eff4ff; font-size: 21px; letter-spacing: -.025em; }
.compact-metrics div { background: linear-gradient(145deg, rgba(30, 45, 71, .82), rgba(17, 29, 49, .82)); }
.compact-metrics strong { color: #eff4ff; }
.stats-panel {
  border-color: rgba(104,133,181,.34);
  border-radius: 8px;
  background: rgba(11, 21, 37, .58);
  box-shadow: inset 0 1px 0 rgba(255,255,255,.04);
}
.stats-title { color: #a9b9cf; }
.result-file, .notice-list div {
  border-color: rgba(104,133,181,.34);
  border-radius: 8px;
  background: rgba(11, 21, 37, .58);
  box-shadow: inset 0 1px 0 rgba(255,255,255,.035);
}
.result-file strong { color: #e8effa; }
.result-file span, .result-file small { color: #94a6c0; }
.cross-result-meta { border-bottom-color: rgba(104,133,181,.26); background: rgba(11, 21, 37, .54); }
.cross-result-meta span, .cross-cell {
  border-color: rgba(104,133,181,.34);
  border-radius: 7px;
  background: rgba(19, 31, 52, .70);
  color: #c8d5e7;
}
.cross-cell:hover { border-color: rgba(159, 187, 241, .58); box-shadow: none; background: rgba(37, 56, 89, .72); }
.cross-cell.matched { border-color: rgba(244,183,64,.42); background: rgba(104,76,20,.22); }
.cross-cell-index { color: #9aacC3; }
.cross-cell-text { color: #e5edf9; }
.result-box, .log-box, pre { color: #cfdbeb; }
.progress-track, .progress-bar-wrap { background: rgba(130, 153, 190, .16); }
.progress-fill, .progress-bar { background: linear-gradient(90deg, #7ad8cd, #8daeff); }
button.primary {
  border-color: rgba(193, 211, 255, .58);
  background: linear-gradient(115deg, rgba(141,174,255,.34), rgba(122,216,205,.13));
  color: #f3f7ff;
}
button.primary:hover:not(:disabled) {
  border-color: rgba(214,225,255,.88);
  background: linear-gradient(115deg, rgba(154,184,255,.47), rgba(122,216,205,.22));
}
.tool-guide-card, .review-conversation-shell {
  background: linear-gradient(145deg, rgba(20, 33, 55, .90), rgba(13, 23, 39, .88));
  border-color: rgba(102,130,175,.38);
}

/* Full legacy-surface audit: page-specific panels, empty states and editors. */
.scope-item, .prompt-card, .cross-summary-box, .cross-empty-state,
.cross-result-card, .cross-header-item, .diff-header-presence span,
.diff-compare-field-list, .diff-compare-field-item, .review-followup-context,
.review-followup-messages, .update-leaving-card {
  border-color: rgba(104,133,181,.36);
  border-radius: 8px;
  background: linear-gradient(145deg, rgba(22, 36, 59, .82), rgba(12, 22, 38, .78));
  color: #dce7f7;
  box-shadow: inset 0 1px 0 rgba(255,255,255,.04);
}
.scope-empty, .cross-empty-state {
  min-height: 58px;
  display: grid;
  place-items: center;
  border-style: dashed;
  color: #9fb0c7;
  background: rgba(13, 24, 42, .48);
}
.prompt-card { gap: 10px; padding: 15px; }
.prompt-card strong { color: #eff4ff; }
.prompt-card small { color: #9fb0c7; }
.prompt-card textarea, .scope-fields textarea, .review-followup-input textarea {
  border-color: rgba(105, 138, 192, .48);
  border-radius: 6px;
  background: #0b1424;
  color: #dbe7f8;
  box-shadow: inset 0 1px 0 rgba(255,255,255,.025);
}
.prompt-card textarea:disabled, .scope-fields textarea:disabled, textarea:disabled,
input:disabled, select:disabled {
  border-color: rgba(93, 119, 161, .38);
  background: #0a1322;
  color: #9baac0;
  opacity: 1;
}
.prompt-badge { border: 1px solid rgba(122,154,208,.34); border-radius: 4px; background: rgba(90,116,167,.16); color: #aebfd8; }
.prompt-badge.modified { border-color: rgba(244,183,64,.34); background: rgba(118,87,22,.22); color: #ffd684; }
.scope-delete { border: 1px solid rgba(255,143,149,.36); border-radius: 5px; background: rgba(147,43,55,.16); color: #ffb4b8; }
.cross-summary-box { min-height: 112px; }
.cross-summary-box span, .cross-summary-box small, .cross-header-item span { color: #9fb0c7; }
.cross-summary-box strong { color: #edf3fc; }
.cross-header-item { transition: border-color .18s ease, background .18s ease; }
.cross-header-item:hover { border-color: rgba(157,188,241,.62); background: rgba(38,57,91,.66); }
.cross-result-card { overflow: hidden; }
.cross-result-meta { border-bottom-color: rgba(104,133,181,.26); background: rgba(10,20,36,.48); }
.cross-result-meta span { border-color: rgba(104,133,181,.34); border-radius: 4px; background: rgba(40,57,88,.34); color: #aebfd7; }
.cross-merge-actions .check { color: #c7d4e6; }
.diff-header-presence strong, .diff-compare-field-item span { color: #dce7f7; }
.diff-compare-field-item { border-radius: 0; border-width: 0 0 1px; border-bottom-color: rgba(104,133,181,.24); box-shadow: none; }
.diff-compare-field-item:last-child { border-bottom: 0; }
.diff-compare-field-item:hover { background: rgba(51,73,112,.42); }
.diff-inline-side { border-color: rgba(104,133,181,.34); box-shadow: none; }
.diff-inline-side:hover { border-color: rgba(157,188,241,.58); box-shadow: none; }
.diff-inline-delete { background: rgba(129,43,56,.17); border-color: rgba(255,143,149,.34); }
.diff-inline-add { background: rgba(43,109,90,.18); border-color: rgba(122,216,205,.30); }
.diff-inline-delete .diff-inline-label { color: #ffabb0; }
.diff-inline-add .diff-inline-label { color: #91e6d8; }
.diff-inline-text, .followup-context-section p { color: #d4e0f0; }
.diff-token-delete { color: #ffb2b6; }
.diff-token-add { color: #9be5c4; }
.summary-line span { border-color: rgba(104,133,181,.34); border-radius: 5px; background: rgba(38,57,88,.30); color: #aebfd7; }
.followup-message.user { background: rgba(93,141,255,.18); color: #e9f1ff; }
.followup-message.assistant { background: rgba(38,57,88,.42); color: #dce7f7; }
.review-followup-context, .review-followup-messages { overflow: auto; }
.review-detail-table td:nth-child(-n + 5), .review-input-table td:nth-child(-n + 3), .review-input-table td:last-child, .diff-preview-table td:nth-child(-n + 4) { color: #b7c6db; }
.pattern-table, .review-input-table, .review-result-table, .review-detail-table, .diff-preview-table { background: transparent; color: #dce7f7; }
.pattern-table td, .review-input-table td, .review-result-table td, .review-detail-table td, .diff-preview-table td { border-bottom-color: rgba(104,133,181,.22); }
.pattern-table th, .review-input-table th, .review-result-table th, .review-detail-table th, .diff-preview-table th { background: rgba(42,61,94,.40); color: #b7c6dc; }
.update-leaving-card h2 { color: #f0f5ff; }
.update-leaving-card p { color: #a6b6cb; }
.tool-guide-dialog {
  width: min(860px, calc(100vw - 32px));
  max-height: calc(100dvh - 32px);
  padding: 0;
  overflow: hidden;
}
.tool-guide-dialog .dialog-card {
  width: 100%;
  max-height: calc(100dvh - 32px);
  display: grid;
  grid-template-rows: auto minmax(0, 1fr);
  gap: 0;
  padding: 0;
  overflow: hidden;
}
.tool-guide-dialog .dialog-header {
  min-width: 0;
  padding: 20px 22px 16px;
  border-bottom: 1px solid rgba(132,160,204,.34);
}
.tool-guide-dialog .dialog-header h3 { color: #f1f5ff; }
.tool-guide-dialog .dialog-kicker { color: #82dfd1; }
.tool-guide-dialog .markdown-guide {
  min-height: 0;
  overflow: auto;
  padding: 20px 22px 24px;
  color: #d6e1f0;
}
.tool-guide-dialog .markdown-guide section { padding: 0 0 18px; border-bottom: 1px solid rgba(104,133,181,.22); }
.tool-guide-dialog .markdown-guide section + section { padding-top: 18px; }
.tool-guide-dialog .markdown-guide section:last-child { padding-bottom: 0; border-bottom: 0; }
.tool-guide-dialog .markdown-guide h4 { color: #eff4ff; }
.tool-guide-dialog .markdown-guide p, .tool-guide-dialog .markdown-guide ul, .tool-guide-dialog .markdown-guide ol { color: #b6c5d9; }
.tool-guide-dialog .markdown-guide code { border-radius: 4px; background: rgba(141,174,255,.14); color: #c9d9ff; }
.compact-rule-table .pattern-actions button,
.pattern-actions button {
  border: 1px solid rgba(126, 157, 208, .38);
  border-radius: 5px;
  background: linear-gradient(145deg, rgba(62, 86, 129, .38), rgba(20, 34, 57, .68));
  color: #c8d8ed;
  box-shadow: none;
}
.compact-rule-table .pattern-actions button:hover,
.pattern-actions button:hover {
  border-color: rgba(134, 225, 211, .7);
  background: linear-gradient(145deg, rgba(73, 111, 151, .5), rgba(27, 51, 76, .78));
  color: #eff8ff;
}
.preprocess-progress-card {
  display: grid;
  gap: 14px;
}
.preprocess-progress-card .card-title { margin-bottom: 0; align-items: flex-start; }
.preprocess-progress-card .card-title > div { display: grid; gap: 4px; }
.preprocess-status-pill {
  flex: 0 0 auto;
  min-width: 54px;
  padding: 5px 9px;
  border: 1px solid rgba(125, 151, 192, .42);
  border-radius: 999px;
  background: rgba(52, 70, 103, .36);
  color: #b7c7de;
  font-size: 12px;
  font-weight: 750;
  text-align: center;
}
.preprocess-status-pill.is-running { border-color: rgba(141, 174, 255, .56); background: rgba(93, 141, 255, .16); color: #c8d8ff; }
.preprocess-status-pill.is-completed { border-color: rgba(122, 216, 205, .46); background: rgba(65, 174, 156, .14); color: #9ae6d6; }
.preprocess-status-pill.is-failed { border-color: rgba(255, 143, 149, .44); background: rgba(147, 43, 55, .18); color: #ffb8bd; }
.preprocess-status-pill.is-paused { border-color: rgba(244, 183, 64, .42); background: rgba(122, 88, 24, .20); color: #ffd486; }
.preprocess-progress-overview {
  display: grid;
  grid-template-columns: minmax(132px, .72fr) minmax(0, 1.8fr) auto;
  align-items: stretch;
  gap: 1px;
  overflow: hidden;
  border: 1px solid rgba(104, 133, 181, .34);
  border-radius: 8px;
  background: rgba(104, 133, 181, .26);
}
.preprocess-progress-overview > div,
.preprocess-progress-percent {
  display: grid;
  align-content: center;
  gap: 4px;
  min-height: 68px;
  padding: 12px 14px;
  background: rgba(10, 20, 36, .58);
}
.preprocess-progress-overview span { color: #91a5c1; font-size: 12px; }
.preprocess-progress-overview strong { color: #eff5ff; font-size: 19px; letter-spacing: -.02em; }
.preprocess-progress-percent {
  min-width: 74px;
  color: #aee8df !important;
  font-weight: 750;
  text-align: right;
  font-variant-numeric: tabular-nums;
}
.preprocess-progress-card .progress {
  height: 7px;
  margin: -3px 0 0;
  border: 0;
  border-radius: 999px;
  background: rgba(125, 151, 192, .18);
}
.preprocess-progress-card .progress > span {
  display: block;
  height: 100%;
  border-radius: inherit;
  background: linear-gradient(90deg, #72d8cb, #91b0ff 70%, #dbc06d);
  box-shadow: 0 0 14px rgba(122, 216, 205, .20);
  transition: width .28s ease;
}
.preprocess-progress-metrics {
  display: grid;
  grid-template-columns: repeat(6, minmax(0, 1fr));
  gap: 8px;
}
.preprocess-progress-metrics > div {
  display: grid;
  gap: 5px;
  min-width: 0;
  padding: 10px 11px;
  border: 1px solid rgba(104, 133, 181, .28);
  border-radius: 6px;
  background: rgba(13, 25, 44, .46);
}
.preprocess-progress-metrics span { color: #8fa3bf; font-size: 11px; }
.preprocess-progress-metrics strong { overflow: hidden; color: #e6eefb; font-size: 16px; font-variant-numeric: tabular-nums; text-overflow: ellipsis; white-space: nowrap; }
.preprocess-events {
  overflow: hidden;
  border: 1px solid rgba(104, 133, 181, .28);
  border-radius: 7px;
  background: rgba(7, 16, 30, .32);
}
.preprocess-events-head {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  padding: 8px 11px;
  border-bottom: 1px solid rgba(104, 133, 181, .20);
  background: rgba(39, 57, 89, .23);
}
.preprocess-events-head strong { color: #dce8f8; font-size: 12px; }
.preprocess-events-head span { color: #8ea5c5; font-size: 11px; }
.preprocess-events-list { display: grid; gap: 0; }
.preprocess-event {
  display: grid;
  grid-template-columns: 8px minmax(0, 1fr);
  align-items: center;
  gap: 8px;
  min-height: 30px;
  padding: 7px 11px;
  border-top: 1px solid rgba(104, 133, 181, .16);
  color: #9fb0c8;
  font-size: 12px;
  line-height: 1.4;
}
.preprocess-event:first-child { border-top: 0; }
.preprocess-event.is-latest { color: #d3e1f4; background: rgba(93, 141, 255, .06); }
.preprocess-event-marker { width: 5px; height: 5px; border-radius: 50%; background: #7188aa; }
.preprocess-event.is-latest .preprocess-event-marker { background: #7ad8cd; box-shadow: 0 0 0 3px rgba(122, 216, 205, .10); }
.preprocess-event span { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.preprocess-events-empty { padding: 11px; color: #8193ae; font-size: 12px; }
.review-request-queue {
  margin: 0 0 14px;
  border: 1px solid rgba(108, 145, 202, .36);
  border-radius: 8px;
  background: rgba(11, 23, 42, .52);
  overflow: hidden;
}
.review-request-queue-head {
  display: flex;
  align-items: baseline;
  justify-content: space-between;
  gap: 12px;
  padding: 10px 12px;
  border-bottom: 1px solid rgba(104, 133, 181, .24);
  background: rgba(37, 56, 90, .28);
}
.review-request-queue-head strong { color: #edf4ff; font-size: 13px; }
.review-request-queue-head span { color: #9fb3cf; font-size: 12px; white-space: nowrap; }
.review-request-columns {
  display: grid;
  grid-template-columns: repeat(3, minmax(0, 1fr));
  gap: 8px;
  max-height: 280px;
  overflow: auto;
  padding: 10px;
}
.review-request-lane {
  min-width: 0;
  border: 1px solid rgba(104, 133, 181, .27);
  border-radius: 6px;
  background: rgba(10, 20, 36, .36);
  overflow: hidden;
}
.review-request-lane-head {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 8px;
  padding: 8px 9px;
  border-bottom: 1px solid rgba(104, 133, 181, .22);
  font-size: 12px;
}
.review-request-lane-head strong { color: #dbe8f9; font-weight: 700; }
.review-request-lane-head span {
  min-width: 20px;
  padding: 1px 6px;
  border: 1px solid rgba(126, 157, 208, .32);
  border-radius: 999px;
  color: #aebfd7;
  font-variant-numeric: tabular-nums;
  text-align: center;
}
.review-request-lane.status-queued .review-request-lane-head { box-shadow: inset 3px 0 0 #7088ad; }
.review-request-lane.status-submitted .review-request-lane-head { box-shadow: inset 3px 0 0 #78a8ff; }
.review-request-lane.status-thinking .review-request-lane-head { box-shadow: inset 3px 0 0 #b28cff; }
.review-request-lane.status-output .review-request-lane-head { box-shadow: inset 3px 0 0 #5ed4c1; }
.review-request-lane.status-retrying .review-request-lane-head { box-shadow: inset 3px 0 0 #e3b85b; }
.review-request-lane.status-completed .review-request-lane-head { box-shadow: inset 3px 0 0 #72c992; }
.review-request-lane.status-failed .review-request-lane-head { box-shadow: inset 3px 0 0 #ee8793; }
.review-request-items { display: grid; gap: 6px; padding: 8px; }
.review-request-item {
  display: grid;
  gap: 3px;
  padding: 7px 8px;
  border: 1px solid rgba(104, 133, 181, .21);
  border-radius: 5px;
  background: rgba(40, 57, 88, .22);
}
.review-request-item strong { color: #dce8f8; font-size: 12px; font-weight: 650; }
.review-request-item span { overflow: hidden; color: #9fb0c7; font-size: 11px; text-overflow: ellipsis; white-space: nowrap; }
.review-request-item em { color: #f0c46e; font-size: 11px; font-style: normal; }
.review-request-more { padding: 3px 2px 0; color: #8296b2; font-size: 11px; }
.review-request-empty { padding: 10px; color: #7f92ad; font-size: 12px; }
@media (max-width: 700px) {
  .preprocess-progress-overview { grid-template-columns: 1fr auto; }
  .preprocess-progress-stage { grid-column: 1 / -1; border-top: 1px solid rgba(104, 133, 181, .26); }
  .preprocess-progress-metrics { grid-template-columns: repeat(3, minmax(0, 1fr)); }
  .review-request-columns { grid-template-columns: repeat(2, minmax(0, 1fr)); }
  .review-request-queue-head { align-items: flex-start; flex-direction: column; gap: 3px; }
  .tool-guide-dialog { width: calc(100vw - 20px); max-height: calc(100dvh - 20px); }
  .tool-guide-dialog .dialog-card { max-height: calc(100dvh - 20px); }
  .tool-guide-dialog .dialog-header { padding: 16px; }
  .tool-guide-dialog .markdown-guide { padding: 16px; }
}
"""


APP_JS = r"""
const $ = (id) => document.getElementById(id);
let asciiPatterns = [];
let promptTemplates = [];
let builtinRuleRows = [];
let builtinRuleDefinitions = [];
let pendingRuleDefinitions = [];
let pendingRuleState = { count: 0, has_pending: false, notice_seen: true, library_seen: true, show_notice_dot: false, show_library_dot: false };
let appUpdateState = {
  supported: false,
  current_version: "",
  latest_version: "",
  update_available: false,
  release_notes: "",
  published_at: "",
  download_url: "",
  asset_name: "",
  message: "",
};
let appUpdateAutoPrompted = false;
let latestResultFile = "";
let availableModels = [];
let currentProviderName = "DeepSeek";
let currentBaseUrl = "";
let lastScanResult = null;
let crossExcelScanState = null;
let crossExcelSearchState = null;
let crossExcelOutputFile = "";
let diffExcelState = {
  cacheFile: "",
  resultId: "",
  previewRecords: [],
  meta: null,
  outputFile: "",
  totalCount: 0,
  matchedCount: 0,
  previewLimit: 200,
  previewOffset: 0,
  taskId: "",
  previewTruncated: false,
};
let diffFieldMatchSettings = {
  headers: null,
  referenceField: "",
  compareFields: [],
  includeUnmatched: false,
};
let feedbackStatusState = { enabled: false, log_path: "output/log.txt", has_log_file: false };
let feedbackScreenshotFile = null;
let currentPageId = "toolGuidePage";
let lastTrackedToolKey = "";
let aiReviewBatch = null;
let aiReviewTaskId = "";
let aiReviewCurrentTask = null;
let aiReviewLogCursor = 0;
let aiReviewTaskPoller = null;
let aiReviewPromptTemplates = [];
let aiReviewDirectionalTemplates = [];
let aiReviewForbiddenTemplates = [];
let aiReviewExcelMappingPresets = [];
let aiReviewSheetNames = [];
let aiReviewColumnsBySheet = {};
let aiReviewExcelMappingState = {};
let aiReviewActiveSheetName = "";
let aiReviewMappingTemplateIssues = [];
let aiReviewIssueResults = [];
let aiReviewFollowupState = { taskId: "", resultId: "", item: null, messages: [] };
let reviewConversationState = {
  sessions: [],
  currentId: "",
  snapshot: null,
  eventSource: null,
  eventCursor: 0,
  refreshTimer: null,
  sourceLanguage: "auto",
  targetLanguages: ["auto"],
  promptTemplateId: "",
  termBaseId: "",
  termBases: [],
  composerSaveQueue: Promise.resolve(),
  languageMode: "source",
  activeTarget: "",
  streamRunId: "",
  streamText: "",
  streamThinkingText: "",
  streamAnswerText: "",
  streamPhase: "reasoning",
  streamActive: false,
  mappingAttachmentId: "",
  mappingEditorAttachmentId: "",
  mappingEditorPresetId: "",
  mappingEditorContext: "",
  mappingEditorApplyAfterSave: false,
};
const reviewLanguages = [
  "自动检测", "无源文", "简体中文", "繁体中文", "英语", "日语", "韩语", "法语", "德语", "西班牙语", "葡萄牙语",
  "意大利语", "俄语", "阿拉伯语", "泰语", "越南语", "印尼语", "土耳其语", "波兰语", "荷兰语", "瑞典语",
  "挪威语", "丹麦语", "芬兰语", "捷克语", "匈牙利语", "罗马尼亚语", "希腊语", "希伯来语", "乌克兰语",
];
const reviewTermBaseLanguageLabels = {
  English: "英语",
  English_United_Kingdom: "英语（英国）",
  French: "法语",
  German: "德语",
  Italian: "意大利语",
  Japanese: "日语",
  Korean: "韩语",
  Spanish: "西班牙语",
  Portuguese: "葡萄牙语",
  Russian: "俄语",
  Arabic: "阿拉伯语",
  Thai: "泰语",
  Vietnamese: "越南语",
  Indonesian: "印尼语",
  Turkish: "土耳其语",
  Polish: "波兰语",
  Dutch: "荷兰语",
  Swedish: "瑞典语",
  Norwegian: "挪威语",
  Danish: "丹麦语",
  Finnish: "芬兰语",
  Czech: "捷克语",
  Hungarian: "匈牙利语",
  Romanian: "罗马尼亚语",
  Greek: "希腊语",
  Hebrew: "希伯来语",
  Ukrainian: "乌克兰语",
  Chinese_PRC: "简体中文",
  Chinese_Taiwan: "繁体中文",
};
const TASK_STATUS_BY_PAGE = {
  overviewPage: {
    taskLabel: "文本预处理工具",
    pill: "空闲",
    pillClass: "",
    stageLabel: "未启动",
    message: "等待开始任务",
    active: false,
  },
  crossExcelPage: {
    taskLabel: "跨Excel搜索与合并",
    pill: "空闲",
    pillClass: "",
    stageLabel: "未启动",
    message: "等待开始任务",
    active: false,
  },
  diffExcelPage: {
    taskLabel: "Diff 工具",
    pill: "空闲",
    pillClass: "",
    stageLabel: "未启动",
    message: "等待开始任务",
    active: false,
  },
  aiReviewTaskPage: {
    taskLabel: "AI 审校工具",
    pill: "空闲",
    pillClass: "",
    stageLabel: "未启动",
    message: "等待开始任务",
    active: false,
  },
  aiReviewSettingsPage: {
    taskLabel: "AI 审校工具",
    pill: "空闲",
    pillClass: "",
    stageLabel: "未启动",
    message: "等待开始任务",
    active: false,
  },
  aiReviewForbiddenPage: {
    taskLabel: "AI 审校工具",
    pill: "空闲",
    pillClass: "",
    stageLabel: "未启动",
    message: "等待开始任务",
    active: false,
  },
};
const PAGE_TASK_LABELS = {
  toolGuidePage: "工具说明",
  overviewPage: "文本预处理工具",
  modelSettingsPage: "模型设置",
  modelStageSettingsPage: "文本预处理工具",
  nontransSettingsPage: "文本预处理工具",
  promptSettingsPage: "文本预处理工具",
  runDetailsPage: "文本预处理工具",
  resultsPage: "文本预处理工具",
  diffExcelPage: "Diff 工具",
  crossExcelPage: "跨Excel搜索与合并",
  aiReviewTaskPage: "AI 审校工具",
  aiReviewSettingsPage: "AI 审校工具",
  aiReviewForbiddenPage: "AI 审校工具",
};
const PAGE_HERO_COPY = {
  toolGuidePage: { title: "工具说明", lede: "先了解每个工具能做什么，再开始任务。" },
  overviewPage: { title: "文本预处理工具", lede: "用于提取术语、识别非译元素，并整理文本预处理结果。" },
  modelSettingsPage: { title: "模型设置", lede: "统一管理当前工具使用的模型连接。" },
  modelStageSettingsPage: { title: "模型阶段设置", lede: "分别控制非译元素、术语召回和术语校验阶段。" },
  nontransSettingsPage: { title: "非译元素设置", lede: "管理检测规则、内置规则库和保护方式。" },
  promptSettingsPage: { title: "提示词设置", lede: "按阶段维护默认提示词。" },
  runDetailsPage: { title: "运行详情", lede: "查看任务过程、统计和日志。" },
  resultsPage: { title: "结果", lede: "在这里查看导出结果和数量概览。" },
  diffExcelPage: { title: "Excel差异比对", lede: "对比两个 Excel 文件或目录，查看差异、导出结果并批量标记。" },
  crossExcelPage: { title: "跨Excel搜索与合并", lede: "跨文件搜索整行内容，并按表头合并结果。" },
  aiReviewTaskPage: { title: "审校任务", lede: "导入文件、确认映射、查看预览并启动审校任务。" },
  aiReviewSettingsPage: { title: "审校设置", lede: "管理 Workflow 参数、阶段深度思考与调试选项。" },
  aiReviewForbiddenPage: { title: "禁用词", lede: "单独管理禁用词开关与禁用词模板。" },
};
const PAGE_ACCORDION_KEYS = {
  overviewPage: "text-preprocess",
  modelStageSettingsPage: "text-preprocess",
  nontransSettingsPage: "text-preprocess",
  promptSettingsPage: "text-preprocess",
  runDetailsPage: "text-preprocess",
  resultsPage: "text-preprocess",
  diffExcelPage: "diff-tool",
  crossExcelPage: "cross-excel-search",
  aiReviewTaskPage: "ai-review-tool",
  aiReviewSettingsPage: "ai-review-tool",
  aiReviewForbiddenPage: "ai-review-tool",
};
const TOOL_GUIDES = {
  textPreprocess: {
    title: "文本预处理工具",
    sections: [
      ["用途", "从游戏文本表中提取术语，识别不应翻译的标签、变量、占位符等非译元素，并导出整理好的结果表。"],
      ["适合处理", ["Excel 文本表", "XLIFF 文件", "带有 HTML 标签、花括号、变量、格式代码的游戏文本"]],
      ["基本用法", ["在“模型设置”中加载模型。", "进入“文本预处理工具”，选择输入目录和待提取列。", "选择运行模式，点击“开始提取”。", "完成后在“总览”中查看进度并打开输出文件。"]],
      ["输出结果", ["术语库", "非译元素正则规则", "失败记录和任务日志"]],
    ],
  },
  aiReview: {
    title: "AI 审校工具",
    sections: [
      ["用途", "由 Workspace Agent 自动理解附件结构、原文、译文和同一行上下文列，再按目标语言分别执行翻译审校与严格校验。"],
      ["适合处理", ["Excel、XLIFF、CSV 等双语文件", "只有译文的文档或直接粘贴文本", "多文件与多目标语言项目"]],
      ["基本用法", ["进入“AI 审校工具”，拖入一个或多个文件，也可以直接输入待审校文本。", "选择提示词、源语言和目标语言后发送。", "检查 Workspace Agent 的识别报告；默认确认后开始审校。", "首批结果会实时显示在会话下方，完成后可打开独立结果文件。"]],
      ["常用设置", ["模型与 API：统一使用工具“模型设置”中的当前配置。", "并发与分包：控制审校速度和单次响应规模。", "禁用词：关联在提示词模板中，由本地确定性规则检查。"]],
    ],
  },
  crossExcel: {
    title: "跨Excel搜索与合并",
    sections: [
      ["用途", "在多个 Excel 文件里快速搜索文本，或把多个 Excel 按相同表头合并成一个结果文件。"],
      ["适合处理", ["多个结构相近的 Excel", "需要全局查找某个词或句子的项目", "需要按表头抽取并合并列的项目"]],
      ["搜索用法", ["选择 Excel 所在目录。", "输入关键词后点击“搜索”。", "结果会按整行显示，点击单元格可以复制内容。"]],
      ["合并用法", ["选择要合并的表头列。", "点击“合并”。", "完成后打开输出文件或输出目录。"]],
    ],
  },
  diffExcel: {
    title: "Diff 工具",
    sections: [
      ["用途", "对比两个 Excel 文件，或两个目录下的同名 Excel 文件，快速找出所有差异单元格。"],
      ["适合处理", ["版本更新前后文本表比对", "校对不同翻译包的改动", "批量定位某批修改是否已经写入表格"]],
      ["基本用法", ["选择路径 A 和路径 B。", "点击“开始比对”。", "需要时导出差异结果，或把预览结果批量标记回原表。"]],
      ["当前版本支持", ["文件对文件", "目录对目录同名配对", "差异导出", "批量高亮回写原表"]],
    ],
  },
};

const TOOL_KEY_BY_PAGE = {
  toolGuidePage: "home_guide",
  overviewPage: "text_preprocess",
  modelStageSettingsPage: "text_preprocess",
  nontransSettingsPage: "text_preprocess",
  promptSettingsPage: "text_preprocess",
  runDetailsPage: "text_preprocess",
  resultsPage: "text_preprocess",
  diffExcelPage: "diff_excel",
  aiReviewTaskPage: "ai_review",
  aiReviewSettingsPage: "ai_review",
  aiReviewForbiddenPage: "ai_review",
  crossExcelPage: "cross_excel",
};

async function trackToolOpen(toolKey) {
  if (!toolKey || toolKey === lastTrackedToolKey) return;
  lastTrackedToolKey = toolKey;
  try {
    await api("/api/telemetry/tool-open", {
      method: "POST",
      body: JSON.stringify({ tool_key: toolKey }),
    });
  } catch (error) {
  }
}

function setPage(pageId) {
  if (pageId === "runDetailsPage" || pageId === "resultsPage") {
    pageId = "overviewPage";
  }
  currentPageId = pageId;
  document.querySelectorAll(".page-section").forEach((section) => {
    section.classList.toggle("active", section.id === pageId);
  });
  document.querySelectorAll(".nav-link").forEach((button) => {
    button.classList.toggle("active", button.dataset.pageTarget === pageId);
  });
  const heroCopy = PAGE_HERO_COPY[pageId] || PAGE_HERO_COPY.overviewPage;
  $("heroTitle").textContent = heroCopy.title;
  $("heroLede").textContent = heroCopy.lede;
  const activeAccordionKey = PAGE_ACCORDION_KEYS[pageId] || "";
  document.querySelectorAll(".nav-accordion").forEach((accordion) => {
    if (!accordion.dataset.accordionKey) return;
    accordion.open = accordion.dataset.accordionKey === activeAccordionKey;
  });
  if (pageId === "nontransSettingsPage" && pendingRuleState.show_library_dot) {
    markPendingRuleSeen({ library_seen: true }).catch(() => {});
  }
  trackToolOpen(TOOL_KEY_BY_PAGE[pageId] || "");
  renderCurrentTaskStatus();
}

function taskStatusForPage(pageId) {
  if (TASK_STATUS_BY_PAGE[pageId]) {
    return TASK_STATUS_BY_PAGE[pageId];
  }
  if (PAGE_ACCORDION_KEYS[pageId] === "text-preprocess") {
    return TASK_STATUS_BY_PAGE.overviewPage;
  }
  return {
    taskLabel: PAGE_TASK_LABELS[pageId] || "当前任务",
    pill: "空闲",
    pillClass: "",
    stageLabel: "未启动",
    message: "等待开始任务",
    active: false,
  };
}

function setTaskStatus(pageId, nextState = {}) {
  const base = taskStatusForPage(pageId);
  TASK_STATUS_BY_PAGE[pageId] = {
    ...base,
    ...nextState,
  };
}

function renderCurrentTaskStatus() {
  const state = taskStatusForPage(currentPageId);
  renderTaskStatus(
    state.taskLabel,
    state.pill,
    state.pillClass,
    state.stageLabel,
    state.message,
  );
}

function renderToolGuide(toolKey = "textPreprocess") {
  const guide = TOOL_GUIDES[toolKey] || TOOL_GUIDES.textPreprocess;
  $("toolGuideDialogTitle").textContent = guide.title;
  $("toolGuideDialogBody").innerHTML = guide.sections.map(([heading, content]) => {
    const body = Array.isArray(content)
      ? `<ul>${content.map((item) => `<li>${escapeHtml(item)}</li>`).join("")}</ul>`
      : `<p>${escapeHtml(content)}</p>`;
    return `<section><h4>${escapeHtml(heading)}</h4>${body}</section>`;
  }).join("");
  $("toolGuideDialog").showModal();
}

function renderTaskStatus(taskLabel, pillText, pillClass, stageLabel, message) {
  const taskTypeLabel = $("taskTypeLabel");
  const statusPill = $("statusPill");
  const stageLabelNode = $("stageLabel");
  const statusMessage = $("statusMessage");
  if (!taskTypeLabel || !statusPill || !stageLabelNode || !statusMessage) return;
  taskTypeLabel.textContent = taskLabel || "文本预处理工具";
  statusPill.textContent = pillText || "空闲";
  statusPill.className = `pill ${pillClass || ""}`.trim();
  stageLabelNode.textContent = stageLabel || "未启动";
  statusMessage.textContent = message || "等待开始任务";
}

function clearFeedbackForm() {
  $("feedbackMessageInput").value = "";
  $("feedbackScreenshotInput").value = "";
  feedbackScreenshotFile = null;
  $("feedbackScreenshotName").textContent = "未选择截图";
  $("feedbackSubmitHint").textContent = "";
  $("feedbackScreenshotDropzone").classList.remove("drag-over");
}

function renderFeedbackStatus(data) {
  feedbackStatusState = {
    enabled: Boolean(data?.enabled),
    log_path: String(data?.log_path || "output/log.txt"),
    has_log_file: Boolean(data?.has_log_file),
  };
  $("feedbackLogPath").textContent = feedbackStatusState.log_path || "output/log.txt";
  $("feedbackLogHint").textContent = feedbackStatusState.has_log_file
    ? "会自动附带当前日志。"
    : "当前还没有生成日志，提交时会只发送问题描述。";
}

async function loadFeedbackStatus() {
  try {
    const data = await api("/api/feedback/status");
    renderFeedbackStatus(data);
  } catch (error) {
    renderFeedbackStatus({});
  }
}

function openFeedbackModal() {
  $("feedbackOverlay").hidden = false;
  $("feedbackSubmitHint").textContent = "";
  $("feedbackScreenshotDropzone").focus();
}

function closeFeedbackModal() {
  $("feedbackOverlay").hidden = true;
  $("feedbackSubmitHint").textContent = "";
}

async function submitFeedback() {
  const message = String($("feedbackMessageInput").value || "").trim();
  if (!message) {
    $("feedbackSubmitHint").textContent = "请先填写问题。";
    return;
  }
  const formData = new FormData();
  formData.append("message", message);
  const file = feedbackScreenshotFile || $("feedbackScreenshotInput").files?.[0];
  if (file) {
    formData.append("screenshot", file);
  }
  const button = $("submitFeedbackButton");
  button.disabled = true;
  $("feedbackSubmitHint").textContent = "正在提交...";
  try {
    await api("/api/feedback/submit", {
      method: "POST",
      body: formData,
    });
    $("feedbackSubmitHint").textContent = "反馈已提交。";
    setTimeout(() => {
      clearFeedbackForm();
      closeFeedbackModal();
    }, 500);
  } catch (error) {
    $("feedbackSubmitHint").textContent = String(error.message || "反馈提交失败，请稍后重试。");
  } finally {
    button.disabled = false;
  }
}

function setFeedbackScreenshotFile(file) {
  if (!file) {
    feedbackScreenshotFile = null;
    $("feedbackScreenshotName").textContent = "未选择截图";
    return;
  }
  feedbackScreenshotFile = file;
  const sizeKb = Math.max(1, Math.round((Number(file.size || 0) / 1024)));
  $("feedbackScreenshotName").textContent = `${file.name} (${sizeKb} KB)`;
}

async function readClipboardScreenshot(event) {
  if ($("feedbackOverlay").hidden) return;
  const clipboardItems = Array.from(event.clipboardData?.items || []);
  const imageItem = clipboardItems.find((item) => String(item.type || "").startsWith("image/"));
  if (!imageItem) return;
  const file = imageItem.getAsFile();
  if (!file) return;
  event.preventDefault();
  const extension = String(file.type || "image/png").split("/")[1] || "png";
  const stampedFile = new File([file], `pasted_screenshot.${extension}`, { type: file.type || "image/png" });
  setFeedbackScreenshotFile(stampedFile);
  $("feedbackSubmitHint").textContent = "已读取剪贴板截图。";
}

function setSubtab(group, targetId) {
  document.querySelectorAll(`[data-subtab-group="${group}"]`).forEach((button) => {
    button.classList.toggle("active", button.dataset.subtabTarget === targetId);
  });
  document.querySelectorAll(`[data-subtab-panel-group="${group}"]`).forEach((panel) => {
    panel.classList.toggle("active", panel.id === targetId);
  });
  if (group === "nontrans" && targetId === "nontransBuiltinPanel" && pendingRuleState.show_library_dot) {
    markPendingRuleSeen({ library_seen: true }).catch(() => {});
  }
}

async function api(path, options = {}) {
  const isFormData = typeof FormData !== "undefined" && options.body instanceof FormData;
  const headers = { ...(options.headers || {}) };
  if (!isFormData && !headers["Content-Type"]) {
    headers["Content-Type"] = "application/json";
  }
  const response = await fetch(path, {
    headers,
    ...options,
  });
  if (!response.ok) {
    const text = await response.text();
    let message = text || response.statusText;
    try {
      const data = JSON.parse(text);
      message = data?.detail || data?.message || message;
    } catch (_) {}
    throw new Error(message);
  }
  return response.json();
}

function settingsPayload() {
  return {
    provider_name: currentProviderName,
    model_name: $("modelName").value,
    api_key: $("apiKey").value,
    base_url: currentBaseUrl,
    timeout_seconds: Number($("timeoutSeconds").value || 90),
    disable_system_proxy: $("disableSystemProxy").checked,
    extraction_mode: $("extractionMode").value,
    source_language: $("sourceLanguage").value,
    nontrans_chunk_char_limit: Number($("nontransLimit").value || 5000),
    nontrans_placeholder_format: $("nontransPlaceholderFormat").value || "<{n}>",
    term_recall_batch_char_limit: Number($("recallLimit").value || 5000),
    term_review_batch_char_limit: Number($("reviewLimit").value || 5000),
    term_review_max_context_chars: Number($("reviewContextLimit").value || 220),
    nontrans_reasoning_effort: $("nontransReasoningEffort").value || "low",
    term_recall_reasoning_effort: $("recallReasoningEffort").value || "low",
    term_review_reasoning_effort: $("termReviewReasoningEffort").value || "low",
    builtin_regex_enabled: $("builtinRegex").checked,
    ai_discovery_enabled: $("aiDiscovery").checked,
    ai_regex_generation_enabled: $("aiRegex").checked,
    numeric_normalization_enabled: $("numericNormalization").checked,
  };
}

function modelConnectionPayload() {
  return {
    provider_name: currentProviderName,
    model_name: $("modelName").value,
    api_key: $("apiKey").value,
    base_url: currentBaseUrl,
    timeout_seconds: Number($("timeoutSeconds").value || 90),
    disable_system_proxy: $("disableSystemProxy").checked,
  };
}

function setHeaderOptions(headers, preferredValue = "") {
  const select = $("headerName");
  const values = Array.isArray(headers) ? headers.map((item) => String(item || "").trim()).filter(Boolean) : [];
  const currentValue = String(preferredValue || select.value || "").trim();
  select.innerHTML = "";

  const placeholder = document.createElement("option");
  placeholder.value = "";
  placeholder.textContent = values.length ? "请选择待提取列" : "请先选择文件夹";
  placeholder.selected = !currentValue;
  select.appendChild(placeholder);

  values.forEach((name) => {
    const option = document.createElement("option");
    option.value = name;
    option.textContent = name;
    if (name === currentValue) {
      option.selected = true;
    }
    select.appendChild(option);
  });
}

function applyScanResult(data) {
  lastScanResult = data || null;
  $("scanResult").textContent = JSON.stringify(data, null, 2);

  if ((data?.file_type || "") === "xliff") {
    setHeaderOptions(["source"], "source");
    $("headerName").value = "source";
    $("headerName").disabled = true;
    return;
  }

  $("headerName").disabled = false;
  setHeaderOptions(data?.headers || [], $("headerName").value || (data?.headers || [])[0] || "");
}

function setCrossExcelHeaderSelection(selected) {
  document.querySelectorAll('#crossExcelHeaderList input[type="checkbox"]').forEach((checkbox) => {
    checkbox.checked = Boolean(selected);
  });
}

function selectedCrossExcelHeaders() {
  return Array.from(document.querySelectorAll('#crossExcelHeaderList input[type="checkbox"]:checked'))
    .map((checkbox) => checkbox.value.trim())
    .filter(Boolean);
}

function renderCrossExcelHeaders(headers) {
  const container = $("crossExcelHeaderList");
  const values = Array.isArray(headers)
    ? headers.map((item) => String(item || "").trim()).filter(Boolean)
    : [];
  container.innerHTML = "";
  if (!values.length) {
    container.innerHTML = '<div class="cross-empty-state">扫描目录后会在这里显示全部表头。</div>';
    return;
  }
  values.forEach((header, index) => {
    const label = document.createElement("label");
    label.className = "cross-header-item";
    label.innerHTML = `
      <input type="checkbox" value="${escapeHtml(header)}" />
      <span>${escapeHtml(header)}</span>`;
    container.appendChild(label);
  });
}

function renderCrossExcelScan(data) {
  crossExcelScanState = data || null;
  $("crossExcelFileCount").textContent = Number(data?.file_count || 0);
  $("crossExcelScanHint").textContent = data
    ? `已发现 ${Number(data.file_count || 0)} 个文件，${Number((data.headers || []).length)} 个表头。`
    : "请先选择目录并扫描。";
  renderCrossExcelHeaders(data?.headers || []);
  renderCrossExcelSearchResults(null);
  $("crossExcelSearchHint").textContent = "";
  setCrossExcelOutput("");
}

function renderCrossExcelSearchResults(data) {
  crossExcelSearchState = data || null;
  const items = Array.isArray(data?.items) ? data.items : [];
  $("crossExcelMatchCount").textContent = items.length;
  $("crossExcelScannedRows").textContent = Number(data?.scanned_rows || 0);
  $("crossExcelTruncatedLabel").textContent = !data
    ? "未搜索"
    : data.truncated
      ? "已截断"
      : "已完成";

  const container = $("crossExcelSearchResults");
  container.innerHTML = "";
  if (!data) {
    container.innerHTML = '<div class="cross-empty-state">执行搜索后，这里会显示命中的行。</div>';
    return;
  }
  if (!items.length) {
    container.innerHTML = '<div class="cross-empty-state">没有找到匹配内容。</div>';
    return;
  }

  items.forEach((item) => {
    const card = document.createElement("article");
    card.className = "cross-result-card";
    const matched = new Set(Array.isArray(item.matched_columns) ? item.matched_columns.map((value) => Number(value)) : []);
    const cellsHtml = (Array.isArray(item.row_values) ? item.row_values : []).map((value, index) => `
      <button
        type="button"
        class="cross-cell ${matched.has(index) ? "matched" : ""}"
        data-copy-value="${escapeHtml(String(value || ""))}">
        <span class="cross-cell-index">第 ${index + 1} 列</span>
        <span class="cross-cell-text">${escapeHtml(String(value || "")) || "&nbsp;"}</span>
      </button>`).join("");
    card.innerHTML = `
      <div class="cross-result-meta">
        <span>${escapeHtml(String(item.file_name || ""))}</span>
        <span>${escapeHtml(String(item.sheet_name || ""))}</span>
        <span>第 ${Number(item.row_index || 0)} 行</span>
      </div>
      <div class="cross-row-grid">${cellsHtml}</div>`;
    card.querySelectorAll(".cross-cell").forEach((button) => {
      button.addEventListener("click", async () => {
        await copyToClipboard(button.dataset.copyValue || "");
        $("crossExcelSearchHint").textContent = "已复制单元格内容";
        setTimeout(() => {
          if ($("crossExcelSearchHint").textContent === "已复制单元格内容") {
            $("crossExcelSearchHint").textContent = "";
          }
        }, 1200);
      });
    });
    container.appendChild(card);
  });
}

function renderDiffExcelSummary() {
  const meta = diffExcelState.meta || {};
  $("diffModeLabel").textContent = String(meta.compare_mode_label || meta.mode_label || "未开始");
  $("diffFilesInA").textContent = Number(meta.files_in_a || 0);
  $("diffFilesInB").textContent = Number(meta.files_in_b || 0);
  $("diffMatchedPairs").textContent = Number(meta.matched_pairs || 0);
  $("diffTotalCount").textContent = Number(diffExcelState.totalCount || 0);
  $("diffVisibleCount").textContent = Number(diffExcelState.matchedCount || 0);
  $("diffOutputFile").textContent = diffExcelState.outputFile || "暂无输出";
  $("diffOutputHint").textContent = diffExcelState.outputFile
    ? "差异结果已导出，可以直接打开。"
    : diffExcelState.previewTruncated
      ? `当前只预览前 ${Number(diffExcelState.previewLimit || 1000)} 条，完整结果保存在本地缓存中。`
      : "比对完成后可导出差异结果。";
  $("exportDiffExcelButton").disabled = !diffExcelState.cacheFile || !Number(diffExcelState.totalCount || 0);
  $("highlightDiffExcelButton").disabled = !diffExcelState.cacheFile || !Number(diffExcelState.totalCount || 0);
  $("openDiffOutputFileButton").disabled = !diffExcelState.outputFile;
  $("openDiffOutputFolderButton").disabled = !diffExcelState.outputFile;
}

function diffCompareMode() {
  return $("diffCompareMode").value === "field_match" ? "field_match" : "position";
}

function resetDiffFieldMatchSettings() {
  diffFieldMatchSettings = {
    headers: null,
    referenceField: "",
    compareFields: [],
    includeUnmatched: false,
  };
}

function syncDiffCompareModeUi() {
  const isFieldMatch = diffCompareMode() === "field_match";
  $("openDiffFieldSettingsButton").hidden = !isFieldMatch;
  if (!isFieldMatch) {
    $("diffExcelHint").textContent = "";
  }
}

function renderDiffFieldSettings(data = diffFieldMatchSettings.headers || {}) {
  const commonHeaders = Array.isArray(data.common_headers) ? data.common_headers : [];
  $("diffHeadersACount").textContent = Number((data.headers_a || []).length || 0);
  $("diffHeadersBCount").textContent = Number((data.headers_b || []).length || 0);
  $("diffHeadersCommonCount").textContent = commonHeaders.length;
  $("diffFieldSettingsSummary").textContent = Number(data.matched_pairs || 0)
    ? `已检测到 ${Number(data.matched_pairs || 0)} 对文件，可选择两侧共同存在的表头。`
    : "没有找到可配对的文件，请先检查路径 A 和路径 B。";

  const reference = $("diffReferenceField");
  reference.innerHTML = '<option value="">请选择参考字段</option>';
  commonHeaders.forEach((header) => {
    const option = document.createElement("option");
    option.value = header;
    option.textContent = header;
    option.selected = header === diffFieldMatchSettings.referenceField;
    reference.appendChild(option);
  });

  const list = $("diffCompareFieldList");
  list.innerHTML = "";
  if (!commonHeaders.length) {
    list.innerHTML = '<div class="empty-cell">没有可用的共同表头</div>';
  }
  commonHeaders.forEach((header) => {
    const label = document.createElement("label");
    label.className = "diff-compare-field-item";
    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.value = header;
    checkbox.checked = diffFieldMatchSettings.compareFields.includes(header);
    checkbox.disabled = header === diffFieldMatchSettings.referenceField;
    const text = document.createElement("span");
    text.textContent = header;
    label.append(checkbox, text);
    list.appendChild(label);
  });
  $("diffIncludeUnmatched").checked = Boolean(diffFieldMatchSettings.includeUnmatched);
}

async function openDiffFieldSettings() {
  const pathA = $("diffPathA").value.trim();
  const pathB = $("diffPathB").value.trim();
  if (!pathA || !pathB) {
    $("diffExcelHint").textContent = "请先选择路径 A 和路径 B。";
    return;
  }
  $("diffFieldSettingsHint").textContent = "正在读取表头...";
  $("diffFieldSettingsOverlay").hidden = false;
  const data = await api(`/api/diff-excel/field-match-headers?path_a=${encodeURIComponent(pathA)}&path_b=${encodeURIComponent(pathB)}`);
  diffFieldMatchSettings.headers = data;
  renderDiffFieldSettings(data);
  $("diffFieldSettingsHint").textContent = "";
}

function closeDiffFieldSettings() {
  $("diffFieldSettingsOverlay").hidden = true;
  $("diffFieldSettingsHint").textContent = "";
}

function saveDiffFieldSettings() {
  const referenceField = $("diffReferenceField").value.trim();
  const compareFields = Array.from(document.querySelectorAll('#diffCompareFieldList input[type="checkbox"]:checked'))
    .map((checkbox) => checkbox.value.trim())
    .filter(Boolean);
  if (!referenceField) {
    $("diffFieldSettingsHint").textContent = "请选择参考字段。";
    return;
  }
  if (!compareFields.length) {
    $("diffFieldSettingsHint").textContent = "请至少选择一个比对字段。";
    return;
  }
  if (compareFields.includes(referenceField)) {
    $("diffFieldSettingsHint").textContent = "参考字段不能同时作为比对字段。";
    return;
  }
  diffFieldMatchSettings.referenceField = referenceField;
  diffFieldMatchSettings.compareFields = compareFields;
  diffFieldMatchSettings.includeUnmatched = $("diffIncludeUnmatched").checked;
  closeDiffFieldSettings();
  $("diffExcelHint").textContent = `已设置参考字段：${referenceField}；已选择 ${compareFields.length} 个比对字段。`;
}

function buildDiffTokens(leftText, rightText) {
  function splitChars(text) {
    return Array.from(String(text || ""));
  }

  function buildLcsMatrix(leftTokens, rightTokens) {
    const rows = leftTokens.length + 1;
    const cols = rightTokens.length + 1;
    const matrix = Array.from({ length: rows }, () => new Uint16Array(cols));
    for (let i = leftTokens.length - 1; i >= 0; i -= 1) {
      for (let j = rightTokens.length - 1; j >= 0; j -= 1) {
        if (leftTokens[i] === rightTokens[j]) {
          matrix[i][j] = matrix[i + 1][j + 1] + 1;
        } else {
          matrix[i][j] = Math.max(matrix[i + 1][j], matrix[i][j + 1]);
        }
      }
    }
    return matrix;
  }

  function tokensToDiffHtml(parts, changedClass) {
    if (!parts.length) {
      return '<span class="diff-inline-empty">无</span>';
    }
    return parts.map((part) => {
      const html = escapeHtml(part.text || "");
      if (!html) {
        return "";
      }
      if (part.changed) {
        return `<span class="diff-token ${changedClass}">${html}</span>`;
      }
      return `<span class="diff-plain-token">${html}</span>`;
    }).join("");
  }

  const leftTokens = splitChars(leftText);
  const rightTokens = splitChars(rightText);
  if (!leftTokens.length && !rightTokens.length) {
    return {
      leftHtml: '<span class="diff-inline-empty">无</span>',
      rightHtml: '<span class="diff-inline-empty">无</span>',
    };
  }

  const matrix = buildLcsMatrix(leftTokens, rightTokens);
  const leftParts = [];
  const rightParts = [];
  let i = 0;
  let j = 0;
  while (i < leftTokens.length && j < rightTokens.length) {
    if (leftTokens[i] === rightTokens[j]) {
      leftParts.push({ text: leftTokens[i], changed: false });
      rightParts.push({ text: rightTokens[j], changed: false });
      i += 1;
      j += 1;
      continue;
    }
    if (matrix[i + 1][j] >= matrix[i][j + 1]) {
      leftParts.push({ text: leftTokens[i], changed: true });
      i += 1;
    } else {
      rightParts.push({ text: rightTokens[j], changed: true });
      j += 1;
    }
  }
  while (i < leftTokens.length) {
    leftParts.push({ text: leftTokens[i], changed: true });
    i += 1;
  }
  while (j < rightTokens.length) {
    rightParts.push({ text: rightTokens[j], changed: true });
    j += 1;
  }

  return {
    leftHtml: tokensToDiffHtml(leftParts, "diff-token-delete"),
    rightHtml: tokensToDiffHtml(rightParts, "diff-token-add"),
  };
}

function renderDiffExcelResults() {
  renderDiffExcelResultHead();
  const body = $("diffExcelBody");
  const records = Array.isArray(diffExcelState.previewRecords) ? diffExcelState.previewRecords : [];
  body.innerHTML = "";
  if (!records.length) {
    body.innerHTML = '<tr><td colspan="5" class="empty-cell">暂无差异结果</td></tr>';
    renderDiffExcelSummary();
    renderDiffPreviewPager();
    return;
  }
  records.forEach((item) => {
    const tr = document.createElement("tr");
    const isFieldMatch = String(item.diff_kind || "") !== "cell";
    const location = isFieldMatch
      ? `${item.reference_field || "参考字段"}：${item.reference_value || ""}${item.compare_field ? `\n比对字段：${item.compare_field}` : ""}`
      : `${item.sheet || ""}\n${item.cell_address || ""}`;
    [item.filename_a || "", item.filename_b || "", isFieldMatch ? `${item.sheet_a || ""}\n${item.sheet_b || ""}` : (item.sheet || ""), location].forEach((value) => {
      const td = document.createElement("td");
      td.textContent = String(value || "");
      tr.appendChild(td);
    });

    const diffTd = document.createElement("td");
    if (String(item.diff_kind || "") === "unmatched_reference") {
      const existsInA = Boolean(item.sheet_a && item.cell_address_a);
      const label = existsInA ? "仅 A 有" : "仅 B 有";
      const side = existsInA ? "A" : "B";
      const style = existsInA ? "diff-inline-delete" : "diff-inline-add";
      diffTd.innerHTML = `
        <div class="diff-inline-card diff-inline-single">
          <div class="diff-inline-side ${style}" title="点击跳转到实际存在的参考字段">
            <div class="diff-inline-label">${label}</div>
            <div class="diff-inline-text">${escapeHtml(String(item.reference_value || ""))}</div>
          </div>
        </div>`;
      diffTd.querySelector(".diff-inline-side").addEventListener("click", () => openDiffExcelCell(item, side));
    } else {
      const tokens = buildDiffTokens(item.value_a || "", item.value_b || "");
      diffTd.innerHTML = `
        <div class="diff-inline-card">
          <div class="diff-inline-side diff-inline-delete" title="点击跳转到文件 A 的对应单元格">
            <div class="diff-inline-label">A 删除</div>
            <div class="diff-inline-text">${tokens.leftHtml}</div>
          </div>
          <div class="diff-inline-side diff-inline-add" title="点击跳转到文件 B 的对应单元格">
            <div class="diff-inline-label">B 添加</div>
            <div class="diff-inline-text">${tokens.rightHtml}</div>
          </div>
        </div>`;
      diffTd.querySelector(".diff-inline-delete").addEventListener("click", () => openDiffExcelCell(item, "A"));
      diffTd.querySelector(".diff-inline-add").addEventListener("click", () => openDiffExcelCell(item, "B"));
    }
    tr.appendChild(diffTd);
    body.appendChild(tr);
  });
  renderDiffExcelSummary();
  renderDiffPreviewPager();
}

function renderDiffExcelResultHead() {
  const head = $("diffExcelHead");
  const isFieldMatch = String(diffExcelState.meta?.compare_mode_label || "") === "按字段匹配";
  const labels = isFieldMatch
    ? ["文件 A", "文件 B", "A / B Sheet", "参考字段与比对字段", "差异对照"]
    : ["文件 A", "文件 B", "Sheet", "单元格", "差异对照"];
  head.innerHTML = "";
  labels.forEach((label) => {
    const th = document.createElement("th");
    th.textContent = label;
    head.appendChild(th);
  });
}

function renderDiffPreviewPager() {
  const total = Number(diffExcelState.matchedCount || 0);
  const limit = Number(diffExcelState.previewLimit || 200);
  const offset = Number(diffExcelState.previewOffset || 0);
  const shown = Array.isArray(diffExcelState.previewRecords) ? diffExcelState.previewRecords.length : 0;
  $("diffPreviewPreviousButton").disabled = offset <= 0;
  $("diffPreviewNextButton").disabled = offset + shown >= total;
  $("diffPreviewPageLabel").textContent = total
    ? `显示 ${offset + 1}-${offset + shown} / ${total}`
    : "暂无结果";
}

async function applyDiffExcelFilter() {
  if (!diffExcelState.cacheFile) {
    diffExcelState.previewRecords = [];
    diffExcelState.matchedCount = 0;
    diffExcelState.previewTruncated = false;
    renderDiffExcelResults();
    return;
  }
  const data = await api(`/api/diff-excel/preview?cache_file=${encodeURIComponent(diffExcelState.cacheFile)}&query=&limit=${encodeURIComponent(String(diffExcelState.previewLimit || 200))}&offset=${encodeURIComponent(String(diffExcelState.previewOffset || 0))}`);
  diffExcelState.previewRecords = Array.isArray(data.records) ? data.records : [];
  diffExcelState.matchedCount = Number(data.matched_count || 0);
  diffExcelState.previewTruncated = Boolean(data.preview_truncated);
  diffExcelState.previewOffset = Number(data.offset || 0);
  renderDiffExcelResults();
}

async function changeDiffPreviewPage(direction) {
  const limit = Number(diffExcelState.previewLimit || 200);
  const nextOffset = Math.max(0, Number(diffExcelState.previewOffset || 0) + (direction * limit));
  if (nextOffset === Number(diffExcelState.previewOffset || 0)) return;
  diffExcelState.previewOffset = nextOffset;
  await applyDiffExcelFilter();
}

function clearDiffExcelState() {
  diffExcelState = {
    cacheFile: "",
    resultId: "",
    previewRecords: [],
    meta: null,
    outputFile: "",
    totalCount: 0,
    matchedCount: 0,
    previewLimit: 200,
    previewOffset: 0,
    taskId: "",
    previewTruncated: false,
  };
  $("diffExcelHint").textContent = "";
  $("diffHighlightHint").textContent = "";
  renderDiffExcelResults();
  setTaskStatus("diffExcelPage", {
    active: false,
    taskLabel: "Diff 工具",
    pill: "空闲",
    pillClass: "",
    stageLabel: "未启动",
    message: "等待开始任务",
  });
  renderCurrentTaskStatus();
}

function refreshDiffPresetColorButtons() {
  const current = String($("diffHighlightColor").value || "#FFD966").toLowerCase();
  document.querySelectorAll(".preset-color-btn").forEach((button) => {
    const color = String(button.dataset.color || "").toLowerCase();
    button.classList.toggle("active", color === current);
    button.style.backgroundColor = color || "#FFD966";
  });
}

async function openDiffExcelCell(item, target) {
  const filePath = target === "A" ? item.file_path_a : item.file_path_b;
  const sheetName = target === "A" ? (item.sheet_a || item.sheet || "") : (item.sheet_b || item.sheet || "");
  const cellAddress = target === "A" ? (item.cell_address_a || item.cell_address || "") : (item.cell_address_b || item.cell_address || "");
  if (!filePath || !sheetName || !cellAddress) {
    $("diffExcelHint").textContent = `文件 ${target} 中没有可跳转的位置。`;
    return;
  }
  await api("/api/diff-excel/open-cell", {
    method: "POST",
    body: JSON.stringify({
      file_path: filePath,
      sheet_name: sheetName,
      cell_address: cellAddress,
    }),
  });
}

function renderAiReviewPreview(items) {
  const body = $("previewBody");
  body.innerHTML = "";
  const rows = Array.isArray(items) ? items : [];
  if (!rows.length) {
    body.innerHTML = '<tr><td colspan="6" class="empty-cell">暂无预览</td></tr>';
    return;
  }
  rows.forEach((item) => {
    const tr = document.createElement("tr");
    const location = String(item.sheet_name || item.segment_id || "");
    const values = [
      item.source_file || "",
      location,
      item.row_number || "",
      item.source_text || "",
      item.target_text || "",
      item.status_note || "",
    ];
    values.forEach((value) => {
      const td = document.createElement("td");
      td.textContent = String(value || "");
      tr.appendChild(td);
    });
    body.appendChild(tr);
  });
}

function setAiReviewFileCard(text = "") {
  const card = $("reviewFilePath");
  const displayText = String(text || "").trim();
  card.textContent = displayText || "请选择待审校文件";
  card.classList.toggle("empty", !displayText);
}

function initializeAiReviewExcelMapping(data) {
  aiReviewSheetNames = Array.isArray(data?.sheet_names) ? data.sheet_names : [];
  aiReviewColumnsBySheet = data?.columns_by_sheet || {};
  aiReviewActiveSheetName = aiReviewSheetNames[0] || "";
  aiReviewExcelMappingState = {};
  aiReviewMappingTemplateIssues = [];
  aiReviewSheetNames.forEach((sheetName) => {
    aiReviewExcelMappingState[sheetName] = { sources: [], targets: {}, infos: {} };
  });
  if ($("excelMappingPresetSelect")) {
    $("excelMappingPresetSelect").value = "";
  }
  if ($("excelMappingPresetHint")) {
    $("excelMappingPresetHint").textContent = "";
  }
  $("excelMappingSummary").textContent = "";
}

function renderAiReviewBatch(data) {
  const batch = data?.batch || null;
  aiReviewBatch = batch;
  const isExcel = String(batch?.file_type || "") === "excel";
  $("reviewBatchCount").textContent = Number(batch?.item_count || 0);
  $("reviewFileHint").textContent = data?.message || "尚未读取文件。";
  $("openExcelMappingButton").disabled = !isExcel;
  setAiReviewFileCard(batch?.metadata?.original_file_path || batch?.filename || "");
  $("sourceLanguageInput").value = String(batch?.source_language || $("sourceLanguageInput").value || "");
  $("targetLanguageInput").value = String(batch?.target_language || $("targetLanguageInput").value || "");
  const excelMapping = batch?.metadata?.excel_mapping || null;
  if (excelMapping) {
    $("excelMappingSummary").textContent = summarizeAiReviewExcelMapping(excelMapping);
  } else if (!isExcel) {
    $("excelMappingSummary").textContent = "";
  }
  if (isExcel && !excelMapping) {
    $("openExcelMappingButton").disabled = !aiReviewSheetNames.length;
  }
  if (isExcel && !excelMapping) {
    renderAiReviewPreview([]);
    return;
  }
  renderAiReviewPreview(data?.preview || []);
}

function setAiReviewTaskStatus(nextState = {}) {
  const pages = ["aiReviewTaskPage", "aiReviewSettingsPage", "aiReviewForbiddenPage"];
  pages.forEach((pageId) => {
    setTaskStatus(pageId, {
      taskLabel: "AI 审校工具",
      ...nextState,
    });
  });
  renderCurrentTaskStatus();
}

function updateAiReviewModeVisibility() {
  const enableAi = $("enableAiReview").checked;
  const enableDirectional = enableAi && $("enableDirectionalReview").checked;
  $("enableDirectionalReview").disabled = !enableAi;
  $("directionalReviewLine").classList.toggle("disabled-line", !enableAi);
  $("directionalTemplatePanel").classList.toggle("hidden", !enableDirectional);
  $("promptTemplateSelect").disabled = !enableAi || enableDirectional;
  $("directionalTemplateSelect").disabled = !enableDirectional;
}

function renderAiReviewPromptTemplateOptions() {
  const select = $("promptTemplateSelect");
  select.innerHTML = "";
  const dialogSelect = $("promptDialogTemplateSelect");
  if (dialogSelect) dialogSelect.innerHTML = "";
  aiReviewPromptTemplates.forEach((item) => {
    select.appendChild(new Option(String(item.name || ""), String(item.id || "")));
    if (dialogSelect) dialogSelect.appendChild(new Option(String(item.name || ""), String(item.id || "")));
  });
  if (!reviewConversationState.promptTemplateId && aiReviewPromptTemplates.length) {
    reviewConversationState.promptTemplateId = String(aiReviewPromptTemplates[0].id || "");
  }
  if ($("reviewPromptChip")) updateReviewComposerChips();
}

function renderAiReviewDirectionalTemplateOptions() {
  const select = $("directionalTemplateSelect");
  select.innerHTML = "";
  aiReviewDirectionalTemplates.forEach((item) => {
    select.appendChild(new Option(String(item.name || ""), String(item.id || "")));
  });
  updateAiReviewModeVisibility();
}

function renderAiReviewForbiddenTemplateOptions() {
  const select = $("forbiddenTemplateSelect");
  select.innerHTML = "";
  aiReviewForbiddenTemplates.forEach((item) => {
    select.appendChild(new Option(String(item.name || ""), String(item.id || "")));
  });
}

function clearAiReviewTaskHint() {
  $("reviewTaskHint").textContent = "";
  aiReviewIssueResults = [];
  aiReviewFollowupState = { taskId: "", resultId: "", item: null, messages: [] };
  closeReviewFollowupDialog();
}

function renderAiReviewProgress(task = {}) {
  const total = Math.max(0, Number(task.progress_total ?? task.total_count ?? 0));
  const current = Math.max(0, Number(task.progress_current ?? task.completed_count ?? 0));
  const safeCurrent = total > 0 ? Math.min(current, total) : current;
  const percent = total > 0 ? Math.min(100, Math.max(0, Math.round((safeCurrent / total) * 100))) : 0;
  const isActive = ["pending", "running"].includes(String(task.status || ""));
  $("reviewProgress").textContent = String(task.status_label || task.status || "尚未开始");
  $("reviewProgressPercent").textContent = total > 0 ? `${percent}%` : isActive ? "0%" : "未开始";
  $("reviewProgressCount").textContent = `${safeCurrent} / ${total}`;
  $("reviewFailedCount").textContent = Number(task.failed_count || 0);
  $("reviewRequestedCount").textContent = Number(task.requested_count || 0);
  $("reviewProgressBar").style.width = total > 0 ? `${percent}%` : "0";
}

const REVIEW_REQUEST_STATUS_LANES = [
  ["queued", "排队中"],
  ["submitted", "已提交"],
  ["thinking", "思考中"],
  ["output", "输出中"],
  ["retrying", "等待重试"],
  ["completed", "已完成"],
  ["failed", "失败"],
];

function renderReviewRequestQueue(task = {}) {
  const root = $("reviewRequestQueue");
  const columns = $("reviewRequestColumns");
  const summary = $("reviewRequestQueueSummary");
  if (!root || !columns || !summary) return;
  const states = Array.isArray(task?.request_states) ? task.request_states : [];
  root.classList.toggle("hidden", !task?.id || !states.length);
  columns.innerHTML = "";
  if (!states.length) {
    summary.textContent = task?.id ? "正在创建请求包…" : "等待创建请求";
    return;
  }
  const totalItems = states.reduce((sum, item) => sum + Math.max(0, Number(item.item_count || 0)), 0);
  const completedRequests = states.filter((item) => String(item.status || "queued") === "completed").length;
  summary.textContent = `已完成 ${completedRequests} / ${states.length} 个请求 · ${totalItems} 条`;
  REVIEW_REQUEST_STATUS_LANES.forEach(([status, label]) => {
    const items = states.filter((item) => String(item.status || "queued") === status);
    if (!items.length && status === "failed") return;
    const lane = document.createElement("section");
    lane.className = `review-request-lane status-${status}`;
    const head = document.createElement("div");
    head.className = "review-request-lane-head";
    const heading = document.createElement("strong");
    heading.textContent = label;
    const count = document.createElement("span");
    count.textContent = status === "completed" ? `${items.length} / ${states.length}` : String(items.length);
    head.append(heading, count);
    const body = document.createElement("div");
    body.className = "review-request-items";
    if (!items.length) {
      const empty = document.createElement("div");
      empty.className = "review-request-empty";
      empty.textContent = "—";
      body.appendChild(empty);
    } else {
      items.slice(0, 3).forEach((item) => {
        const entry = document.createElement("div");
        entry.className = "review-request-item";
        const title = document.createElement("strong");
        title.textContent = `请求 ${item.package_index}/${item.package_total} · ${item.item_count} 条`;
        const location = document.createElement("span");
        location.title = String(item.location || "待审校单元");
        location.textContent = location.title;
        entry.append(title, location);
        if (Number(item.attempt_count || 0) > 1) {
          const retry = document.createElement("em");
          retry.textContent = `第 ${item.attempt_count} 次请求`;
          entry.appendChild(retry);
        }
        body.appendChild(entry);
      });
      if (items.length > 3) {
        const more = document.createElement("div");
        more.className = "review-request-more";
        more.textContent = `另有 ${items.length - 3} 个请求，已折叠显示`;
        body.appendChild(more);
      }
    }
    lane.append(head, body);
    columns.appendChild(lane);
  });
}

function summarizeAiReviewExcelMapping(mapping) {
  const sheets = Array.isArray(mapping?.sheets) ? mapping.sheets : [];
  const count = sheets.reduce((total, sheet) => total + Number((sheet.mappings || []).length || 0), 0);
  return count > 0 ? `已配置 ${count} 组原文 / 译文列` : "";
}

function renderAiReviewExcelMappingPresets() {
  const select = $("excelMappingPresetSelect");
  const selectedId = String(select.value || "");
  select.innerHTML = "";
  select.appendChild(new Option("选择模板", ""));
  aiReviewExcelMappingPresets.forEach((item) => {
    const option = new Option(String(item.name || ""), String(item.id || ""));
    option.title = String(item.name || "");
    select.appendChild(option);
  });
  select.value = selectedId;
  $("deleteExcelMappingPresetButton").disabled = !select.value;
}

async function loadAiReviewExcelMappingPresets(selectedId = "") {
  const data = await api("/api/ai-review/excel-mapping-presets");
  aiReviewExcelMappingPresets = Array.isArray(data.presets) ? data.presets : [];
  renderAiReviewExcelMappingPresets();
  if (selectedId) {
    $("excelMappingPresetSelect").value = String(selectedId);
  }
  $("deleteExcelMappingPresetButton").disabled = !$("excelMappingPresetSelect").value;
}

function excelMappingColumnLetter(columnIndex) {
  let value = Number(columnIndex) + 1;
  let result = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    value = Math.floor((value - 1) / 26);
  }
  return result || "?";
}

function renderAiReviewMappingTemplateIssues() {
  const hint = $("excelMappingPresetHint");
  hint.textContent = aiReviewMappingTemplateIssues.join(" ");
}

function clearAiReviewMappingTemplateIssues() {
  if (!aiReviewMappingTemplateIssues.length) return;
  aiReviewMappingTemplateIssues = [];
  renderAiReviewMappingTemplateIssues();
}

function aiReviewColumnLabel(sheetName, columnIndex) {
  const columns = Array.isArray(aiReviewColumnsBySheet[sheetName]) ? aiReviewColumnsBySheet[sheetName] : [];
  const column = columns.find((item) => Number(item.index) === Number(columnIndex));
  if (!column) {
    return `第 ${Number(columnIndex) + 1} 列`;
  }
  return `${column.letter || ""}${column.header ? ` ${column.header}` : ""}`.trim();
}

function syncAiReviewActiveSheetMapping() {
  const sheetName = aiReviewActiveSheetName;
  if (!sheetName) return;
  const sources = [];
  const targets = {};
  const infos = {};
  $("excelMappingColumns").querySelectorAll(".mapping-row").forEach((row) => {
    const columnIndex = Number(row.dataset.columnIndex || 0);
    if (row.querySelector(".mapping-source").checked) {
      sources.push(columnIndex);
    }
    const targetValue = row.querySelector(".mapping-target").value;
    if (targetValue !== "") {
      targets[columnIndex] = Number(targetValue);
    }
    const infoValue = row.querySelector(".mapping-info").value;
    if (infoValue !== "" && targetValue === "") {
      infos[columnIndex] = {
        sourceColumns: [Number(infoValue)],
        category: row.querySelector(".mapping-info-category").value.trim(),
      };
    }
  });
  aiReviewExcelMappingState[sheetName] = { sources, targets, infos };
}

function renderAiReviewExcelSheetTabs() {
  const container = $("excelSheetTabs");
  container.innerHTML = "";
  aiReviewSheetNames.forEach((sheetName) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = `sheet-tab ${sheetName === aiReviewActiveSheetName ? "active" : ""}`.trim();
    button.textContent = String(sheetName || "");
    button.addEventListener("click", () => {
      clearAiReviewTaskHint();
      syncAiReviewActiveSheetMapping();
      aiReviewActiveSheetName = sheetName;
      renderAiReviewExcelMappingDialog();
    });
    container.appendChild(button);
  });
}

function renderAiReviewExcelMappingColumns() {
  const sheetName = aiReviewActiveSheetName;
  const columns = Array.isArray(aiReviewColumnsBySheet[sheetName]) ? aiReviewColumnsBySheet[sheetName] : [];
  const sheetState = aiReviewExcelMappingState[sheetName] || { sources: [], targets: {}, infos: {} };
  const container = $("excelMappingColumns");
  container.innerHTML = "";
  if (!columns.length) {
    container.innerHTML = '<div class="cross-empty-state">当前 sheet 没有可用列。</div>';
    return;
  }
  columns.forEach((column) => {
    const row = document.createElement("div");
    row.className = "mapping-row";
    row.dataset.columnIndex = String(column.index);
    const sourceChecked = Array.isArray(sheetState.sources) && sheetState.sources.includes(column.index);
    const targetValue = sheetState.targets[column.index] === undefined ? "" : String(sheetState.targets[column.index]);
    const infoValue = sheetState.infos[column.index]?.sourceColumns?.[0] === undefined
      ? ""
      : String(sheetState.infos[column.index].sourceColumns[0]);
    row.innerHTML = `
      <div class="mapping-col-id">${escapeHtml(String(column.letter || ""))}</div>
      <div class="mapping-header">${escapeHtml(String(column.header || ""))}</div>
      <label class="check-line"><input class="mapping-source" type="checkbox" ${sourceChecked ? "checked" : ""} /><span>原文</span></label>
      <select class="mapping-target"><option value="">非译文列</option></select>
      <select class="mapping-info"><option value="">非信息列</option></select>
      <input class="mapping-info-category" type="text" placeholder="信息类别" value="${escapeHtml(String(sheetState.infos[column.index]?.category || ""))}" />`;
    const targetSelect = row.querySelector(".mapping-target");
    const infoSelect = row.querySelector(".mapping-info");
    const sourceList = Array.isArray(sheetState.sources) ? sheetState.sources : [];
    sourceList.forEach((sourceIndex) => {
      targetSelect.appendChild(new Option(`译文 -> ${aiReviewColumnLabel(sheetName, sourceIndex)}`, String(sourceIndex)));
      infoSelect.appendChild(new Option(`信息 -> ${aiReviewColumnLabel(sheetName, sourceIndex)}`, String(sourceIndex)));
    });
    targetSelect.value = targetValue;
    infoSelect.value = infoValue;
    targetSelect.disabled = !sourceList.length || sourceChecked;
    infoSelect.disabled = !sourceList.length || sourceChecked;
    const categoryInput = row.querySelector(".mapping-info-category");
    categoryInput.disabled = infoSelect.disabled || !infoSelect.value;
    row.querySelector(".mapping-source").addEventListener("change", () => {
      clearAiReviewTaskHint();
      clearAiReviewMappingTemplateIssues();
      syncAiReviewActiveSheetMapping();
      renderAiReviewExcelMappingDialog();
    });
    targetSelect.addEventListener("change", () => {
      clearAiReviewTaskHint();
      clearAiReviewMappingTemplateIssues();
      if (targetSelect.value !== "") {
        infoSelect.value = "";
        categoryInput.disabled = true;
      }
      syncAiReviewActiveSheetMapping();
      renderAiReviewExcelMappingDialog();
    });
    infoSelect.addEventListener("change", () => {
      clearAiReviewTaskHint();
      clearAiReviewMappingTemplateIssues();
      if (infoSelect.value !== "") {
        targetSelect.value = "";
        categoryInput.disabled = false;
      } else {
        categoryInput.disabled = infoSelect.disabled;
      }
      syncAiReviewActiveSheetMapping();
      renderAiReviewExcelMappingDialog();
    });
    categoryInput.addEventListener("input", () => {
      clearAiReviewTaskHint();
      clearAiReviewMappingTemplateIssues();
      syncAiReviewActiveSheetMapping();
      $("applyExcelMappingButton").disabled = validateAiReviewExcelMapping(buildAiReviewExcelMapping()).length > 0;
    });
    container.appendChild(row);
  });
}

function renderAiReviewExcelMappingDialog() {
  renderAiReviewExcelSheetTabs();
  renderAiReviewExcelMappingColumns();
  renderAiReviewMappingTemplateIssues();
  const mapping = buildAiReviewExcelMapping();
  $("applyExcelMappingButton").disabled = validateAiReviewExcelMapping(mapping).length > 0;
}

function buildAiReviewExcelMapping() {
  syncAiReviewActiveSheetMapping();
  const sheets = aiReviewSheetNames.map((sheetName) => {
    const sheetState = aiReviewExcelMappingState[sheetName] || { sources: [], targets: {}, infos: {} };
    const mappings = (sheetState.sources || []).map((sourceColumn) => {
      const targetEntry = Object.entries(sheetState.targets || {}).find(([, sourceIndex]) => Number(sourceIndex) === Number(sourceColumn));
      const infoColumns = Object.entries(sheetState.infos || {})
        .filter(([, info]) => Array.isArray(info.sourceColumns) && info.sourceColumns.includes(sourceColumn))
        .map(([column, info]) => ({ column: Number(column), category: String(info.category || "") }));
      return {
        source_column: Number(sourceColumn),
        target_column: targetEntry ? Number(targetEntry[0]) : null,
        info_columns: infoColumns,
      };
    });
    return { sheet_name: sheetName, mappings };
  });
  return {
    source_language: $("mappingSourceLanguageInput").value.trim(),
    target_language: $("mappingTargetLanguageInput").value.trim(),
    sheets,
  };
}

function validateAiReviewExcelMapping(mapping, { includeTemplateIssues = true } = {}) {
  const errors = includeTemplateIssues ? [...aiReviewMappingTemplateIssues] : [];
  (mapping.sheets || []).forEach((sheet) => {
    const usedTargets = new Set();
    (sheet.mappings || []).forEach((item) => {
      const label = aiReviewColumnLabel(sheet.sheet_name, item.source_column);
      if (item.target_column === null || item.target_column === undefined) {
        errors.push(`${sheet.sheet_name} 的 ${label} 缺少译文列`);
      } else if (usedTargets.has(item.target_column)) {
        errors.push(`${sheet.sheet_name} 的 ${aiReviewColumnLabel(sheet.sheet_name, item.target_column)} 被多个原文列共用`);
      } else if (item.target_column === item.source_column) {
        errors.push(`${sheet.sheet_name} 的 ${label} 不能同时作为译文列`);
      }
      (item.info_columns || []).forEach((info) => {
        if (info.column === item.target_column) {
          errors.push(`${sheet.sheet_name} 的 ${aiReviewColumnLabel(sheet.sheet_name, info.column)} 不能同时作为译文列和信息列`);
        }
        if (info.column === item.source_column) {
          errors.push(`${sheet.sheet_name} 的 ${aiReviewColumnLabel(sheet.sheet_name, info.column)} 不能同时作为原文列和信息列`);
        }
      });
      usedTargets.add(item.target_column);
    });
  });
  if (!(mapping.sheets || []).some((sheet) => Array.isArray(sheet.mappings) && sheet.mappings.length)) {
    errors.push("请至少选择一组原文列和译文列");
  }
  return errors;
}

function applyAiReviewExcelMappingPresetToState(mapping) {
  $("mappingSourceLanguageInput").value = String(mapping?.source_language || "");
  $("mappingTargetLanguageInput").value = String(mapping?.target_language || "");
  const issues = [];
  const nextState = {};
  aiReviewSheetNames.forEach((sheetName) => {
    nextState[sheetName] = { sources: [], targets: {}, infos: {} };
  });
  (mapping?.sheets || []).forEach((sheet, sheetIndex) => {
    const sheetName = aiReviewSheetNames[sheetIndex];
    if (!sheetName) {
      issues.push(`模板需要第 ${sheetIndex + 1} 个工作表，当前文件只有 ${aiReviewSheetNames.length} 个工作表。`);
      return;
    }
    const validColumns = new Set((aiReviewColumnsBySheet[sheetName] || []).map((column) => Number(column.index)));
    (sheet.mappings || []).forEach((item) => {
      const sourceColumn = Number(item.source_column);
      const targetColumn = Number(item.target_column);
      if (!validColumns.has(sourceColumn)) {
        issues.push(`第 ${sheetIndex + 1} 个工作表缺少模板需要的 ${excelMappingColumnLetter(sourceColumn)} 列。`);
        return;
      }
      if (!validColumns.has(targetColumn)) {
        issues.push(`第 ${sheetIndex + 1} 个工作表缺少模板需要的 ${excelMappingColumnLetter(targetColumn)} 列。`);
        return;
      }
      nextState[sheetName].sources.push(sourceColumn);
      nextState[sheetName].targets[targetColumn] = sourceColumn;
      (item.info_columns || []).forEach((info) => {
        const infoColumn = Number(info.column);
        if (!validColumns.has(infoColumn)) {
          issues.push(`第 ${sheetIndex + 1} 个工作表缺少模板需要的 ${excelMappingColumnLetter(infoColumn)} 列。`);
          return;
        }
        const current = nextState[sheetName].infos[infoColumn] || { sourceColumns: [], category: String(info.category || "") };
        if (!current.sourceColumns.includes(sourceColumn)) {
          current.sourceColumns.push(sourceColumn);
        }
        if (info.category && !current.category) {
          current.category = String(info.category);
        }
        nextState[sheetName].infos[infoColumn] = current;
      });
    });
  });
  aiReviewExcelMappingState = nextState;
  aiReviewMappingTemplateIssues = [...new Set(issues)];
}

async function openAiReviewExcelMappingDialog() {
  if (!aiReviewBatch?.id || !aiReviewSheetNames.length) {
    throw new Error("请先读取 Excel 文件");
  }
  clearAiReviewTaskHint();
  await loadAiReviewExcelMappingPresets();
  reviewConversationState.mappingEditorContext = "legacy";
  reviewConversationState.mappingEditorAttachmentId = "";
  reviewConversationState.mappingEditorPresetId = "";
  reviewConversationState.mappingEditorApplyAfterSave = false;
  $("excelMappingDialogTitle").textContent = "Excel 映射";
  $("excelMappingDialogSubtitle").textContent = "按工作表选择原文列、译文列和信息列。";
  $("applyExcelMappingButton").textContent = "确认读取";
  $("mappingSourceLanguageInput").value = $("sourceLanguageInput").value || "";
  $("mappingTargetLanguageInput").value = $("targetLanguageInput").value || "";
  renderAiReviewExcelMappingDialog();
  $("excelMappingDialog").showModal();
}

async function applyAiReviewExcelMapping() {
  if (reviewConversationState.mappingEditorContext === "conversation") {
    const mapping = buildAiReviewExcelMapping();
    const errors = validateAiReviewExcelMapping(mapping);
    if (errors.length) throw new Error(errors[0]);
    const presetId = $("excelMappingPresetSelect").value || reviewConversationState.mappingEditorPresetId;
    if (!presetId) {
      reviewConversationState.mappingEditorApplyAfterSave = true;
      openAiReviewExcelMappingPresetDialog();
      return;
    }
    const preset = aiReviewExcelMappingPresets.find((item) => String(item.id || "") === String(presetId));
    if (!preset) throw new Error("映射模板不存在，请保存为新模板");
    await api("/api/ai-review/excel-mapping-presets", {
      method: "POST",
      body: JSON.stringify({ id: presetId, name: preset.name, mapping }),
    });
    await configureConversationAttachmentWithPreset(presetId);
    return;
  }
  if (!aiReviewBatch?.id) {
    throw new Error("请先读取 Excel 文件");
  }
  const mapping = buildAiReviewExcelMapping();
  const errors = validateAiReviewExcelMapping(mapping);
  if (errors.length) {
    throw new Error(errors[0]);
  }
  const data = await api("/api/ai-review/select-excel-mapping", {
    method: "POST",
    body: JSON.stringify({
      batch_id: aiReviewBatch.id,
      mapping,
    }),
  });
  $("sourceLanguageInput").value = mapping.source_language || "";
  $("targetLanguageInput").value = mapping.target_language || "";
  $("excelMappingSummary").textContent = summarizeAiReviewExcelMapping(mapping);
  $("excelMappingDialog").close();
  clearAiReviewTaskHint();
  renderAiReviewBatch(data);
}

function openAiReviewExcelMappingPresetDialog() {
  const mapping = buildAiReviewExcelMapping();
  const errors = validateAiReviewExcelMapping(mapping, { includeTemplateIssues: false });
  if (errors.length) {
    throw new Error(errors[0]);
  }
  $("excelMappingPresetNameInput").value = "";
  $("excelMappingPresetSaveHint").textContent = "";
  $("excelMappingPresetDialog").showModal();
  $("excelMappingPresetNameInput").focus();
}

async function saveAiReviewExcelMappingPreset() {
  const mapping = buildAiReviewExcelMapping();
  const errors = validateAiReviewExcelMapping(mapping, { includeTemplateIssues: false });
  if (errors.length) {
    throw new Error(errors[0]);
  }
  const name = $("excelMappingPresetNameInput").value.trim();
  if (!name) {
    throw new Error("请输入模板名称");
  }
  const existing = aiReviewExcelMappingPresets.find((item) => String(item.name || "").trim().toLowerCase() === name.toLowerCase());
  let presetId = null;
  if (existing) {
    if (!window.confirm(`已存在“${existing.name}”，是否覆盖？`)) {
      return;
    }
    presetId = String(existing.id || "") || null;
  }
  const data = await api("/api/ai-review/excel-mapping-presets", {
    method: "POST",
    body: JSON.stringify({
      id: presetId,
      name,
      mapping,
    }),
  });
  const savedId = String(data?.preset?.id || "");
  reviewConversationState.mappingEditorPresetId = savedId;
  await loadAiReviewExcelMappingPresets(savedId);
  $("excelMappingPresetDialog").close();
  $("excelMappingPresetHint").textContent = "映射模板已保存。";
  if (reviewConversationState.mappingEditorContext === "conversation" && reviewConversationState.mappingEditorApplyAfterSave) {
    reviewConversationState.mappingEditorApplyAfterSave = false;
    await configureConversationAttachmentWithPreset(savedId);
  }
}

async function configureConversationAttachmentWithPreset(presetId) {
  const attachmentId = reviewConversationState.mappingEditorAttachmentId;
  if (!attachmentId || !presetId) throw new Error("映射附件或模板无效");
  await api(`/api/ai-review/conversations/${encodeURIComponent(reviewConversationState.currentId)}/attachments/${encodeURIComponent(attachmentId)}`, {
    method: "PATCH",
    body: JSON.stringify({ mode: "preset", preset_id: presetId }),
  });
  $("excelMappingDialog").close();
  reviewConversationState.mappingEditorContext = "";
  reviewConversationState.mappingEditorAttachmentId = "";
  reviewConversationState.mappingEditorPresetId = "";
  await refreshCurrentReviewConversation();
}

async function applyAiReviewExcelMappingPreset() {
  const presetId = $("excelMappingPresetSelect").value;
  $("deleteExcelMappingPresetButton").disabled = !presetId;
  if (!presetId) {
    clearAiReviewMappingTemplateIssues();
    return;
  }
  const data = await api(`/api/ai-review/excel-mapping-presets/${encodeURIComponent(presetId)}`);
  applyAiReviewExcelMappingPresetToState(data?.preset?.mapping || {});
  renderAiReviewExcelMappingDialog();
}

async function deleteAiReviewExcelMappingPreset() {
  const presetId = $("excelMappingPresetSelect").value;
  if (!presetId) return;
  if (!window.confirm("确定删除这个映射模板吗？")) {
    return;
  }
  await api(`/api/ai-review/excel-mapping-presets/${encodeURIComponent(presetId)}`, {
    method: "DELETE",
  });
  await loadAiReviewExcelMappingPresets();
  $("excelMappingPresetHint").textContent = "映射模板已删除。";
}

function fillAiReviewPromptDialog(template = {}) {
  $("promptTemplateId").value = String(template.id || "");
  $("promptNameInput").value = String(template.name || "");
  $("systemPromptInput").value = String(template.system_prompt || "");
  $("userPromptInput").value = String(template.user_prompt || "");
  $("promptForbiddenWordsInput").value = String(template.forbidden_words_text || "");
  $("promptDialogTemplateSelect").value = String(template.id || "");
  $("deletePromptButton").disabled = Boolean(template.is_default) || !template.id;
}

async function openAiReviewPromptDialog() {
  const templateId = $("promptTemplateSelect").value;
  if (!templateId) {
    return;
  }
  const data = await api(`/api/ai-review/prompt-templates/${encodeURIComponent(templateId)}`);
  fillAiReviewPromptDialog(data.template || {});
  $("promptDialog").showModal();
}

function newAiReviewPromptTemplate() {
  const defaultTemplate = aiReviewPromptTemplates.find((item) => item.is_default) || aiReviewPromptTemplates[0] || {};
  fillAiReviewPromptDialog({
    id: "",
    name: "新建模板",
    // A new template starts from the proven complete schema rather than a bare
    // {text} placeholder, so users can edit review rules without reconstructing
    // the JSON contract themselves.
    system_prompt: String(defaultTemplate.system_prompt || ""),
    user_prompt: String(defaultTemplate.user_prompt || "{text}"),
    forbidden_words_text: "",
    is_default: false,
  });
}

async function saveAiReviewPromptTemplate() {
  const data = await api("/api/ai-review/prompt-templates", {
    method: "POST",
    body: JSON.stringify({
      id: $("promptTemplateId").value || null,
      name: $("promptNameInput").value || "未命名模板",
      system_prompt: $("systemPromptInput").value,
      user_prompt: $("userPromptInput").value,
      forbidden_words_text: $("promptForbiddenWordsInput").value,
    }),
  });
  await loadAiReviewPromptTemplates();
  $("promptTemplateSelect").value = String(data?.template?.id || "");
  reviewConversationState.promptTemplateId = String(data?.template?.id || reviewConversationState.promptTemplateId || "");
  updateReviewComposerChips();
  await saveReviewConversationComposerSettings();
  $("promptDialog").close();
  $("reviewSettingsHint").textContent = data.message || "提示词模板已保存";
}

async function resetAiReviewPromptTemplate() {
  const data = await api("/api/ai-review/prompt-templates/reset-default", {
    method: "POST",
    body: "{}",
  });
  await loadAiReviewPromptTemplates();
  $("promptTemplateSelect").value = String(data?.template?.id || $("promptTemplateSelect").value || "");
  fillAiReviewPromptDialog(data.template || {});
}

async function deleteAiReviewPromptTemplate() {
  const templateId = $("promptTemplateId").value;
  if (!templateId) {
    return;
  }
  if (!window.confirm("确定删除这个提示词模板吗？")) {
    return;
  }
  const data = await api(`/api/ai-review/prompt-templates/${encodeURIComponent(templateId)}`, {
    method: "DELETE",
  });
  await loadAiReviewPromptTemplates();
  const fallbackId = String(data?.fallback_template?.id || aiReviewPromptTemplates[0]?.id || "");
  reviewConversationState.promptTemplateId = fallbackId;
  if (reviewConversationState.snapshot?.session) {
    reviewConversationState.snapshot.session.prompt_template_id = fallbackId || null;
  }
  if ($("promptTemplateSelect")) $("promptTemplateSelect").value = fallbackId;
  updateReviewComposerChips();
  renderReviewConversationList();
  $("promptDialog").close();
}

function renderDirectionalEditorItems(items) {
  const container = $("directionalItems");
  container.innerHTML = "";
  const rows = Array.isArray(items) ? items : [];
  rows.forEach((item) => {
    const row = document.createElement("label");
    row.className = "directional-item";
    row.innerHTML = `
      <input type="checkbox" ${item.enabled !== false ? "checked" : ""} />
      <input type="text" value="${escapeHtml(String(item.name || ""))}" placeholder="输入审校项名称" />`;
    container.appendChild(row);
  });
}

function appendDirectionalEditorItem(name = "", enabled = true) {
  const current = Array.from($("directionalItems").querySelectorAll(".directional-item")).map((row) => ({
    enabled: row.querySelector('input[type="checkbox"]').checked,
    name: row.querySelector('input[type="text"]').value,
  }));
  current.push({ name, enabled });
  renderDirectionalEditorItems(current);
}

function fillDirectionalDialog(template = {}) {
  $("directionalTemplateId").value = String(template.id || "");
  $("directionalNameInput").value = String(template.name || "");
  renderDirectionalEditorItems(template.items || []);
}

async function openDirectionalDialog() {
  const templateId = $("directionalTemplateSelect").value;
  if (!templateId) {
    return;
  }
  const data = await api(`/api/ai-review/directional-templates/${encodeURIComponent(templateId)}`);
  fillDirectionalDialog(data.template || {});
  $("directionalDialog").showModal();
}

function newDirectionalTemplate() {
  fillDirectionalDialog({
    id: "",
    name: "新建定向模板",
    items: [{ name: "", enabled: true }],
  });
}

async function saveDirectionalTemplateFromDialog() {
  const items = Array.from($("directionalItems").querySelectorAll(".directional-item")).map((row) => ({
    enabled: row.querySelector('input[type="checkbox"]').checked,
    name: row.querySelector('input[type="text"]').value,
  }));
  const data = await api("/api/ai-review/directional-templates", {
    method: "POST",
    body: JSON.stringify({
      id: $("directionalTemplateId").value || null,
      name: $("directionalNameInput").value || "未命名定向模板",
      items,
    }),
  });
  await loadAiReviewDirectionalTemplates();
  $("directionalTemplateSelect").value = String(data?.template?.id || "");
  $("directionalDialog").close();
  $("reviewSettingsHint").textContent = data.message || "定向模板已保存";
}

function fillForbiddenDialog(template = {}) {
  $("forbiddenTemplateId").value = String(template.id || "");
  $("forbiddenNameInput").value = String(template.name || "");
  $("forbiddenWordsInput").value = String(template.words_text || "");
}

async function openForbiddenDialog() {
  const templateId = $("forbiddenTemplateSelect").value;
  if (!templateId) {
    return;
  }
  const data = await api(`/api/ai-review/forbidden-templates/${encodeURIComponent(templateId)}`);
  fillForbiddenDialog(data.template || {});
  $("forbiddenDialog").showModal();
}

function newForbiddenTemplate() {
  fillForbiddenDialog({
    id: "",
    name: "新建禁用词模板",
    words_text: "",
  });
}

async function saveForbiddenTemplateFromDialog() {
  const data = await api("/api/ai-review/forbidden-templates", {
    method: "POST",
    body: JSON.stringify({
      id: $("forbiddenTemplateId").value || null,
      name: $("forbiddenNameInput").value || "未命名禁用词模板",
      words_text: $("forbiddenWordsInput").value,
    }),
  });
  await loadAiReviewForbiddenTemplates();
  $("forbiddenTemplateSelect").value = String(data?.template?.id || "");
  $("forbiddenDialog").close();
  $("reviewForbiddenHint").textContent = data.message || "禁用词模板已保存";
}

function renderAiReviewResults(task, results) {
  aiReviewCurrentTask = task || null;
  renderAiReviewResultHead(task || {});
  const outputFile = String(task?.output_path || task?.output_file || "");
  renderAiReviewProgress(task || {});
  renderReviewRequestQueue(task || {});
  $("outputPath").textContent = outputFile || "暂无输出";
  $("outputPanel").classList.toggle("hidden", !outputFile);
  $("openOutputFileButton").disabled = !outputFile;
  $("openReviewDetailButton").disabled = !task?.id;

  const body = $("reviewResultBody");
  body.innerHTML = "";
  const items = Array.isArray(results) ? results : [];
  if (!items.length) {
    const colspan = Number($("reviewResultHead").querySelectorAll("th").length || 6);
    body.innerHTML = `<tr><td colspan="${colspan}" class="empty-cell">暂无审校结果</td></tr>`;
    return;
  }
  const config = task?.config || {};
  const isDirectional = config.mode === "directional";
  const isForbiddenOnly = config.mode === "forbidden_only";
  const hasForbidden = Boolean(config.enable_forbidden_check);
  const reviewTypes = Array.isArray(config.review_types) ? config.review_types : [];
  items.forEach((item) => {
    const tr = document.createElement("tr");
    let cells = [];
    if (isForbiddenOnly) {
      cells = [
        item.source_text || "",
        item.target_text || "",
        item.matched_words || "",
      ];
    } else if (isDirectional) {
      const checks = item.checks || {};
      cells = [
        item.source_text || "",
        item.target_text || "",
        item.suggestion || "",
        ...reviewTypes.map((reviewType) => {
          const key = String(reviewType?.key || "");
          return item.error_message || checks[key] || "";
        }),
      ];
    } else {
      cells = [
        item.source_text || "",
        item.target_text || "",
        item.has_issue === null ? "" : item.has_issue ? "是" : "否",
        item.issue_type || "",
        item.error_message || item.issue || "",
        item.suggestion || "",
      ];
    }
    if (hasForbidden && !isForbiddenOnly) {
      cells.push(item.matched_words || "");
    }
    cells.forEach((value) => {
      const td = document.createElement("td");
      td.textContent = String(value || "");
      tr.appendChild(td);
    });
    body.appendChild(tr);
  });
}

function renderAiReviewResultHead(task) {
  const config = task?.config || {};
  const isDirectional = config.mode === "directional";
  const isForbiddenOnly = config.mode === "forbidden_only";
  const hasForbidden = Boolean(config.enable_forbidden_check);
  const reviewTypes = Array.isArray(config.review_types) ? config.review_types : [];
  const headers = isForbiddenOnly
    ? ["原文", "译文", "禁用词检查情况"]
    : isDirectional
      ? ["原文", "译文", "修改建议", ...reviewTypes.map((item) => String(item?.key || ""))]
      : ["原文", "译文", "是否有问题", "问题类型", "问题说明", "修改建议"];
  if (hasForbidden && !isForbiddenOnly) {
    headers.push("禁用词检查情况");
  }
  const table = $("reviewResultBody").closest("table");
  table.style.minWidth = "0";
  table.style.width = "100%";
  table.querySelector("colgroup")?.remove();
  const head = $("reviewResultHead");
  head.innerHTML = "";
  const tr = document.createElement("tr");
  headers.forEach((header) => {
    const th = document.createElement("th");
    th.textContent = header;
    tr.appendChild(th);
  });
  head.appendChild(tr);
}

function reviewDetailExtraHeaders(task) {
  return ["问题类型", "问题说明"];
}

function reviewDetailExtraValues(task, item) {
  const config = task?.config || {};
  const isDirectional = config.mode === "directional";
  const isForbiddenOnly = config.mode === "forbidden_only";
  const hasForbidden = Boolean(config.enable_forbidden_check);
  if (isForbiddenOnly) {
    return ["禁用词", item.matched_words || ""];
  }
  if (isDirectional) {
    const checks = item.checks || {};
    const reviewTypes = Array.isArray(config.review_types) ? config.review_types : [];
    const issues = reviewTypes.map((reviewType) => {
      const key = String(reviewType?.key || "");
      const value = String(item.error_message || checks[key] || "").trim();
      return value ? { key, value } : null;
    }).filter(Boolean);
    if (hasForbidden && String(item.matched_words || "").trim()) {
      issues.push({ key: "禁用词", value: String(item.matched_words || "").trim() });
    }
    return [
      issues.map((issue) => issue.key).join("\n"),
      issues.map((issue) => `${issue.key}：${issue.value}`).join("\n"),
    ];
  }
  const types = [];
  const descriptions = [];
  if (item.issue_type || item.error_message || item.issue) {
    types.push(item.issue_type || "审校问题");
    descriptions.push(item.error_message || item.issue || "");
  }
  if (hasForbidden && String(item.matched_words || "").trim()) {
    types.push("禁用词");
    descriptions.push(`禁用词：${String(item.matched_words || "").trim()}`);
  }
  return [types.join("\n"), descriptions.join("\n")];
}

function createReviewDiffCell(html, mode) {
  const td = document.createElement("td");
  td.className = "review-diff-cell";
  const sideClass = mode === "add" ? "diff-inline-add" : "diff-inline-delete";
  const label = mode === "add" ? "建议" : "原译文";
  td.innerHTML = `
    <div class="diff-inline-side ${sideClass}">
      <div class="diff-inline-label">${label}</div>
      <div class="diff-inline-text">${html}</div>
    </div>`;
  return td;
}

function renderReviewDetailRows(task, results) {
  const groups = $("reviewDetailGroups");
  const extraHeaders = reviewDetailExtraHeaders(task);
  const headers = ["原文", "原译文", "修改建议", ...extraHeaders];
  const items = Array.isArray(results) ? results : [];
  aiReviewIssueResults = items;
  const grouped = new Map();
  items.forEach((item) => {
    const filename = String(item?.source_file || "直接输入").trim() || "直接输入";
    if (!grouped.has(filename)) grouped.set(filename, []);
    grouped.get(filename).push(item);
  });
  $("reviewDetailSummary").textContent = items.length
    ? `共 ${items.length} 条问题 · ${grouped.size} 个文件`
    : "暂无问题条目";
  groups.innerHTML = "";
  if (!items.length) {
    groups.innerHTML = '<div class="empty-cell">暂无问题条目</div>';
    return;
  }
  grouped.forEach((fileItems, filename) => {
    const section = document.createElement("section");
    section.className = "review-detail-group";
    const groupHead = document.createElement("div");
    groupHead.className = "review-detail-group-head";
    const title = document.createElement("h4");
    title.textContent = filename;
    const count = document.createElement("span");
    count.textContent = `${fileItems.length} 条问题`;
    groupHead.append(title, count);
    section.appendChild(groupHead);

    const tableWrap = document.createElement("div");
    tableWrap.className = "review-detail-group-table-wrap";
    const table = document.createElement("table");
    table.className = "pattern-table review-detail-table";
    const head = document.createElement("thead");
    const headRow = document.createElement("tr");
    headers.forEach((header) => {
      const th = document.createElement("th");
      th.textContent = header;
      headRow.appendChild(th);
    });
    head.appendChild(headRow);
    const body = document.createElement("tbody");
    fileItems.forEach((item) => {
      const tr = document.createElement("tr");
      const sourceTd = document.createElement("td");
      sourceTd.textContent = String(item.source_text || "");
      tr.appendChild(sourceTd);

      const tokens = buildDiffTokens(item.target_text || "", item.suggestion || "");
      tr.appendChild(createReviewDiffCell(tokens.leftHtml, "delete"));
      tr.appendChild(createReviewDiffCell(tokens.rightHtml, "add"));

      const extraValues = reviewDetailExtraValues(task, item);
      extraValues.forEach((value, index) => {
        const td = document.createElement("td");
        if (index === extraValues.length - 1) td.className = "review-detail-issue-cell";
        td.textContent = String(value || "");
        if (index === extraValues.length - 1) {
          const followupButton = document.createElement("button");
          followupButton.type = "button";
          followupButton.className = "secondary review-followup-inline-button";
          followupButton.textContent = "追问";
          followupButton.addEventListener("click", () => openReviewFollowupDialog(item.id).catch((error) => {
            $("reviewFollowupHint").textContent = error.message;
          }));
          td.appendChild(followupButton);
        }
        tr.appendChild(td);
      });
      body.appendChild(tr);
    });
    table.append(head, body);
    tableWrap.appendChild(table);
    section.append(tableWrap);
    groups.appendChild(section);
  });
}

async function openReviewDetailDialog() {
  if (!aiReviewTaskId) {
    $("reviewTaskHint").textContent = "请先开始审校任务";
    return;
  }
  const data = await api(`/api/ai-review/tasks/${encodeURIComponent(aiReviewTaskId)}/issue-results`);
  renderReviewDetailRows(data.task || aiReviewCurrentTask || {}, data.results || []);
  $("reviewDetailOverlay").hidden = false;
}

function closeReviewDetailDialog() {
  $("reviewDetailOverlay").hidden = true;
}

function reviewFollowupIssueText(item) {
  const task = aiReviewCurrentTask || {};
  const [issueType, issueText] = reviewDetailExtraValues(task, item || {});
  return { issueType, issueText };
}

function renderReviewFollowupContext(item) {
  const container = $("reviewFollowupContext");
  const issue = reviewFollowupIssueText(item || {});
  const sections = [
    ["原文", item?.source_text || ""],
    ["译文", item?.target_text || ""],
    ["修改建议", item?.suggestion || ""],
    ["问题类型", issue.issueType || ""],
    ["问题说明", issue.issueText || ""],
  ].filter(([, value]) => String(value || "").trim());
  container.innerHTML = `
    <div class="followup-context-card">
      ${sections.map(([label, value]) => `
        <div class="followup-context-section">
          <span>${escapeHtml(label)}</span>
          <p>${escapeHtml(String(value || ""))}</p>
        </div>
      `).join("")}
    </div>`;
}

function renderReviewFollowupMessages(messages) {
  const box = $("reviewFollowupMessages");
  const items = Array.isArray(messages) ? messages : [];
  box.innerHTML = "";
  if (!items.length) {
    box.innerHTML = '<div class="followup-message-empty">还没有追问记录，可以直接输入问题。</div>';
    return;
  }
  items.forEach((message) => {
    const div = document.createElement("div");
    const role = String(message.role || "") === "user" ? "user" : "assistant";
    div.className = `followup-message ${role}`;
    div.textContent = String(message.content || "");
    box.appendChild(div);
  });
  box.scrollTop = box.scrollHeight;
}

function renderReviewFollowupDialog(data) {
  const item = data?.item || aiReviewFollowupState.item || {};
  const messages = Array.isArray(data?.messages) ? data.messages : [];
  aiReviewFollowupState.item = item;
  aiReviewFollowupState.messages = messages;
  $("reviewFollowupSummary").textContent = item?.row_number ? `第 ${item.row_number} 行` : "围绕当前问题继续提问";
  renderReviewFollowupContext(item);
  renderReviewFollowupMessages(messages);
}

async function openReviewFollowupDialog(resultId) {
  if (!aiReviewTaskId || !resultId) return;
  aiReviewFollowupState = { taskId: aiReviewTaskId, resultId: String(resultId), item: null, messages: [] };
  $("reviewFollowupInput").value = "";
  $("reviewFollowupHint").textContent = "正在读取对话...";
  $("reviewFollowupOverlay").hidden = false;
  try {
    const data = await api(`/api/ai-review/tasks/${encodeURIComponent(aiReviewTaskId)}/followup/${encodeURIComponent(resultId)}`);
    renderReviewFollowupDialog(data);
    $("reviewFollowupHint").textContent = "";
    $("reviewFollowupInput").focus();
  } catch (error) {
    $("reviewFollowupHint").textContent = error.message || "追问读取失败";
  }
}

function closeReviewFollowupDialog() {
  $("reviewFollowupOverlay").hidden = true;
}

async function sendReviewFollowupMessage() {
  const message = $("reviewFollowupInput").value.trim();
  if (!message) {
    $("reviewFollowupHint").textContent = "请输入追问内容";
    return;
  }
  const taskId = aiReviewFollowupState.taskId || aiReviewTaskId;
  const resultId = aiReviewFollowupState.resultId;
  if (!taskId || !resultId) {
    $("reviewFollowupHint").textContent = "当前没有可追问的条目";
    return;
  }
  $("sendReviewFollowupButton").disabled = true;
  $("reviewFollowupHint").textContent = "发送中...";
  try {
    const data = await api(`/api/ai-review/tasks/${encodeURIComponent(taskId)}/followup/${encodeURIComponent(resultId)}`, {
      method: "POST",
      body: JSON.stringify({ message }),
    });
    $("reviewFollowupInput").value = "";
    renderReviewFollowupDialog(data);
    $("reviewFollowupHint").textContent = "";
  } finally {
    $("sendReviewFollowupButton").disabled = false;
  }
}

function appendAiReviewLogs(logs) {
  const list = $("reviewLogList");
  const items = Array.isArray(logs) ? logs : [];
  items.forEach((item) => {
    aiReviewLogCursor = Math.max(aiReviewLogCursor, Number(item.id || 0));
    const li = document.createElement("li");
    li.textContent = `${String(item.created_at || "")} ${String(item.message || "")}`.trim();
    if (String(item.level || "").toLowerCase() === "error") {
      li.classList.add("error");
    } else if (String(item.level || "").toLowerCase() === "debug") {
      li.classList.add("debug");
    }
    list.appendChild(li);
  });
}

async function loadAiReviewPromptTemplates() {
  const data = await api("/api/ai-review/prompt-templates");
  aiReviewPromptTemplates = Array.isArray(data.templates) ? data.templates : [];
  renderAiReviewPromptTemplateOptions();
}

async function loadAiReviewDirectionalTemplates() {
  const data = await api("/api/ai-review/directional-templates");
  aiReviewDirectionalTemplates = Array.isArray(data.templates) ? data.templates : [];
  renderAiReviewDirectionalTemplateOptions();
}

async function loadAiReviewForbiddenTemplates() {
  const data = await api("/api/ai-review/forbidden-templates");
  aiReviewForbiddenTemplates = Array.isArray(data.templates) ? data.templates : [];
  renderAiReviewForbiddenTemplateOptions();
}

async function loadReviewConversations(preferredId = "") {
  const data = await api("/api/ai-review/conversations");
  reviewConversationState.sessions = Array.isArray(data.sessions) ? data.sessions : [];
  renderReviewConversationList();
  let nextId = preferredId || reviewConversationState.currentId || reviewConversationState.sessions[0]?.id || "";
  if (!nextId) {
    const created = await createReviewConversation();
    nextId = created?.id || "";
  }
  if (nextId) await openReviewConversation(nextId);
}

async function loadReviewTermBases() {
  const data = await api("/api/ai-review/term-bases");
  reviewConversationState.termBases = Array.isArray(data.term_bases) ? data.term_bases : [];
  renderReviewTermBaseOptions();
  updateReviewComposerChips();
}

async function createReviewConversation() {
  const promptId = reviewConversationState.promptTemplateId || aiReviewPromptTemplates[0]?.id || null;
  const data = await api("/api/ai-review/conversations", {
    method: "POST",
    body: JSON.stringify({
      title: "新审校",
      prompt_template_id: promptId,
      term_base_id: null,
      source_language: "auto",
      target_languages: ["auto"],
      auto_start: false,
    }),
  });
  const session = data.session || {};
  reviewConversationState.currentId = String(session.id || "");
  reviewConversationState.sessions.unshift(session);
  renderReviewConversationList();
  if (session.id) await openReviewConversation(session.id);
  return session;
}

function renderReviewConversationList() {
  const list = $("reviewConversationList");
  if (!list) return;
  list.innerHTML = "";
  reviewConversationState.sessions.forEach((session) => {
    const row = document.createElement("div");
    row.className = `review-session-row ${session.id === reviewConversationState.currentId ? "active" : ""}`.trim();
    const button = document.createElement("button");
    button.type = "button";
    button.className = "review-session-item";
    const title = document.createElement("strong");
    title.textContent = session.title || "新审校";
    const status = document.createElement("span");
    status.textContent = reviewSessionStatusLabel(session.status);
    button.append(title, status);
    button.addEventListener("click", () => openReviewConversation(session.id).catch(showReviewConversationError));
    const renameButton = document.createElement("button");
    renameButton.type = "button";
    renameButton.className = "review-session-rename";
    renameButton.title = "重命名会话";
    renameButton.setAttribute("aria-label", `重命名 ${session.title || "新审校"}`);
    renameButton.textContent = "✎";
    renameButton.addEventListener("click", () => beginRenameReviewSessionItem(session, row, button, renameButton));
    row.append(button, renameButton);
    list.appendChild(row);
  });
}

function beginRenameReviewSessionItem(session, row, sessionButton, renameButton) {
  const input = document.createElement("input");
  input.className = "review-session-inline-input";
  input.type = "text";
  input.maxLength = 120;
  input.value = session.title || "新审校";
  sessionButton.replaceWith(input);
  const saveButton = renameButton.cloneNode(true);
  saveButton.textContent = "✓";
  saveButton.title = "保存名称";
  saveButton.setAttribute("aria-label", "保存会话名称");
  renameButton.replaceWith(saveButton);
  const save = async () => {
    const title = input.value.trim();
    if (!title) throw new Error("会话名称不能为空");
    await renameReviewConversation(session.id, title);
  };
  saveButton.addEventListener("click", () => save().catch(showReviewConversationError));
  input.addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      save().catch(showReviewConversationError);
    } else if (event.key === "Escape") {
      event.preventDefault();
      renderReviewConversationList();
    }
  });
  input.focus();
  input.select();
}

async function renameReviewConversation(sessionId, title) {
  const data = await api(`/api/ai-review/conversations/${encodeURIComponent(sessionId)}`, {
    method: "PATCH",
    body: JSON.stringify({ title }),
  });
  const updated = data.session || {};
  const sessionIndex = reviewConversationState.sessions.findIndex((item) => item.id === sessionId);
  if (sessionIndex >= 0) reviewConversationState.sessions[sessionIndex] = updated;
  if (reviewConversationState.snapshot?.session?.id === sessionId) {
    reviewConversationState.snapshot.session = updated;
    renderReviewConversationSnapshot();
  }
  renderReviewConversationList();
  return updated;
}

async function saveReviewConversationAutoStart(enabled) {
  const sessionId = String(reviewConversationState.currentId || "");
  if (!sessionId) return;
  const data = await api(`/api/ai-review/conversations/${encodeURIComponent(sessionId)}`, {
    method: "PATCH",
    body: JSON.stringify({ auto_start: Boolean(enabled) }),
  });
  const updated = data.session || {};
  const sessionIndex = reviewConversationState.sessions.findIndex((item) => item.id === sessionId);
  if (sessionIndex >= 0) reviewConversationState.sessions[sessionIndex] = updated;
  if (reviewConversationState.snapshot?.session?.id === sessionId) {
    reviewConversationState.snapshot.session = updated;
  }
  renderReviewConversationList();
}

function applyReviewConversationSessionUpdate(updated) {
  if (!updated?.id) return;
  const sessionIndex = reviewConversationState.sessions.findIndex((item) => item.id === updated.id);
  if (sessionIndex >= 0) reviewConversationState.sessions[sessionIndex] = updated;
  if (reviewConversationState.snapshot?.session?.id === updated.id) {
    reviewConversationState.snapshot.session = updated;
  }
  renderReviewConversationList();
}

function saveReviewConversationComposerSettings() {
  const sessionId = String(reviewConversationState.currentId || "");
  if (!sessionId) return Promise.resolve();
  const payload = {
    prompt_template_id: reviewConversationState.promptTemplateId || null,
    source_language: reviewConversationState.sourceLanguage || "auto",
    target_languages: [...(reviewConversationState.targetLanguages || ["auto"])],
    term_base_id: reviewConversationState.termBaseId || null,
  };
  const save = async () => {
    const data = await api(`/api/ai-review/conversations/${encodeURIComponent(sessionId)}`, {
      method: "PATCH",
      body: JSON.stringify(payload),
    });
    applyReviewConversationSessionUpdate(data.session || {});
  };
  // Language chips can be clicked rapidly. Serialize writes so a slower earlier
  // response cannot overwrite the newest selection for this session.
  reviewConversationState.composerSaveQueue = reviewConversationState.composerSaveQueue
    .catch(() => undefined)
    .then(save);
  return reviewConversationState.composerSaveQueue;
}

async function openReviewConversation(sessionId) {
  if (!sessionId) return;
  reviewConversationState.currentId = String(sessionId);
  reviewConversationState.streamRunId = "";
  reviewConversationState.streamText = "";
  reviewConversationState.streamThinkingText = "";
  reviewConversationState.streamAnswerText = "";
  reviewConversationState.streamPhase = "reasoning";
  reviewConversationState.streamActive = false;
  reviewConversationState.eventCursor = 0;
  const snapshot = await api(`/api/ai-review/conversations/${encodeURIComponent(sessionId)}`);
  reviewConversationState.snapshot = snapshot;
  const session = snapshot.session || {};
  reviewConversationState.sourceLanguage = session.source_language || "auto";
  reviewConversationState.targetLanguages = Array.isArray(session.target_languages) && session.target_languages.length ? session.target_languages : ["auto"];
  reviewConversationState.promptTemplateId = session.prompt_template_id || aiReviewPromptTemplates[0]?.id || "";
  reviewConversationState.termBaseId = session.term_base_id || "";
  reviewConversationState.activeTarget = reviewConversationState.activeTarget || snapshot.task_results?.[0]?.target_language || "";
  renderReviewConversationList();
  renderReviewConversationSnapshot();
  connectReviewConversationEvents(sessionId);
}

function renderReviewConversationSnapshot() {
  const snapshot = reviewConversationState.snapshot || {};
  const session = snapshot.session || {};
  $("reviewConversationTitle").textContent = session.title || "新审校";
  $("reviewConversationStatus").textContent = reviewSessionStatusLabel(session.status);
  $("reviewAutoStart").checked = Boolean(session.auto_start);
  const liveHints = {
    inspecting: "Workspace Agent 正在识别内容…",
    reviewing: "Task Workflow 正在审校…",
    validating: "正在校验审校结果…",
  };
  $("reviewConversationHint").textContent = liveHints[session.status] || "";
  updateReviewComposerChips();
  renderReviewAttachments(snapshot.attachments || []);
  renderReviewMessages(snapshot.messages || [], snapshot.questions || []);
  renderReviewConversationResults(snapshot.task_results || []);
}

function renderWorkspaceReportPreview(message) {
  const plan = message.payload || {};
  const files = Array.isArray(plan.file_summaries) ? plan.file_summaries : [];
  if (!files.length) return null;
  const totalUnits = (plan.targets || []).reduce(
    (total, target) => total + (
      Array.isArray(target?.units) ? target.units.length : Number(target?.unit_count || 0)
    ), 0,
  );
  const preview = document.createElement("section");
  preview.className = "review-workspace-preview";
  const previewHead = document.createElement("div");
  previewHead.className = "review-workspace-preview-head";
  const headCopy = document.createElement("div");
  const title = document.createElement("strong");
  title.textContent = "识别预览";
  const subtitle = document.createElement("span");
  subtitle.textContent = "请核对列映射与同一行样例";
  headCopy.append(title, subtitle);
  const count = document.createElement("span");
  count.className = "review-workspace-total";
  count.textContent = `${totalUnits} 条待审校`;
  previewHead.append(headCopy, count);
  preview.appendChild(previewHead);

  files.forEach((file) => {
    const fileCard = document.createElement("section");
    fileCard.className = "review-workspace-file-card";
    const fileHead = document.createElement("div");
    fileHead.className = "review-workspace-file-head";
    const fileIcon = document.createElement("span");
    fileIcon.className = "review-workspace-file-icon";
    fileIcon.textContent = String(file.file_type || "file").slice(0, 4).toUpperCase();
    const fileCopy = document.createElement("div");
    const filename = document.createElement("strong");
    filename.textContent = file.filename || "未命名文件";
    const fileMeta = document.createElement("span");
    fileMeta.textContent = file.mappings?.length ? `${file.mappings.length} 组映射` : (file.structure_label || "未形成审校映射");
    fileCopy.append(filename, fileMeta);
    fileHead.append(fileIcon, fileCopy);
    fileCard.appendChild(fileHead);

    const mappings = Array.isArray(file.mappings) ? file.mappings : [];
    if (!mappings.length) {
      const empty = document.createElement("div");
      empty.className = "review-workspace-empty";
      empty.textContent = file.structure_label || "尚未形成可执行映射";
      fileCard.appendChild(empty);
    }
    mappings.forEach((mapping) => {
      const mappingCard = document.createElement("div");
      mappingCard.className = "review-workspace-mapping";
      const mappingHead = document.createElement("div");
      mappingHead.className = "review-workspace-mapping-head";
      const language = document.createElement("span");
      language.className = "review-workspace-language";
      language.textContent = mapping.language || "自动识别";
      const mappingCount = document.createElement("span");
      mappingCount.textContent = `${mapping.count || 0} 条`;
      mappingHead.append(language, mappingCount);
      const flow = document.createElement("div");
      flow.className = "review-workspace-flow";
      const makeEndpoint = (label, value, muted = false) => {
        const endpoint = document.createElement("div");
        endpoint.className = `review-workspace-endpoint${muted ? " is-empty" : ""}`;
        const endpointLabel = document.createElement("span");
        endpointLabel.textContent = label;
        const endpointValue = document.createElement("strong");
        endpointValue.textContent = value || (label === "原文" ? "无源文" : "未识别");
        endpoint.append(endpointLabel, endpointValue);
        return endpoint;
      };
      flow.append(
        makeEndpoint("原文", mapping.source_location, !mapping.source_location),
        Object.assign(document.createElement("span"), { className: "review-workspace-arrow", textContent: "→" }),
        makeEndpoint("译文", mapping.target_location),
      );
      mappingCard.append(mappingHead, flow);
      if (mapping.reference_locations) {
        const reference = document.createElement("div");
        reference.className = "review-workspace-reference";
        reference.textContent = `同一行上下文：${mapping.reference_locations}`;
        mappingCard.appendChild(reference);
      }
      const samples = Array.isArray(mapping.samples) ? mapping.samples : [];
      if (samples.length) {
        const samplesGrid = document.createElement("div");
        samplesGrid.className = "review-workspace-samples";
        samples.slice(0, 3).forEach((sample, index) => {
          const sampleCard = document.createElement("article");
          sampleCard.className = "review-workspace-sample";
          const sampleMeta = document.createElement("span");
          sampleMeta.textContent = `样例 ${index + 1}${sample.position ? ` · ${sample.position}` : ""}`;
          const values = document.createElement("div");
          values.className = "review-workspace-sample-values";
          const source = document.createElement("div");
          source.className = "review-workspace-sample-source";
          source.textContent = sample.source || "无源文";
          const target = document.createElement("div");
          target.className = "review-workspace-sample-target";
          target.textContent = sample.target || "（空译文）";
          values.append(source, target);
          sampleCard.append(sampleMeta, values);
          samplesGrid.appendChild(sampleCard);
        });
        mappingCard.appendChild(samplesGrid);
      }
      fileCard.appendChild(mappingCard);
    });
    preview.appendChild(fileCard);
  });
  if ((plan.assumptions || []).length || (plan.warnings || []).length) {
  const notes = document.createElement("details");
  notes.className = "review-workspace-notes";
  notes.open = true;
    const summary = document.createElement("summary");
    summary.textContent = "查看识别依据与注意事项";
    const noteCopy = document.createElement("div");
    noteCopy.className = "review-workspace-note-copy";
    [...(plan.assumptions || []), ...(plan.warnings || [])].slice(0, 5).forEach((note) => {
      const row = document.createElement("p");
      row.textContent = note;
      noteCopy.appendChild(row);
    });
    notes.append(summary, noteCopy);
    preview.appendChild(notes);
  }
  return preview;
}

function renderReviewMessages(messages, questions) {
  const container = $("reviewConversationMessages");
  container.innerHTML = "";
  messages.forEach((message) => {
    const item = document.createElement("div");
    item.className = `review-chat-message ${message.role === "user" ? "user" : "assistant"} ${message.kind === "error" ? "error" : ""} ${message.kind === "workspace_report" ? "workspace-report" : ""}`.trim();
    const workspacePreview = message.kind === "workspace_report" ? renderWorkspaceReportPreview(message) : null;
    if (workspacePreview) {
      item.appendChild(workspacePreview);
    } else {
      const content = document.createElement("div");
      content.textContent = message.content || "";
      item.appendChild(content);
    }
    const attachments = Array.isArray(message.payload?.attachments) ? message.payload.attachments : [];
    if (attachments.length) {
      const files = document.createElement("div");
      files.className = "review-message-files";
      attachments.forEach((attachment) => {
        const file = document.createElement("span");
        file.className = "review-message-file";
        file.textContent = `▤ ${attachment.original_filename || "文件"}`;
        files.appendChild(file);
      });
      item.appendChild(files);
    }
    const thinkingText = String(message.payload?._thinking_text || "");
    if (message.kind === "workspace_report" && thinkingText) {
      const details = document.createElement("details");
      details.className = "review-thinking-details";
      const summary = document.createElement("summary");
      summary.textContent = `查看思考过程 · ${thinkingText.length} 字`;
      const thinkingOutput = document.createElement("div");
      thinkingOutput.className = "review-stream-output";
      thinkingOutput.textContent = thinkingText;
      details.append(summary, thinkingOutput);
      item.appendChild(details);
    }
    const meta = document.createElement("div");
    meta.className = "review-chat-message-meta";
    meta.textContent = String(message.created_at || "").replace("T", " ");
    item.appendChild(meta);
    container.appendChild(item);
  });
  if (reviewConversationState.streamActive) {
    const item = document.createElement("div");
    item.className = "review-chat-message assistant workspace-report";
    const title = document.createElement("strong");
    title.textContent = reviewConversationState.streamText
      ? `${reviewConversationState.streamPhase === "reasoning" ? "Workspace Agent 正在思考" : "Workspace Agent 正在生成方案"} · 已接收 ${reviewConversationState.streamText.length} 字`
      : "Workspace Agent 已连接，等待首个输出…";
    const details = document.createElement("details");
    details.className = "review-thinking-details";
    details.open = true;
    const summary = document.createElement("summary");
    summary.textContent = reviewConversationState.streamPhase === "reasoning" ? "实时思考过程" : "实时模型输出";
    const output = document.createElement("div");
    output.className = "review-stream-output";
    const liveText = reviewConversationState.streamPhase === "reasoning"
      ? reviewConversationState.streamThinkingText
      : reviewConversationState.streamAnswerText;
    output.textContent = liveText.slice(-5000) || "正在分析文件结构…";
    details.append(summary, output);
    item.append(title, details);
    container.appendChild(item);
    window.requestAnimationFrame(() => {
      output.scrollTop = output.scrollHeight;
      container.scrollTop = container.scrollHeight;
    });
  }
  const latestRun = reviewConversationState.snapshot?.workspace_runs?.[0] || null;
  const latestRunId = String(latestRun?.id || "");
  const hasExecutablePlan = Boolean((latestRun?.plan?.targets || []).some((target) => Array.isArray(target?.units) && target.units.length));
  questions.filter((question) => question.status === "pending" && (!latestRunId || question.run_id === latestRunId)).forEach((question) => {
    const item = document.createElement("div");
    item.className = "review-chat-message assistant";
    const prompt = document.createElement("div");
    prompt.textContent = question.prompt || "请确认识别结果";
    const actions = document.createElement("div");
    actions.className = "review-decision-actions";
    const canConfirm = hasExecutablePlan && Boolean(String(question.recommended_label || "").trim()) && Boolean(String(question.recommended_value || "").trim());
    if (canConfirm) {
      const confirmRow = document.createElement("div");
      confirmRow.className = "review-decision-confirm-row";
      const confirmButton = document.createElement("button");
      confirmButton.type = "button";
      confirmButton.className = "primary";
      confirmButton.textContent = question.recommended_label;
      confirmButton.addEventListener("click", () => submitReviewDecision(question, "confirm", question.recommended_value).catch(showReviewConversationError));
      confirmRow.appendChild(confirmButton);
      actions.appendChild(confirmRow);
    }
    const otherInput = document.createElement("input");
    otherInput.placeholder = canConfirm ? "其他：自行输入" : "请补充目标语种、译文内容或文件位置";
    const otherButton = document.createElement("button");
    otherButton.type = "button";
    otherButton.className = "secondary";
    otherButton.textContent = "提交";
    otherButton.addEventListener("click", () => submitReviewDecision(question, "other", otherInput.value).catch(showReviewConversationError));
    const otherRow = document.createElement("div");
    otherRow.className = "review-decision-other-row";
    otherRow.append(otherInput, otherButton);
    actions.appendChild(otherRow);
    item.append(prompt, actions);
    container.appendChild(item);
  });
  container.scrollTop = container.scrollHeight;
}

async function submitReviewDecision(question, action, answer) {
  if (action === "other" && !String(answer || "").trim()) throw new Error("请输入补充说明");
  await api(`/api/ai-review/conversations/${encodeURIComponent(reviewConversationState.currentId)}/decision`, {
    method: "POST",
    body: JSON.stringify({ run_id: question.run_id, question_id: question.id, action, answer: answer || action }),
  });
  await refreshCurrentReviewConversation();
}

function renderReviewAttachments(attachments) {
  const container = $("reviewAttachmentChips");
  container.innerHTML = "";
  attachments.filter((attachment) => !attachment.sent_at).forEach((attachment) => {
    const chip = document.createElement("span");
    chip.className = "review-attachment-chip";
    const openButton = document.createElement("button");
    openButton.type = "button";
    openButton.className = "review-attachment-open";
    const modeLabel = attachment.mapping_mode === "preset" ? "映射模板" : "AI 识别";
    openButton.textContent = `${attachment.original_filename} · ${modeLabel}`;
    openButton.addEventListener("click", () => openReviewAttachmentMappingDialog(attachment).catch(showReviewConversationError));
    const removeButton = document.createElement("button");
    removeButton.type = "button";
    removeButton.className = "review-attachment-remove";
    removeButton.title = "移除文件";
    removeButton.textContent = "×";
    removeButton.addEventListener("click", () => removeReviewPendingAttachment(attachment.id).catch(showReviewConversationError));
    chip.append(openButton, removeButton);
    container.appendChild(chip);
  });
}

function renderReviewConversationResults(taskResults) {
  const card = $("reviewConversationResultCard");
  const tasks = Array.isArray(taskResults) ? taskResults : [];
  card.classList.toggle("hidden", !tasks.length);
  const tabs = $("reviewTargetTabs");
  tabs.innerHTML = "";
  if (!tasks.length) {
    reviewConversationState.activeTarget = "";
    aiReviewTaskId = "";
    aiReviewCurrentTask = null;
    renderAiReviewResults({}, []);
    return;
  }
  if (!tasks.some((item) => item.target_language === reviewConversationState.activeTarget)) {
    reviewConversationState.activeTarget = tasks[0].target_language;
  }
  tasks.forEach((item) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = `review-target-tab ${item.target_language === reviewConversationState.activeTarget ? "active" : ""}`.trim();
    button.textContent = item.target_language || "自动识别";
    button.addEventListener("click", () => {
      reviewConversationState.activeTarget = item.target_language;
      renderReviewConversationResults(tasks);
    });
    tabs.appendChild(button);
  });
  const active = tasks.find((item) => item.target_language === reviewConversationState.activeTarget) || tasks[0];
  aiReviewTaskId = String(active?.task?.id || "");
  aiReviewCurrentTask = active?.task || null;
  renderAiReviewResults(active?.task || {}, active?.results || []);
}

async function refreshCurrentReviewConversation() {
  if (!reviewConversationState.currentId) return;
  const snapshot = await api(`/api/ai-review/conversations/${encodeURIComponent(reviewConversationState.currentId)}`);
  reviewConversationState.snapshot = snapshot;
  const sessionIndex = reviewConversationState.sessions.findIndex((item) => item.id === snapshot.session?.id);
  if (sessionIndex >= 0) reviewConversationState.sessions[sessionIndex] = snapshot.session;
  renderReviewConversationSnapshot();
}

function connectReviewConversationEvents(sessionId) {
  if (reviewConversationState.eventSource) reviewConversationState.eventSource.close();
  const source = new EventSource(`/api/ai-review/conversations/${encodeURIComponent(sessionId)}/events?after_id=${reviewConversationState.eventCursor}`);
  reviewConversationState.eventSource = source;
  const eventNames = [
    "message.created", "attachment.uploaded", "attachment.deleted", "attachment.inspected", "workspace.started", "workspace.ready",
    "workspace.needs_input", "workspace.failed", "workspace.question", "review.started", "review.progress",
    "review.request_status", "review.package_completed", "review.completed",
  ];
  eventNames.forEach((name) => source.addEventListener(name, (event) => {
    reviewConversationState.eventCursor = Math.max(reviewConversationState.eventCursor, Number(event.lastEventId || 0));
    if (["workspace.ready", "workspace.needs_input", "workspace.failed"].includes(name)) {
      reviewConversationState.streamActive = false;
      reviewConversationState.streamText = "";
      reviewConversationState.streamThinkingText = "";
      reviewConversationState.streamAnswerText = "";
    }
    if (name === "workspace.started") renderReviewConversationResults([]);
    scheduleReviewConversationRefresh();
  }));
  source.addEventListener("workspace.output_started", (event) => {
    reviewConversationState.eventCursor = Math.max(reviewConversationState.eventCursor, Number(event.lastEventId || 0));
    const payload = JSON.parse(event.data || "{}").payload || {};
    reviewConversationState.streamRunId = String(payload.run_id || "");
    reviewConversationState.streamText = "";
    reviewConversationState.streamThinkingText = "";
    reviewConversationState.streamAnswerText = "";
    reviewConversationState.streamPhase = "reasoning";
    reviewConversationState.streamActive = true;
    renderReviewMessages(reviewConversationState.snapshot?.messages || [], reviewConversationState.snapshot?.questions || []);
  });
  source.addEventListener("workspace.delta", (event) => {
    reviewConversationState.eventCursor = Math.max(reviewConversationState.eventCursor, Number(event.lastEventId || 0));
    const payload = JSON.parse(event.data || "{}").payload || {};
    if (reviewConversationState.streamRunId && payload.run_id !== reviewConversationState.streamRunId) return;
    reviewConversationState.streamRunId = String(payload.run_id || reviewConversationState.streamRunId || "");
    reviewConversationState.streamActive = true;
    const delta = String(payload.delta || "");
    const phase = payload.phase === "content" ? "content" : "reasoning";
    reviewConversationState.streamPhase = phase;
    reviewConversationState.streamText += delta;
    if (phase === "reasoning") reviewConversationState.streamThinkingText += delta;
    else reviewConversationState.streamAnswerText += delta;
    renderReviewMessages(reviewConversationState.snapshot?.messages || [], reviewConversationState.snapshot?.questions || []);
  });
  source.onerror = () => {
    if (reviewConversationState.currentId !== sessionId) source.close();
  };
}

function scheduleReviewConversationRefresh() {
  if (reviewConversationState.refreshTimer) window.clearTimeout(reviewConversationState.refreshTimer);
  reviewConversationState.refreshTimer = window.setTimeout(() => refreshCurrentReviewConversation().catch(showReviewConversationError), 180);
}

async function uploadReviewConversationFiles(files) {
  if (!files?.length) return;
  if (!reviewConversationState.currentId) await createReviewConversation();
  const form = new FormData();
  Array.from(files).forEach((file) => form.append("files", file));
  $("reviewConversationHint").textContent = `正在添加 ${files.length} 个文件…`;
  const response = await fetch(`/api/ai-review/conversations/${encodeURIComponent(reviewConversationState.currentId)}/attachments`, { method: "POST", body: form });
  const data = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error(data.detail || data.message || "文件上传失败");
  $("reviewConversationHint").textContent = "文件已添加";
  await refreshCurrentReviewConversation();
}

async function chooseReviewConversationFiles() {
  const data = await api("/api/dialog/select-review-files");
  const paths = Array.isArray(data.file_paths) ? data.file_paths.filter(Boolean) : [];
  if (!paths.length) return;
  if (!reviewConversationState.currentId) await createReviewConversation();
  $("reviewConversationHint").textContent = `正在添加 ${paths.length} 个文件…`;
  for (const filePath of paths) {
    await api(`/api/ai-review/conversations/${encodeURIComponent(reviewConversationState.currentId)}/attachments/local`, {
      method: "POST",
      body: JSON.stringify({ file_path: filePath }),
    });
  }
  $("reviewConversationHint").textContent = "文件已添加，可在对话中溯源打开原文件";
  await refreshCurrentReviewConversation();
}

async function sendReviewConversationMessage() {
  if (!reviewConversationState.currentId) await createReviewConversation();
  const text = $("reviewComposerInput").value.trim();
  const attachments = (reviewConversationState.snapshot?.attachments || []).filter((item) => !item.sent_at);
  if (!text && !attachments.length) throw new Error("请输入待审校文本或添加文件");
  $("sendReviewConversationButton").disabled = true;
  $("reviewConversationHint").textContent = "Workspace Agent 正在识别内容…";
  renderReviewConversationResults([]);
  try {
    await api(`/api/ai-review/conversations/${encodeURIComponent(reviewConversationState.currentId)}/messages`, {
      method: "POST",
      body: JSON.stringify({
        text,
        prompt_template_id: reviewConversationState.promptTemplateId || null,
        term_base_id: reviewConversationState.termBaseId || null,
        source_language: reviewConversationState.sourceLanguage,
        target_languages: reviewConversationState.targetLanguages,
        auto_start: $("reviewAutoStart").checked,
      }),
    });
    $("reviewComposerInput").value = "";
    await refreshCurrentReviewConversation();
  } finally {
    $("sendReviewConversationButton").disabled = false;
  }
}

function beginRenameCurrentReviewConversation() {
  const input = $("reviewConversationTitleInput");
  if (!input.classList.contains("hidden")) {
    saveCurrentReviewConversationTitle().catch(showReviewConversationError);
    return;
  }
  input.value = reviewConversationState.snapshot?.session?.title || "新审校";
  $("reviewConversationTitle").classList.add("hidden");
  input.classList.remove("hidden");
  $("renameReviewConversationButton").textContent = "✓";
  input.focus();
  input.select();
}

function cancelRenameCurrentReviewConversation() {
  $("reviewConversationTitleInput").classList.add("hidden");
  $("reviewConversationTitle").classList.remove("hidden");
  $("renameReviewConversationButton").textContent = "✎";
}

async function saveCurrentReviewConversationTitle() {
  const title = $("reviewConversationTitleInput").value.trim();
  if (!title) throw new Error("会话名称不能为空");
  await renameReviewConversation(reviewConversationState.currentId, title);
  cancelRenameCurrentReviewConversation();
}

async function removeReviewPendingAttachment(attachmentId) {
  await api(`/api/ai-review/conversations/${encodeURIComponent(reviewConversationState.currentId)}/attachments/${encodeURIComponent(attachmentId)}`, {
    method: "DELETE",
  });
  await refreshCurrentReviewConversation();
}

async function openReviewAttachmentMappingDialog(attachment) {
  reviewConversationState.mappingAttachmentId = String(attachment.id || "");
  $("reviewAttachmentMappingTitle").textContent = attachment.original_filename || "文件导入方式";
  const suffix = String(attachment.original_filename || "FILE").split(".").pop().toUpperCase().slice(0, 5);
  document.querySelector(".review-mapping-file-icon").textContent = suffix || "FILE";
  if (!aiReviewExcelMappingPresets.length) await loadAiReviewExcelMappingPresets();
  const select = $("reviewAttachmentPresetSelect");
  select.innerHTML = "";
  aiReviewExcelMappingPresets.forEach((preset) => {
    const option = document.createElement("option");
    option.value = preset.id;
    option.textContent = preset.name || "未命名模板";
    select.appendChild(option);
  });
  select.insertBefore(new Option("不套用模板，从当前结构开始", ""), select.firstChild);
  const requestedMode = attachment.mapping_mode === "preset" ? "preset" : "ai";
  document.querySelectorAll('input[name="reviewAttachmentMappingMode"]').forEach((input) => {
    input.checked = input.value === requestedMode;
  });
  select.value = attachment.mapping_preset_id || select.options[0]?.value || "";
  const isExcel = /\.(xlsx|xlsm)$/i.test(attachment.original_filename || "");
  const presetRadio = document.querySelector('input[name="reviewAttachmentMappingMode"][value="preset"]');
  presetRadio.disabled = !isExcel;
  if (!isExcel && requestedMode === "preset") {
    document.querySelector('input[name="reviewAttachmentMappingMode"][value="ai"]').checked = true;
  }
  updateReviewAttachmentMappingControls();
  $("reviewAttachmentMappingHint").textContent = isExcel ? "" : "该格式仅支持 AI 识别；映射模板只适用于 Excel。";
  $("reviewAttachmentMappingDialog").showModal();
}

function updateReviewAttachmentMappingControls() {
  const mode = document.querySelector('input[name="reviewAttachmentMappingMode"]:checked')?.value || "ai";
  const presetRadio = document.querySelector('input[name="reviewAttachmentMappingMode"][value="preset"]');
  document.querySelectorAll(".review-import-mode-option").forEach((option) => {
    const input = option.querySelector('input[name="reviewAttachmentMappingMode"]');
    option.classList.toggle("selected", Boolean(input?.checked));
    option.classList.toggle("disabled", Boolean(input?.disabled));
  });
  $("reviewAttachmentPresetField").classList.toggle("hidden", mode !== "preset");
  $("reviewAttachmentPresetSelect").disabled = mode !== "preset" || Boolean(presetRadio?.disabled);
  $("editReviewAttachmentMappingButton").disabled = mode !== "preset" || Boolean(presetRadio?.disabled);
}

async function openReviewAttachmentMappingEditor() {
  const attachmentId = reviewConversationState.mappingAttachmentId;
  if (!attachmentId) throw new Error("没有可编辑的附件");
  $("reviewAttachmentMappingHint").textContent = "正在读取工作表结构…";
  const data = await api(`/api/ai-review/conversations/${encodeURIComponent(reviewConversationState.currentId)}/attachments/${encodeURIComponent(attachmentId)}/structure`);
  const attachment = data.attachment || {};
  const structure = data.manifest?.structure || attachment.manifest?.structure || {};
  const sheets = Array.isArray(structure.sheets) ? structure.sheets : [];
  if (!sheets.length) throw new Error("没有读取到可映射的工作表结构");
  aiReviewSheetNames = sheets.map((sheet) => String(sheet.name || ""));
  aiReviewColumnsBySheet = Object.fromEntries(sheets.map((sheet) => [String(sheet.name || ""), Array.isArray(sheet.columns) ? sheet.columns : []]));
  aiReviewExcelMappingState = Object.fromEntries(aiReviewSheetNames.map((sheetName) => [sheetName, { sources: [], targets: {}, infos: {} }]));
  aiReviewActiveSheetName = aiReviewSheetNames[0] || "";
  aiReviewMappingTemplateIssues = [];
  reviewConversationState.mappingEditorAttachmentId = attachmentId;
  reviewConversationState.mappingEditorPresetId = $("reviewAttachmentPresetSelect").value || "";
  reviewConversationState.mappingEditorContext = "conversation";
  $("mappingSourceLanguageInput").value = reviewConversationState.sourceLanguage === "auto" ? "" : reviewConversationState.sourceLanguage;
  $("mappingTargetLanguageInput").value = reviewConversationState.targetLanguages.filter((value) => value !== "auto").join(", ");
  await loadAiReviewExcelMappingPresets(reviewConversationState.mappingEditorPresetId);
  if (reviewConversationState.mappingEditorPresetId) {
    const presetData = await api(`/api/ai-review/excel-mapping-presets/${encodeURIComponent(reviewConversationState.mappingEditorPresetId)}`);
    applyAiReviewExcelMappingPresetToState(presetData?.preset?.mapping || {});
  }
  $("excelMappingDialogTitle").textContent = attachment.original_filename || "Excel 映射";
  $("excelMappingDialogSubtitle").textContent = "已读取当前文件结构。按工作表选择原文列、译文列和补充信息列。";
  $("applyExcelMappingButton").textContent = "使用当前映射";
  renderAiReviewExcelMappingDialog();
  $("reviewAttachmentMappingDialog").close();
  $("excelMappingDialog").showModal();
}

async function saveReviewAttachmentMapping() {
  const attachmentId = reviewConversationState.mappingAttachmentId;
  const mode = document.querySelector('input[name="reviewAttachmentMappingMode"]:checked')?.value || "ai";
  const presetId = mode === "preset" ? $("reviewAttachmentPresetSelect").value : null;
  if (mode === "preset" && !presetId) throw new Error("请点击“读取结构并编辑”，完成映射后保存模板");
  await api(`/api/ai-review/conversations/${encodeURIComponent(reviewConversationState.currentId)}/attachments/${encodeURIComponent(attachmentId)}`, {
    method: "PATCH",
    body: JSON.stringify({ mode, preset_id: presetId }),
  });
  $("reviewAttachmentMappingDialog").close();
  await refreshCurrentReviewConversation();
}

async function deleteCurrentReviewConversation() {
  if (!reviewConversationState.currentId) return;
  if (!window.confirm("删除后会同时清理本会话的附件和审校结果，确定继续吗？")) return;
  await api(`/api/ai-review/conversations/${encodeURIComponent(reviewConversationState.currentId)}`, {
    method: "DELETE",
    body: JSON.stringify({ confirm: true }),
  });
  if (reviewConversationState.eventSource) reviewConversationState.eventSource.close();
  reviewConversationState.currentId = "";
  reviewConversationState.snapshot = null;
  reviewConversationState.activeTarget = "";
  await loadReviewConversations();
}

function openReviewLanguagePopover(mode, anchor) {
  reviewConversationState.languageMode = mode;
  $("reviewLanguagePopoverTitle").textContent = mode === "source" ? "源语言" : "目标语言";
  $("reviewLanguagePopoverHint").textContent = mode === "source" ? "单选" : "可多选";
  $("reviewLanguageSearch").value = "";
  renderReviewLanguageOptions();
  const popover = $("reviewLanguagePopover");
  popover.classList.remove("hidden");
  const composer = $("reviewComposer");
  const wantedLeft = Math.max(0, anchor.offsetLeft - 6);
  const maxLeft = Math.max(0, composer.clientWidth - popover.offsetWidth - 14);
  popover.style.left = `${Math.min(wantedLeft, maxLeft)}px`;
  $("reviewLanguageSearch").focus();
}

function closeReviewLanguagePopover() {
  $("reviewLanguagePopover").classList.add("hidden");
}

function currentReviewTermBase() {
  return reviewConversationState.termBases.find((item) => item.id === reviewConversationState.termBaseId) || null;
}

function reviewTermBaseLanguageLabel(value) {
  return reviewTermBaseLanguageLabels[value] || String(value || "").replaceAll("_", " ");
}

function toggleReviewTermBasePopover() {
  closeReviewLanguagePopover();
  renderReviewTermBaseOptions();
  const popover = $("reviewTermBasePopover");
  popover.classList.toggle("hidden");
  if (!popover.classList.contains("hidden")) {
    const anchor = $("reviewTermBaseChip");
    const composer = $("reviewComposer");
    const wantedLeft = Math.max(0, anchor.offsetLeft - 6);
    const maxLeft = Math.max(0, composer.clientWidth - popover.offsetWidth - 14);
    popover.style.left = `${Math.min(wantedLeft, maxLeft)}px`;
  }
}

function closeReviewTermBasePopover() {
  $("reviewTermBasePopover").classList.add("hidden");
}

function renderReviewTermBaseOptions() {
  const container = $("reviewTermBaseOptions");
  if (!container) return;
  container.innerHTML = "";
  if (!reviewConversationState.termBases.length) {
    const empty = document.createElement("p");
    empty.className = "hint";
    empty.textContent = "尚未缓存术语表";
    container.appendChild(empty);
    return;
  }
  reviewConversationState.termBases.forEach((termBase) => {
    const row = document.createElement("div");
    row.className = `review-term-base-option ${termBase.id === reviewConversationState.termBaseId ? "selected" : ""}`.trim();
    row.setAttribute("role", "menuitem");
    row.tabIndex = 0;
    const check = document.createElement("span");
    check.className = "review-term-base-option-check";
    check.textContent = termBase.id === reviewConversationState.termBaseId ? "✓" : "";
    const copy = document.createElement("span");
    copy.className = "review-term-base-option-copy";
    const name = document.createElement("strong");
    name.textContent = termBase.filename || "未命名术语表";
    const meta = document.createElement("small");
    const noteLabel = termBase.has_entry_note ? " · 含备注" : "";
    const languages = (termBase.languages || []).map(reviewTermBaseLanguageLabel);
    meta.textContent = `${Number(termBase.entry_count || 0)} 条 · ${languages.join(" / ")}${noteLabel}`;
    copy.append(name, meta);
    const remove = document.createElement("button");
    remove.type = "button";
    remove.className = "review-term-base-delete";
    remove.title = `删除 ${termBase.filename || "术语表"}`;
    remove.setAttribute("aria-label", remove.title);
    remove.textContent = "×";
    const select = () => selectReviewTermBase(termBase.id).catch(showReviewConversationError);
    row.addEventListener("click", (event) => {
      if (!remove.contains(event.target)) select();
    });
    row.addEventListener("keydown", (event) => {
      if (event.key === "Enter" || event.key === " ") {
        event.preventDefault();
        select();
      }
    });
    remove.addEventListener("click", (event) => {
      event.stopPropagation();
      deleteReviewTermBase(termBase).catch(showReviewConversationError);
    });
    row.append(check, copy, remove);
    container.appendChild(row);
  });
}

async function selectReviewTermBase(termBaseId) {
  reviewConversationState.termBaseId = String(termBaseId || "");
  updateReviewComposerChips();
  renderReviewTermBaseOptions();
  closeReviewTermBasePopover();
  await saveReviewConversationComposerSettings();
}

async function uploadReviewTermBase(file) {
  if (!file) return;
  const form = new FormData();
  form.append("file", file);
  $("reviewConversationHint").textContent = `正在读取术语表 ${file.name}…`;
  const data = await api("/api/ai-review/term-bases", { method: "POST", body: form });
  const termBase = data.term_base || {};
  await loadReviewTermBases();
  await selectReviewTermBase(termBase.id || "");
  $("reviewConversationHint").textContent = `已缓存并选择 ${termBase.filename || file.name}`;
}

async function deleteReviewTermBase(termBase) {
  if (!window.confirm(`确定删除术语表“${termBase.filename || "未命名术语表"}”吗？所有引用它的会话将改为不使用术语表。`)) return;
  await api(`/api/ai-review/term-bases/${encodeURIComponent(termBase.id)}`, { method: "DELETE" });
  if (reviewConversationState.termBaseId === termBase.id) reviewConversationState.termBaseId = "";
  await loadReviewTermBases();
  await refreshCurrentReviewConversation();
  $("reviewConversationHint").textContent = "术语表已删除";
}

function renderReviewLanguageOptions() {
  const search = $("reviewLanguageSearch").value.trim().toLowerCase();
  const container = $("reviewLanguageOptions");
  container.innerHTML = "";
  reviewLanguages
    .filter((language) => reviewConversationState.languageMode === "source" || language !== "无源文")
    .filter((language) => !search || language.toLowerCase().includes(search))
    .forEach((language) => {
    const value = language === "自动检测" ? "auto" : language === "无源文" ? "none" : language;
    const label = document.createElement("label");
    label.className = "review-language-option";
    const input = document.createElement("input");
    input.type = "checkbox";
    input.name = "review-language";
    input.value = value;
    input.checked = reviewConversationState.languageMode === "source"
      ? reviewConversationState.sourceLanguage === value
      : reviewConversationState.targetLanguages.includes(value);
    label.classList.toggle("selected", input.checked);
    const check = document.createElement("span");
    check.className = "review-language-check";
    check.textContent = "✓";
    label.append(input, check, document.createTextNode(language));
    label.addEventListener("click", (event) => {
      event.preventDefault();
      if (reviewConversationState.languageMode === "source") {
        reviewConversationState.sourceLanguage = value;
        updateReviewComposerChips();
        closeReviewLanguagePopover();
        saveReviewConversationComposerSettings().catch(showReviewConversationError);
        return;
      }
      let selected = [...reviewConversationState.targetLanguages];
      if (value === "auto") {
        selected = ["auto"];
      } else {
        selected = selected.filter((item) => item !== "auto");
        selected = selected.includes(value) ? selected.filter((item) => item !== value) : [...selected, value];
        if (!selected.length) selected = ["auto"];
      }
      reviewConversationState.targetLanguages = selected;
      updateReviewComposerChips();
      renderReviewLanguageOptions();
      saveReviewConversationComposerSettings().catch(showReviewConversationError);
    });
    container.appendChild(label);
  });
}

function updateReviewComposerChips() {
  const prompt = aiReviewPromptTemplates.find((item) => item.id === reviewConversationState.promptTemplateId) || aiReviewPromptTemplates[0];
  $("reviewPromptChip").textContent = `提示词：${prompt?.name || "默认"}`;
  const sourceLabel = reviewConversationState.sourceLanguage === "auto"
    ? "自动"
    : reviewConversationState.sourceLanguage === "none" ? "无" : reviewConversationState.sourceLanguage;
  $("reviewSourceLanguageChip").textContent = `源语言：${sourceLabel}`;
  const targets = reviewConversationState.targetLanguages.map((value) => value === "auto" ? "自动" : value);
  $("reviewTargetLanguageChip").textContent = `目标语言：${targets.join("、")}`;
  const termBase = currentReviewTermBase();
  $("reviewTermBaseChip").textContent = termBase?.filename || "术语表";
  $("reviewTermBaseChip").title = termBase
    ? `${termBase.filename} · ${Number(termBase.entry_count || 0)} 条术语`
    : "选择术语表";
}

async function openReviewConversationPromptDialog() {
  const templateId = reviewConversationState.promptTemplateId || aiReviewPromptTemplates[0]?.id;
  if (!templateId) throw new Error("没有可用提示词模板");
  $("promptTemplateSelect").value = templateId;
  const data = await api(`/api/ai-review/prompt-templates/${encodeURIComponent(templateId)}`);
  fillAiReviewPromptDialog(data.template || {});
  $("promptDialog").showModal();
}

function reviewSessionStatusLabel(status) {
  return ({ draft: "等待输入", inspecting: "正在识别文件", needs_input: "等待确认", ready: "识别完成", reviewing: "正在审校", validating: "正在校验", completed: "已完成", partial: "部分完成", failed: "失败" })[status] || status || "等待输入";
}

function reviewAttachmentStatusLabel(status) {
  return ({ uploaded: "待识别", ready: "已识别", needs_reader: "格式不支持", failed: "读取失败" })[status] || status || "待识别";
}

function showReviewConversationError(error) {
  $("reviewConversationHint").textContent = error?.message || String(error || "操作失败");
}

async function chooseAiReviewFile() {
  $("reviewFileInput").click();
}

async function uploadAiReviewFile(file) {
  if (!file) {
    return;
  }
  resetAiReviewTaskView();
  setAiReviewFileCard(file.name || "");
  $("reviewFileHint").textContent = "正在读取文件...";
  const formData = new FormData();
  formData.append("file", file);
  const response = await fetch("/api/ai-review/file/upload", {
    method: "POST",
    body: formData,
  });
  const data = await response.json().catch(() => ({}));
  if (!response.ok) {
    throw new Error(data.detail || data.message || response.statusText || "读取失败");
  }
  if (String(data.file_type || "") === "excel") {
    initializeAiReviewExcelMapping(data);
    $("openExcelMappingButton").disabled = !aiReviewSheetNames.length;
  } else {
    aiReviewSheetNames = [];
    aiReviewColumnsBySheet = {};
    aiReviewExcelMappingState = {};
    aiReviewActiveSheetName = "";
    $("excelMappingSummary").textContent = "";
  }
  renderAiReviewBatch(data);
}

async function loadAiReviewFile(filePath) {
  resetAiReviewTaskView();
  const data = await api("/api/ai-review/file/load", {
    method: "POST",
    body: JSON.stringify({ file_path: filePath }),
  });
  $("reviewFileHint").textContent = data.message || "读取完成";
  if (String(data.file_type || "") === "excel") {
    initializeAiReviewExcelMapping(data);
    $("openExcelMappingButton").disabled = !aiReviewSheetNames.length;
  } else {
    aiReviewSheetNames = [];
    aiReviewColumnsBySheet = {};
    aiReviewExcelMappingState = {};
    aiReviewActiveSheetName = "";
    $("excelMappingSummary").textContent = "";
  }
  renderAiReviewBatch(data);
}

function resetAiReviewTaskView() {
  aiReviewTaskId = "";
  aiReviewCurrentTask = null;
  aiReviewLogCursor = 0;
  if (aiReviewTaskPoller) {
    window.clearInterval(aiReviewTaskPoller);
    aiReviewTaskPoller = null;
  }
  $("reviewLogList").innerHTML = "";
  renderAiReviewResultHead({});
  $("reviewResultBody").innerHTML = '<tr><td colspan="6" class="empty-cell">暂无审校结果</td></tr>';
  $("outputPanel").classList.add("hidden");
  $("outputPath").textContent = "暂无输出";
  $("openOutputFileButton").disabled = true;
  $("openReviewDetailButton").disabled = true;
  renderAiReviewProgress({});
  $("reviewTaskHint").textContent = "";
}

async function startAiReviewTask() {
  if (!aiReviewBatch?.id) {
    $("reviewTaskHint").textContent = "请先读取文件";
    return;
  }
  const mode = $("enableAiReview").checked
    ? ($("enableDirectionalReview").checked ? "directional" : "normal")
    : "normal";
  const data = await api("/api/ai-review/start", {
    method: "POST",
    body: JSON.stringify({
      batch_id: aiReviewBatch.id,
      prompt_template_id: $("promptTemplateSelect").value || null,
      source_language: $("sourceLanguageInput").value || "",
      target_language: $("targetLanguageInput").value || "",
      mode,
      directional_template_id: $("directionalTemplateSelect").value || null,
      enable_ai_review: $("enableAiReview").checked,
      enable_forbidden_check: $("enableForbiddenCheck").checked,
      forbidden_template_id: $("forbiddenTemplateSelect").value || null,
    }),
  });
  aiReviewTaskId = String(data?.task?.id || "");
  aiReviewLogCursor = 0;
  $("reviewLogList").innerHTML = "";
  $("reviewTaskHint").textContent = data.message || "审校任务已启动";
  renderAiReviewProgress(data?.task || {});
  setAiReviewTaskStatus({
    active: true,
    pill: "运行中",
    pillClass: "running",
    stageLabel: "审校中",
    message: "正在执行 AI 审校",
  });
  await pollAiReviewTask();
  if (aiReviewTaskPoller) {
    window.clearInterval(aiReviewTaskPoller);
  }
  aiReviewTaskPoller = window.setInterval(pollAiReviewTask, 1500);
}

async function pollAiReviewTask() {
  if (!aiReviewTaskId) {
    return;
  }
  const data = await api(`/api/ai-review/tasks/${encodeURIComponent(aiReviewTaskId)}`);
  renderAiReviewResults(data.task || {}, data.results || []);
  const task = data.task || {};
  const status = String(task.status || "");
  const done = ["completed", "completed_with_errors", "failed"].includes(status);
  const hasFailedItems = Number(task.failed_count || 0) > 0 || status === "completed_with_errors";
  setAiReviewTaskStatus({
    active: !done,
    pill: status === "failed" ? "失败" : hasFailedItems ? "有失败" : done ? "空闲" : "运行中",
    pillClass: status === "failed" || hasFailedItems ? "failed" : done ? "" : "running",
    stageLabel: String(task.status_label || task.status || "审校中"),
    message: String(task.message || task.status_label || "正在执行 AI 审校"),
  });
  const logs = await api(`/api/ai-review/tasks/${encodeURIComponent(aiReviewTaskId)}/logs?after_id=${aiReviewLogCursor}`);
  appendAiReviewLogs(logs.logs || []);
  if (done && aiReviewTaskPoller) {
    window.clearInterval(aiReviewTaskPoller);
    aiReviewTaskPoller = null;
  }
}

async function openAiReviewOutputDir() {
  await api("/api/ai-review/outputs/open-folder", {
    method: "POST",
    body: "{}",
  });
}

async function openAiReviewOutputFile() {
  const filePath = $("outputPath").textContent.trim();
  if (!filePath || filePath === "暂无输出") {
    return;
  }
  await api("/api/ai-review/outputs/open-file", {
    method: "POST",
    body: JSON.stringify({ file_path: filePath }),
  });
}

function setCrossExcelOutput(filePath, message = "") {
  crossExcelOutputFile = String(filePath || "").trim();
  $("crossExcelOutputFile").textContent = crossExcelOutputFile || "暂无输出";
  $("crossExcelOutputHint").textContent = message || (crossExcelOutputFile ? "合并完成，可以打开文件或输出目录。" : "合并结果会保存到工具 output 目录。");
  $("openCrossExcelOutputButton").disabled = !crossExcelOutputFile;
  $("openCrossExcelOutputFileButton").disabled = !crossExcelOutputFile;
}

async function copyToClipboard(text) {
  const value = String(text || "");
  if (navigator.clipboard?.writeText) {
    await navigator.clipboard.writeText(value);
    return;
  }
  const textarea = document.createElement("textarea");
  textarea.value = value;
  textarea.setAttribute("readonly", "readonly");
  textarea.style.position = "fixed";
  textarea.style.opacity = "0";
  document.body.appendChild(textarea);
  textarea.select();
  document.execCommand("copy");
  document.body.removeChild(textarea);
}

async function loadSettings() {
  const data = await api("/api/settings");
  currentProviderName = data.provider_name || "DeepSeek";
  renderModelOptions([], data.model_name || "");
  $("apiKey").value = data.api_key || "";
  currentBaseUrl = data.base_url || "";
  $("timeoutSeconds").value = data.timeout_seconds || 90;
  $("disableSystemProxy").checked = data.disable_system_proxy !== false;
  $("sourceLanguage").value = data.source_language || "中文";
  $("extractionMode").value = data.extraction_mode || "terms";
  $("nontransPlaceholderFormat").value = data.nontrans_placeholder_format || "<{n}>";
  const nontrans = data.nontrans_stage_settings || {};
  const recall = data.term_recall_stage_settings || {};
  const review = data.term_review_stage_settings || {};
  const aiReview = data.ai_review_stage_settings || {};
  $("nontransLimit").value = Number(nontrans.chunk_char_limit || 0) > 0 ? nontrans.chunk_char_limit : 5000;
  $("recallLimit").value = Number(recall.batch_request_char_limit || 0) > 0 ? recall.batch_request_char_limit : 5000;
  $("reviewLimit").value = Number(review.batch_request_char_limit || 0) > 0 ? review.batch_request_char_limit : 5000;
  $("reviewContextLimit").value = review.max_context_chars || 220;
  const normalizeReasoningEffort = (value) => ["low", "medium", "high"].includes(value) ? value : "low";
  $("nontransReasoningEffort").value = normalizeReasoningEffort(nontrans.reasoning_effort);
  $("recallReasoningEffort").value = normalizeReasoningEffort(recall.reasoning_effort);
  $("termReviewReasoningEffort").value = normalizeReasoningEffort(review.reasoning_effort);
  if ($("reviewAiLimit")) {
    const configuredLimit = Number(aiReview.batch_request_char_limit || 0);
    $("reviewAiLimit").value = configuredLimit > 0 ? configuredLimit : 1500;
  }
  $("reviewMaxItems").value = aiReview.max_items_per_request || 20;
  $("reviewReasoningEffort").value = normalizeReasoningEffort(aiReview.reasoning_effort === "max" ? "high" : aiReview.reasoning_effort);
  $("reviewDebugPayloadLogging").checked = Boolean(aiReview.debug_payload_logging);
  $("builtinRegex").checked = nontrans.builtin_regex_enabled !== false;
  $("aiDiscovery").checked = nontrans.ai_discovery_enabled !== false;
  $("aiRegex").checked = nontrans.ai_regex_generation_enabled !== false;
  $("numericNormalization").checked = data.numeric_normalization_enabled !== false;
  applyPendingRuleState(data.pending_nontrans_rules || {});
  setHeaderOptions([], "");
}

function renderModelOptions(models, preferredModel = "") {
  availableModels = Array.isArray(models) ? [...models] : [];
  const select = $("modelName");
  const currentValue = String(preferredModel || select.value || "").trim();
  select.innerHTML = "";

  const placeholder = document.createElement("option");
  placeholder.value = "";
  placeholder.textContent = availableModels.length ? "请选择模型" : "保存 API Key 后加载模型";
  placeholder.disabled = availableModels.length > 0;
  placeholder.selected = !currentValue;
  select.appendChild(placeholder);

  const finalModels = [...availableModels];
  if (currentValue && !finalModels.includes(currentValue)) {
    finalModels.unshift(currentValue);
  }

  finalModels.forEach((name) => {
    const option = document.createElement("option");
    option.value = String(name || "");
    option.textContent = String(name || "");
    if (String(name || "") === currentValue) {
      option.selected = true;
    }
    select.appendChild(option);
  });
}

async function refreshModelList() {
  const data = await api("/api/providers/models", {
    method: "POST",
    body: JSON.stringify(modelConnectionPayload()),
  });
  const currentModel = $("modelName").value.trim();
  const nextModel = currentModel && (data.models || []).includes(currentModel)
    ? currentModel
    : data.selected_model && (data.models || []).includes(data.selected_model)
      ? data.selected_model
      : (data.models || [])[0] || "";
  renderModelOptions(data.models || [], nextModel);
  return data;
}

async function saveModelConnection() {
  try {
    await api("/api/settings", { method: "POST", body: JSON.stringify(modelConnectionPayload()) });
    const modelData = await refreshModelList();
    await api("/api/settings", { method: "POST", body: JSON.stringify(modelConnectionPayload()) });
    $("modelConnectionHint").textContent = modelData.message || "已保存并刷新模型列表";
  } catch (error) {
    renderModelOptions([], $("modelName").value);
    $("modelConnectionHint").textContent = `保存成功，但模型列表刷新失败：${error.message}`;
  }
  setTimeout(() => ($("modelConnectionHint").textContent = ""), 2600);
}

async function saveSettings() {
  await api("/api/settings", { method: "POST", body: JSON.stringify(settingsPayload()) });
  $("saveHint").textContent = "设置已保存";
  setTimeout(() => ($("saveHint").textContent = ""), 1800);
}

function reviewSettingsPayload() {
  return {
    ai_review_batch_char_limit: Number($("reviewAiLimit").value || 1500),
    ai_review_max_items_per_request: Number($("reviewMaxItems").value || 20),
    ai_review_debug_payload_logging: $("reviewDebugPayloadLogging").checked,
    ai_review_reasoning_effort: $("reviewReasoningEffort").value || "low",
  };
}

async function saveReviewSettings() {
  await api("/api/settings", { method: "POST", body: JSON.stringify(reviewSettingsPayload()) });
  $("reviewSettingsHint").textContent = "审校设置已保存";
  setTimeout(() => ($("reviewSettingsHint").textContent = ""), 1800);
}

function renderPromptTemplates() {
  const groups = {
    recall: $("promptRecallPanel"),
    review: $("promptReviewPanel"),
    nontrans: $("promptNontransPanel"),
  };
  Object.values(groups).forEach((panel) => {
    panel.innerHTML = "";
  });
  promptTemplates.forEach((item) => {
    const key = String(item.key || "");
    let groupKey = "nontrans";
    if (key.startsWith("candidate_")) {
      groupKey = "recall";
    } else if (key.startsWith("classification_")) {
      groupKey = "review";
    }
    const list = groups[groupKey];
    const card = document.createElement("div");
    card.className = "prompt-card";
    card.innerHTML = `
      <header>
        <div>
          <strong>${escapeHtml(item.label || item.key || "")}</strong>
          ${item.description ? `<small>${escapeHtml(item.description || "")}</small>` : ""}
        </div>
        <div class="prompt-card-actions">
          <button type="button" class="secondary mini-button prompt-reset-button" data-key="${escapeHtml(item.key || "")}">恢复默认</button>
          <span class="prompt-badge ${item.is_default ? "" : "modified"}">${item.is_default ? "默认" : "已修改"}</span>
        </div>
      </header>
      <textarea data-key="${escapeHtml(item.key || "")}" spellcheck="false">${escapeHtml(item.value || "")}</textarea>`;
    card.querySelector("textarea").addEventListener("input", (event) => {
      item.value = event.target.value;
      item.is_default = false;
      const badge = card.querySelector(".prompt-badge");
      badge.textContent = "已修改";
      badge.classList.add("modified");
    });
    card.querySelector(".prompt-reset-button").addEventListener("click", () => resetSinglePromptTemplate(item.key));
    list.appendChild(card);
  });
}

async function loadPromptTemplates() {
  const data = await api("/api/prompt-templates");
  promptTemplates = data.templates || [];
  renderPromptTemplates();
}

async function savePromptTemplates() {
  const templates = {};
  promptTemplates.forEach((item) => {
    templates[item.key] = item.value || "";
  });
  await api("/api/prompt-templates", {
    method: "POST",
    body: JSON.stringify({ templates }),
  });
  $("promptTemplateHint").textContent = "提示词已保存";
  setTimeout(() => ($("promptTemplateHint").textContent = ""), 1800);
  await loadPromptTemplates();
}

async function resetPromptTemplates() {
  if (!window.confirm("恢复默认提示词会覆盖当前编辑内容，确定继续吗？")) return;
  await api("/api/prompt-templates/reset", {
    method: "POST",
    body: JSON.stringify({ keys: promptTemplates.map((item) => item.key) }),
  });
  $("promptTemplateHint").textContent = "默认提示词已恢复";
  setTimeout(() => ($("promptTemplateHint").textContent = ""), 1800);
  await loadPromptTemplates();
}

async function resetSinglePromptTemplate(key) {
  const localSnapshot = new Map(
    promptTemplates.map((item) => [
      item.key,
      {
        ...item,
        value: item.value || "",
        is_default: Boolean(item.is_default),
      },
    ]),
  );
  const data = await api("/api/prompt-templates/reset", {
    method: "POST",
    body: JSON.stringify({ keys: [key] }),
  });
  const resetItem = (data.templates || []).find((item) => item.key === key) || null;
  promptTemplates = promptTemplates.map((item) => {
    if (item.key === key) {
      return resetItem || item;
    }
    return localSnapshot.get(item.key) || item;
  });
  $("promptTemplateHint").textContent = "该提示词已恢复默认";
  setTimeout(() => ($("promptTemplateHint").textContent = ""), 1800);
  renderPromptTemplates();
}

function renderAsciiPatterns() {
  const body = $("asciiPatternBody");
  body.innerHTML = "";
  if (!asciiPatterns.length) {
    body.innerHTML = '<tr><td colspan="5" class="empty-cell">暂无规则</td></tr>';
    return;
  }
  asciiPatterns.forEach((item, index) => {
    const row = document.createElement("tr");
    row.innerHTML = `
      <td><input type="checkbox" data-field="enabled" ${item.enabled !== false ? "checked" : ""}></td>
      <td><input type="text" data-field="name" value="${escapeHtml(item.name || "")}"></td>
      <td><input type="text" data-field="pattern" value="${escapeHtml(item.pattern || "")}"></td>
      <td>${index + 1}</td>
      <td class="pattern-actions">
        <button type="button" data-action="up">上移</button>
        <button type="button" data-action="down">下移</button>
        <button type="button" data-action="delete">删除</button>
      </td>`;
    row.querySelectorAll("input").forEach((input) => {
      input.addEventListener("change", () => syncAsciiPatternFromRow(index, row));
      input.addEventListener("input", () => syncAsciiPatternFromRow(index, row));
    });
    row.querySelector('[data-action="up"]').addEventListener("click", () => moveAsciiPattern(index, -1));
    row.querySelector('[data-action="down"]').addEventListener("click", () => moveAsciiPattern(index, 1));
    row.querySelector('[data-action="delete"]').addEventListener("click", () => {
      asciiPatterns.splice(index, 1);
      renderAsciiPatterns();
    });
    body.appendChild(row);
  });
}

function escapeHtml(value) {
  return String(value).replace(/[&<>"']/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[char]));
}

function syncAsciiPatternFromRow(index, row) {
  asciiPatterns[index] = {
    name: row.querySelector('[data-field="name"]').value,
    pattern: row.querySelector('[data-field="pattern"]').value,
    enabled: row.querySelector('[data-field="enabled"]').checked,
    order_index: index + 1,
  };
}

function moveAsciiPattern(index, offset) {
  const target = index + offset;
  if (target < 0 || target >= asciiPatterns.length) return;
  const [item] = asciiPatterns.splice(index, 1);
  asciiPatterns.splice(target, 0, item);
  renderAsciiPatterns();
}

async function loadAsciiPatterns() {
  const data = await api("/api/ascii-candidate-patterns");
  asciiPatterns = (data.patterns || []).map((item, index) => ({
    name: item.name || item.pattern || "",
    pattern: item.pattern || "",
    enabled: item.enabled !== false,
    order_index: item.order_index || index + 1,
  })).sort((a, b) => Number(a.order_index || 0) - Number(b.order_index || 0));
  renderAsciiPatterns();
}

function addAsciiPattern() {
  asciiPatterns.push({ name: "新规则", pattern: "", enabled: true, order_index: asciiPatterns.length + 1 });
  renderAsciiPatterns();
}

async function saveAsciiPatterns() {
  await api("/api/ascii-candidate-patterns", {
    method: "POST",
    body: JSON.stringify({ patterns: asciiPatterns }),
  });
  $("asciiPatternHint").textContent = "规则已保存";
  setTimeout(() => ($("asciiPatternHint").textContent = ""), 1800);
  await loadAsciiPatterns();
}

function normalizeBuiltinRule(item = {}, index = 0) {
  return {
    rule_id: String(item.rule_id || `custom_rule_${String(index + 1).padStart(3, "0")}`),
    name: String(item.name || ""),
    role: String(item.role || "empty"),
    regex: String(item.regex || item.pattern || ""),
    element_type: String(item.element_type || "other"),
    enabled: item.enabled !== false,
    examples: Array.isArray(item.examples) ? item.examples.map((value) => String(value || "")).filter(Boolean) : [],
  };
}

function normalizePendingRule(item = {}, index = 0) {
  return {
    cache_id: String(item.cache_id || `pending_rule_${String(index + 1).padStart(4, "0")}`),
    rule_id: String(item.rule_id || item.cache_id || `pending_rule_${String(index + 1).padStart(4, "0")}`),
    name: String(item.name || item.rule_id || item.cache_id || ""),
    role: String(item.role || "empty"),
    regex: String(item.regex || item.pattern || ""),
    element_type: String(item.element_type || "other"),
    enabled: item.enabled !== false,
    selected: item.selected !== false,
    examples: Array.isArray(item.examples) ? item.examples.map((value) => String(value || "")).filter(Boolean) : [],
  };
}

function renderBuiltinRuleEditors() {
  const body = $("builtinRuleEditorBody");
  body.innerHTML = "";
  if (!builtinRuleDefinitions.length) {
    body.innerHTML = '<tr><td colspan="5" class="empty-cell">暂无规则</td></tr>';
    return;
  }
  builtinRuleDefinitions.forEach((item, index) => {
    const row = document.createElement("tr");
    row.innerHTML = `
      <td><input type="checkbox" data-field="enabled" ${item.enabled !== false ? "checked" : ""}></td>
      <td><input type="text" data-field="name" value="${escapeHtml(item.name || "")}" placeholder="规则标题"></td>
      <td><input type="text" data-field="regex" value="${escapeHtml(item.regex || "")}" placeholder="正则表达式"></td>
      <td>
        <select data-field="role">
          <option value="open" ${item.role === "open" ? "selected" : ""}>开始</option>
          <option value="close" ${item.role === "close" ? "selected" : ""}>结束</option>
          <option value="empty" ${item.role === "empty" ? "selected" : ""}>空</option>
        </select>
      </td>
      <td class="pattern-actions"><button type="button" data-action="delete">删除</button></td>`;
    row.querySelectorAll("input, select").forEach((input) => {
      input.addEventListener("change", () => syncBuiltinRuleFromEditor(index, row));
      input.addEventListener("input", () => syncBuiltinRuleFromEditor(index, row));
    });
    row.querySelector('[data-action="delete"]').addEventListener("click", () => {
      builtinRuleDefinitions.splice(index, 1);
      renderBuiltinRuleEditors();
    });
    body.appendChild(row);
  });
}

function syncBuiltinRuleFromEditor(index, row) {
  builtinRuleDefinitions[index] = {
    ...builtinRuleDefinitions[index],
    name: row.querySelector('[data-field="name"]').value.trim(),
    role: row.querySelector('[data-field="role"]').value,
    regex: row.querySelector('[data-field="regex"]').value.trim(),
    enabled: row.querySelector('[data-field="enabled"]').checked,
  };
}

function addBuiltinRule() {
  builtinRuleDefinitions.push(normalizeBuiltinRule({}, builtinRuleDefinitions.length));
  renderBuiltinRuleEditors();
}

function renderBuiltinRules(ruleCount, rowCount) {
  $("builtinRuleCount").textContent = ruleCount || 0;
  $("builtinRuleRowCount").textContent = rowCount || 0;
  const body = $("builtinRuleBody");
  body.innerHTML = "";
  if (!builtinRuleRows.length) {
    body.innerHTML = '<tr><td colspan="6" class="empty-cell">暂无内置规则</td></tr>';
    return;
  }
  builtinRuleRows.forEach((item, index) => {
    const row = document.createElement("tr");
    const examples = (item.examples || []).slice(0, 3);
    const examplesHtml = examples.length
      ? `<div class="examples-scroll">${examples.map((value) => `<div class="examples-item">${escapeHtml(value || "")}</div>`).join("")}</div>`
      : "";
    row.innerHTML = `
      <td>${item.order_index || index + 1}</td>
      <td>${escapeHtml(item.name || "")}</td>
      <td>${escapeHtml(item.role_label || item.role || "")}</td>
      <td>${escapeHtml(item.element_type_label || item.element_type || "")}</td>
      <td class="regex-cell">${escapeHtml(item.regex || "")}</td>
      <td class="examples-cell">${examplesHtml}</td>`;
    body.appendChild(row);
  });
}

async function loadBuiltinRules() {
  const data = await api("/api/nontrans-builtin-rules");
  builtinRuleDefinitions = (data.rules || []).map((item, index) => normalizeBuiltinRule(item, index));
  builtinRuleRows = data.rows || [];
  renderBuiltinRuleEditors();
  renderBuiltinRules(data.rule_count || 0, data.row_count || builtinRuleRows.length);
}

function renderPendingRuleEditors() {
  const body = $("pendingRuleEditorBody");
  if (!body) return;
  body.innerHTML = "";
  if (!pendingRuleDefinitions.length) {
    body.innerHTML = '<tr><td colspan="5" class="empty-cell">当前没有待导入的规则</td></tr>';
    return;
  }
  pendingRuleDefinitions.forEach((item, index) => {
    const row = document.createElement("tr");
    const examples = (item.examples || []).slice(0, 3);
    const examplesHtml = examples.length
      ? `<div class="examples-scroll">${examples.map((value) => `<div class="examples-item">${escapeHtml(value || "")}</div>`).join("")}</div>`
      : "";
    row.innerHTML = `
      <td><input type="checkbox" data-field="selected" ${item.selected !== false ? "checked" : ""}></td>
      <td><input type="text" data-field="name" value="${escapeHtml(item.name || "")}" placeholder="规则标题"></td>
      <td><input type="text" data-field="regex" value="${escapeHtml(item.regex || "")}" placeholder="正则表达式"></td>
      <td>
        <select data-field="role">
          <option value="open" ${item.role === "open" ? "selected" : ""}>开始</option>
          <option value="close" ${item.role === "close" ? "selected" : ""}>结束</option>
          <option value="empty" ${item.role === "empty" ? "selected" : ""}>空</option>
        </select>
      </td>
      <td class="examples-cell">${examplesHtml}</td>`;
    row.querySelectorAll("input, select").forEach((input) => {
      input.addEventListener("change", () => syncPendingRuleFromEditor(index, row));
      input.addEventListener("input", () => syncPendingRuleFromEditor(index, row));
    });
    body.appendChild(row);
  });
}

function syncPendingRuleFromEditor(index, row) {
  pendingRuleDefinitions[index] = {
    ...pendingRuleDefinitions[index],
    selected: row.querySelector('[data-field="selected"]').checked,
    name: row.querySelector('[data-field="name"]').value.trim(),
    role: row.querySelector('[data-field="role"]').value,
    regex: row.querySelector('[data-field="regex"]').value.trim(),
  };
}

function setAllPendingRuleSelection(selected) {
  pendingRuleDefinitions = pendingRuleDefinitions.map((item) => ({ ...item, selected: Boolean(selected) }));
  renderPendingRuleEditors();
}

function ensurePendingRuleNavDot(pageTarget, visible) {
  const button = document.querySelector(`.nav-link[data-page-target="${pageTarget}"]`);
  if (!button) return;
  let wrapper = button.querySelector(".link-with-dot");
  let label = button.dataset.labelText;
  if (!label) {
    label = button.textContent.trim();
    button.dataset.labelText = label;
  }
  if (!wrapper) {
    button.textContent = "";
    wrapper = document.createElement("span");
    wrapper.className = "link-with-dot";
    const text = document.createElement("span");
    text.className = "link-text";
    text.textContent = label;
    const dot = document.createElement("span");
    dot.className = "badge-dot";
    dot.hidden = true;
    wrapper.appendChild(text);
    wrapper.appendChild(dot);
    button.appendChild(wrapper);
  }
  const dot = wrapper.querySelector(".badge-dot");
  if (dot) dot.hidden = !visible;
}

function ensurePendingRuleSubtabDot(targetId, visible) {
  const button = document.querySelector(`.subnav-link[data-subtab-target="${targetId}"]`);
  if (!button) return;
  let wrapper = button.querySelector(".link-with-dot");
  let label = button.dataset.labelText;
  if (!label) {
    label = button.textContent.trim();
    button.dataset.labelText = label;
  }
  if (!wrapper) {
    button.textContent = "";
    wrapper = document.createElement("span");
    wrapper.className = "link-with-dot";
    const text = document.createElement("span");
    text.className = "link-text";
    text.textContent = label;
    const dot = document.createElement("span");
    dot.className = "badge-dot";
    dot.hidden = true;
    wrapper.appendChild(text);
    wrapper.appendChild(dot);
    button.appendChild(wrapper);
  }
  const dot = wrapper.querySelector(".badge-dot");
  if (dot) dot.hidden = !visible;
}

function renderPendingRuleIndicators() {
  const noticeButton = $("pendingRuleNoticeButton");
  if (noticeButton) {
    noticeButton.hidden = !pendingRuleState.show_notice_button;
  }
  ensurePendingRuleNavDot("nontransSettingsPage", pendingRuleState.show_library_dot);
  ensurePendingRuleSubtabDot("nontransBuiltinPanel", pendingRuleState.show_library_dot);
}

function applyPendingRuleState(data) {
  pendingRuleState = {
    count: Number(data?.count || 0),
    has_pending: Boolean(data?.has_pending),
    notice_seen: Boolean(data?.notice_seen),
    library_seen: Boolean(data?.library_seen),
    show_notice_dot: Boolean(data?.show_notice_dot),
    show_library_dot: Boolean(data?.show_library_dot),
    show_notice_button: Boolean(data?.show_notice_button),
  };
  pendingRuleDefinitions = (data?.rules || []).map((item, index) => normalizePendingRule(item, index));
  renderPendingRuleEditors();
  renderPendingRuleIndicators();
}

function renderAppUpdateNotice() {
  const button = $("updateNoticeButton");
  if (!button) return;
  button.hidden = !appUpdateState.update_available;
}

function applyAppUpdateState(data) {
  appUpdateState = {
    supported: Boolean(data?.supported),
    current_version: String(data?.current_version || ""),
    latest_version: String(data?.latest_version || ""),
    update_available: Boolean(data?.update_available),
    release_notes: String(data?.release_notes || ""),
    published_at: String(data?.published_at || ""),
    download_url: String(data?.download_url || ""),
    asset_name: String(data?.asset_name || ""),
    message: String(data?.message || ""),
  };
  renderAppUpdateNotice();
  if (appUpdateState.update_available && !appUpdateAutoPrompted) {
    appUpdateAutoPrompted = true;
    window.setTimeout(openAppUpdateModal, 180);
  }
}

async function loadAppUpdateInfo() {
  try {
    const data = await api("/api/app-update");
    applyAppUpdateState(data);
  } catch (error) {
    applyAppUpdateState({
      supported: false,
      current_version: "",
      latest_version: "",
      update_available: false,
      message: error.message,
    });
  }
}

function openAppUpdateModal() {
  $("appUpdateCurrentVersion").textContent = appUpdateState.current_version || "-";
  $("appUpdateLatestVersion").textContent = appUpdateState.latest_version || "-";
  $("appUpdateSummary").textContent = appUpdateState.message
    ? appUpdateState.message
    : `当前版本 ${appUpdateState.current_version || "-"}，最新版本 ${appUpdateState.latest_version || "-"}。`;
  $("appUpdateReleaseNotes").textContent = appUpdateState.release_notes || "暂无更新日志";
  $("appUpdateHint").textContent = "";
  $("confirmAppUpdateButton").disabled = !appUpdateState.update_available || !appUpdateState.download_url;
  $("appUpdateOverlay").hidden = false;
}

function closeAppUpdateModal() {
  $("appUpdateOverlay").hidden = true;
  $("appUpdateHint").textContent = "";
}

function leavePageForAppUpdate(message) {
  const safeMessage = escapeHtml(message || "更新已开始，请查看弹出的更新窗口。");
  document.body.innerHTML = `
    <div class="update-leaving-screen">
      <div class="update-leaving-card">
        <h2>正在更新</h2>
        <p>${safeMessage}</p>
      </div>
    </div>`;
  const tryCloseWindow = () => {
    try {
      window.open("", "_self");
      window.close();
    } catch (error) {
      console.warn(error);
    }
  };
  setTimeout(tryCloseWindow, 200);
  setTimeout(() => {
    tryCloseWindow();
    window.location.replace("about:blank");
  }, 1200);
}

async function startAppUpdate() {
  $("appUpdateHint").textContent = "正在启动更新...";
  $("confirmAppUpdateButton").disabled = true;
  try {
    const data = await api("/api/app-update/start", {
      method: "POST",
      body: JSON.stringify({}),
    });
    const message = data.message || "更新已开始，当前页面会关闭，并弹出更新窗口显示进度。";
    $("appUpdateHint").textContent = message;
    setTimeout(() => {
      leavePageForAppUpdate(message);
    }, 180);
  } catch (error) {
    $("appUpdateHint").textContent = error.message;
    $("confirmAppUpdateButton").disabled = false;
  }
}

async function loadPendingRules() {
  const data = await api("/api/nontrans-pending-rules");
  applyPendingRuleState(data);
}

async function markPendingRuleSeen(payload) {
  const data = await api("/api/nontrans-pending-rules/seen", {
    method: "POST",
    body: JSON.stringify(payload),
  });
  applyPendingRuleState(data);
}

function openPendingRuleModal() {
  $("pendingRuleOverlay").hidden = false;
  $("pendingRuleHint").textContent = "";
  if (pendingRuleState.show_notice_dot) {
    markPendingRuleSeen({ notice_seen: true }).catch(() => {});
  }
}

async function clearPendingRules() {
  const data = await api("/api/nontrans-pending-rules/clear", {
    method: "POST",
    body: "{}",
  });
  applyPendingRuleState(data);
  return data;
}

async function closePendingRuleModal() {
  if (pendingRuleState.has_pending) {
    await clearPendingRules();
  }
  $("pendingRuleOverlay").hidden = true;
  $("pendingRuleHint").textContent = "";
}

async function importPendingRules() {
  const selectedRules = pendingRuleDefinitions
    .filter((item) => item.selected !== false)
    .map((item) => ({
      cache_id: item.cache_id,
      rule_id: item.rule_id,
      name: item.name,
      role: item.role,
      element_type: item.element_type,
      pattern: item.regex,
      enabled: item.enabled !== false,
      examples: item.examples || [],
    }));
  if (!selectedRules.length) {
    $("pendingRuleHint").textContent = "请先勾选要导入的规则";
    return;
  }
  const data = await api("/api/nontrans-pending-rules/import", {
    method: "POST",
    body: JSON.stringify({ rules: selectedRules }),
  });
  $("pendingRuleHint").textContent = "已导入到内置规则库";
  builtinRuleDefinitions = (data.rules || []).map((item, index) => normalizeBuiltinRule(item, index));
  builtinRuleRows = data.rows || [];
  renderBuiltinRuleEditors();
  renderBuiltinRules(data.rule_count || 0, data.row_count || builtinRuleRows.length);
  applyPendingRuleState(data.pending_nontrans_rules || {});
  if (!pendingRuleState.has_pending) {
    closePendingRuleModal();
  }
}

async function saveBuiltinRules() {
  await api("/api/nontrans-builtin-rules", {
    method: "POST",
    body: JSON.stringify({
      rules: builtinRuleDefinitions.map((item, index) => ({
        rule_id: item.rule_id || `custom_rule_${String(index + 1).padStart(3, "0")}`,
        name: item.name || "",
        role: item.role || "empty",
        element_type: item.element_type || "other",
        pattern: item.regex || "",
        enabled: item.enabled !== false,
        examples: item.examples || [],
      })),
    }),
  });
  $("builtinRuleHint").textContent = "规则库已保存";
  setTimeout(() => ($("builtinRuleHint").textContent = ""), 1800);
  await loadBuiltinRules();
}

async function scanFolder() {
  const folder = $("folderPath").value.trim();
  if (!folder) {
    $("scanResult").textContent = "请先填写输入目录";
    return;
  }
  try {
    const data = await api(`/api/scan?folder_path=${encodeURIComponent(folder)}`);
    applyScanResult(data);
  } catch (error) {
    lastScanResult = null;
    $("headerName").disabled = false;
    setHeaderOptions([], "");
    $("scanResult").textContent = error.message;
  }
}

async function chooseFolder() {
  try {
    const data = await api("/api/dialog/select-folder", {
      method: "POST",
      body: "{}",
    });
    if (data.cancelled || !data.folder_path) {
      return;
    }
    $("folderPath").value = data.folder_path;
    await scanFolder();
  } catch (error) {
    $("scanResult").textContent = error.message;
  }
}

async function chooseCrossExcelFolder() {
  try {
    const data = await api("/api/dialog/select-folder", {
      method: "POST",
      body: "{}",
    });
    if (data.cancelled || !data.folder_path) {
      return;
    }
    $("crossExcelFolderPath").value = data.folder_path;
    await scanCrossExcelFolder();
  } catch (error) {
    $("crossExcelScanHint").textContent = error.message;
  }
}

async function chooseDiffFolder(side) {
  const data = await api("/api/dialog/select-folder", { method: "POST", body: "{}" });
  if (!data.cancelled && data.folder_path) {
    if (side === "A") {
      $("diffPathA").value = data.folder_path;
    } else {
      $("diffPathB").value = data.folder_path;
    }
    resetDiffFieldMatchSettings();
  }
}

async function chooseDiffFile(side) {
  const data = await api("/api/dialog/select-review-file");
  if (!data.cancelled && data.file_path) {
    if (side === "A") {
      $("diffPathA").value = data.file_path;
    } else {
      $("diffPathB").value = data.file_path;
    }
    resetDiffFieldMatchSettings();
  }
}

async function scanCrossExcelFolder() {
  const folder = $("crossExcelFolderPath").value.trim();
  if (!folder) {
    $("crossExcelScanHint").textContent = "请先填写输入目录";
    return;
  }
  setTaskStatus("crossExcelPage", {
    active: true,
    taskLabel: "跨Excel搜索与合并",
    pill: "运行中",
    pillClass: "running",
    stageLabel: "扫描目录",
    message: "正在扫描文件和表头",
  });
  renderCurrentTaskStatus();
  try {
    const data = await api(`/api/cross-excel/scan?folder_path=${encodeURIComponent(folder)}`);
    renderCrossExcelScan(data);
    setTaskStatus("crossExcelPage", {
      active: false,
      taskLabel: "跨Excel搜索与合并",
      pill: "空闲",
      pillClass: "",
      stageLabel: "扫描完成",
      message: `已发现 ${Number(data.file_count || 0)} 个文件，${Number((data.headers || []).length)} 个表头`,
    });
    renderCurrentTaskStatus();
  } catch (error) {
    crossExcelScanState = null;
    $("crossExcelFileCount").textContent = "0";
    $("crossExcelScanHint").textContent = error.message;
    renderCrossExcelHeaders([]);
    renderCrossExcelSearchResults(null);
    setCrossExcelOutput("");
    setTaskStatus("crossExcelPage", {
      active: false,
      taskLabel: "跨Excel搜索与合并",
      pill: "失败",
      pillClass: "failed",
      stageLabel: "扫描失败",
      message: error.message,
    });
    renderCurrentTaskStatus();
  }
}

async function searchCrossExcel() {
  const folder = $("crossExcelFolderPath").value.trim();
  const query = $("crossExcelQuery").value.trim();
  if (!folder) {
    $("crossExcelSearchHint").textContent = "请先选择目录";
    return;
  }
  if (!query) {
    $("crossExcelSearchHint").textContent = "请输入要搜索的内容";
    return;
  }
  $("crossExcelSearchHint").textContent = "搜索中...";
  setTaskStatus("crossExcelPage", {
    active: true,
    taskLabel: "跨Excel搜索与合并",
    pill: "运行中",
    pillClass: "running",
    stageLabel: "全局搜索",
    message: `正在搜索：${query}`,
  });
  renderCurrentTaskStatus();
  try {
    const data = await api("/api/cross-excel/search", {
      method: "POST",
      body: JSON.stringify({
        folder_path: folder,
        query,
        limit: Number($("crossExcelLimit").value || 300),
      }),
    });
    renderCrossExcelSearchResults(data);
    $("crossExcelSearchHint").textContent = data.truncated ? "结果过多，已按上限截断显示" : "搜索完成";
    setTaskStatus("crossExcelPage", {
      active: false,
      taskLabel: "跨Excel搜索与合并",
      pill: "空闲",
      pillClass: "",
      stageLabel: "搜索完成",
      message: data.truncated
        ? `已命中 ${Number((data.items || []).length)} 条，结果已截断`
        : `已命中 ${Number((data.items || []).length)} 条`,
    });
    renderCurrentTaskStatus();
  } catch (error) {
    renderCrossExcelSearchResults(null);
    $("crossExcelSearchHint").textContent = error.message;
    setTaskStatus("crossExcelPage", {
      active: false,
      taskLabel: "跨Excel搜索与合并",
      pill: "失败",
      pillClass: "failed",
      stageLabel: "搜索失败",
      message: error.message,
    });
    renderCurrentTaskStatus();
  }
}

async function mergeCrossExcel() {
  const folder = $("crossExcelFolderPath").value.trim();
  if (!folder) {
    $("crossExcelOutputHint").textContent = "请先选择目录";
    return;
  }
  const headers = selectedCrossExcelHeaders();
  if (!headers.length) {
    $("crossExcelOutputHint").textContent = "请至少勾选一个表头";
    return;
  }
  $("crossExcelOutputHint").textContent = "合并中...";
  setTaskStatus("crossExcelPage", {
    active: true,
    taskLabel: "跨Excel搜索与合并",
    pill: "运行中",
    pillClass: "running",
    stageLabel: "按表头合并",
    message: `正在合并 ${headers.length} 个表头`,
  });
  renderCurrentTaskStatus();
  try {
    const data = await api("/api/cross-excel/merge", {
      method: "POST",
      body: JSON.stringify({
        folder_path: folder,
        headers,
        apply_format: $("crossExcelApplyFormat").checked,
      }),
    });
    setCrossExcelOutput(
      data.output_file || "",
      `合并完成，共输出 ${Number(data.row_count || 0)} 行，${Number(data.column_count || 0)} 列。`,
    );
    setTaskStatus("crossExcelPage", {
      active: false,
      taskLabel: "跨Excel搜索与合并",
      pill: "空闲",
      pillClass: "",
      stageLabel: "合并完成",
      message: `已输出 ${Number(data.row_count || 0)} 行到结果文件`,
    });
    renderCurrentTaskStatus();
  } catch (error) {
    setCrossExcelOutput("", error.message);
    setTaskStatus("crossExcelPage", {
      active: false,
      taskLabel: "跨Excel搜索与合并",
      pill: "失败",
      pillClass: "failed",
      stageLabel: "合并失败",
      message: error.message,
    });
    renderCurrentTaskStatus();
  }
}

async function startDiffExcel() {
  const pathA = $("diffPathA").value.trim();
  const pathB = $("diffPathB").value.trim();
  if (!pathA || !pathB) {
    $("diffExcelHint").textContent = "请先选择路径 A 和路径 B。";
    return;
  }
  const compareMode = diffCompareMode();
  if (compareMode === "field_match" && (!diffFieldMatchSettings.referenceField || !diffFieldMatchSettings.compareFields.length)) {
    $("diffExcelHint").textContent = "请先打开字段设置，选择参考字段和比对字段。";
    return;
  }
  $("diffExcelHint").textContent = "比对中...";
  setTaskStatus("diffExcelPage", {
    active: true,
    taskLabel: "Diff 工具",
    pill: "运行中",
    pillClass: "running",
    stageLabel: compareMode === "field_match" ? "按字段匹配" : "按位置比对",
    message: "正在读取并比较两个路径",
  });
  renderCurrentTaskStatus();
  try {
    const started = await api("/api/diff-excel/compare", {
      method: "POST",
      body: JSON.stringify({
        path_a: pathA,
        path_b: pathB,
        compare_mode: compareMode,
        reference_field: diffFieldMatchSettings.referenceField,
        compare_fields: diffFieldMatchSettings.compareFields,
        include_unmatched: diffFieldMatchSettings.includeUnmatched,
      }),
    });
    const taskId = String(started.task_id || "");
    if (!taskId) throw new Error("未能启动比对任务。");
    diffExcelState.taskId = taskId;
    let data = null;
    while (true) {
      const task = await api(`/api/diff-excel/task/${encodeURIComponent(taskId)}`);
      const message = String(task.message || "正在比对");
      $("diffExcelHint").textContent = message;
      setTaskStatus("diffExcelPage", {
        active: true,
        taskLabel: "Diff 工具",
        pill: "运行中",
        pillClass: "running",
        stageLabel: compareMode === "field_match" ? "按字段匹配" : "按位置比对",
        message,
      });
      renderCurrentTaskStatus();
      if (task.status === "completed") {
        data = task.result || {};
        break;
      }
      if (task.status === "failed") {
        throw new Error(String(task.error || "比对失败"));
      }
      await new Promise((resolve) => window.setTimeout(resolve, 500));
    }
    diffExcelState.cacheFile = String(data.cache_file || "");
    diffExcelState.resultId = String(data.result_id || "");
    diffExcelState.previewOffset = 0;
    diffExcelState.previewRecords = [];
    diffExcelState.meta = data.meta || null;
    diffExcelState.outputFile = "";
    diffExcelState.totalCount = Number(data.total_count || 0);
    diffExcelState.matchedCount = Number(data.total_count || 0);
    diffExcelState.previewLimit = 200;
    diffExcelState.previewTruncated = Boolean(data.preview_truncated);
    const meta = diffExcelState.meta || {};
    const fieldSummary = compareMode === "field_match"
      ? `有效参考行 ${Number(meta.valid_reference_rows || 0)}，已配对 ${Number(meta.paired_reference_rows || 0)}，未匹配 ${Number(meta.unmatched_reference_rows || 0)}，跳过字段 ${Number(meta.skipped_field_count || 0)}。`
      : "";
    await applyDiffExcelFilter();
    $("diffExcelHint").textContent = `比对完成，共找到 ${diffExcelState.totalCount} 处差异。${fieldSummary}`;
    setTaskStatus("diffExcelPage", {
      active: false,
      taskLabel: "Diff 工具",
      pill: "空闲",
      pillClass: "",
      stageLabel: "比对完成",
      message: `已找到 ${diffExcelState.totalCount} 处差异`,
    });
    renderCurrentTaskStatus();
  } catch (error) {
    $("diffExcelHint").textContent = error.message;
    clearDiffExcelState();
    setTaskStatus("diffExcelPage", {
      active: false,
      taskLabel: "Diff 工具",
      pill: "失败",
      pillClass: "failed",
      stageLabel: "比对失败",
      message: error.message,
    });
    renderCurrentTaskStatus();
  }
}

async function exportDiffExcel() {
  if (!diffExcelState.cacheFile || !Number(diffExcelState.totalCount || 0)) {
    $("diffExcelHint").textContent = "当前没有可导出的差异。";
    return;
  }
  $("diffExcelHint").textContent = "导出中...";
  try {
    const data = await api("/api/diff-excel/export", {
      method: "POST",
      body: JSON.stringify({
        cache_file: diffExcelState.cacheFile,
        query: "",
        output_file: "",
      }),
    });
    diffExcelState.outputFile = String(data.output_file || "");
    $("diffExcelHint").textContent = "差异结果已导出。";
    renderDiffExcelSummary();
  } catch (error) {
    $("diffExcelHint").textContent = error.message;
  }
}

async function highlightDiffExcel() {
  if (!(Array.isArray(diffExcelState.previewRecords) && diffExcelState.previewRecords.length)) {
    $("diffHighlightHint").textContent = "当前没有可标记的预览结果。";
    return;
  }
  $("diffHighlightHint").textContent = "标记中...";
  try {
    const data = await api("/api/diff-excel/highlight", {
      method: "POST",
      body: JSON.stringify({
        cache_file: diffExcelState.cacheFile,
        query: "",
        target: $("diffMarkTarget").value,
        color_hex: $("diffHighlightColor").value || "#FFD966",
      }),
    });
    $("diffHighlightHint").textContent = `已标记 ${Number(data.changed_cells || 0)} 个单元格，涉及 ${Number(data.workbook_count || 0)} 个文件。`;
  } catch (error) {
    $("diffHighlightHint").textContent = error.message;
  }
}

async function startTask() {
  const payload = {
    folder_path: $("folderPath").value.trim(),
    header_name: $("headerName").value.trim(),
    source_language: $("sourceLanguage").value.trim() || "中文",
    file_type: "",
    export_review_sheet: true,
    extraction_mode: $("extractionMode").value,
    single_item_char_limit: 500,
    batch_request_char_limit: Number($("recallLimit").value || 3000),
    resume: false,
  };
  $("errorPanel").hidden = true;
  await saveSettings();
  await api("/api/tasks/start", {
    method: "POST",
    body: JSON.stringify(payload),
  });
  await refreshStatus();
}

async function resumeTask() {
  $("errorPanel").hidden = true;
  await api("/api/tasks/resume", { method: "POST", body: "{}" });
  await refreshStatus();
}

async function stopTask() {
  if (!window.confirm("确定停止当前任务吗？停止后可以继续上次任务。")) return;
  await api("/api/tasks/stop", { method: "POST", body: "{}" });
  await refreshStatus();
}

async function clearRuntimeCache() {
  if (!window.confirm("清空缓存后将无法继续上次任务，确定要清空吗？")) return;
  await api("/api/tasks/clear-cache", { method: "POST", body: "{}" });
  await refreshStatus();
}

function renderResultSummary(data) {
  latestResultFile = data.exists && data.output_file ? data.output_file : "";
  $("resultOutputFile").textContent = data.output_file || "暂无输出";
  $("resultOutputFileMirror").textContent = data.output_file || "暂无输出";
  $("resultState").textContent = data.error
    ? data.error
    : data.exists
      ? "结果文件已生成"
      : "等待任务完成";
  $("resultStateMirror").textContent = $("resultState").textContent;
  $("resultTermCount").textContent = data.term_library_count || 0;
  $("resultTermCountMirror").textContent = data.term_library_count || 0;
  $("resultFailureCount").textContent = data.failure_count || 0;
  $("resultFailureCountMirror").textContent = data.failure_count || 0;
  $("resultRegexCount").textContent = data.nontrans_regex_count || 0;
  $("resultRegexCountMirror").textContent = data.nontrans_regex_count || 0;
  $("downloadResultButton").disabled = !latestResultFile;
  $("openResultFolderButton").disabled = !latestResultFile;
}

async function refreshResults(outputFile = "") {
  const query = outputFile ? `?output_file=${encodeURIComponent(outputFile)}` : "";
  const data = await api(`/api/results/summary${query}`);
  renderResultSummary(data);
}

function formatProgressPercent(current, total, isRunning) {
  const safeCurrent = Math.max(0, Number(current || 0));
  const safeTotal = Math.max(0, Number(total || 0));
  if (safeTotal <= 0) {
    return isRunning ? "0%" : "未开始";
  }
  const percent = Math.min(100, Math.max(0, Math.round((safeCurrent / safeTotal) * 100)));
  return `${percent}%`;
}

function getPreprocessProgressState(data) {
  const current = Math.max(0, Number(data.progress_current || 0));
  const total = Math.max(0, Number(data.progress_total || 0));
  if (data.is_running) return { label: "运行中", className: "is-running" };
  if (data.last_error) return { label: "失败", className: "is-failed" };
  if (total > 0 && current >= total) return { label: "已完成", className: "is-completed" };
  if (data.can_resume) return { label: "已暂停", className: "is-paused" };
  return { label: "空闲", className: "" };
}

function renderPreprocessRecentEvents(data) {
  const container = $("preprocessRecentEvents");
  if (!container) return;
  const logs = Array.isArray(data.logs) ? data.logs.filter(Boolean).slice(-3).reverse() : [];
  container.replaceChildren();
  if (!logs.length) {
    const empty = document.createElement("span");
    empty.className = "preprocess-events-empty";
    empty.textContent = data.message || "暂无任务事件";
    container.appendChild(empty);
    return;
  }
  logs.forEach((entry, index) => {
    const item = document.createElement("div");
    item.className = "preprocess-event";
    const marker = document.createElement("i");
    marker.className = "preprocess-event-marker";
    const copy = document.createElement("span");
    copy.textContent = String(entry);
    item.append(marker, copy);
    if (index === 0) item.classList.add("is-latest");
    container.appendChild(item);
  });
}

function renderPreprocessProgress(data) {
  const current = Math.max(0, Number(data.progress_current || 0));
  const total = Math.max(0, Number(data.progress_total || 0));
  const state = getPreprocessProgressState(data);
  const stateElement = $("preprocessProgressState");
  stateElement.textContent = state.label;
  stateElement.className = `preprocess-status-pill ${state.className}`.trim();
  $("progressText").textContent = `${current} / ${total}`;
  $("preprocessProgressStage").textContent = data.stage_label || data.stage || "未启动";
  $("preprocessProgressMessage").textContent = data.message || (data.is_running ? "任务正在处理。" : "等待开始任务。");
  $("preprocessProgressPercent").textContent = formatProgressPercent(current, total, data.is_running);
  $("preprocessLlmRequests").textContent = Number((data.stats || {}).llm_request_count || 0);
  $("progressBar").style.width = total > 0 ? `${Math.min(100, Math.round((current / total) * 100))}%` : "0";
  renderPreprocessRecentEvents(data);
}

async function refreshStatus() {
  const data = await api("/api/status");
  const preprocessState = getPreprocessProgressState(data);
  setTaskStatus("overviewPage", {
    active: Boolean(data.is_running),
    taskLabel: "文本预处理工具",
    pill: preprocessState.label,
    pillClass: data.is_running ? "running" : data.last_error ? "failed" : "",
    stageLabel: data.stage_label || data.stage || "未启动",
    message: data.message || (data.is_running ? "任务运行中" : "等待开始任务"),
  });
  renderCurrentTaskStatus();
  renderPreprocessProgress(data);
  $("batchText").textContent = `${data.current_batch || 0} / ${data.total_batches || 0}`;
  $("successText").textContent = data.success_count || 0;
  $("failureText").textContent = data.failure_count || 0;
  $("retryText").textContent = data.retry_count || 0;
  $("concurrencyText").textContent = data.current_concurrency || 0;
  const stats = data.stats || {};
  $("statSourceRecords").textContent = stats.source_record_count || 0;
  $("statNontransCandidateRecords").textContent = stats.nontrans_candidate_record_count || 0;
  $("statNontransElements").textContent = stats.nontrans_element_count || 0;
  $("statNontransRegexRows").textContent = stats.nontrans_regex_row_count || 0;
  $("statSegments").textContent = stats.segment_count || 0;
  $("statCandidates").textContent = stats.candidate_count || 0;
  $("statApproved").textContent = stats.approved_count || 0;
  $("statRecallableRecords").textContent = stats.term_recall_recallable_record_count || 0;
  $("statUniqueRecallTexts").textContent = stats.term_recall_unique_text_count || 0;
  $("statDedupedRecords").textContent = stats.term_recall_deduped_record_count || 0;
  $("statDedupeSavingsPercent").textContent = `${stats.term_recall_dedupe_savings_percent || 0}%`;
  $("statRecallChunkBatches").textContent = stats.term_recall_chunk_batch_count || 0;
  $("statProtectedChanged").textContent = stats.nontrans_protected_changed_record_count || 0;
  $("statCleanLostMeaningful").textContent = stats.term_recall_clean_lost_meaningful_record_count || 0;
  $("statNumericNormalized").textContent = stats.term_recall_numeric_normalized_count || 0;
  $("statLlmRequests").textContent = stats.llm_request_count || 0;
  $("statLlmLatencyTotal").textContent = stats.llm_latency_ms_total || 0;
  $("statLlmLatencyAvg").textContent = stats.llm_latency_ms_avg || 0;
  $("statLlmPromptChars").textContent = stats.llm_prompt_char_count || 0;
  $("statLlmTokens").textContent = stats.llm_total_token_count || 0;
  $("outputFile").value = data.output_file || "";
  $("lastError").value = data.last_error || "";
  $("errorPanel").hidden = !data.last_error;
  $("errorPanelText").textContent = data.last_error || "";
  $("logBox").textContent = (data.logs && data.logs.length) ? data.logs.join("\n") : "暂无日志";
  $("startButton").disabled = Boolean(data.is_running);
  $("resumeButton").disabled = Boolean(data.is_running) || !data.can_resume;
  $("stopButton").disabled = !data.is_running;
  $("clearCacheButton").disabled = Boolean(data.is_running) || !data.can_resume;
  if (data.output_file) {
    await refreshResults(data.output_file);
  } else {
    renderResultSummary({});
  }
}

$("saveModelConnectionButton").addEventListener("click", saveModelConnection);
$("saveSettingsButton").addEventListener("click", saveSettings);
$("savePromptTemplatesButton").addEventListener("click", savePromptTemplates);
$("resetPromptTemplatesButton").addEventListener("click", resetPromptTemplates);
$("addAsciiPatternButton").addEventListener("click", addAsciiPattern);
$("saveAsciiPatternsButton").addEventListener("click", saveAsciiPatterns);
$("addBuiltinRuleButton").addEventListener("click", addBuiltinRule);
$("saveBuiltinRulesButton").addEventListener("click", saveBuiltinRules);
$("updateNoticeButton").addEventListener("click", openAppUpdateModal);
$("pendingRuleNoticeButton").addEventListener("click", openPendingRuleModal);
$("closeAppUpdateModalButton").addEventListener("click", closeAppUpdateModal);
$("cancelAppUpdateButton").addEventListener("click", closeAppUpdateModal);
$("confirmAppUpdateButton").addEventListener("click", startAppUpdate);
$("closePendingRuleModalButton").addEventListener("click", closePendingRuleModal);
$("pendingRuleSelectAllButton").addEventListener("click", () => setAllPendingRuleSelection(true));
$("pendingRuleClearSelectionButton").addEventListener("click", () => setAllPendingRuleSelection(false));
$("confirmPendingRuleImportButton").addEventListener("click", importPendingRules);
$("chooseFolderButton").addEventListener("click", chooseFolder);
$("chooseCrossExcelFolderButton").addEventListener("click", chooseCrossExcelFolder);
$("chooseDiffPathAButton").addEventListener("click", () => chooseDiffFolder("A").catch((error) => {
  $("diffExcelHint").textContent = error.message;
}));
$("chooseDiffPathBButton").addEventListener("click", () => chooseDiffFolder("B").catch((error) => {
  $("diffExcelHint").textContent = error.message;
}));
$("chooseDiffFileAButton").addEventListener("click", () => chooseDiffFile("A").catch((error) => {
  $("diffExcelHint").textContent = error.message;
}));
$("chooseDiffFileBButton").addEventListener("click", () => chooseDiffFile("B").catch((error) => {
  $("diffExcelHint").textContent = error.message;
}));
$("scanButton").addEventListener("click", scanFolder);
$("scanCrossExcelButton").addEventListener("click", scanCrossExcelFolder);
$("searchCrossExcelButton").addEventListener("click", searchCrossExcel);
$("mergeCrossExcelButton").addEventListener("click", mergeCrossExcel);
$("startDiffExcelButton").addEventListener("click", startDiffExcel);
$("clearDiffExcelButton").addEventListener("click", clearDiffExcelState);
$("diffPreviewPreviousButton").addEventListener("click", () => changeDiffPreviewPage(-1).catch((error) => { $("diffExcelHint").textContent = error.message; }));
$("diffPreviewNextButton").addEventListener("click", () => changeDiffPreviewPage(1).catch((error) => { $("diffExcelHint").textContent = error.message; }));
$("diffCompareMode").addEventListener("change", syncDiffCompareModeUi);
$("diffPathA").addEventListener("input", resetDiffFieldMatchSettings);
$("diffPathB").addEventListener("input", resetDiffFieldMatchSettings);
$("diffPathA").addEventListener("change", resetDiffFieldMatchSettings);
$("diffPathB").addEventListener("change", resetDiffFieldMatchSettings);
$("openDiffFieldSettingsButton").addEventListener("click", () => openDiffFieldSettings().catch((error) => {
  $("diffFieldSettingsHint").textContent = error.message;
}));
$("closeDiffFieldSettingsButton").addEventListener("click", closeDiffFieldSettings);
$("cancelDiffFieldSettingsButton").addEventListener("click", closeDiffFieldSettings);
$("saveDiffFieldSettingsButton").addEventListener("click", saveDiffFieldSettings);
$("diffReferenceField").addEventListener("change", () => {
  diffFieldMatchSettings.referenceField = $("diffReferenceField").value.trim();
  diffFieldMatchSettings.compareFields = diffFieldMatchSettings.compareFields.filter((field) => field !== diffFieldMatchSettings.referenceField);
  renderDiffFieldSettings();
});
$("exportDiffExcelButton").addEventListener("click", exportDiffExcel);
$("highlightDiffExcelButton").addEventListener("click", highlightDiffExcel);
$("diffHighlightColor").addEventListener("input", refreshDiffPresetColorButtons);
document.querySelectorAll(".preset-color-btn").forEach((button) => {
  button.addEventListener("click", () => {
    const color = String(button.dataset.color || "").trim();
    if (!color) return;
    $("diffHighlightColor").value = color;
    refreshDiffPresetColorButtons();
  });
});
$("selectAllCrossHeadersButton").addEventListener("click", () => setCrossExcelHeaderSelection(true));
$("clearCrossHeadersButton").addEventListener("click", () => setCrossExcelHeaderSelection(false));
$("newReviewConversationButton").addEventListener("click", () => createReviewConversation().catch(showReviewConversationError));
$("renameReviewConversationButton").addEventListener("click", beginRenameCurrentReviewConversation);
$("reviewConversationTitleInput").addEventListener("keydown", (event) => {
  if (event.key === "Enter") {
    event.preventDefault();
    saveCurrentReviewConversationTitle().catch(showReviewConversationError);
  } else if (event.key === "Escape") {
    event.preventDefault();
    cancelRenameCurrentReviewConversation();
  }
});
$("deleteReviewConversationButton").addEventListener("click", () => deleteCurrentReviewConversation().catch(showReviewConversationError));
$("saveReviewAttachmentMappingButton").addEventListener("click", () => saveReviewAttachmentMapping().catch((error) => {
  $("reviewAttachmentMappingHint").textContent = error?.message || String(error || "保存失败");
}));
$("editReviewAttachmentMappingButton").addEventListener("click", () => openReviewAttachmentMappingEditor().catch((error) => {
  $("reviewAttachmentMappingHint").textContent = error?.message || String(error || "读取结构失败");
}));
document.querySelectorAll('input[name="reviewAttachmentMappingMode"]').forEach((input) => {
  input.addEventListener("change", updateReviewAttachmentMappingControls);
});
$("reviewUploadChip").addEventListener("click", () => chooseReviewConversationFiles().catch(showReviewConversationError));
$("reviewTermBaseChip").addEventListener("click", toggleReviewTermBasePopover);
$("uploadReviewTermBaseButton").addEventListener("click", () => {
  closeReviewTermBasePopover();
  $("reviewTermBaseFileInput").click();
});
$("clearReviewTermBaseButton").addEventListener("click", () => selectReviewTermBase("").catch(showReviewConversationError));
$("reviewTermBaseFileInput").addEventListener("change", () => {
  const file = $("reviewTermBaseFileInput").files?.[0];
  uploadReviewTermBase(file).catch(showReviewConversationError).finally(() => ($("reviewTermBaseFileInput").value = ""));
});
$("reviewConversationFileInput").addEventListener("change", () => {
  const files = $("reviewConversationFileInput").files;
  uploadReviewConversationFiles(files).catch(showReviewConversationError).finally(() => ($("reviewConversationFileInput").value = ""));
});
$("reviewComposer").addEventListener("dragover", (event) => {
  event.preventDefault();
  $("reviewComposer").classList.add("dragging");
});
$("reviewComposer").addEventListener("dragleave", () => $("reviewComposer").classList.remove("dragging"));
$("reviewComposer").addEventListener("drop", (event) => {
  event.preventDefault();
  $("reviewComposer").classList.remove("dragging");
  uploadReviewConversationFiles(event.dataTransfer?.files).catch(showReviewConversationError);
});
$("sendReviewConversationButton").addEventListener("click", () => sendReviewConversationMessage().catch(showReviewConversationError));
$("reviewAutoStart").addEventListener("change", (event) => {
  saveReviewConversationAutoStart(event.currentTarget.checked).catch((error) => {
    event.currentTarget.checked = Boolean(reviewConversationState.snapshot?.session?.auto_start);
    showReviewConversationError(error);
  });
});
$("reviewComposerInput").addEventListener("keydown", (event) => {
  if (event.key === "Enter" && !event.shiftKey && !event.isComposing) {
    event.preventDefault();
    if ($("sendReviewConversationButton").disabled) return;
    sendReviewConversationMessage().catch(showReviewConversationError);
  }
});
$("reviewPromptChip").addEventListener("click", () => openReviewConversationPromptDialog().catch(showReviewConversationError));
$("reviewSourceLanguageChip").addEventListener("click", (event) => openReviewLanguagePopover("source", event.currentTarget));
$("reviewTargetLanguageChip").addEventListener("click", (event) => openReviewLanguagePopover("target", event.currentTarget));
$("reviewLanguageSearch").addEventListener("input", renderReviewLanguageOptions);
$("closeReviewLanguagePopoverButton").addEventListener("click", closeReviewLanguagePopover);
document.addEventListener("pointerdown", (event) => {
  const popover = $("reviewLanguagePopover");
  if (!popover.classList.contains("hidden") && !popover.contains(event.target)
    && !$("reviewSourceLanguageChip").contains(event.target) && !$("reviewTargetLanguageChip").contains(event.target)) {
    closeReviewLanguagePopover();
  }
  const termPopover = $("reviewTermBasePopover");
  if (!termPopover.classList.contains("hidden") && !termPopover.contains(event.target)
    && !$("reviewTermBaseChip").contains(event.target)) {
    closeReviewTermBasePopover();
  }
});
$("promptDialogTemplateSelect").addEventListener("change", async () => {
  const templateId = $("promptDialogTemplateSelect").value;
  if (!templateId) return;
  const data = await api(`/api/ai-review/prompt-templates/${encodeURIComponent(templateId)}`);
  reviewConversationState.promptTemplateId = templateId;
  fillAiReviewPromptDialog(data.template || {});
  updateReviewComposerChips();
  await saveReviewConversationComposerSettings();
});
$("chooseReviewFileButton").addEventListener("click", () => chooseAiReviewFile().catch((error) => {
  $("reviewTaskHint").textContent = error.message;
}));
$("reviewFileInput").addEventListener("change", () => {
  const file = $("reviewFileInput").files?.[0];
  uploadAiReviewFile(file).catch((error) => {
    $("reviewTaskHint").textContent = error.message;
    $("reviewFileHint").textContent = error.message;
  }).finally(() => {
    $("reviewFileInput").value = "";
  });
});
$("openExcelMappingButton").addEventListener("click", () => openAiReviewExcelMappingDialog().catch((error) => {
  $("reviewTaskHint").textContent = error.message;
}));
$("startReviewButton").addEventListener("click", () => startAiReviewTask().catch((error) => {
  $("reviewTaskHint").textContent = error.message;
  setAiReviewTaskStatus({
    active: false,
    pill: "失败",
    pillClass: "failed",
    stageLabel: "审校失败",
    message: error.message,
  });
}));
$("openReviewSettingsButton").addEventListener("click", () => setPage("aiReviewSettingsPage"));
$("openReviewForbiddenButton").addEventListener("click", () => setPage("aiReviewForbiddenPage"));
$("closeToolGuideDialogButton").addEventListener("click", () => $("toolGuideDialog").close());
$("feedbackEntryButton").addEventListener("click", openFeedbackModal);
$("closeFeedbackModalButton").addEventListener("click", closeFeedbackModal);
$("cancelFeedbackButton").addEventListener("click", closeFeedbackModal);
$("chooseFeedbackScreenshotButton").addEventListener("click", () => $("feedbackScreenshotInput").click());
$("feedbackScreenshotDropzone").addEventListener("click", () => {
  $("feedbackScreenshotDropzone").focus();
  $("feedbackSubmitHint").textContent = "现在可以直接按 Ctrl+V 粘贴截图，或把图片拖到这里。";
});
$("feedbackScreenshotInput").addEventListener("change", () => {
  const file = $("feedbackScreenshotInput").files?.[0];
  setFeedbackScreenshotFile(file || null);
});
$("feedbackScreenshotDropzone").addEventListener("dragover", (event) => {
  event.preventDefault();
  $("feedbackScreenshotDropzone").classList.add("drag-over");
});
$("feedbackScreenshotDropzone").addEventListener("dragleave", () => {
  $("feedbackScreenshotDropzone").classList.remove("drag-over");
});
$("feedbackScreenshotDropzone").addEventListener("drop", (event) => {
  event.preventDefault();
  $("feedbackScreenshotDropzone").classList.remove("drag-over");
  const files = Array.from(event.dataTransfer?.files || []);
  const imageFile = files.find((file) => String(file.type || "").startsWith("image/"));
  if (!imageFile) {
    $("feedbackSubmitHint").textContent = "这里只能接收图片文件。";
    return;
  }
  setFeedbackScreenshotFile(imageFile);
  $("feedbackSubmitHint").textContent = "已读取拖入的截图。";
});
$("openFeedbackLogButton").addEventListener("click", async () => {
  try {
    await api("/api/feedback/open-log", { method: "POST", body: "{}" });
  } catch (error) {
    $("feedbackSubmitHint").textContent = error.message;
  }
});
$("submitFeedbackButton").addEventListener("click", submitFeedback);
$("enableAiReview").addEventListener("change", updateAiReviewModeVisibility);
$("enableDirectionalReview").addEventListener("change", updateAiReviewModeVisibility);
$("reviewAiLimit").addEventListener("change", () => saveReviewSettings().catch((error) => {
  $("reviewSettingsHint").textContent = error.message;
}));
$("editPromptButton").addEventListener("click", () => openAiReviewPromptDialog().catch((error) => {
  $("reviewSettingsHint").textContent = error.message;
}));
$("newPromptButton").addEventListener("click", newAiReviewPromptTemplate);
$("savePromptButton").addEventListener("click", () => saveAiReviewPromptTemplate().catch((error) => {
  $("reviewSettingsHint").textContent = error.message;
}));
$("resetPromptButton").addEventListener("click", () => resetAiReviewPromptTemplate().catch((error) => {
  $("reviewSettingsHint").textContent = error.message;
}));
$("deletePromptButton").addEventListener("click", () => deleteAiReviewPromptTemplate().catch((error) => {
  $("reviewSettingsHint").textContent = error.message;
}));
$("cancelPromptDialogButton").addEventListener("click", () => $("promptDialog").close());
$("closePromptDialogButton").addEventListener("click", () => $("promptDialog").close());
$("editDirectionalButton").addEventListener("click", () => openDirectionalDialog().catch((error) => {
  $("reviewSettingsHint").textContent = error.message;
}));
$("newDirectionalButton").addEventListener("click", newDirectionalTemplate);
$("addDirectionalItemButton").addEventListener("click", () => appendDirectionalEditorItem("", true));
$("saveDirectionalButton").addEventListener("click", () => saveDirectionalTemplateFromDialog().catch((error) => {
  $("reviewSettingsHint").textContent = error.message;
}));
$("cancelDirectionalDialogButton").addEventListener("click", () => $("directionalDialog").close());
$("closeDirectionalDialogButton").addEventListener("click", () => $("directionalDialog").close());
$("editForbiddenButton").addEventListener("click", () => openForbiddenDialog().catch((error) => {
  $("reviewForbiddenHint").textContent = error.message;
}));
$("newForbiddenButton").addEventListener("click", newForbiddenTemplate);
$("saveForbiddenButton").addEventListener("click", () => saveForbiddenTemplateFromDialog().catch((error) => {
  $("reviewForbiddenHint").textContent = error.message;
}));
$("cancelForbiddenDialogButton").addEventListener("click", () => $("forbiddenDialog").close());
$("closeForbiddenDialogButton").addEventListener("click", () => $("forbiddenDialog").close());
$("applyExcelMappingButton").addEventListener("click", () => applyAiReviewExcelMapping().catch((error) => {
  $("reviewTaskHint").textContent = error.message;
}));
$("saveExcelMappingPresetButton").addEventListener("click", () => {
  try {
    openAiReviewExcelMappingPresetDialog();
  } catch (error) {
    $("excelMappingPresetHint").textContent = error.message;
  }
});
$("confirmExcelMappingPresetButton").addEventListener("click", () => saveAiReviewExcelMappingPreset().catch((error) => {
  $("excelMappingPresetSaveHint").textContent = error.message;
}));
$("saveReviewAgentSettingsButton").addEventListener("click", () => saveReviewSettings().catch((error) => {
  $("reviewSettingsHint").textContent = error.message;
}));
$("excelMappingPresetNameInput").addEventListener("keydown", (event) => {
  if (event.key !== "Enter") return;
  event.preventDefault();
  saveAiReviewExcelMappingPreset().catch((error) => {
    $("excelMappingPresetSaveHint").textContent = error.message;
  });
});
$("excelMappingPresetSelect").addEventListener("change", () => applyAiReviewExcelMappingPreset().catch((error) => {
  $("excelMappingPresetHint").textContent = error.message;
}));
$("deleteExcelMappingPresetButton").addEventListener("click", () => deleteAiReviewExcelMappingPreset().catch((error) => {
  $("excelMappingPresetHint").textContent = error.message;
}));
$("cancelExcelMappingPresetDialogButton").addEventListener("click", () => $("excelMappingPresetDialog").close());
$("closeExcelMappingPresetDialogButton").addEventListener("click", () => $("excelMappingPresetDialog").close());
$("cancelExcelMappingDialogButton").addEventListener("click", () => $("excelMappingDialog").close());
$("closeExcelMappingDialogButton").addEventListener("click", () => $("excelMappingDialog").close());
$("openOutputDirButton").addEventListener("click", () => openAiReviewOutputDir().catch((error) => {
  $("reviewTaskHint").textContent = error.message;
}));
$("openOutputFileButton").addEventListener("click", () => openAiReviewOutputFile().catch((error) => {
  $("reviewTaskHint").textContent = error.message;
}));
$("openReviewDetailButton").addEventListener("click", () => openReviewDetailDialog().catch((error) => {
  $("reviewTaskHint").textContent = error.message;
}));
$("closeReviewDetailButton").addEventListener("click", closeReviewDetailDialog);
$("reviewDetailOverlay").addEventListener("click", (event) => {
  if (event.target === $("reviewDetailOverlay")) {
    closeReviewDetailDialog();
  }
});
$("closeReviewFollowupButton").addEventListener("click", closeReviewFollowupDialog);
$("reviewFollowupOverlay").addEventListener("click", (event) => {
  if (event.target === $("reviewFollowupOverlay")) {
    closeReviewFollowupDialog();
  }
});
$("sendReviewFollowupButton").addEventListener("click", () => sendReviewFollowupMessage().catch((error) => {
  $("reviewFollowupHint").textContent = error.message;
  $("sendReviewFollowupButton").disabled = false;
}));
$("reviewFollowupInput").addEventListener("keydown", (event) => {
  if (event.key === "Enter" && !event.shiftKey) {
    event.preventDefault();
    sendReviewFollowupMessage().catch((error) => {
      $("reviewFollowupHint").textContent = error.message;
      $("sendReviewFollowupButton").disabled = false;
    });
  }
});
$("startButton").addEventListener("click", startTask);
$("resumeButton").addEventListener("click", resumeTask);
$("stopButton").addEventListener("click", stopTask);
$("clearCacheButton").addEventListener("click", clearRuntimeCache);
$("crossExcelQuery").addEventListener("keydown", (event) => {
  if (event.key === "Enter") {
    event.preventDefault();
    searchCrossExcel();
  }
});
$("openCrossExcelOutputButton").addEventListener("click", async () => {
  if (!crossExcelOutputFile) return;
  await api(`/api/results/open-folder?output_file=${encodeURIComponent(crossExcelOutputFile)}`, {
    method: "POST",
    body: "{}",
  });
});
$("openCrossExcelOutputFileButton").addEventListener("click", async () => {
  if (!crossExcelOutputFile) return;
  await api(`/api/results/open-file?output_file=${encodeURIComponent(crossExcelOutputFile)}`, {
    method: "POST",
    body: "{}",
  });
});
$("openDiffOutputFolderButton").addEventListener("click", async () => {
  if (!diffExcelState.outputFile) return;
  await api(`/api/results/open-folder?output_file=${encodeURIComponent(diffExcelState.outputFile)}`, {
    method: "POST",
    body: "{}",
  });
});
$("openDiffOutputFileButton").addEventListener("click", async () => {
  if (!diffExcelState.outputFile) return;
  await api(`/api/results/open-file?output_file=${encodeURIComponent(diffExcelState.outputFile)}`, {
    method: "POST",
    body: "{}",
  });
});
$("downloadResultButton").addEventListener("click", async () => {
  if (!latestResultFile) return;
  await api(`/api/results/open-file?output_file=${encodeURIComponent(latestResultFile)}`, {
    method: "POST",
    body: "{}",
  });
});
$("openResultFolderButton").addEventListener("click", async () => {
  if (!latestResultFile) return;
  await api(`/api/results/open-folder?output_file=${encodeURIComponent(latestResultFile)}`, {
    method: "POST",
    body: "{}",
  });
});

document.querySelectorAll(".nav-link, .shortcut-button").forEach((button) => {
  button.addEventListener("click", () => setPage(button.dataset.pageTarget));
});

document.querySelectorAll(".tool-guide-card").forEach((button) => {
  button.addEventListener("click", () => renderToolGuide(button.dataset.toolGuide));
});

document.querySelectorAll(".subnav-link").forEach((button) => {
  button.addEventListener("click", () => setSubtab(button.dataset.subtabGroup, button.dataset.subtabTarget));
});
document.addEventListener("paste", (event) => {
  readClipboardScreenshot(event).catch((error) => {
    $("feedbackSubmitHint").textContent = error.message;
  });
});
$("feedbackOverlay").addEventListener("click", (event) => {
  if (event.target === $("feedbackOverlay")) {
    closeFeedbackModal();
  }
});

setPage("toolGuidePage");
renderCrossExcelHeaders([]);
renderCrossExcelSearchResults(null);
setCrossExcelOutput("");
refreshDiffPresetColorButtons();
syncDiffCompareModeUi();
updateAiReviewModeVisibility();
renderCurrentTaskStatus();
Promise.all([
  loadFeedbackStatus(),
  loadSettings(),
  loadAsciiPatterns(),
  loadPromptTemplates(),
  loadBuiltinRules(),
  loadPendingRules(),
  loadAppUpdateInfo(),
  Promise.all([loadAiReviewPromptTemplates(), loadReviewTermBases()]).then(() => loadReviewConversations()),
  loadAiReviewDirectionalTemplates(),
  loadAiReviewForbiddenTemplates(),
]).then(refreshStatus).catch((error) => {
  $("statusMessage").textContent = error.message;
});
setInterval(refreshStatus, 1500);
"""


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="127.0.0.1", port=8765, access_log=False)

