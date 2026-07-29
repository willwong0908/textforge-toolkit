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
    max_concurrency: Optional[int] = None
    disable_system_proxy: Optional[bool] = None
    extraction_mode: Optional[str] = None
    source_language: Optional[str] = None
    nontrans_chunk_char_limit: Optional[int] = None
    nontrans_placeholder_format: Optional[str] = None
    term_recall_batch_char_limit: Optional[int] = None
    term_review_batch_char_limit: Optional[int] = None
    term_review_max_context_chars: Optional[int] = None
    ai_review_batch_char_limit: Optional[int] = None
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
    app = FastAPI(title="AI Term Extractor WebUI")
    init_ai_review_db()

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
            "max_concurrency": provider_settings.max_concurrency if provider_settings else 1,
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
            if payload.max_concurrency is not None:
                provider_settings.max_concurrency = max(1, int(payload.max_concurrency))
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
        if payload.term_recall_enable_thinking is not None:
            recall["enable_thinking"] = bool(payload.term_recall_enable_thinking)
        settings.input_defaults["term_recall_stage_settings"] = recall

        review = dict(settings.input_defaults.get("term_review_stage_settings", {}) or {})
        if payload.term_review_batch_char_limit is not None:
            review["batch_request_char_limit"] = int(payload.term_review_batch_char_limit)
        if payload.term_review_max_context_chars is not None:
            review["max_context_chars"] = int(payload.term_review_max_context_chars)
        if payload.term_review_enable_thinking is not None:
            review["enable_thinking"] = bool(payload.term_review_enable_thinking)
        settings.input_defaults["term_review_stage_settings"] = review

        ai_review = dict(settings.input_defaults.get("ai_review_stage_settings", {}) or {})
        if payload.ai_review_batch_char_limit is not None:
            ai_review["batch_request_char_limit"] = int(payload.ai_review_batch_char_limit)
        if payload.ai_review_enable_thinking is not None:
            ai_review["enable_thinking"] = bool(payload.ai_review_enable_thinking)
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
        if payload.max_concurrency is not None:
            provider_settings.max_concurrency = max(1, int(payload.max_concurrency))
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
            result = run_diff_excel_compare_to_cache(
                payload.path_a,
                payload.path_b,
                compare_mode=payload.compare_mode,
                reference_field=payload.reference_field,
                compare_fields=payload.compare_fields,
                include_unmatched=payload.include_unmatched,
                ignore_case=False,
                trim_whitespace=False,
            )
            track_event("diff.compare.success")
            return result
        except Exception as exc:
            track_event("diff.compare.fail")
            raise HTTPException(status_code=400, detail=str(exc)) from exc

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
    async def diff_excel_preview(cache_file: str, query: str = "", limit: int = 1000):
        try:
            return read_diff_excel_cached_preview(cache_file, query=query, limit=limit)
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
            "max_concurrency": settings.get("max_concurrency", 6),
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
        preset = save_excel_mapping_preset(payload.name, payload.mapping, payload.id)
        return {"ok": True, "message": "Excel 列映射预设已保存", "preset": preset}

    @app.delete("/api/ai-review/excel-mapping-presets/{preset_id}")
    async def ai_review_delete_excel_mapping_preset(preset_id: str):
        try:
            delete_excel_mapping_preset(preset_id)
        except ValueError as exc:
            raise HTTPException(status_code=404, detail=str(exc)) from exc
        return {"ok": True, "message": "Excel 列映射预设已删除"}

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
            delete_ai_review_prompt_template(template_id)
        except ValueError as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc
        return {"ok": True, "message": "提示词模板已删除"}

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
            <button class="nav-link nav-link-sub" data-page-target="runDetailsPage">运行详情</button>
            <button class="nav-link nav-link-sub" data-page-target="resultsPage">结果</button>
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
            <button class="nav-link nav-link-sub" data-page-target="aiReviewForbiddenPage">禁用词</button>
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
      <div class="sidebar-status">
        <div class="status-block">
          <small class="status-caption">当前任务</small>
          <strong id="taskTypeLabel">文本预处理工具</strong>
        </div>
        <div class="status-block">
          <small class="status-caption">任务状态</small>
          <span id="statusPill" class="pill">空闲</span>
          <strong id="stageLabel">未启动</strong>
          <small id="statusMessage">等待开始任务</small>
        </div>
        <small class="service-tip">服务地址：<code>http://127.0.0.1:8765</code></small>
      </div>
      <div class="sidebar-actions-panel">
        <button id="feedbackEntryButton" class="feedback-entry-button" type="button">我要反馈</button>
      </div>
    </aside>

    <main class="content">
      <header class="hero">
        <div>
          <p class="eyebrow">WebUI</p>
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
        <div class="inline-notice">本工具会统计匿名的功能触发次数，不收集文本内容、文件名、路径、账号或密钥等敏感信息。</div>
        <div class="tool-guide-grid">
          <button class="tool-guide-card" type="button" data-tool-guide="textPreprocess">文本预处理工具</button>
          <button class="tool-guide-card" type="button" data-tool-guide="diffExcel">Diff 工具</button>
          <button class="tool-guide-card" type="button" data-tool-guide="aiReview">AI 审校工具</button>
          <button class="tool-guide-card" type="button" data-tool-guide="crossExcel">跨Excel搜索与合并</button>
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

          <section class="card">
            <div class="card-title">
              <h3>当前状态</h3>
              <p></p>
            </div>
            <div class="metrics hero-metrics">
              <div><span>进度</span><strong id="progressText">0 / 0</strong></div>
              <div><span>批次</span><strong id="batchText">0 / 0</strong></div>
              <div><span>成功</span><strong id="successText">0</strong></div>
              <div><span>失败</span><strong id="failureText">0</strong></div>
              <div><span>重试</span><strong id="retryText">0</strong></div>
              <div><span>并发</span><strong id="concurrencyText">0</strong></div>
            </div>
            <div class="progress"><span id="progressBar"></span></div>
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
            <label>最大并发<input id="maxConcurrency" type="number" min="1" value="6" /></label>
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
            <label>单次处理长度<input id="nontransLimit" type="number" min="200" value="3000" /></label>
            <label class="check"><input id="nontransThinking" type="checkbox" /> 深度思考</label>
          </section>
          <section class="card compact-card">
            <div class="card-title">
              <h3>术语召回阶段</h3>
              <p></p>
            </div>
            <label>单次处理长度<input id="recallLimit" type="number" min="200" value="3000" /></label>
            <label class="check"><input id="recallThinking" type="checkbox" /> 深度思考</label>
          </section>
          <section class="card compact-card">
            <div class="card-title">
              <h3>术语校验阶段</h3>
              <p></p>
            </div>
            <label>单次处理长度<input id="reviewLimit" type="number" min="200" value="3000" /></label>
            <label>上下文长度<input id="reviewContextLimit" type="number" min="50" max="2000" value="220" /></label>
            <label class="check"><input id="reviewThinking" type="checkbox" /> 深度思考</label>
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

      <section id="runDetailsPage" class="page-section">
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

      <section id="resultsPage" class="page-section">
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
            <p>左侧显示 A 中被删掉的内容，右侧显示 B 中新增的内容。</p>
          </div>
          <div class="pattern-table-wrap">
            <table class="pattern-table">
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
        <div class="grid dashboard-grid">
          <section class="card">
            <div class="card-title">
              <h3>读取文件</h3>
              <p>支持 Excel 与 XLIFF。读取后可预览前 5 条并开始审校。</p>
            </div>
            <div class="grid two">
              <div class="field">
                <span class="field-label">当前文件</span>
                <div class="file-card-row">
                  <div id="reviewFilePath" class="file-card empty">请选择待审校文件</div>
                  <button id="chooseReviewFileButton" class="mini-button" type="button">选择文件</button>
                  <input id="reviewFileInput" type="file" accept=".xlsx,.xlsm,.xlf,.xliff" hidden />
                </div>
              </div>
              <div class="cross-summary-box">
                <span>读取状态</span>
                <strong id="reviewBatchCount">0</strong>
                <small id="reviewFileHint">尚未读取文件。</small>
              </div>
            </div>
            <div class="actions">
              <button id="openExcelMappingButton" class="secondary" type="button" disabled>导入文本</button>
            </div>
            <p id="excelMappingSummary" class="hint"></p>
          </section>
        </div>

        <div class="grid dashboard-grid">
          <section class="card">
            <div class="card-title">
              <h3>读取预览</h3>
              <p>预览前 5 条，检查原文和译文是否对应。</p>
            </div>
            <div class="pattern-table-wrap">
              <table class="pattern-table">
                <thead>
                  <tr>
                    <th>来源文件</th>
                    <th>sheet / segment ID</th>
                    <th>原始行号</th>
                    <th>原文</th>
                    <th>译文</th>
                    <th>提示</th>
                  </tr>
                </thead>
                <tbody id="previewBody">
                  <tr><td colspan="6" class="empty-cell">暂无预览</td></tr>
                </tbody>
              </table>
            </div>
          </section>
        </div>

        <div class="grid dashboard-grid">
          <section class="card">
            <div class="card-title">
              <h3>任务操作</h3>
              <p>确认预览无误后开始审校。</p>
            </div>
            <div class="grid two hidden">
              <label>原文语种<input id="sourceLanguageInput" placeholder="例如 English" /></label>
              <label>译文语种<input id="targetLanguageInput" placeholder="例如 简体中文" /></label>
            </div>
            <div class="actions">
              <button id="startReviewButton" class="primary" type="button">开始审校</button>
              <button id="openReviewSettingsButton" class="secondary" type="button">打开审校设置</button>
              <button id="openReviewForbiddenButton" class="secondary" type="button">打开禁用词</button>
            </div>
            <span id="reviewTaskHint" class="hint"></span>
          </section>

          <section class="card">
            <div class="card-title">
              <h3>审校进度</h3>
              <p id="reviewProgress">尚未开始</p>
            </div>
            <div class="metrics hero-metrics">
              <div><span>进度</span><strong id="reviewProgressPercent">未开始</strong></div>
              <div><span>已处理</span><strong id="reviewProgressCount">0 / 0</strong></div>
              <div><span>失败</span><strong id="reviewFailedCount">0</strong></div>
              <div><span>请求</span><strong id="reviewRequestedCount">0</strong></div>
            </div>
            <div class="progress"><span id="reviewProgressBar"></span></div>
            <div id="outputPanel" class="result-file hidden">
              <span>结果文件</span>
              <strong id="outputPath">暂无输出</strong>
            </div>
            <div class="actions">
              <button id="openOutputDirButton" class="secondary" type="button">打开输出目录</button>
              <button id="openOutputFileButton" class="secondary" type="button" disabled>打开结果文件</button>
            </div>
            <div class="review-log-wrap">
              <ol id="reviewLogList" class="review-log-list"></ol>
            </div>
          </section>
        </div>

        <section class="card">
          <div class="card-title">
            <div>
              <h3>审校结果预览</h3>
              <p>展示前 20 条结果，完整结果会自动保存为 Excel。</p>
            </div>
            <button id="openReviewDetailButton" class="secondary" type="button" disabled>详情</button>
          </div>
          <div class="pattern-table-wrap">
            <table class="pattern-table">
              <thead id="reviewResultHead">
                <tr>
                  <th>原文</th>
                  <th>译文</th>
                  <th>是否有问题</th>
                  <th>问题类型</th>
                  <th>问题说明</th>
                  <th>修改建议</th>
                </tr>
              </thead>
              <tbody id="reviewResultBody">
                <tr><td colspan="6" class="empty-cell">暂无审校结果</td></tr>
              </tbody>
            </table>
          </div>
        </section>
      </section>

      <section id="aiReviewSettingsPage" class="page-section">
        <div class="grid one">
          <section class="card">
            <div class="card-title">
              <h3>审校方式</h3>
              <p>选择审校模式、提示词模板和模型思考深度。</p>
            </div>
            <div class="actions review-toggle-row">
              <label class="check-line"><input id="enableAiReview" type="checkbox" checked /><span>启用 AI 审校</span></label>
              <label id="directionalReviewLine" class="check-line"><input id="enableDirectionalReview" type="checkbox" /><span>启用定向审校</span></label>
              <label class="check-line"><input id="reviewAiThinking" type="checkbox" /><span>深度思考</span></label>
            </div>
            <div class="grid two">
              <label>提示词模板<select id="promptTemplateSelect"></select></label>
              <label id="directionalTemplatePanel" class="hidden">定向审校模板<select id="directionalTemplateSelect"></select></label>
              <label>单次请求字符上限<input id="reviewAiLimit" type="number" min="200" value="3000" /></label>
            </div>
            <div class="actions">
              <button id="editPromptButton" class="secondary" type="button">编辑提示词</button>
              <button id="editDirectionalButton" class="secondary" type="button">编辑定向</button>
            </div>
            <div class="notice-list">
              <div>定向审校会按所选项目生成结果列。</div>
              <div>模型连接请在“模型设置”中调整。</div>
            </div>
            <span id="reviewSettingsHint" class="hint"></span>
          </section>
        </div>
      </section>

      <section id="aiReviewForbiddenPage" class="page-section">
        <div class="grid dashboard-grid">
          <section class="card">
            <div class="card-title">
              <h3>禁用词开关</h3>
              <p>禁用词会在译文中直接检查命中内容，可与 AI 审校一起使用。</p>
            </div>
            <div class="actions review-toggle-row">
              <label class="check-line"><input id="enableForbiddenCheck" type="checkbox" /><span>启用禁用词</span></label>
            </div>
            <div class="grid one">
              <label id="forbiddenTemplatePanel">禁用词模板<select id="forbiddenTemplateSelect"></select></label>
            </div>
            <div class="actions">
              <button id="editForbiddenButton" class="secondary" type="button">编辑禁用词</button>
            </div>
            <span id="reviewForbiddenHint" class="hint"></span>
          </section>
        </div>
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
      <div class="pattern-table-wrap review-detail-wrap">
        <table class="pattern-table review-detail-table">
          <thead id="reviewDetailHead">
            <tr><th>原文</th><th>原译文</th><th>修改建议</th></tr>
          </thead>
          <tbody id="reviewDetailBody">
            <tr><td colspan="3" class="empty-cell">暂无问题条目</td></tr>
          </tbody>
        </table>
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
          <h2>Excel 映射</h2>
          <p>按 sheet 选择原文列、译文列和信息列。</p>
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
      <div class="grid two hidden">
        <div class="field">
          <label for="excelMappingPresetSelect">映射预设</label>
          <select id="excelMappingPresetSelect"></select>
        </div>
        <div class="field">
          <label for="excelMappingPresetNameInput">预设名称</label>
          <input id="excelMappingPresetNameInput" type="text" placeholder="输入预设名称" />
        </div>
      </div>
      <div class="actions hidden">
        <button id="applyExcelMappingPresetButton" class="secondary" type="button">应用预设</button>
        <button id="saveExcelMappingPresetButton" class="secondary" type="button">保存预设</button>
        <button id="deleteExcelMappingPresetButton" class="danger" type="button">删除预设</button>
      </div>
      <div id="excelSheetTabs" class="sheet-tabs"></div>
      <div id="excelMappingColumns" class="excel-mapping-columns"></div>
      <div class="dialog-actions">
        <button id="applyExcelMappingButton" type="button">确认读取</button>
        <button id="cancelExcelMappingDialogButton" class="secondary" type="button">取消</button>
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
}
.review-detail-table {
  min-width: 1280px;
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
  grid-template-columns: repeat(2, minmax(0, 1fr));
  gap: 10px;
}
.diff-inline-card.diff-inline-single { grid-template-columns: minmax(0, 1fr); }
.diff-inline-side {
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
  word-break: break-word;
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
@media (max-width: 980px) {
  .shell { grid-template-columns: 1fr; }
  .sidebar { position: static; }
  .hero, .grid.two, .stage-grid, .metrics, .prompt-grid, .dashboard-grid, .scope-list, .compact-metrics, .result-metrics, .cross-row-grid { grid-template-columns: 1fr; }
  .sticky-actions { position: static; }
  .content { padding: 18px; }
  .hero-side { justify-items: stretch; }
  .modal-card { padding: 18px; }
  .modal-footer { align-items: stretch; }
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
  previewLimit: 1000,
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
let aiReviewIssueResults = [];
let aiReviewFollowupState = { taskId: "", resultId: "", item: null, messages: [] };
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
  aiReviewSettingsPage: { title: "审校设置", lede: "管理 AI 审校方式、定向审校模板与深度思考。" },
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
      ["基本用法", ["在“模型设置”中加载模型。", "进入“文本预处理工具”，选择输入目录和待提取列。", "选择运行模式，点击“开始提取”。", "完成后在“结果”页打开输出文件。"]],
      ["输出结果", ["术语库", "非译元素正则规则", "失败记录和任务日志"]],
    ],
  },
  aiReview: {
    title: "AI 审校工具",
    sections: [
      ["用途", "检查译文相对原文是否存在问题，适合做翻译质检、定向问题检查和禁用词检查。"],
      ["适合处理", ["Excel 双语表", "XLIFF 文件", "需要按问题类型输出审校结果的项目"]],
      ["基本用法", ["进入“AI 审校工具”，导入待审校文件。", "Excel 文件需要先确认原文列和译文列；XLIFF 会自动读取。", "在“审校设置”中选择普通审校或定向审校。", "点击“开始审校”，完成后打开结果文件。"]],
      ["常用设置", ["单次请求字符上限：控制每次发给模型的文本长度。", "深度思考：用于更复杂的审校，但会更慢。", "禁用词：单独检查译文中是否出现指定词。"]],
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
  currentPageId = pageId;
  document.querySelectorAll(".page-section").forEach((section) => {
    section.classList.toggle("active", section.id === pageId);
  });
  document.querySelectorAll(".nav-link").forEach((button) => {
    button.classList.toggle("active", button.dataset.pageTarget === pageId);
  });
  $("taskTypeLabel").textContent = PAGE_TASK_LABELS[pageId] || "文本预处理工具";
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
  $("taskTypeLabel").textContent = taskLabel || "文本预处理工具";
  $("statusPill").textContent = pillText || "空闲";
  $("statusPill").className = `pill ${pillClass || ""}`.trim();
  $("stageLabel").textContent = stageLabel || "未启动";
  $("statusMessage").textContent = message || "等待开始任务";
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
    max_concurrency: Number($("maxConcurrency").value || 1),
    disable_system_proxy: $("disableSystemProxy").checked,
    extraction_mode: $("extractionMode").value,
    source_language: $("sourceLanguage").value,
    nontrans_chunk_char_limit: Number($("nontransLimit").value || 3000),
    nontrans_placeholder_format: $("nontransPlaceholderFormat").value || "<{n}>",
    term_recall_batch_char_limit: Number($("recallLimit").value || 3000),
    term_review_batch_char_limit: Number($("reviewLimit").value || 3000),
    term_review_max_context_chars: Number($("reviewContextLimit").value || 220),
    nontrans_enable_thinking: $("nontransThinking").checked,
    term_recall_enable_thinking: $("recallThinking").checked,
    term_review_enable_thinking: $("reviewThinking").checked,
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
    max_concurrency: Number($("maxConcurrency").value || 1),
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
  $("diffModeLabel").textContent = String(meta.mode_label || "未开始");
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
  $("highlightDiffExcelButton").disabled = !(Array.isArray(diffExcelState.previewRecords) && diffExcelState.previewRecords.length);
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
}

function renderDiffExcelResultHead() {
  const head = $("diffExcelHead");
  const isFieldMatch = String(diffExcelState.meta?.mode_label || "") === "按字段匹配";
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

async function applyDiffExcelFilter() {
  if (!diffExcelState.cacheFile) {
    diffExcelState.previewRecords = [];
    diffExcelState.matchedCount = 0;
    diffExcelState.previewTruncated = false;
    renderDiffExcelResults();
    return;
  }
  const data = await api(`/api/diff-excel/preview?cache_file=${encodeURIComponent(diffExcelState.cacheFile)}&query=&limit=${encodeURIComponent(String(diffExcelState.previewLimit || 1000))}`);
  diffExcelState.previewRecords = Array.isArray(data.records) ? data.records : [];
  diffExcelState.matchedCount = Number(data.matched_count || 0);
  diffExcelState.previewTruncated = Boolean(data.preview_truncated);
  renderDiffExcelResults();
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
    previewLimit: 1000,
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
  aiReviewSheetNames.forEach((sheetName) => {
    aiReviewExcelMappingState[sheetName] = { sources: [], targets: {}, infos: {} };
  });
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
  aiReviewPromptTemplates.forEach((item) => {
    select.appendChild(new Option(String(item.name || ""), String(item.id || "")));
  });
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

function summarizeAiReviewExcelMapping(mapping) {
  const sheets = Array.isArray(mapping?.sheets) ? mapping.sheets : [];
  const count = sheets.reduce((total, sheet) => total + Number((sheet.mappings || []).length || 0), 0);
  return count > 0 ? `已配置 ${count} 组原文 / 译文列` : "";
}

function renderAiReviewExcelMappingPresets() {
  const select = $("excelMappingPresetSelect");
  select.innerHTML = "";
  select.appendChild(new Option("选择预设", ""));
  aiReviewExcelMappingPresets.forEach((item) => {
    select.appendChild(new Option(String(item.name || ""), String(item.id || "")));
  });
}

async function loadAiReviewExcelMappingPresets() {
  const data = await api("/api/ai-review/excel-mapping-presets");
  aiReviewExcelMappingPresets = Array.isArray(data.presets) ? data.presets : [];
  renderAiReviewExcelMappingPresets();
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
      syncAiReviewActiveSheetMapping();
      renderAiReviewExcelMappingDialog();
    });
    targetSelect.addEventListener("change", () => {
      clearAiReviewTaskHint();
      if (targetSelect.value !== "") {
        infoSelect.value = "";
        categoryInput.disabled = true;
      }
      syncAiReviewActiveSheetMapping();
    });
    infoSelect.addEventListener("change", () => {
      clearAiReviewTaskHint();
      if (infoSelect.value !== "") {
        targetSelect.value = "";
        categoryInput.disabled = false;
      } else {
        categoryInput.disabled = infoSelect.disabled;
      }
      syncAiReviewActiveSheetMapping();
    });
    categoryInput.addEventListener("input", () => {
      clearAiReviewTaskHint();
      syncAiReviewActiveSheetMapping();
    });
    container.appendChild(row);
  });
}

function renderAiReviewExcelMappingDialog() {
  renderAiReviewExcelSheetTabs();
  renderAiReviewExcelMappingColumns();
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

function validateAiReviewExcelMapping(mapping) {
  const errors = [];
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
  const nextState = {};
  aiReviewSheetNames.forEach((sheetName) => {
    nextState[sheetName] = { sources: [], targets: {}, infos: {} };
  });
  (mapping?.sheets || []).forEach((sheet) => {
    if (!nextState[sheet.sheet_name]) return;
    const validColumns = new Set((aiReviewColumnsBySheet[sheet.sheet_name] || []).map((column) => Number(column.index)));
    (sheet.mappings || []).forEach((item) => {
      if (!validColumns.has(Number(item.source_column)) || !validColumns.has(Number(item.target_column))) return;
      nextState[sheet.sheet_name].sources.push(Number(item.source_column));
      nextState[sheet.sheet_name].targets[Number(item.target_column)] = Number(item.source_column);
      (item.info_columns || []).forEach((info) => {
        if (!validColumns.has(Number(info.column))) return;
        const current = nextState[sheet.sheet_name].infos[Number(info.column)] || { sourceColumns: [], category: String(info.category || "") };
        if (!current.sourceColumns.includes(Number(item.source_column))) {
          current.sourceColumns.push(Number(item.source_column));
        }
        if (info.category && !current.category) {
          current.category = String(info.category);
        }
        nextState[sheet.sheet_name].infos[Number(info.column)] = current;
      });
    });
  });
  aiReviewExcelMappingState = nextState;
}

async function openAiReviewExcelMappingDialog() {
  if (!aiReviewBatch?.id || !aiReviewSheetNames.length) {
    throw new Error("请先读取 Excel 文件");
  }
  clearAiReviewTaskHint();
  $("mappingSourceLanguageInput").value = $("sourceLanguageInput").value || "";
  $("mappingTargetLanguageInput").value = $("targetLanguageInput").value || "";
  renderAiReviewExcelMappingDialog();
  $("excelMappingDialog").showModal();
}

async function applyAiReviewExcelMapping() {
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

async function saveAiReviewExcelMappingPreset() {
  const mapping = buildAiReviewExcelMapping();
  const errors = validateAiReviewExcelMapping(mapping);
  if (errors.length) {
    throw new Error(errors[0]);
  }
  const data = await api("/api/ai-review/excel-mapping-presets", {
    method: "POST",
    body: JSON.stringify({
      id: null,
      name: $("excelMappingPresetNameInput").value.trim() || "未命名映射预设",
      mapping,
    }),
  });
  await loadAiReviewExcelMappingPresets();
  $("excelMappingPresetSelect").value = String(data?.preset?.id || "");
}

async function applyAiReviewExcelMappingPreset() {
  const presetId = $("excelMappingPresetSelect").value;
  if (!presetId) return;
  const data = await api(`/api/ai-review/excel-mapping-presets/${encodeURIComponent(presetId)}`);
  applyAiReviewExcelMappingPresetToState(data?.preset?.mapping || {});
  $("excelMappingPresetNameInput").value = String(data?.preset?.name || "");
  renderAiReviewExcelMappingDialog();
}

async function deleteAiReviewExcelMappingPreset() {
  const presetId = $("excelMappingPresetSelect").value;
  if (!presetId) return;
  if (!window.confirm("确定删除这个映射预设吗？")) {
    return;
  }
  await api(`/api/ai-review/excel-mapping-presets/${encodeURIComponent(presetId)}`, {
    method: "DELETE",
  });
  $("excelMappingPresetNameInput").value = "";
  await loadAiReviewExcelMappingPresets();
}

function fillAiReviewPromptDialog(template = {}) {
  $("promptTemplateId").value = String(template.id || "");
  $("promptNameInput").value = String(template.name || "");
  $("systemPromptInput").value = String(template.system_prompt || "");
  $("userPromptInput").value = String(template.user_prompt || "");
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
  fillAiReviewPromptDialog({
    id: "",
    name: "新建模板",
    system_prompt: "",
    user_prompt: "{text}",
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
    }),
  });
  await loadAiReviewPromptTemplates();
  $("promptTemplateSelect").value = String(data?.template?.id || "");
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
  await api(`/api/ai-review/prompt-templates/${encodeURIComponent(templateId)}`, {
    method: "DELETE",
  });
  await loadAiReviewPromptTemplates();
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
  const head = $("reviewDetailHead");
  const body = $("reviewDetailBody");
  const extraHeaders = reviewDetailExtraHeaders(task);
  const headers = ["原文", "原译文", "修改建议", ...extraHeaders];
  head.innerHTML = "";
  const headRow = document.createElement("tr");
  headers.forEach((header) => {
    const th = document.createElement("th");
    th.textContent = header;
    headRow.appendChild(th);
  });
  head.appendChild(headRow);

  body.innerHTML = "";
  const items = Array.isArray(results) ? results : [];
  aiReviewIssueResults = items;
  $("reviewDetailSummary").textContent = items.length ? `共 ${items.length} 条问题` : "暂无问题条目";
  if (!items.length) {
    body.innerHTML = `<tr><td colspan="${headers.length}" class="empty-cell">暂无问题条目</td></tr>`;
    return;
  }
  items.forEach((item) => {
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
      if (index === extraValues.length - 1) {
        td.className = "review-detail-issue-cell";
      }
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
  $("maxConcurrency").value = data.max_concurrency || 1;
  $("disableSystemProxy").checked = data.disable_system_proxy !== false;
  $("sourceLanguage").value = data.source_language || "中文";
  $("extractionMode").value = data.extraction_mode || "terms";
  $("nontransPlaceholderFormat").value = data.nontrans_placeholder_format || "<{n}>";
  const nontrans = data.nontrans_stage_settings || {};
  const recall = data.term_recall_stage_settings || {};
  const review = data.term_review_stage_settings || {};
  const aiReview = data.ai_review_stage_settings || {};
  $("nontransLimit").value = nontrans.chunk_char_limit || 3000;
  $("recallLimit").value = recall.batch_request_char_limit || 3000;
  $("reviewLimit").value = review.batch_request_char_limit || 3000;
  $("reviewContextLimit").value = review.max_context_chars || 220;
  $("nontransThinking").checked = Boolean(nontrans.enable_thinking);
  $("recallThinking").checked = Boolean(recall.enable_thinking);
  $("reviewThinking").checked = Boolean(review.enable_thinking);
  if ($("reviewAiThinking")) {
    $("reviewAiThinking").checked = Boolean(aiReview.enable_thinking);
  }
  if ($("reviewAiLimit")) {
    $("reviewAiLimit").value = aiReview.batch_request_char_limit || 3000;
  }
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
    ai_review_batch_char_limit: Number($("reviewAiLimit").value || 3000),
    ai_review_enable_thinking: $("reviewAiThinking").checked,
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
    const data = await api("/api/diff-excel/compare", {
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
    diffExcelState.cacheFile = String(data.cache_file || "");
    diffExcelState.resultId = String(data.result_id || "");
    diffExcelState.previewRecords = Array.isArray(data.preview_records) ? data.preview_records : [];
    diffExcelState.meta = data.meta || null;
    diffExcelState.outputFile = "";
    diffExcelState.totalCount = Number(data.total_count || 0);
    diffExcelState.matchedCount = Number(data.total_count || 0);
    diffExcelState.previewLimit = Number(data.preview_limit || 1000);
    diffExcelState.previewTruncated = Boolean(data.preview_truncated);
    const meta = diffExcelState.meta || {};
    const fieldSummary = compareMode === "field_match"
      ? `有效参考行 ${Number(meta.valid_reference_rows || 0)}，已配对 ${Number(meta.paired_reference_rows || 0)}，未匹配 ${Number(meta.unmatched_reference_rows || 0)}，跳过字段 ${Number(meta.skipped_field_count || 0)}。`
      : "";
    $("diffExcelHint").textContent = diffExcelState.previewTruncated
      ? `比对完成，共找到 ${diffExcelState.totalCount} 处差异，当前仅预览前 ${diffExcelState.previewLimit} 条。`
      : `比对完成，共找到 ${diffExcelState.totalCount} 处差异。${fieldSummary}`;
    renderDiffExcelResults();
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

async function refreshStatus() {
  const data = await api("/api/status");
  setTaskStatus("overviewPage", {
    active: Boolean(data.is_running),
    taskLabel: "文本预处理工具",
    pill: data.is_running ? "运行中" : data.last_error ? "失败" : "空闲",
    pillClass: data.is_running ? "running" : data.last_error ? "failed" : "",
    stageLabel: data.stage_label || data.stage || "未启动",
    message: data.message || (data.is_running ? "任务运行中" : "等待开始任务"),
  });
  renderCurrentTaskStatus();
  $("progressText").textContent = formatProgressPercent(data.progress_current, data.progress_total, data.is_running);
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
  const total = Number(data.progress_total || 0);
  const current = Number(data.progress_current || 0);
  $("progressBar").style.width = total > 0 ? `${Math.min(100, Math.round((current / total) * 100))}%` : "0";
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
$("reviewAiThinking").addEventListener("change", () => saveReviewSettings().catch((error) => {
  $("reviewSettingsHint").textContent = error.message;
}));
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
$("saveExcelMappingPresetButton").addEventListener("click", () => saveAiReviewExcelMappingPreset().catch((error) => {
  $("reviewTaskHint").textContent = error.message;
}));
$("applyExcelMappingPresetButton").addEventListener("click", () => applyAiReviewExcelMappingPreset().catch((error) => {
  $("reviewTaskHint").textContent = error.message;
}));
$("deleteExcelMappingPresetButton").addEventListener("click", () => deleteAiReviewExcelMappingPreset().catch((error) => {
  $("reviewTaskHint").textContent = error.message;
}));
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
  loadAiReviewPromptTemplates(),
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

