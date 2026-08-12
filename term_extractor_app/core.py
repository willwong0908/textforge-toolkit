"""Core business logic helpers."""

from __future__ import annotations

import importlib
import json
import os
import re
import time
import xml.etree.ElementTree as ET
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any, Callable, Dict, Iterable, List, Optional, Sequence, Tuple

import pandas as pd

from .excel_streaming import header_map_from_values, is_streamable_excel, open_streaming_workbook

from .constants import (
    APPROVED_DECISION,
    DEFAULT_RECALL_SCOPES,
    FAILURE_COLUMNS,
    FAILURE_SHEET,
    GENERIC_TERM_BLACKLIST,
    NONTRANS_REGEX_COLUMNS,
    NONTRANS_REGEX_SHEET,
    NOISE_FRAGMENTS,
    REJECTED_DECISION,
    REVIEW_COLUMNS,
    REVIEW_DECISION,
    REVIEW_SHEET,
    TERM_LIBRARY_COLUMNS,
    TERM_LIBRARY_SHEET,
    TERM_TYPE_CHOICES,
)
from .models import CandidateTerm, FolderScanResult, ProtectedText, ReviewedTerm, SourceRecord, TermRecallCleanRecord, TextSegment


_BAIDU_STOPWORDS_CACHE_FILE = Path(__file__).resolve().parent.parent / "vendor_cache" / "baidu_stopwords.txt"
_NLTK_DATA_DIR = Path(__file__).resolve().parent.parent / "vendor_cache" / "nltk_data"
_GENERIC_STOPWORD_CACHE: Optional[set[str]] = None
_AUTO_RECALL_SCOPE_NAME = "自动"
_PROGRAM_TOKEN_SUFFIXES = (
    "id",
    "ids",
    "key",
    "keys",
    "path",
    "icon",
    "desc",
    "text",
    "value",
    "values",
    "param",
    "params",
    "level",
    "levels",
    "name",
    "names",
    "type",
    "types",
    "flag",
    "flags",
    "index",
    "indices",
    "count",
    "counts",
    "code",
    "token",
)
_COMMON_RICH_TEXT_TAGS = ("color", "sprite", "b", "i", "u", "size", "font", "material")


def _parse_version_parts(version_text: str) -> Tuple[int, ...]:
    parts = []
    for chunk in str(version_text).split("."):
        digits = "".join(ch for ch in chunk if ch.isdigit())
        if not digits:
            break
        parts.append(int(digits))
    return tuple(parts)


def _excel_engine_for(file_path: str) -> str:
    suffix = Path(file_path).suffix.lower()
    if suffix == ".xls":
        return "xlrd"
    return "openpyxl"


def _validate_excel_dependency(file_path: str, engine: str) -> None:
    suffix = Path(file_path).suffix.lower()
    if engine == "xlrd" and suffix == ".xls":
        try:
            xlrd = importlib.import_module("xlrd")
        except ImportError as exc:
            raise RuntimeError("读取 .xls 文件需要 xlrd 依赖，但当前程序未包含该依赖。") from exc
        current_version = _parse_version_parts(getattr(xlrd, "__version__", "0"))
        required_version = (2, 0, 1)
        if current_version < required_version:
            raise RuntimeError(
                "当前程序内置的 xlrd 版本过旧（{0}），读取 .xls 文件需要 2.0.1 或更高版本。".format(
                    getattr(xlrd, "__version__", "未知")
                )
            )


def _open_excel_file(file_path: str) -> pd.ExcelFile:
    engine = _excel_engine_for(file_path)
    _validate_excel_dependency(file_path, engine)
    try:
        return pd.ExcelFile(file_path, engine=engine)
    except Exception as exc:
        raise RuntimeError("读取 Excel 文件失败：{0}。文件：{1}".format(exc, os.path.basename(file_path))) from exc


def preserve_unique(items: Iterable[str]) -> List[str]:
    seen = set()
    ordered = []
    for item in items:
        normalized = str(item).strip()
        if not normalized or normalized in seen:
            continue
        seen.add(normalized)
        ordered.append(normalized)
    return ordered


def normalize_recall_scopes(raw_scopes: Any) -> List[Dict[str, Any]]:
    if not isinstance(raw_scopes, list):
        raw_scopes = []
    scopes: List[Dict[str, Any]] = []
    seen = set()

    for item in raw_scopes:
        if isinstance(item, dict):
            name = str(item.get("name", "")).strip()
            enabled = bool(item.get("enabled", True))
            description = str(item.get("description", "")).strip()
        else:
            name = str(item or "").strip()
            enabled = True
            description = ""
        if not name or name in seen:
            continue
        seen.add(name)
        scopes.append({"name": name, "enabled": enabled, "description": description})

    return scopes


def _effective_enabled_recall_scopes(raw_scopes: Any) -> List[Dict[str, Any]]:
    scopes = [item for item in normalize_recall_scopes(raw_scopes) if item.get("enabled", True)]
    if scopes:
        return scopes
    return [{"name": _AUTO_RECALL_SCOPE_NAME, "enabled": True, "description": ""}]


def recall_scope_auto_enabled(raw_scopes: Any) -> bool:
    return any(
        str(item.get("name", "")).strip() == _AUTO_RECALL_SCOPE_NAME
        for item in _effective_enabled_recall_scopes(raw_scopes)
    )


def format_enabled_recall_scopes(raw_scopes: Any) -> str:
    if recall_scope_auto_enabled(raw_scopes):
        return "- 自动：不预设术语类别，由模型自行判断。"

    scopes = _effective_enabled_recall_scopes(raw_scopes)

    lines = []
    for scope in scopes:
        name = str(scope.get("name", "")).strip()
        description = str(scope.get("description", "")).strip()
        if description:
            lines.append("- {0}: {1}".format(name, description))
        else:
            lines.append("- {0}".format(name))
    return "\n".join(lines)


def enabled_recall_scope_names(raw_scopes: Any) -> List[str]:
    scopes = _effective_enabled_recall_scopes(raw_scopes)
    return preserve_unique(
        [
            str(item.get("name", "")).strip()
            for item in scopes
            if str(item.get("name", "")).strip() != _AUTO_RECALL_SCOPE_NAME
        ]
    )


def allowed_term_types_from_recall_scopes(raw_scopes: Any) -> Optional[List[str]]:
    if recall_scope_auto_enabled(raw_scopes):
        return []
    names = enabled_recall_scope_names(raw_scopes)
    return preserve_unique(names + ["non_term"])


def detect_file_type(folder_path: str) -> str:
    file_types = set()
    for filename in os.listdir(folder_path):
        lower_name = filename.lower()
        if lower_name.endswith((".xlsx", ".xls")):
            file_types.add("excel")
        elif lower_name.endswith(".csv"):
            file_types.add("csv")
        elif lower_name.endswith(".xliff"):
            file_types.add("xliff")
    if not file_types:
        raise ValueError("所选文件夹里没有可处理的 Excel、CSV 或 XLIFF 文件。")
    if len(file_types) > 1:
        raise ValueError("文件夹里包含多种文件类型，请分开处理。")
    return list(file_types)[0]


def scan_folder(folder_path: str) -> FolderScanResult:
    if not os.path.isdir(folder_path):
        raise ValueError("文件夹路径无效。")
    file_type = detect_file_type(folder_path)
    files = sorted(
        [
            filename
            for filename in os.listdir(folder_path)
            if (
                (file_type == "excel" and filename.lower().endswith((".xlsx", ".xls")))
                or (file_type == "csv" and filename.lower().endswith(".csv"))
                or (file_type == "xliff" and filename.lower().endswith(".xliff"))
            )
        ]
    )
    headers: List[str] = ["source"] if file_type == "xliff" else []
    if file_type in ("excel", "csv"):
        headers = get_available_headers(folder_path, file_type)
    return FolderScanResult(
        folder_path=folder_path,
        file_type=file_type,
        file_count=len(files),
        headers=headers,
        files=files,
    )


def get_available_headers(folder_path: str, file_type: str) -> List[str]:
    headers = []
    for filename in sorted(os.listdir(folder_path)):
        file_path = os.path.join(folder_path, filename)
        try:
            if file_type == "excel" and filename.lower().endswith((".xlsx", ".xls")):
                if is_streamable_excel(file_path):
                    with open_streaming_workbook(file_path) as workbook:
                        for sheet in workbook.worksheets:
                            first_row = next(sheet.iter_rows(min_row=1, max_row=1, values_only=True), ())
                            headers.extend(header_map_from_values(first_row).keys())
                else:
                    xls = _open_excel_file(file_path)
                    try:
                        for sheet_name in xls.sheet_names:
                            df = pd.read_excel(xls, sheet_name=sheet_name, nrows=0)
                            headers.extend([str(column) for column in df.columns])
                    finally:
                        xls.close()
            elif file_type == "csv" and filename.lower().endswith(".csv"):
                df = _read_csv(file_path)
                headers.extend([str(column) for column in df.columns])
        except Exception:
            continue
    return preserve_unique(headers)


def _read_csv(file_path: str) -> pd.DataFrame:
    last_error = None
    for encoding in ("utf-8-sig", "utf-8", "gbk", "gb18030"):
        try:
            return pd.read_csv(file_path, encoding=encoding)
        except Exception as exc:
            last_error = exc
    raise last_error


def _clean_text(value: object) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if not text:
        return ""
    return re.sub(r"\s+", " ", text).strip()


def read_source_records(
    folder_path: str,
    file_type: str,
    header_name: str,
    progress_callback: Optional[Callable[[Dict[str, object]], None]] = None,
) -> Tuple[List[SourceRecord], List[str]]:
    records: List[SourceRecord] = []
    processed_files: List[str] = []
    scan_result = scan_folder(folder_path)

    for index, filename in enumerate(scan_result.files, start=1):
        file_path = os.path.join(folder_path, filename)
        if progress_callback:
            progress_callback(
                {
                    "stage": "READING_FILES",
                    "current_file": filename,
                    "current": index,
                    "total": scan_result.file_count,
                    "message": "正在读取 {0}".format(filename),
                }
            )

        if file_type == "excel":
            records.extend(_read_excel_records(file_path, header_name))
        elif file_type == "csv":
            records.extend(_read_csv_records(file_path, header_name))
        else:
            records.extend(_read_xliff_records(file_path))
        processed_files.append(filename)

    unique_records = []
    seen_keys = set()
    for record in records:
        key = (
            record.file_name,
            record.sheet_or_unit,
            record.row_index,
            record.column_name,
            record.text,
        )
        if not record.text or key in seen_keys:
            continue
        seen_keys.add(key)
        unique_records.append(record)
    return unique_records, processed_files


def _read_excel_records(file_path: str, header_name: str) -> List[SourceRecord]:
    collected: List[SourceRecord] = []
    if is_streamable_excel(file_path):
        with open_streaming_workbook(file_path) as workbook:
            for sheet in workbook.worksheets:
                first_row = next(sheet.iter_rows(min_row=1, max_row=1, values_only=True), ())
                columns = header_map_from_values(first_row)
                column_index = columns.get(header_name)
                if column_index is None:
                    continue
                for row_index, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
                    value = row[column_index] if column_index < len(row) else None
                    text = _clean_text(value)
                    if not text:
                        continue
                    collected.append(
                        SourceRecord(
                            record_id="{0}:{1}:{2}:{3}".format(
                                os.path.basename(file_path), sheet.title, row_index, header_name
                            ),
                            file_name=os.path.basename(file_path),
                            source_type="excel",
                            sheet_or_unit=str(sheet.title),
                            row_index=row_index,
                            column_name=str(header_name),
                            text=text,
                        )
                    )
        return collected

    xls = _open_excel_file(file_path)
    try:
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            if header_name not in df.columns:
                continue
            for row_index, value in enumerate(df[header_name].tolist(), start=2):
                text = _clean_text(value)
                if not text:
                    continue
                collected.append(
                    SourceRecord(
                        record_id="{0}:{1}:{2}:{3}".format(
                            os.path.basename(file_path), sheet_name, row_index, header_name
                        ),
                        file_name=os.path.basename(file_path),
                        source_type="excel",
                        sheet_or_unit=str(sheet_name),
                        row_index=row_index,
                        column_name=str(header_name),
                        text=text,
                    )
                )
    finally:
        xls.close()
    return collected


def _read_csv_records(file_path: str, header_name: str) -> List[SourceRecord]:
    df = _read_csv(file_path)
    if header_name not in df.columns:
        return []
    collected = []
    for row_index, value in enumerate(df[header_name].tolist(), start=2):
        text = _clean_text(value)
        if not text:
            continue
        collected.append(
            SourceRecord(
                record_id="{0}:{1}:{2}".format(os.path.basename(file_path), row_index, header_name),
                file_name=os.path.basename(file_path),
                source_type="csv",
                sheet_or_unit="",
                row_index=row_index,
                column_name=str(header_name),
                text=text,
            )
        )
    return collected


def _read_xliff_records(file_path: str) -> List[SourceRecord]:
    collected: List[SourceRecord] = []
    tree = ET.parse(file_path)
    root = tree.getroot()
    namespaces = {"xliff": "urn:oasis:names:tc:xliff:document:1.2"}
    nodes = root.findall(".//xliff:trans-unit", namespaces)
    if not nodes:
        nodes = [node for node in root.iter() if node.tag.lower().endswith("trans-unit")]

    for order, unit in enumerate(nodes, start=1):
        unit_id = unit.attrib.get("id") or str(order)
        source_node = None
        for child in unit.iter():
            if child.tag.lower().endswith("source"):
                source_node = child
                break
        if source_node is None:
            continue
        text = _clean_text("".join(source_node.itertext()))
        if not text:
            continue
        collected.append(
            SourceRecord(
                record_id="{0}:{1}".format(os.path.basename(file_path), unit_id),
                file_name=os.path.basename(file_path),
                source_type="xliff",
                sheet_or_unit=str(unit_id),
                row_index=order,
                column_name="source",
                text=text,
            )
        )
    return collected


def _normalize_ascii_filter_terms(values: Optional[Sequence[str]]) -> Tuple[set[str], set[str]]:
    original_terms = {str(item).strip() for item in list(values or []) if str(item).strip()}
    lowered_terms = {item.lower() for item in original_terms}
    return original_terms, lowered_terms


def _strip_placeholder_blocks(text: str) -> str:
    cleaned = str(text or "")
    patterns = (
        re.compile(r"\{\{[^{}]{1,160}\}\}"),
        re.compile(r"\{[^{}]{1,160}\}"),
        re.compile(r"\[(?:/?color|BTN_[A-Za-z0-9_]+)[^\]]{0,80}\]", re.IGNORECASE),
    )
    for _ in range(4):
        previous = cleaned
        for pattern in patterns:
            cleaned = pattern.sub(" ", cleaned)
        if cleaned == previous:
            break
    return cleaned


def _strip_markup_tags(text: str) -> str:
    cleaned = str(text or "")
    cleaned = re.sub(r"</?(?:{0})\b[^<>]*?/?>".format("|".join(_COMMON_RICH_TEXT_TAGS)), "", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"</?[A-Za-z][A-Za-z0-9]*(?:\s+[^<>]{0,120})?/?>", "", cleaned)
    return cleaned


def _should_strip_ascii_fragment(
    token: str,
    blacklist_terms: set[str],
    blacklist_lower: set[str],
    whitelist_terms: set[str],
    whitelist_lower: set[str],
) -> bool:
    normalized = str(token or "").strip()
    if not normalized:
        return False
    lowered = normalized.lower()
    if normalized in whitelist_terms or lowered in whitelist_lower:
        return False
    if normalized in blacklist_terms or lowered in blacklist_lower:
        return True
    if re.fullmatch(r"[A-Z]{2,5}", normalized):
        return False
    if re.fullmatch(r"[A-Za-z][A-Za-z0-9]{0,3}", normalized):
        return False
    if re.fullmatch(r"T\d{1,2}", normalized, re.IGNORECASE):
        return False
    if re.fullmatch(r"[A-Za-z][A-Za-z0-9:+-]{2,20}", normalized):
        return False
    if looks_like_placeholder_or_program_token(normalized):
        return True
    if "_" in normalized or "." in normalized or "/" in normalized or "\\" in normalized:
        return True
    return False


def build_clean_text_view(
    text: str,
    ascii_filter_blacklist: Optional[Sequence[str]] = None,
    ascii_filter_whitelist: Optional[Sequence[str]] = None,
) -> str:
    cleaned = str(text or "")
    if not cleaned:
        return ""

    blacklist_terms, blacklist_lower = _normalize_ascii_filter_terms(ascii_filter_blacklist)
    whitelist_terms, whitelist_lower = _normalize_ascii_filter_terms(ascii_filter_whitelist)

    cleaned = _strip_markup_tags(cleaned)
    cleaned = _strip_placeholder_blocks(cleaned)
    cleaned = re.sub(r"%(?:\d+\$)?[#0\- +']*(?:\d+)?(?:\.\d+)?[A-Za-z]", " ", cleaned)

    def replace_ascii_token(match: re.Match[str]) -> str:
        token = match.group(0)
        if _should_strip_ascii_fragment(token, blacklist_terms, blacklist_lower, whitelist_terms, whitelist_lower):
            return " "
        return token

    cleaned = re.sub(r"[A-Za-z][A-Za-z0-9_./:\\-]{2,}", replace_ascii_token, cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    return cleaned


def build_term_recall_dedupe_key(text: object) -> str:
    value = str(text or "")
    value = re.sub(r"\s+", " ", value)
    return value.strip()


def has_meaningful_clean_text(text: str) -> bool:
    compact = str(text or "").strip()
    if not compact:
        return False
    if re.search(r"[\u4e00-\u9fff]", compact):
        return True
    if re.search(r"[A-Za-z]{2,}", compact):
        return True
    if re.search(r"[A-Za-z]\d|\d[A-Za-z]", compact):
        return True
    return False


_NUMERIC_NORMALIZATION_PATTERN = re.compile(
    r"(?<![A-Za-z])"
    r"(?:[+-]?\d+(?:\.\d+)?%?|[xX]\d+(?:\.\d+)?|\d+(?:\.\d+)?[xX]|\d+(?:\.\d+)?\s*[-~～]\s*\d+(?:\.\d+)?)"
    r"(?:\s*(?:秒|分钟|分|小时|天|日|周|个月|月|年|点|次|个|颗|枚|件|名|层|级|星|%))?"
    r"(?![A-Za-z])"
)


def build_numeric_normalized_key(text: str) -> str:
    cleaned = _NUMERIC_NORMALIZATION_PATTERN.sub("", str(text or ""))
    cleaned = re.sub(r"\s+", "", cleaned)
    return cleaned.strip()


def apply_duplicate_group_numeric_normalization(texts: Sequence[str]) -> List[str]:
    originals = [str(item or "") for item in texts]
    groups: Dict[str, List[int]] = defaultdict(list)
    for index, text in enumerate(originals):
        key = build_numeric_normalized_key(text)
        if key and key != re.sub(r"\s+", "", text).strip():
            groups[key].append(index)

    normalized = list(originals)
    for key, indexes in groups.items():
        distinct_originals = {originals[index] for index in indexes}
        if len(indexes) < 2 or len(distinct_originals) < 2:
            continue
        for index in indexes:
            normalized[index] = key
    return normalized


def build_term_recall_clean_records(
    protected_items: Sequence[Tuple[str, ProtectedText]],
    placeholder_pattern: str = r"<\d+>",
    numeric_normalization_enabled: bool = True,
    ascii_filter_blacklist: Optional[Sequence[str]] = None,
    ascii_filter_whitelist: Optional[Sequence[str]] = None,
) -> Tuple[List[TermRecallCleanRecord], Dict[str, List[str]]]:
    protected_texts = [item.protected_text for _, item in protected_items]
    without_placeholders = [
        build_clean_text_view(
            re.sub(placeholder_pattern, "", str(text or "")).strip(),
            ascii_filter_blacklist=ascii_filter_blacklist,
            ascii_filter_whitelist=ascii_filter_whitelist,
        )
        for text in protected_texts
    ]
    numeric_normalized = (
        apply_duplicate_group_numeric_normalization(without_placeholders)
        if numeric_normalization_enabled
        else without_placeholders
    )

    clean_records: List[TermRecallCleanRecord] = []
    dedupe_map: Dict[str, List[str]] = defaultdict(list)
    for (source_record_id, protected), clean_text in zip(protected_items, numeric_normalized):
        dedupe_key = build_term_recall_dedupe_key(clean_text)
        record = TermRecallCleanRecord(
            source_record_id=source_record_id,
            original_text=protected.original_text,
            nontrans_protected_text=protected.protected_text,
            term_recall_clean_text=clean_text,
            dedupe_key=dedupe_key,
        )
        clean_records.append(record)
        if dedupe_key:
            dedupe_map[dedupe_key].append(source_record_id)
    return clean_records, dict(dedupe_map)


def segment_source_records(
    source_records: Sequence[SourceRecord],
    single_item_char_limit: int,
    context_window_chars: int = 40,
    ascii_filter_blacklist: Optional[Sequence[str]] = None,
    ascii_filter_whitelist: Optional[Sequence[str]] = None,
) -> List[TextSegment]:
    segments: List[TextSegment] = []
    for record in source_records:
        chunks = _split_long_text(record.text, limit=max(50, int(single_item_char_limit or 50)))
        total_segments = len(chunks)
        for index, (start_offset, chunk) in enumerate(chunks, start=1):
            before = record.text[max(0, start_offset - context_window_chars) : start_offset]
            after_start = start_offset + len(chunk)
            after = record.text[after_start : after_start + context_window_chars]
            context_text = "{0}{1}{2}".format(before, chunk, after).strip()
            llm_text = build_clean_text_view(
                chunk,
                ascii_filter_blacklist=ascii_filter_blacklist,
                ascii_filter_whitelist=ascii_filter_whitelist,
            )
            llm_context_text = build_clean_text_view(
                context_text or chunk,
                ascii_filter_blacklist=ascii_filter_blacklist,
                ascii_filter_whitelist=ascii_filter_whitelist,
            )
            segments.append(
                TextSegment(
                    segment_id="{0}#seg{1}".format(record.record_id, index),
                    source_record_id=record.record_id,
                    file_name=record.file_name,
                    source_type=record.source_type,
                    sheet_or_unit=record.sheet_or_unit,
                    row_index=record.row_index,
                    column_name=record.column_name,
                    segment_index=index,
                    total_segments=total_segments,
                    text=chunk,
                    context_text=context_text or chunk,
                    llm_text=llm_text,
                    llm_context_text=llm_context_text or llm_text,
                )
            )
    return segments


def _split_long_text(text: str, limit: int) -> List[Tuple[int, str]]:
    text = _clean_text(text)
    if not text:
        return []
    if len(text) <= limit:
        return [(0, text)]

    chunks = []
    start = 0
    separators = "。！？；;,.，、\n"
    min_chunk = max(20, limit // 3)

    while start < len(text):
        end = min(start + limit, len(text))
        if end >= len(text):
            chunk = text[start:].strip()
            if chunk:
                chunks.append((start, chunk))
            break

        split_at = -1
        for pointer in range(end, start + min_chunk, -1):
            if text[pointer - 1] in separators:
                split_at = pointer
                break
        if split_at == -1:
            split_at = end

        chunk = text[start:split_at].strip()
        if chunk:
            chunks.append((start, chunk))
        start = split_at

    return chunks


def build_source_location(segment: TextSegment) -> Dict[str, Any]:
    return {
        "file_name": segment.file_name,
        "source_type": segment.source_type,
        "sheet_or_unit": segment.sheet_or_unit,
        "row_index": segment.row_index,
        "column_name": segment.column_name,
        "segment_index": segment.segment_index,
        "total_segments": segment.total_segments,
    }


def format_source_location(location: Dict[str, Any]) -> str:
    parts = [str(location.get("file_name", "")).strip()]
    sheet_or_unit = str(location.get("sheet_or_unit", "")).strip()
    if sheet_or_unit:
        parts.append(sheet_or_unit)
    row_index = int(location.get("row_index", 0) or 0)
    if row_index > 0:
        parts.append("行 {0}".format(row_index))
    column_name = str(location.get("column_name", "")).strip()
    if column_name:
        parts.append("列 {0}".format(column_name))
    segment_index = int(location.get("segment_index", 1) or 1)
    total_segments = int(location.get("total_segments", 1) or 1)
    if total_segments > 1:
        parts.append("分段 {0}/{1}".format(segment_index, total_segments))
    return " | ".join(part for part in parts if part)


def _read_stopword_file(path: Path) -> set[str]:
    try:
        return _parse_stopword_text(path.read_text(encoding="utf-8"))
    except Exception:
        return set()


def _parse_stopword_text(text: str) -> set[str]:
    return {line.strip() for line in str(text).splitlines() if line.strip()}


def _load_baidu_stopwords() -> set[str]:
    return _read_stopword_file(_BAIDU_STOPWORDS_CACHE_FILE)


def _load_nltk_stopwords() -> set[str]:
    try:
        import nltk
        from nltk.corpus import stopwords
    except ImportError:
        return set()

    try:
        nltk.data.path.insert(0, str(_NLTK_DATA_DIR))
        words = stopwords.words("english")
    except LookupError:
        return set()
    except Exception:
        return set()

    return {str(item).strip().lower() for item in words if str(item).strip()}


def get_generic_term_blacklist() -> set[str]:
    global _GENERIC_STOPWORD_CACHE
    if _GENERIC_STOPWORD_CACHE is not None:
        return _GENERIC_STOPWORD_CACHE

    combined = {str(item).strip() for item in GENERIC_TERM_BLACKLIST if str(item).strip()}
    combined.update(_load_baidu_stopwords())
    combined.update(_load_nltk_stopwords())

    normalized = set()
    for item in combined:
        normalized.add(item)
        normalized.add(item.lower())

    _GENERIC_STOPWORD_CACHE = normalized
    return _GENERIC_STOPWORD_CACHE


def is_blacklisted_term(token: str) -> bool:
    normalized = str(token or "").strip()
    if not normalized:
        return False
    blacklist = get_generic_term_blacklist()
    return normalized in blacklist or normalized.lower() in blacklist


def looks_like_placeholder_or_program_token(token: str) -> bool:
    normalized = str(token or "").strip()
    if not normalized:
        return False

    lowered = normalized.lower()
    if any(character in normalized for character in "{}[]<>"):
        return True
    if re.search(r"%(?:\d+\$)?[#0\- +']*(?:\d+)?(?:\.\d+)?[A-Za-z]", normalized):
        return True
    if re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*=[^=\s]+", normalized):
        return True
    if re.fullmatch(r"[A-Z][A-Z0-9]*(?:_[A-Z0-9]+)+", normalized):
        return True
    if re.fullmatch(r"[A-Za-z]+(?:_[A-Za-z0-9]+)+", normalized):
        return True
    if re.fullmatch(r"[A-Za-z0-9_./\\:-]+\.(?:png|jpg|jpeg|webp|prefab|asset|json|xml|txt|csv|wav|mp3|ogg|fbx|uasset)", normalized, re.IGNORECASE):
        return True
    if "\\" in normalized or "/" in normalized:
        if re.search(r"[A-Za-z]", normalized):
            return True

    is_camel_or_pascal = bool(
        re.fullmatch(r"[a-z]+(?:[A-Z][A-Za-z0-9]+)+", normalized)
        or re.fullmatch(r"[A-Z][a-z0-9]+(?:[A-Z][A-Za-z0-9]+)+", normalized)
    )
    if is_camel_or_pascal and any(lowered.endswith(suffix) for suffix in _PROGRAM_TOKEN_SUFFIXES):
        return True

    if re.fullmatch(r"[A-Za-z][A-Za-z0-9]*", normalized):
        if any(lowered.endswith(suffix) for suffix in _PROGRAM_TOKEN_SUFFIXES) and len(normalized) >= 5:
            return True

    return False


def _is_program_token_character(character: str) -> bool:
    return character.isalnum() or character in "_:./\\-%$"


def _find_embedded_program_token_evidence(token: str, source_text: str) -> str:
    normalized = str(token or "").strip()
    text = str(source_text or "")
    if not normalized or not text or normalized not in text:
        return ""

    embedded_fragments: List[str] = []
    for match in re.finditer(re.escape(normalized), text):
        fragment = _extract_enclosing_program_fragment(text, match.start(), match.end(), normalized)
        if not fragment:
            return ""
        embedded_fragments.append(fragment)

    if not embedded_fragments:
        return ""
    return sorted(preserve_unique(embedded_fragments), key=lambda item: (len(item), item))[0]


def _extract_enclosing_program_fragment(text: str, start: int, end: int, token: str) -> str:
    for open_mark, close_mark in (("{{", "}}"), ("{", "}"), ("[", "]"), ("<", ">")):
        fragment = _find_wrapped_fragment(text, start, end, token, open_mark, close_mark)
        if fragment:
            return fragment

    left = start
    while left > 0 and _is_program_token_character(text[left - 1]):
        left -= 1
    right = end
    while right < len(text) and _is_program_token_character(text[right]):
        right += 1

    fragment = text[left:right].strip()
    if fragment and fragment != token and token in fragment and looks_like_placeholder_or_program_token(fragment):
        return fragment
    return ""


def _find_wrapped_fragment(text: str, start: int, end: int, token: str, open_mark: str, close_mark: str) -> str:
    search_start = max(0, start - 80)
    search_end = min(len(text), end + 80)
    open_index = text.rfind(open_mark, search_start, start + len(open_mark))
    if open_index < 0:
        return ""

    close_index = text.find(close_mark, end, search_end)
    if close_index < 0:
        return ""

    fragment = text[open_index : close_index + len(close_mark)].strip()
    if not fragment or len(fragment) > 80 or "\n" in fragment or token not in fragment:
        return ""
    if open_mark == "<":
        first_close = fragment.find(">")
        last_open = fragment.rfind("<")
        if 0 <= first_close < last_open:
            inner_text = fragment[first_close + 1 : last_open]
            if token in inner_text:
                return ""
    if looks_like_placeholder_or_program_token(fragment):
        return fragment
    return ""


def evidence_shows_program_token_usage(token: str, evidence_text: str, source_text: str = "") -> bool:
    normalized = str(token or "").strip()
    evidence = str(evidence_text or "").strip()
    if evidence and normalized and normalized in evidence:
        if evidence != normalized and looks_like_placeholder_or_program_token(evidence):
            return True
    return bool(_find_embedded_program_token_evidence(normalized, source_text))


def evidence_suggests_partial_term(surface_form: str, evidence_text: str) -> bool:
    candidate = str(surface_form or "").strip()
    evidence = str(evidence_text or "").strip()
    if not candidate or not evidence or candidate == evidence:
        return False
    if len(candidate) < 3:
        return False
    if not re.fullmatch(r"[\u4e00-\u9fffA-Za-z]+", evidence):
        return False
    start = evidence.find(candidate)
    if start < 0:
        return False
    end = start + len(candidate)
    left_char = evidence[start - 1] if start > 0 else ""
    right_char = evidence[end] if end < len(evidence) else ""
    def is_term_char(ch: str) -> bool:
        return bool(ch) and (ch.isalnum() or ("\u4e00" <= ch <= "\u9fff"))
    return is_term_char(left_char) or is_term_char(right_char)


def merge_candidate_terms(candidates: Sequence[CandidateTerm], max_contexts: int = 3) -> List[CandidateTerm]:
    grouped: Dict[str, List[CandidateTerm]] = defaultdict(list)
    for candidate in candidates:
        if candidate.surface_form:
            grouped[candidate.surface_form].append(candidate)

    merged_candidates: List[CandidateTerm] = []
    for surface_form, group in grouped.items():
        first = group[0]
        recall_sources = []
        source_record_ids = []
        segment_ids = []
        source_locations: List[Dict[str, Any]] = []
        sample_contexts: List[str] = []
        evidence_text = ""
        occurrence_count = 0

        for candidate in group:
            for source in str(candidate.recall_source or "").split("+"):
                source = source.strip()
                if source:
                    recall_sources.append(source)
            source_record_ids.extend(candidate.source_record_ids or [candidate.source_record_id])
            segment_ids.extend(candidate.segment_ids or [candidate.segment_id])
            source_locations.extend(candidate.source_locations or [])
            if candidate.sample_contexts:
                sample_contexts.extend(candidate.sample_contexts)
            elif candidate.context_text:
                sample_contexts.append(candidate.context_text)
            if not evidence_text and candidate.evidence_text:
                evidence_text = candidate.evidence_text
            occurrence_count += max(1, int(candidate.occurrence_count or 1))

        unique_contexts = _pick_representative_contexts(surface_form, sample_contexts, max_contexts=max_contexts)
        merged_candidates.append(
            CandidateTerm(
                candidate_id="merged:{0}".format(surface_form),
                surface_form=surface_form,
                source_record_id=first.source_record_id,
                segment_id=first.segment_id,
                recall_source="+".join(sorted(set(recall_sources))) or first.recall_source,
                context_text="\n\n".join(unique_contexts) or first.context_text,
                evidence_text=evidence_text or surface_form,
                source_record_ids=preserve_unique(source_record_ids),
                segment_ids=preserve_unique(segment_ids),
                source_locations=_dedupe_source_locations(source_locations),
                sample_contexts=unique_contexts,
                occurrence_count=occurrence_count,
            )
        )

    return sorted(merged_candidates, key=lambda item: (item.surface_form, item.candidate_id))


def _pick_representative_contexts(surface_form: str, contexts: Sequence[str], max_contexts: int = 3) -> List[str]:
    scored_contexts = []
    for context in preserve_unique(contexts):
        compact = str(context or "").strip()
        if not compact:
            continue
        score = 0
        if surface_form in compact:
            score += 100
            score += max(0, 120 - len(compact))
        score += compact.count(surface_form) * 10
        scored_contexts.append((score, compact))

    scored_contexts.sort(key=lambda item: (-item[0], len(item[1]), item[1]))
    return [context for _, context in scored_contexts[: max(1, int(max_contexts or 1))]]


def _dedupe_source_locations(locations: Sequence[Dict[str, Any]]) -> List[Dict[str, Any]]:
    unique_locations = []
    seen_keys = set()
    for location in locations:
        normalized = dict(location or {})
        key = (
            normalized.get("file_name", ""),
            normalized.get("source_type", ""),
            normalized.get("sheet_or_unit", ""),
            int(normalized.get("row_index", 0) or 0),
            normalized.get("column_name", ""),
            int(normalized.get("segment_index", 1) or 1),
            int(normalized.get("total_segments", 1) or 1),
        )
        if key in seen_keys:
            continue
        seen_keys.add(key)
        unique_locations.append(normalized)
    return unique_locations


def sanitize_candidate(token: str) -> str:
    token = str(token or "").strip()
    token = token.strip("“”\"'《》〈〉【】[]（）()")
    token = re.sub(r"\s+", " ", token).strip()
    token = _trim_candidate_condition_noise(token)
    if len(token) < 2 or len(token) > 24:
        return ""
    return token


def _trim_candidate_condition_noise(token: str) -> str:
    cleaned = str(token or "").strip()
    if not cleaned:
        return ""

    for _ in range(3):
        previous = cleaned
        cleaned = re.sub(r"^(?:阶段结束时|任务完成时|活动结束时|目标为|要求为|需要|需|请|将|通过)\s*", "", cleaned)
        cleaned = re.sub(r"^(?:全服|本服|所在公会的?|公会的?|玩家|队伍)\s*", "", cleaned)
        cleaned = re.sub(r"^(?:累计|总计|至少|成功|完成)\s*", "", cleaned)
        cleaned = re.sub(r"^(?:占领|达到|解锁|提升|升级|搜索|使用|消耗|获得)\s*", "", cleaned)
        cleaned = re.sub(r"^\d[\d,]*(?:个|位|次|级|%)?(?:及以上|或以上|以上)?\s*", "", cleaned)
        cleaned = re.sub(r"^T\d{1,2}(?:及以上|或以上|以上)?\s*", "", cleaned, flags=re.IGNORECASE)
        cleaned = re.sub(r"^[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩⅪⅫ]+(?:及以上|或以上|以上)?\s*", "", cleaned)
        cleaned = re.sub(r"\s*\d[\d,]*(?:次|个|位|级|%)?$", "", cleaned)
        cleaned = re.sub(r"\s*(?:被占领|被击败|已完成|可提升|可获得|即可获得|达到.*|完成.*)$", "", cleaned)
        cleaned = cleaned.strip(" ：:，,。.;；、")
        if cleaned == previous:
            break
    return cleaned.strip()


def looks_like_unsalvageable_action_candidate(token: str) -> bool:
    normalized = re.sub(r"\s+", "", str(token or "").strip())
    if not normalized:
        return False
    return False


def looks_like_generic_grade_candidate(token: str) -> bool:
    normalized = re.sub(r"\s+", "", str(token or "").strip())
    if not normalized:
        return False
    roman = "IVXivxⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩⅪⅫ"
    if re.fullmatch(r"(?:Tier|T)(?:\d{{1,2}}|[{0}]+)?".format(roman), normalized, re.IGNORECASE):
        return True
    return False


def looks_like_sentence_fragment_candidate(token: str) -> bool:
    normalized = re.sub(r"\s+", "", str(token or "").strip())
    if not normalized:
        return False
    if looks_like_unsalvageable_action_candidate(normalized):
        return True
    if normalized.startswith(("通过", "为了", "需要", "请", "将", "在此处")):
        return True
    if any(marker in normalized for marker in ("可提升", "可获得", "即可获得", "后解锁", "时获得")):
        return True
    if re.search(r"(?:达到|占领|击败|完成|使用|消耗|搜索|解锁).{2,}(?:后|时|可|即可)", normalized):
        return True
    return False


def is_noise_candidate(token: str, source_text: str = "", evidence_text: str = "") -> bool:
    normalized = str(token or "").strip()
    if not token:
        return True
    if is_blacklisted_term(normalized):
        return True
    if looks_like_sentence_fragment_candidate(token):
        return True
    if looks_like_generic_grade_candidate(token):
        return True
    if any(mark in str(token) for mark in ("，", ",", "。", ".")):
        return True
    if evidence_shows_program_token_usage(token, evidence_text, source_text):
        return True
    if looks_like_placeholder_or_program_token(token):
        return True
    if any(fragment in token for fragment in NOISE_FRAGMENTS):
        return True
    if re.fullmatch(r"[\W_]+", token):
        return True
    if re.fullmatch(r"[0-9IVXivx]+", token):
        return True
    if re.fullmatch(r"[A-Za-z]{1,2}", token):
        return True
    if re.fullmatch(r"[A-Za-z][A-Za-z'\-]{2,}", normalized):
        lowered = normalized.lower()
        if lowered in {
            "starting",
            "ended",
            "finish",
            "play",
            "jump",
            "pocket",
            "available",
            "altogether",
            "through",
            "wrapped",
            "another",
            "three",
            "thank",
            "with",
            "without",
            "into",
            "onto",
            "upon",
            "your",
            "you",
            "they",
            "them",
            "their",
            "this",
            "that",
            "these",
            "those",
            "then",
            "than",
            "here",
            "there",
            "will",
            "would",
            "should",
            "could",
            "have",
            "has",
            "had",
            "been",
            "being",
            "does",
            "did",
            "done",
            "were",
            "was",
            "are",
            "is",
            "am",
            "the",
            "and",
            "for",
            "from",
            "into",
            "onto",
            "over",
            "under",
            "after",
            "before",
            "about",
            "across",
            "around",
            "between",
            "during",
            "throughout",
            "within",
            "across",
            "via",
            "per",
            "all",
        }:
            return True
    return False


def build_candidate_recall_batches(
    segments: Sequence[TextSegment],
    user_prompt_template: str,
    source_language: str,
    batch_char_limit: int,
    system_prompt_template: str = "",
    enabled_recall_scopes: str = "",
    retry_feedback_by_segment_id: Optional[Dict[str, Sequence[str]]] = None,
) -> List[Dict[str, Any]]:
    retry_feedback_by_segment_id = retry_feedback_by_segment_id or {}
    prompt_suffix = ""
    if retry_feedback_by_segment_id:
        prompt_suffix = (
            "\nRetry guidance:\n"
            "- Some items include `retry_notes` from the previous failed attempt.\n"
            "- Fix only those mistakes and return exact valid surface_forms.\n"
            "- For each id, always return one result object. If there is no valid term, return an empty `surface_forms` array.\n"
            "- If all apparent terms are placeholders, tags, format fragments, numbers, or noise, return an empty array.\n"
        )
    return _build_numbered_prompt_batches(
        items=segments,
        batch_char_limit=batch_char_limit,
        payload_builder=lambda segment, request_id: {
            "id": request_id,
            "text": segment.llm_text or segment.text,
            **(
                {"retry_notes": list(retry_feedback_by_segment_id.get(segment.segment_id, []) or [])[:3]}
                if retry_feedback_by_segment_id.get(segment.segment_id)
                else {}
            ),
        },
        prompt_builder=lambda payloads: (
            user_prompt_template.format(
                source_language=source_language,
                batch_items_json=_dump_json(payloads),
            )
            + prompt_suffix
        ),
        message_builder=lambda prompt: build_llm_messages(
            system_prompt_template.format(
                source_language=source_language,
                enabled_recall_scopes=enabled_recall_scopes,
            )
            if system_prompt_template
            else "",
            prompt,
        ),
    )


def build_chunk_term_recall_batches(
    clean_records: Sequence[TermRecallCleanRecord],
    user_prompt_template: str,
    source_language: str,
    batch_char_limit: int,
    system_prompt_template: str = "",
    enabled_recall_scopes: str = "",
) -> List[Dict[str, Any]]:
    unique_records = []
    seen_keys = set()
    for record in clean_records:
        if not record.dedupe_key or record.dedupe_key in seen_keys:
            continue
        seen_keys.add(record.dedupe_key)
        unique_records.append(record)

    return _build_numbered_prompt_batches(
        items=unique_records,
        batch_char_limit=batch_char_limit,
        payload_builder=lambda record, request_id: {
            "id": request_id,
            "dedupe_key": record.dedupe_key,
            "text": record.term_recall_clean_text,
        },
        prompt_builder=lambda payloads: user_prompt_template.format(
            source_language=source_language,
            batch_items_json=_dump_json(payloads),
        ),
        message_builder=lambda prompt: build_llm_messages(
            system_prompt_template.format(
                source_language=source_language,
                enabled_recall_scopes=enabled_recall_scopes,
            )
            if system_prompt_template
            else "",
            prompt,
        ),
    )


def build_review_batches(
    candidates: Sequence[CandidateTerm],
    user_prompt_template: str,
    batch_char_limit: int,
    system_prompt_template: str = "",
    enabled_recall_scopes: str = "",
    term_type_choices: Optional[Sequence[str]] = None,
) -> List[Dict[str, Any]]:
    if term_type_choices == []:
        choices_text = "自由判断；如果不是术语，返回 non_term。"
    else:
        choices_text = ", ".join(term_type_choices or TERM_TYPE_CHOICES)
    return _build_numbered_prompt_batches(
        items=candidates,
        batch_char_limit=batch_char_limit,
        payload_builder=lambda candidate, request_id: {
            "id": request_id,
            "surface_form": candidate.surface_form,
            "context_text": candidate.context_text,
            "clean_context_text": build_clean_text_view(candidate.context_text),
        },
        prompt_builder=lambda payloads: user_prompt_template.format(
            batch_items_json=_dump_json(payloads),
        ),
        message_builder=lambda prompt: build_llm_messages(
            system_prompt_template.format(
                term_type_choices=choices_text,
                enabled_recall_scopes=enabled_recall_scopes,
            )
            if system_prompt_template
            else "",
            prompt,
        ),
    )


def format_recall_retry_feedback(raw_issues: Sequence[str]) -> List[str]:
    notes: List[str] = []
    for issue in raw_issues or []:
        issue_text = str(issue or "").strip()
        if not issue_text:
            continue
        notes.append("上次错误：{0}".format(issue_text))
        corrective_note = _build_recall_retry_note(issue_text)
        if corrective_note:
            notes.append(corrective_note)
    return preserve_unique(notes)


def format_review_retry_feedback(
    raw_issues: Sequence[str],
    allowed_term_types: Optional[Sequence[str]] = None,
) -> List[str]:
    notes: List[str] = []
    for issue in raw_issues or []:
        issue_text = str(issue or "").strip()
        if not issue_text:
            continue
        notes.append("上次错误：{0}".format(issue_text))
        corrective_note = _build_review_retry_note(issue_text, allowed_term_types=allowed_term_types)
        if corrective_note:
            notes.append(corrective_note)
    return preserve_unique(notes)


def build_failure_guidance(
    stage: str,
    task_type: str,
    reason: str,
    allowed_term_types: Optional[Sequence[str]] = None,
) -> str:
    reason_text = str(reason or "").strip()
    if not reason_text:
        return ""

    if stage == "RECALLING_CANDIDATES" or task_type.startswith("candidate_recall"):
        guidance = _build_recall_retry_note(reason_text)
        if guidance:
            return guidance
    if stage == "REVIEWING_CANDIDATES" or task_type.startswith("candidate_review"):
        guidance = _build_review_retry_note(reason_text, allowed_term_types=allowed_term_types)
        if guidance:
            return guidance
    return ""


def _build_recall_retry_note(issue: str) -> str:
    notes: List[str] = []
    if "模型遗漏编号" in issue or "模型遗漏了该编号" in issue:
        notes.append("必须覆盖每个输入 id；即使没有术语，也要返回该 id 对应的空 `surface_forms` 数组。")
    if "返回了重复 id" in issue:
        notes.append("每个 id 只能返回一次，不要重复编号，也不要新增未请求的编号。")
    if "缺少 id" in issue:
        notes.append("每个结果对象都必须带上正确的 `id`，不能省略。")
    if "未请求的 id" in issue:
        notes.append("不要新增未请求的编号；只返回当前输入里已有的 id。")
    if "不是 JSON 对象" in issue:
        notes.append("数组中的每一项都必须是 JSON 对象，不能返回字符串、数字或其他结构。")
    if "非逐字候选" in issue:
        notes.append("候选词必须是原文中连续逐字出现的核心术语；不要改写、补词、缩写或把整句动作提示改造成术语。")
    if "条件整段" in issue:
        notes.append(
            "不要返回“阶段结束时”“加入”“被击杀”“达到XX”这类动作或条件整段；只保留其中真正稳定的核心名词，没有就返回空数组。"
        )
    if "代码/噪声规则过滤" in issue:
        notes.append(
            "上次返回的内容像占位符、标签、整句提示或普通提示语；请忽略代码和噪声，若文本没有稳定术语就直接返回空数组。"
        )
    if "缺少 surface_forms 字段" in issue or "surface_forms 不是数组" in issue:
        notes.append(
            "返回结构必须包含 `surface_forms` 数组；顶层保持 JSON，且每个条目都要有 `id` 和 `surface_forms`。"
        )
    if "不是合法 JSON" in issue or "不是 JSON 对象" in issue:
        notes.append("只返回合法 JSON，不要解释、不要 Markdown 代码块、不要额外文本。")
    return " ".join(preserve_unique(notes))


def _build_review_retry_note(
    issue: str,
    allowed_term_types: Optional[Sequence[str]] = None,
) -> str:
    if allowed_term_types == []:
        allowed_text = "自由判断；如果不是术语，固定返回 non_term"
    else:
        allowed_text = "、".join(list(allowed_term_types or TERM_TYPE_CHOICES))
    if "不是合法 JSON" in issue or "缺少 items 数组" in issue or "不是 JSON 对象" in issue:
        return (
            "只返回一个合法 JSON 对象，顶层必须是 `{\\\"items\\\":[...]}`；不要解释、不要 Markdown 代码块、不要额外文本。"
        )
    if "缺少 id" in issue:
        return "每个结果对象都必须带上正确的 `id`，不能省略。"
    if "未请求的 id" in issue:
        return "不要新增未请求的编号；只返回当前输入里已有的 id。"
    if "模型遗漏了该编号" in issue:
        return "必须覆盖每个输入 id；即使判为非术语，也要返回该 id 的完整结果对象。"
    if "返回了重复 id" in issue:
        return "每个 id 只能返回一次，不要重复编号，也不要新增未请求的编号。"
    if "surface_form 与候选词不一致" in issue:
        return "返回的 `surface_form` 必须与输入候选词完全一致，不能改写、截断、补词或归一化。"
    if "不支持的术语类别" in issue:
        return "`term_type` 只能使用允许集合中的值：{0}。如果不是术语，固定返回 `non_term`。".format(allowed_text)
    if "decision" in issue:
        return "`decision` 只能返回 approved、review、rejected；不要返回置信度或风险标签。"
    return ""


def _build_numbered_prompt_batches(
    items: Sequence[Any],
    batch_char_limit: int,
    payload_builder: Callable[[Any, str], Dict[str, Any]],
    prompt_builder: Callable[[List[Dict[str, Any]]], str],
    message_builder: Optional[Callable[[str], List[Dict[str, str]]]] = None,
) -> List[Dict[str, Any]]:
    effective_limit = max(200, int(batch_char_limit or 200))
    batches: List[Dict[str, Any]] = []
    current_items: List[Any] = []
    current_payloads: List[Dict[str, Any]] = []

    for item in items:
        candidate_payload = payload_builder(item, str(len(current_payloads) + 1))
        candidate_payloads = current_payloads + [candidate_payload]
        candidate_prompt = prompt_builder(candidate_payloads)

        if current_payloads and len(candidate_prompt) > effective_limit:
            batches.append(_finalize_prompt_batch(current_items, current_payloads, prompt_builder, message_builder))
            current_items = [item]
            current_payloads = [payload_builder(item, "1")]
        else:
            current_items.append(item)
            current_payloads = candidate_payloads

    if current_payloads:
        batches.append(_finalize_prompt_batch(current_items, current_payloads, prompt_builder, message_builder))
    return batches


def _finalize_prompt_batch(
    items: Sequence[Any],
    payloads: Sequence[Dict[str, Any]],
    prompt_builder: Callable[[List[Dict[str, Any]]], str],
    message_builder: Optional[Callable[[str], List[Dict[str, str]]]] = None,
) -> Dict[str, Any]:
    payload_list = list(payloads)
    prompt = prompt_builder(payload_list)
    return {
        "prompt": prompt,
        "payload_json": _dump_json(payload_list),
        "messages": message_builder(prompt) if message_builder else [{"role": "user", "content": prompt}],
        "items": [
            {
                "request_id": str(payload["id"]),
                "item": original_item,
                "payload": dict(payload),
            }
            for payload, original_item in zip(payload_list, items)
        ],
    }


def build_llm_messages(system_prompt: str, user_prompt: str) -> List[Dict[str, str]]:
    messages: List[Dict[str, str]] = []
    system_text = str(system_prompt or "").strip()
    user_text = str(user_prompt or "").strip()
    if system_text:
        messages.append({"role": "system", "content": system_text})
    messages.append({"role": "user", "content": user_text})
    return messages


def _dump_json(payload: Sequence[Dict[str, Any]]) -> str:
    return json.dumps(list(payload), ensure_ascii=False, indent=2)


def build_review_retry_batches(
    candidates: Sequence[CandidateTerm],
    user_prompt_template: str,
    batch_char_limit: int,
    system_prompt_template: str = "",
    enabled_recall_scopes: str = "",
    term_type_choices: Optional[Sequence[str]] = None,
    retry_feedback_by_candidate_id: Optional[Dict[str, Sequence[str]]] = None,
    max_items_per_batch: Optional[int] = None,
    max_context_chars: int = 220,
) -> List[Dict[str, Any]]:
    retry_feedback_by_candidate_id = retry_feedback_by_candidate_id or {}
    if term_type_choices == []:
        choices_text = "自由判断；如果不是术语，返回 non_term。"
    else:
        choices_text = ", ".join(term_type_choices or TERM_TYPE_CHOICES)
    prompt_suffix = ""
    if retry_feedback_by_candidate_id:
        prompt_suffix = (
            "\nRetry guidance:\n"
            "- Some items include `retry_notes` from the previous failed attempt.\n"
            "- Fix only those mistakes.\n"
            "- Keep `surface_form` exactly equal to the input candidate.\n"
            "- `decision` must be one of approved, review, rejected.\n"
            "- If a fixed term type list is provided, `term_type` must use that list. If the list says free classification, choose the most suitable concise type yourself.\n"
            "- Do not output confidence, evidence_text, risk_hints, risk_flags, or extra fields.\n"
            "- Return JSON only.\n"
        )
    effective_limit = max(200, int(batch_char_limit or 200))
    effective_max_items = max(1, int(max_items_per_batch or 0)) if max_items_per_batch else 0

    batches: List[Dict[str, Any]] = []
    current_items: List[CandidateTerm] = []
    current_payloads: List[Dict[str, Any]] = []

    for candidate in candidates:
        payload = {
            "id": str(len(current_payloads) + 1),
            "surface_form": candidate.surface_form,
            "occurrence_count": max(1, int(candidate.occurrence_count or 1)),
            "contexts": preserve_unique(
                [
                    build_review_context_text(
                        candidate,
                        max_chars=max_context_chars,
                        max_contexts=2,
                    )
                ]
            ),
        }
        retry_notes = list(retry_feedback_by_candidate_id.get(candidate.candidate_id, []) or [])[:3]
        if retry_notes:
            payload["retry_notes"] = retry_notes

        candidate_payloads = current_payloads + [payload]
        candidate_prompt = user_prompt_template.format(
            batch_items_json=_dump_json(candidate_payloads),
        ) + prompt_suffix

        exceeds_limit = len(candidate_prompt) > effective_limit
        exceeds_item_cap = bool(effective_max_items and len(candidate_payloads) > effective_max_items)
        if current_payloads and (exceeds_limit or exceeds_item_cap):
            batches.append(
                _finalize_prompt_batch(
                    current_items,
                    current_payloads,
                    lambda payloads: user_prompt_template.format(
                        batch_items_json=_dump_json(payloads),
                    ) + prompt_suffix,
                    lambda prompt: build_llm_messages(
                        system_prompt_template.format(
                            term_type_choices=choices_text,
                            enabled_recall_scopes=enabled_recall_scopes,
                        )
                        if system_prompt_template
                        else "",
                        prompt,
                    ),
                )
            )
            current_items = [candidate]
            current_payloads = [
                {
                    "id": "1",
                    "surface_form": candidate.surface_form,
                    "occurrence_count": max(1, int(candidate.occurrence_count or 1)),
                    "contexts": preserve_unique(
                        [
                            build_review_context_text(
                                candidate,
                                max_chars=max_context_chars,
                                max_contexts=2,
                            )
                        ]
                    ),
                    **({"retry_notes": retry_notes} if retry_notes else {}),
                }
            ]
        else:
            current_items.append(candidate)
            current_payloads = candidate_payloads

    if current_payloads:
        batches.append(
            _finalize_prompt_batch(
                current_items,
                current_payloads,
                lambda payloads: user_prompt_template.format(
                    batch_items_json=_dump_json(payloads),
                ) + prompt_suffix,
                lambda prompt: build_llm_messages(
                    system_prompt_template.format(
                        term_type_choices=choices_text,
                        enabled_recall_scopes=enabled_recall_scopes,
                    )
                    if system_prompt_template
                    else "",
                    prompt,
                ),
            )
        )
    return batches


def build_review_context_text(candidate: CandidateTerm, max_chars: int = 220, max_contexts: int = 1) -> str:
    contexts = list(candidate.sample_contexts or [])
    if not contexts and candidate.context_text:
        contexts = [candidate.context_text]

    chosen_contexts = _pick_representative_contexts(
        candidate.surface_form,
        contexts,
        max_contexts=max(1, int(max_contexts or 1)),
    )
    compact_contexts = preserve_unique(
        [
            _compact_context_around_term(context, candidate.surface_form, max_chars=max_chars)
            for context in chosen_contexts
            if str(context or "").strip()
        ]
    )
    if compact_contexts:
        return "\n\n".join(compact_contexts)

    fallback_text = candidate.context_text or candidate.surface_form
    return _compact_context_around_term(fallback_text, candidate.surface_form, max_chars=max_chars)


def _compact_context_around_term(context: str, surface_form: str, max_chars: int = 220) -> str:
    compact = re.sub(r"\s+", " ", str(context or "")).strip()
    if not compact:
        return ""

    effective_limit = max(40, int(max_chars or 40))
    if len(compact) <= effective_limit:
        return compact

    term = str(surface_form or "").strip()
    if term and term in compact:
        pivot = compact.find(term)
        left_budget = max(10, (effective_limit - len(term)) // 2)
        start = max(0, pivot - left_budget)
        end = min(len(compact), start + effective_limit)
        if end - start < effective_limit:
            start = max(0, end - effective_limit)
        snippet = compact[start:end].strip()
        if start > 0:
            snippet = "..." + snippet
        if end < len(compact):
            snippet = snippet + "..."
        return snippet

    if effective_limit <= 3:
        return compact[:effective_limit]
    return compact[: effective_limit - 3].strip() + "..."


def clean_api_response(response: str) -> str:
    cleaned = re.sub(r"<think>.*?</think>", "", response or "", flags=re.DOTALL | re.IGNORECASE)
    cleaned = re.sub(r"^```json\s*", "", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"^```\s*", "", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\s*```$", "", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"[\x00-\x08\x0b-\x1f]+", " ", cleaned)
    return cleaned.strip()


def parse_candidate_terms_from_response(response: str, source_text: str) -> List[str]:
    cleaned = clean_api_response(response)
    data = _load_json_fragment(cleaned)
    return _parse_candidate_terms_payload(data, source_text)


def parse_chunk_term_recall_response(
    response: str,
    clean_records_by_request_id: Dict[str, TermRecallCleanRecord],
) -> Dict[str, Any]:
    cleaned = clean_api_response(response)
    data = _load_json_fragment(cleaned)
    raw_items = _extract_batch_items(data, "chunk 术语召回")

    source_texts = {
        request_id: record.term_recall_clean_text
        for request_id, record in clean_records_by_request_id.items()
    }
    resolved: Dict[str, List[str]] = {request_id: [] for request_id in clean_records_by_request_id}
    rejected_terms: List[str] = []
    debug_items: List[Dict[str, Any]] = []

    seen_ids = set()
    for index, item in enumerate(raw_items, start=1):
        if not isinstance(item, dict):
            debug_items.append(
                {
                    "request_id": "",
                    "raw_surface": "",
                    "sanitized": "",
                    "accepted_request_ids": [],
                    "rejected_reason": "invalid_item_type",
                    "matched_request_ids": [],
                    "noise_request_ids": [],
                    "sample_source_text": "",
                }
            )
            continue

        request_id = str(item.get("id", "")).strip()
        if not request_id:
            debug_items.append(
                {
                    "request_id": "",
                    "raw_surface": "",
                    "sanitized": "",
                    "accepted_request_ids": [],
                    "rejected_reason": "missing_id",
                    "matched_request_ids": [],
                    "noise_request_ids": [],
                    "sample_source_text": "",
                }
            )
            continue

        if request_id in seen_ids:
            debug_items.append(
                {
                    "request_id": request_id,
                    "raw_surface": "",
                    "sanitized": "",
                    "accepted_request_ids": [],
                    "rejected_reason": "duplicate_id",
                    "matched_request_ids": [],
                    "noise_request_ids": [],
                    "sample_source_text": source_texts.get(request_id, ""),
                }
            )
            continue
        seen_ids.add(request_id)

        if request_id not in clean_records_by_request_id:
            debug_items.append(
                {
                    "request_id": request_id,
                    "raw_surface": "",
                    "sanitized": "",
                    "accepted_request_ids": [],
                    "rejected_reason": "unexpected_id",
                    "matched_request_ids": [],
                    "noise_request_ids": [],
                    "sample_source_text": "",
                }
            )
            continue

        source_text = source_texts.get(request_id, "")
        raw_surface_forms = item.get("surface_forms")
        if raw_surface_forms is None:
            if "surface_form" in item:
                raw_surface_forms = [item.get("surface_form")]
            elif "terms" in item:
                raw_surface_forms = item.get("terms")
            else:
                debug_items.append(
                    {
                        "request_id": request_id,
                        "raw_surface": "",
                        "sanitized": "",
                        "accepted_request_ids": [],
                        "rejected_reason": "missing_surface_forms",
                        "matched_request_ids": [],
                        "noise_request_ids": [],
                        "sample_source_text": source_text,
                    }
                )
                continue

        if not isinstance(raw_surface_forms, list):
            debug_items.append(
                {
                    "request_id": request_id,
                    "raw_surface": str(raw_surface_forms),
                    "sanitized": "",
                    "accepted_request_ids": [],
                    "rejected_reason": "surface_forms_not_list",
                    "matched_request_ids": [],
                    "noise_request_ids": [],
                    "sample_source_text": source_text,
                }
            )
            continue

        for raw_surface in raw_surface_forms:
            raw_surface_text = str(raw_surface or "").strip()
            candidate = sanitize_candidate(raw_surface)
            if not candidate:
                debug_items.append(
                    {
                        "request_id": request_id,
                        "raw_surface": raw_surface_text,
                        "sanitized": "",
                        "accepted_request_ids": [],
                        "rejected_reason": "sanitize_empty",
                        "matched_request_ids": [],
                        "noise_request_ids": [],
                        "sample_source_text": source_text,
                    }
                )
                continue

            if candidate not in source_text:
                rejected_terms.append(raw_surface_text)
                debug_items.append(
                    {
                        "request_id": request_id,
                        "raw_surface": raw_surface_text,
                        "sanitized": candidate,
                        "accepted_request_ids": [],
                        "rejected_reason": "not_found_in_clean_text",
                        "matched_request_ids": [],
                        "noise_request_ids": [],
                        "sample_source_text": source_text,
                    }
                )
                continue

            if is_noise_candidate(candidate, source_text=source_text):
                rejected_terms.append(raw_surface_text)
                debug_items.append(
                    {
                        "request_id": request_id,
                        "raw_surface": raw_surface_text,
                        "sanitized": candidate,
                        "accepted_request_ids": [],
                        "rejected_reason": "noise_filtered",
                        "matched_request_ids": [],
                        "noise_request_ids": [request_id],
                        "sample_source_text": source_text,
                    }
                )
                continue

            resolved[request_id].append(candidate)
            debug_items.append(
                {
                    "request_id": request_id,
                    "raw_surface": raw_surface_text,
                    "sanitized": candidate,
                    "accepted_request_ids": [request_id],
                    "rejected_reason": "",
                    "matched_request_ids": [request_id],
                    "noise_request_ids": [],
                    "sample_source_text": source_text,
                }
            )

    return {
        "resolved": {
            request_id: preserve_unique(terms)
            for request_id, terms in resolved.items()
            if terms
        },
        "rejected_terms": preserve_unique(rejected_terms),
        "debug_items": debug_items,
    }


def build_candidate_terms_from_chunk_recall(
    resolved_terms_by_request_id: Dict[str, Sequence[str]],
    clean_records_by_request_id: Dict[str, TermRecallCleanRecord],
    source_records_by_id: Dict[str, SourceRecord],
) -> List[CandidateTerm]:
    candidates: List[CandidateTerm] = []
    for request_id, terms in resolved_terms_by_request_id.items():
        clean_record = clean_records_by_request_id.get(request_id)
        if clean_record is None:
            continue
        source_record = source_records_by_id.get(clean_record.source_record_id)
        for term in preserve_unique([str(item) for item in terms]):
            context_text = clean_record.original_text or clean_record.term_recall_clean_text
            location = {
                "file_name": source_record.file_name if source_record else "",
                "source_type": source_record.source_type if source_record else "",
                "sheet_or_unit": source_record.sheet_or_unit if source_record else "",
                "row_index": source_record.row_index if source_record else 0,
                "column_name": source_record.column_name if source_record else "",
                "segment_index": 1,
                "total_segments": 1,
            }
            candidates.append(
                CandidateTerm(
                    candidate_id="{0}:chunk_llm:{1}".format(clean_record.source_record_id, term),
                    surface_form=term,
                    source_record_id=clean_record.source_record_id,
                    segment_id=clean_record.source_record_id,
                    recall_source="chunk_llm",
                    context_text=context_text,
                    evidence_text=term,
                    source_record_ids=[clean_record.source_record_id],
                    segment_ids=[clean_record.source_record_id],
                    source_locations=[location],
                    sample_contexts=[context_text],
                    occurrence_count=1,
                )
            )
    return candidates


def _parse_candidate_terms_payload(data: Any, source_text: str) -> List[str]:
    results: List[str] = []
    if not isinstance(data, list):
        return results

    for item in data:
        if isinstance(item, dict):
            surface_form = item.get("surface_form") or item.get("term") or item.get("术语名")
        else:
            surface_form = item
        candidate = sanitize_candidate(surface_form)
        if candidate and candidate in source_text and not is_noise_candidate(candidate, source_text=source_text):
            results.append(candidate)
    return preserve_unique(results)


def parse_candidate_batch_response(
    response: str,
    segments_by_request_id: Dict[str, TextSegment],
) -> Dict[str, Any]:
    cleaned = clean_api_response(response)
    data = _load_json_fragment(cleaned)
    data = _extract_batch_items(data, "候选召回")

    resolved: Dict[str, List[str]] = {}
    item_issues: Dict[str, List[str]] = defaultdict(list)
    item_warnings: Dict[str, List[str]] = defaultdict(list)
    batch_issues: List[str] = []
    seen_ids = set()

    for index, item in enumerate(data, start=1):
        if not isinstance(item, dict):
            batch_issues.append("第 {0} 项不是 JSON 对象".format(index))
            continue

        item_id = str(item.get("id", "")).strip()
        if not item_id:
            batch_issues.append("第 {0} 项缺少 id".format(index))
            continue
        if item_id not in segments_by_request_id:
            batch_issues.append("返回了未请求的 id：{0}".format(item_id))
            continue
        if item_id in seen_ids:
            item_issues[item_id].append("模型返回了重复 id")
            continue

        seen_ids.add(item_id)
        segment = segments_by_request_id[item_id]
        resolved[item_id] = _parse_surface_forms_field(item, segment.text, item_issues[item_id], item_warnings[item_id])

    for expected_id in segments_by_request_id:
        if expected_id not in seen_ids:
            item_issues[expected_id].append("模型遗漏了该编号")

    return {
        "resolved": resolved,
        "unresolved_ids": [item_id for item_id, issues in item_issues.items() if issues],
        "item_issues": {item_id: issues for item_id, issues in item_issues.items() if issues},
        "item_warnings": {item_id: warnings for item_id, warnings in item_warnings.items() if warnings},
        "batch_issues": batch_issues,
    }


def _parse_surface_forms_field(
    payload: Dict[str, Any],
    source_text: str,
    issues: List[str],
    warnings: Optional[List[str]] = None,
) -> List[str]:
    raw_surface_forms = payload.get("surface_forms")
    if raw_surface_forms is None:
        if "surface_form" in payload:
            raw_surface_forms = [payload.get("surface_form")]
        elif "terms" in payload:
            raw_surface_forms = payload.get("terms")
        else:
            issues.append("缺少 surface_forms 字段")
            return []

    if not isinstance(raw_surface_forms, list):
        issues.append("surface_forms 不是数组")
        return []

    valid_terms = []
    invalid_terms = []
    invalid_noise_terms = []
    invalid_sentence_terms = []
    for raw_term in raw_surface_forms:
        raw_text = str(raw_term or "").strip()
        if looks_like_unsalvageable_action_candidate(raw_text):
            invalid_terms.append(raw_text)
            invalid_sentence_terms.append(raw_text)
            continue
        candidate = sanitize_candidate(raw_term)
        if candidate and candidate in source_text and not is_noise_candidate(candidate, source_text=source_text):
            valid_terms.append(candidate)
        else:
            invalid_text = raw_text
            invalid_terms.append(invalid_text)
            if any(mark in invalid_text for mark in ("，", ",", "。", ".")):
                invalid_sentence_terms.append(invalid_text)
            elif candidate and is_noise_candidate(candidate, source_text=source_text):
                invalid_noise_terms.append(invalid_text)

    if invalid_terms:
        if valid_terms:
            if warnings is not None:
                warnings.append("已过滤无效、非逐字或噪声候选词")
        else:
            if invalid_sentence_terms and len(invalid_sentence_terms) == len(invalid_terms):
                issues.append("模型提取了条件整段，没有核心术语")
            elif invalid_noise_terms and len(invalid_noise_terms) == len(invalid_terms):
                issues.append("模型返回候选全部被代码/噪声规则过滤")
            else:
                issues.append("模型遗漏编号或返回非逐字候选")
    return preserve_unique(valid_terms)


def parse_review_response(
    response: str,
    candidate: CandidateTerm,
    source_location: Dict[str, Any],
    allowed_term_types: Optional[Sequence[str]] = None,
) -> ReviewedTerm:
    cleaned = clean_api_response(response)
    data = _load_json_fragment(cleaned)
    if not isinstance(data, dict):
        raise ValueError("模型返回内容不是合法的 JSON 对象。")
    return parse_review_payload(data, candidate, source_location, allowed_term_types=allowed_term_types)


def parse_review_payload(
    data: Dict[str, Any],
    candidate: CandidateTerm,
    source_location: Dict[str, Any],
    allowed_term_types: Optional[Sequence[str]] = None,
) -> ReviewedTerm:
    response_surface = str(data.get("surface_form", "")).strip()
    if response_surface != candidate.surface_form:
        raise ValueError("模型返回的 surface_form 与候选词不一致。")
    if not candidate_has_traceable_evidence(candidate):
        raise ValueError("候选词无法逐字追溯到来源文本或清洗文本。")

    term_type = str(data.get("term_type", "")).strip()
    valid_term_types = None if allowed_term_types == [] else set(allowed_term_types or TERM_TYPE_CHOICES)
    if valid_term_types is not None and term_type not in valid_term_types:
        raise ValueError("模型返回了不支持的术语类别：{0}".format(term_type or "空值"))

    decision = _normalize_review_decision(data)
    if decision == REJECTED_DECISION and term_type != "non_term":
        raise ValueError("模型返回 rejected 时 term_type 必须为 non_term。")
    if decision in {APPROVED_DECISION, REVIEW_DECISION} and term_type == "non_term":
        raise ValueError("模型返回 approved/review 时 term_type 不能为 non_term。")

    decision_reason = str(data.get("reason", "")).strip()

    return ReviewedTerm(
        review_id="{0}:review".format(candidate.candidate_id),
        surface_form=candidate.surface_form,
        term_type=term_type,
        decision=decision,
        source_locations=candidate.source_locations or [source_location],
        decision_reason=decision_reason,
        sample_contexts=candidate.sample_contexts or [candidate.context_text],
        occurrence_count=max(1, int(candidate.occurrence_count or 1)),
    )


def candidate_has_traceable_evidence(candidate: CandidateTerm) -> bool:
    surface_form = str(candidate.surface_form or "").strip()
    if not surface_form:
        return False
    evidence_sources = [
        candidate.evidence_text,
        candidate.context_text,
        *list(candidate.sample_contexts or []),
    ]
    for source_text in evidence_sources:
        text = str(source_text or "")
        if surface_form in text:
            return True
        if surface_form in build_clean_text_view(text):
            return True
    return False


def parse_review_batch_response(
    response: str,
    items_by_request_id: Dict[str, Dict[str, Any]],
    allowed_term_types: Optional[Sequence[str]] = None,
) -> Dict[str, Any]:
    cleaned = clean_api_response(response)
    data = _load_json_fragment(cleaned)
    data = _extract_batch_items(data, "术语判定")

    resolved: Dict[str, ReviewedTerm] = {}
    item_issues: Dict[str, List[str]] = defaultdict(list)
    batch_issues: List[str] = []
    seen_ids = set()

    for index, item in enumerate(data, start=1):
        if not isinstance(item, dict):
            batch_issues.append("第 {0} 项不是 JSON 对象".format(index))
            continue

        item_id = str(item.get("id", "")).strip()
        if not item_id:
            batch_issues.append("第 {0} 项缺少 id".format(index))
            continue
        if item_id not in items_by_request_id:
            batch_issues.append("返回了未请求的 id：{0}".format(item_id))
            continue
        if item_id in seen_ids:
            item_issues[item_id].append("模型返回了重复 id")
            continue

        seen_ids.add(item_id)
        candidate = items_by_request_id[item_id]["candidate"]
        source_location = items_by_request_id[item_id]["source_location"]
        try:
            resolved[item_id] = parse_review_payload(
                item,
                candidate,
                source_location,
                allowed_term_types=allowed_term_types,
            )
        except Exception as exc:
            item_issues[item_id].append(str(exc))

    for expected_id in items_by_request_id:
        if expected_id not in seen_ids:
            item_issues[expected_id].append("模型遗漏了该编号")

    return {
        "resolved": resolved,
        "unresolved_ids": [item_id for item_id, issues in item_issues.items() if issues],
        "item_issues": {item_id: issues for item_id, issues in item_issues.items() if issues},
        "batch_issues": batch_issues,
    }


def _normalize_review_decision(data: Dict[str, Any]) -> str:
    raw_decision = str(data.get("decision", "")).strip().lower()
    if raw_decision:
        mapping = {
            "approve": APPROVED_DECISION,
            "approved": APPROVED_DECISION,
            "pass": APPROVED_DECISION,
            "review": REVIEW_DECISION,
            "needs_review": REVIEW_DECISION,
            "rejected": REJECTED_DECISION,
            "reject": REJECTED_DECISION,
        }
        decision = mapping.get(raw_decision)
        if decision:
            return decision
        raise ValueError("decision 不是合法值：{0}".format(raw_decision))

    # Compatibility for old cached/test responses. New prompts must use `decision`.
    is_valid_term = _coerce_bool(data.get("is_valid_term", False))
    term_type = str(data.get("term_type", "")).strip()
    if not is_valid_term or term_type == "non_term":
        return REJECTED_DECISION
    return APPROVED_DECISION


def _coerce_bool(value: object) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return bool(value)

    text = str(value or "").strip().lower()
    if text in {"true", "1", "yes", "y"}:
        return True
    if text in {"false", "0", "no", "n", ""}:
        return False
    raise ValueError("is_valid_term 不是合法布尔值。")


def aggregate_reviewed_terms(
    reviewed_terms: Sequence[ReviewedTerm],
    single_occurrence_approved_policy: str = "send_to_review",
) -> Tuple[List[Dict[str, object]], List[Dict[str, object]]]:
    approved_rows, _ = _aggregate_rows(
        [item for item in reviewed_terms if item.decision == APPROVED_DECISION],
        review_mode=False,
        allow_post_review=False,
        single_occurrence_approved_policy="allow_to_library",
    )
    return approved_rows, []


def _aggregate_rows(
    items: Sequence[ReviewedTerm],
    review_mode: bool,
    allow_post_review: bool,
    single_occurrence_approved_policy: str = "send_to_review",
) -> Tuple[List[Dict[str, object]], List[Dict[str, object]]]:
    grouped: Dict[Tuple[str, str], List[ReviewedTerm]] = defaultdict(list)
    for item in items:
        grouped[(item.surface_form, item.term_type)].append(item)

    main_rows = []
    review_rows = []
    for (surface_form, term_type), group in grouped.items():
        occurrence_count = sum(max(1, item.occurrence_count) for item in group)
        source_locations = []
        sample_contexts = []
        reasons = []
        review_reasons = []
        for item in group:
            source_locations.extend(item.source_locations)
            sample_contexts.extend(item.sample_contexts)
            if item.decision_reason:
                reasons.append(item.decision_reason)
            if item.review_reason:
                review_reasons.append(item.review_reason)

        effective_review_mode = review_mode
        if (
            allow_post_review
            and not review_mode
            and occurrence_count <= 1
            and str(single_occurrence_approved_policy or "send_to_review") != "allow_to_library"
        ):
            effective_review_mode = True
            review_reasons.append("单次出现的候选默认进入待审核")

        row = {
            "术语原文": surface_form,
            "术语类别": term_type,
            "出现次数": occurrence_count,
            "来源文件": "、".join(sorted({loc.get("file_name", "") for loc in source_locations if loc.get("file_name")})),
            "来源位置": "\n".join(format_source_location(location) for location in source_locations[:10]),
            "示例上下文": "\n\n".join(preserve_unique(sample_contexts)[:3]),
            "判定理由": "；".join(preserve_unique(reasons))[:500],
        }
        if effective_review_mode:
            row["待审核原因"] = "；".join(preserve_unique(review_reasons))[:500] or "需要人工确认"
            review_rows.append(row)
        else:
            main_rows.append(row)

    main_rows.sort(key=lambda item: (-int(item["出现次数"]), str(item["术语原文"])))
    review_rows.sort(key=lambda item: (-int(item["出现次数"]), str(item["术语原文"])))
    return main_rows, review_rows


def _load_json_fragment(cleaned: str):
    if not str(cleaned or "").strip():
        raise ValueError("模型返回为空，无法解析 JSON。") from None
    try:
        return json.loads(cleaned)
    except json.JSONDecodeError as exc:
        match = re.search(r"(\[.*\]|\{.*\})", cleaned, flags=re.DOTALL)
        if not match:
            raise ValueError("模型返回内容不是合法 JSON，未找到完整 JSON 片段。") from None
        try:
            return json.loads(match.group(1))
        except json.JSONDecodeError:
            detail = str(exc).strip() or "疑似截断、缺失括号或转义错误"
            raise ValueError("模型返回内容不是合法 JSON，{0}。".format(detail)) from exc


def _extract_batch_items(data: Any, stage_name: str) -> List[Any]:
    if isinstance(data, list):
        return data
    if isinstance(data, dict):
        for key in ("items", "results", "data"):
            value = data.get(key)
            if isinstance(value, list):
                return value
        raise ValueError("模型返回内容不是合法的{0} JSON 对象，缺少 items 数组。".format(stage_name))
    raise ValueError("模型返回内容不是合法的{0} JSON 数组。".format(stage_name))


def reviewed_terms_from_runtime(items: Sequence[Dict[str, Any]]) -> List[ReviewedTerm]:
    return [ReviewedTerm.from_dict(item) for item in items]


def source_records_from_runtime(items: Sequence[Dict[str, Any]]) -> List[SourceRecord]:
    return [SourceRecord.from_dict(item) for item in items]


def segments_from_runtime(items: Sequence[Dict[str, Any]]) -> List[TextSegment]:
    return [TextSegment.from_dict(item) for item in items]


def candidate_terms_from_runtime(items: Sequence[Dict[str, Any]]) -> List[CandidateTerm]:
    return [CandidateTerm.from_dict(item) for item in items]


def save_results_to_excel(
    approved_rows: List[Dict[str, object]],
    review_rows: List[Dict[str, object]],
    failed_items: List[Dict[str, object]],
    output_folder: Path,
    export_review_sheet: bool = False,
    nontrans_regex_rows: Optional[List[Dict[str, object]]] = None,
) -> Path:
    output_folder.mkdir(parents=True, exist_ok=True)
    filename = output_folder / "游戏术语提取结果_{0}.xlsx".format(time.strftime("%Y%m%d%H%M%S"))
    approved_df = pd.DataFrame(approved_rows, columns=TERM_LIBRARY_COLUMNS)
    review_df = pd.DataFrame(review_rows, columns=REVIEW_COLUMNS)
    failure_df = pd.DataFrame(failed_items, columns=FAILURE_COLUMNS)
    nontrans_columns = list((nontrans_regex_rows or [{}])[0].keys()) if nontrans_regex_rows else list(NONTRANS_REGEX_COLUMNS[:2]) + list(NONTRANS_REGEX_COLUMNS[-2:])
    nontrans_df = pd.DataFrame(nontrans_regex_rows or [], columns=nontrans_columns)

    with pd.ExcelWriter(filename, engine="openpyxl") as writer:
        approved_df.to_excel(writer, sheet_name=TERM_LIBRARY_SHEET, index=False)
        if export_review_sheet and review_rows:
            review_df.to_excel(writer, sheet_name=REVIEW_SHEET, index=False)
        failure_df.to_excel(writer, sheet_name=FAILURE_SHEET, index=False)
        nontrans_df.to_excel(writer, sheet_name=NONTRANS_REGEX_SHEET, index=False)
    return filename


def save_nontrans_regex_to_excel(
    nontrans_regex_rows: List[Dict[str, object]],
    output_folder: Path,
) -> Path:
    output_folder.mkdir(parents=True, exist_ok=True)
    filename = output_folder / "非译元素正则_{0}.xlsx".format(time.strftime("%Y%m%d%H%M%S"))
    nontrans_columns = list(nontrans_regex_rows[0].keys()) if nontrans_regex_rows else list(NONTRANS_REGEX_COLUMNS[:2]) + list(NONTRANS_REGEX_COLUMNS[-2:])
    nontrans_df = pd.DataFrame(nontrans_regex_rows, columns=nontrans_columns)
    with pd.ExcelWriter(filename, engine="openpyxl") as writer:
        nontrans_df.to_excel(writer, sheet_name=NONTRANS_REGEX_SHEET, index=False)
    return filename


