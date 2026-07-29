"""Excel diff helpers for the Yeehe toolkit."""

from __future__ import annotations

import json
from collections import defaultdict
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Callable, Dict, Iterable, Iterator, List, Sequence, Tuple
from uuid import uuid4

from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

from .storage import get_app_paths


SUPPORTED_EXTENSIONS = {".xlsx", ".xlsm"}
DIFF_RESULT_PREFIX = "excel_diff_result_"
DIFF_PREVIEW_LIMIT = 1000
POSITION_MODE = "position"
FIELD_MATCH_MODE = "field_match"


@dataclass(frozen=True)
class DiffRecord:
    filename_a: str
    filename_b: str
    sheet: str
    cell_address: str
    value_a: str
    value_b: str
    file_path_a: str
    file_path_b: str
    sheet_a: str = ""
    sheet_b: str = ""
    cell_address_a: str = ""
    cell_address_b: str = ""
    row_a: int = 0
    row_b: int = 0
    reference_field: str = ""
    reference_value: str = ""
    compare_field: str = ""
    diff_kind: str = "cell"
    highlight_scope_a: str = "cell"
    highlight_scope_b: str = "cell"
    row_cells_a: tuple[str, ...] = ()
    row_cells_b: tuple[str, ...] = ()

    @property
    def pair_label(self) -> str:
        if self.filename_a == self.filename_b:
            return self.filename_a
        return f"{self.filename_a} <> {self.filename_b}"

    @property
    def search_blob(self) -> str:
        return " ".join(
            [
                self.filename_a,
                self.filename_b,
                self.sheet,
                self.sheet_a,
                self.sheet_b,
                self.cell_address,
                self.cell_address_a,
                self.cell_address_b,
                self.reference_field,
                self.reference_value,
                self.compare_field,
                self.value_a,
                self.value_b,
            ]
        ).lower()

    def to_dict(self) -> Dict[str, object]:
        return {
            "filename_a": self.filename_a,
            "filename_b": self.filename_b,
            "sheet": self.sheet,
            "cell_address": self.cell_address,
            "value_a": self.value_a,
            "value_b": self.value_b,
            "file_path_a": self.file_path_a,
            "file_path_b": self.file_path_b,
            "sheet_a": self.sheet_a,
            "sheet_b": self.sheet_b,
            "cell_address_a": self.cell_address_a,
            "cell_address_b": self.cell_address_b,
            "row_a": self.row_a,
            "row_b": self.row_b,
            "reference_field": self.reference_field,
            "reference_value": self.reference_value,
            "compare_field": self.compare_field,
            "diff_kind": self.diff_kind,
            "highlight_scope_a": self.highlight_scope_a,
            "highlight_scope_b": self.highlight_scope_b,
            "row_cells_a": list(self.row_cells_a),
            "row_cells_b": list(self.row_cells_b),
        }

    @classmethod
    def from_dict(cls, item: Dict[str, object]) -> "DiffRecord":
        sheet = str(item.get("sheet", "") or "")
        address = str(item.get("cell_address", "") or "")
        return cls(
            filename_a=str(item.get("filename_a", "") or ""),
            filename_b=str(item.get("filename_b", "") or ""),
            sheet=sheet,
            cell_address=address,
            value_a=str(item.get("value_a", "") or ""),
            value_b=str(item.get("value_b", "") or ""),
            file_path_a=str(item.get("file_path_a", "") or ""),
            file_path_b=str(item.get("file_path_b", "") or ""),
            sheet_a=str(item["sheet_a"]) if "sheet_a" in item else sheet,
            sheet_b=str(item["sheet_b"]) if "sheet_b" in item else sheet,
            cell_address_a=str(item["cell_address_a"]) if "cell_address_a" in item else address,
            cell_address_b=str(item["cell_address_b"]) if "cell_address_b" in item else address,
            row_a=int(item.get("row_a", 0) or 0),
            row_b=int(item.get("row_b", 0) or 0),
            reference_field=str(item.get("reference_field", "") or ""),
            reference_value=str(item.get("reference_value", "") or ""),
            compare_field=str(item.get("compare_field", "") or ""),
            diff_kind=str(item.get("diff_kind", "cell") or "cell"),
            highlight_scope_a=str(item.get("highlight_scope_a", "cell") or "cell"),
            highlight_scope_b=str(item.get("highlight_scope_b", "cell") or "cell"),
            row_cells_a=tuple(str(value) for value in (item.get("row_cells_a") or [])),
            row_cells_b=tuple(str(value) for value in (item.get("row_cells_b") or [])),
        )


@dataclass(frozen=True)
class CompareMeta:
    mode_label: str
    files_in_a: int
    files_in_b: int
    matched_pairs: int
    diff_count: int
    valid_reference_rows: int = 0
    paired_reference_rows: int = 0
    unmatched_reference_rows: int = 0
    skipped_field_count: int = 0

    def to_dict(self) -> Dict[str, object]:
        return {
            "mode_label": self.mode_label,
            "files_in_a": self.files_in_a,
            "files_in_b": self.files_in_b,
            "matched_pairs": self.matched_pairs,
            "diff_count": self.diff_count,
            "valid_reference_rows": self.valid_reference_rows,
            "paired_reference_rows": self.paired_reference_rows,
            "unmatched_reference_rows": self.unmatched_reference_rows,
            "skipped_field_count": self.skipped_field_count,
        }


@dataclass(frozen=True)
class FieldRow:
    sheet_name: str
    row_number: int
    reference_address: str
    reference_value: str
    header_columns: Dict[str, int]
    values: Dict[str, str]
    nonempty_cells: tuple[str, ...]


def normalize_path(value: str) -> str:
    text = str(value or "").strip().strip("\"'")
    return str(Path(text).expanduser()) if text else ""


def format_cell_value(value: object) -> str:
    if value is None or (isinstance(value, float) and value != value):
        return ""
    return str(value)


def compare_text_values(val_a: object, val_b: object, *, ignore_case: bool = False, trim_whitespace: bool = True) -> bool:
    text_a = format_cell_value(val_a)
    text_b = format_cell_value(val_b)
    if trim_whitespace:
        text_a, text_b = text_a.strip(), text_b.strip()
    if ignore_case:
        text_a, text_b = text_a.casefold(), text_b.casefold()
    return text_a != text_b


def _is_supported_excel_file(path: Path) -> bool:
    return path.is_file() and not path.name.startswith("~$") and path.suffix.lower() in SUPPORTED_EXTENSIONS


def scan_excel_files(path: str) -> List[str]:
    normalized = normalize_path(path)
    if not normalized:
        return []
    root = Path(normalized)
    if not root.exists():
        return []
    if _is_supported_excel_file(root):
        return [str(root.resolve())]
    if not root.is_dir():
        return []
    return sorted(str(item.resolve()) for item in root.rglob("*") if _is_supported_excel_file(item))


def match_excel_files(files_a: Sequence[str], files_b: Sequence[str]) -> List[Tuple[str, str, str]]:
    index_a: Dict[str, str] = {}
    index_b: Dict[str, str] = {}
    for file_path in files_a:
        index_a.setdefault(Path(file_path).name, file_path)
    for file_path in files_b:
        index_b.setdefault(Path(file_path).name, file_path)
    return [(index_a[name], index_b[name], name) for name in sorted(set(index_a) & set(index_b))]


def _header_columns(sheet) -> Dict[str, int]:
    result: Dict[str, int] = {}
    for column in range(1, sheet.max_column + 1):
        header = format_cell_value(sheet.cell(1, column).value).strip()
        if header and header not in result:
            result[header] = column
    return result


def _collect_headers(file_path: str) -> set[str]:
    workbook = load_workbook(file_path, data_only=True, read_only=True)
    try:
        return {header for sheet in workbook.worksheets for header in _header_columns(sheet)}
    finally:
        workbook.close()


def _resolve_pairs(path_a: str, path_b: str) -> tuple[List[Tuple[str, str, str]], str, int, int]:
    file_a, file_b = Path(normalize_path(path_a)), Path(normalize_path(path_b))
    is_file_a, is_file_b = _is_supported_excel_file(file_a), _is_supported_excel_file(file_b)
    if is_file_a and is_file_b:
        return [(str(file_a), str(file_b), file_a.name)], "文件对文件", 1, 1
    files_a, files_b = scan_excel_files(str(file_a)), scan_excel_files(str(file_b))
    mode_label = "目录对目录" if file_a.is_dir() or file_b.is_dir() else "混合模式"
    return match_excel_files(files_a, files_b), mode_label, len(files_a), len(files_b)


def scan_field_match_headers(path_a: str, path_b: str) -> Dict[str, object]:
    pairs, mode_label, files_in_a, files_in_b = _resolve_pairs(path_a, path_b)
    headers_a: set[str] = set()
    headers_b: set[str] = set()
    for file_a, file_b, _name in pairs:
        headers_a.update(_collect_headers(file_a))
        headers_b.update(_collect_headers(file_b))
    return {
        "headers_a": sorted(headers_a),
        "headers_b": sorted(headers_b),
        "common_headers": sorted(headers_a & headers_b),
        "matched_pairs": len(pairs),
        "files_in_a": files_in_a,
        "files_in_b": files_in_b,
        "mode_label": mode_label,
    }


def _position_record(file_a: str, file_b: str, sheet_name: str, row: int, col: int, value_a: object, value_b: object) -> DiffRecord:
    address = f"{get_column_letter(col)}{row}"
    return DiffRecord(
        filename_a=Path(file_a).name, filename_b=Path(file_b).name, sheet=sheet_name, cell_address=address,
        value_a=format_cell_value(value_a), value_b=format_cell_value(value_b),
        file_path_a=str(Path(file_a).resolve()), file_path_b=str(Path(file_b).resolve()),
        sheet_a=sheet_name, sheet_b=sheet_name, cell_address_a=address, cell_address_b=address, row_a=row, row_b=row,
    )


def iter_compare_excel_records(file_a: str, file_b: str, *, ignore_case: bool = False, trim_whitespace: bool = True) -> Iterator[DiffRecord]:
    workbook_a, workbook_b = load_workbook(file_a, data_only=True), load_workbook(file_b, data_only=True)
    try:
        for sheet_name in sorted(set(workbook_a.sheetnames) & set(workbook_b.sheetnames)):
            sheet_a, sheet_b = workbook_a[sheet_name], workbook_b[sheet_name]
            for row in range(1, max(sheet_a.max_row, sheet_b.max_row) + 1):
                for col in range(1, max(sheet_a.max_column, sheet_b.max_column) + 1):
                    value_a, value_b = sheet_a.cell(row, col).value, sheet_b.cell(row, col).value
                    if compare_text_values(value_a, value_b, ignore_case=ignore_case, trim_whitespace=trim_whitespace):
                        yield _position_record(file_a, file_b, sheet_name, row, col, value_a, value_b)
    finally:
        workbook_a.close()
        workbook_b.close()


def compare_excel_files(file_a: str, file_b: str, *, ignore_case: bool = False, trim_whitespace: bool = True) -> List[DiffRecord]:
    return list(iter_compare_excel_records(file_a, file_b, ignore_case=ignore_case, trim_whitespace=trim_whitespace))


def _field_rows(workbook, reference_field: str, compare_fields: Sequence[str]) -> List[FieldRow]:
    rows: List[FieldRow] = []
    wanted = [reference_field, *compare_fields]
    for sheet in workbook.worksheets:
        header_columns = _header_columns(sheet)
        reference_column = header_columns.get(reference_field)
        if not reference_column:
            continue
        for row_number in range(2, sheet.max_row + 1):
            reference_value = format_cell_value(sheet.cell(row_number, reference_column).value)
            if not reference_value:
                continue
            values = {field: format_cell_value(sheet.cell(row_number, header_columns[field]).value) for field in wanted if field in header_columns}
            nonempty = tuple(
                f"{get_column_letter(column)}{row_number}"
                for column in range(1, sheet.max_column + 1)
                if format_cell_value(sheet.cell(row_number, column).value)
            )
            rows.append(FieldRow(sheet.title, row_number, f"{get_column_letter(reference_column)}{row_number}", reference_value, header_columns, values, nonempty))
    return rows


def _field_diff_record(file_a: str, file_b: str, row_a: FieldRow | None, row_b: FieldRow | None, reference_field: str, compare_field: str, *, unmatched: bool = False) -> DiffRecord:
    ref_value = row_a.reference_value if row_a else row_b.reference_value if row_b else ""
    address_a = row_a.reference_address if unmatched and row_a else (f"{get_column_letter(row_a.header_columns[compare_field])}{row_a.row_number}" if row_a and compare_field in row_a.header_columns else "")
    address_b = row_b.reference_address if unmatched and row_b else (f"{get_column_letter(row_b.header_columns[compare_field])}{row_b.row_number}" if row_b and compare_field in row_b.header_columns else "")
    return DiffRecord(
        filename_a=Path(file_a).name, filename_b=Path(file_b).name,
        sheet=row_a.sheet_name if row_a else row_b.sheet_name if row_b else "",
        cell_address=address_a or address_b,
        value_a="" if unmatched else (row_a.values.get(compare_field, "") if row_a else ""),
        value_b="" if unmatched else (row_b.values.get(compare_field, "") if row_b else ""),
        file_path_a=str(Path(file_a).resolve()), file_path_b=str(Path(file_b).resolve()),
        sheet_a=row_a.sheet_name if row_a else "", sheet_b=row_b.sheet_name if row_b else "",
        cell_address_a=address_a, cell_address_b=address_b,
        row_a=row_a.row_number if row_a else 0, row_b=row_b.row_number if row_b else 0,
        reference_field=reference_field, reference_value=ref_value, compare_field=compare_field,
        diff_kind="unmatched_reference" if unmatched else "field_value",
        highlight_scope_a="row" if unmatched and row_a else "cell",
        highlight_scope_b="row" if unmatched and row_b else "cell",
        row_cells_a=row_a.nonempty_cells if unmatched and row_a else (),
        row_cells_b=row_b.nonempty_cells if unmatched and row_b else (),
    )


def iter_compare_field_records(file_a: str, file_b: str, reference_field: str, compare_fields: Sequence[str], include_unmatched: bool, stats: Dict[str, int]) -> Iterator[DiffRecord]:
    workbook_a, workbook_b = load_workbook(file_a, data_only=True), load_workbook(file_b, data_only=True)
    try:
        rows_a, rows_b = _field_rows(workbook_a, reference_field, compare_fields), _field_rows(workbook_b, reference_field, compare_fields)
        stats["valid_reference_rows"] += len(rows_a) + len(rows_b)
        by_value_a: Dict[str, List[FieldRow]] = defaultdict(list)
        by_value_b: Dict[str, List[FieldRow]] = defaultdict(list)
        for row in rows_a: by_value_a[row.reference_value].append(row)
        for row in rows_b: by_value_b[row.reference_value].append(row)
        for value in dict.fromkeys([*by_value_a, *by_value_b]):
            left, right = by_value_a.get(value, []), by_value_b.get(value, [])
            paired = min(len(left), len(right))
            stats["paired_reference_rows"] += paired
            for index in range(paired):
                row_a, row_b = left[index], right[index]
                for field in compare_fields:
                    if field not in row_a.header_columns or field not in row_b.header_columns:
                        stats["skipped_field_count"] += 1
                        continue
                    if compare_text_values(row_a.values.get(field, ""), row_b.values.get(field, ""), trim_whitespace=False):
                        yield _field_diff_record(file_a, file_b, row_a, row_b, reference_field, field)
            remaining = [*(left[paired:]), *(right[paired:])]
            stats["unmatched_reference_rows"] += len(remaining)
            if include_unmatched:
                for row in left[paired:]:
                    yield _field_diff_record(file_a, file_b, row, None, reference_field, reference_field, unmatched=True)
                for row in right[paired:]:
                    yield _field_diff_record(file_a, file_b, None, row, reference_field, reference_field, unmatched=True)
    finally:
        workbook_a.close()
        workbook_b.close()


def compare_paths(path_a: str, path_b: str, *, ignore_case: bool = False, trim_whitespace: bool = True, progress_callback: Callable[[str], None] | None = None) -> tuple[List[DiffRecord], CompareMeta]:
    result = run_compare_to_cache(path_a, path_b, ignore_case=ignore_case, trim_whitespace=trim_whitespace, progress_callback=progress_callback)
    records = list(iter_cached_diff_records(str(result["cache_file"])))
    meta_data = result["meta"]
    return records, CompareMeta(**meta_data)


def excel_color_from_hex(color_hex: str) -> str:
    cleaned = str(color_hex or "").replace("#", "").strip()
    if len(cleaned) == 6: return f"FF{cleaned.upper()}"
    if len(cleaned) == 8: return cleaned.upper()
    raise ValueError(f"无法识别的颜色值: {color_hex}")


def apply_highlight_to_records(records: Iterable[DiffRecord], target: str, color_hex: str) -> tuple[int, int]:
    if target not in {"A", "B"}: raise ValueError("target 必须是 A 或 B")
    grouped: Dict[str, Dict[str, set[str]]] = {}
    for record in records:
        is_a = target == "A"
        workbook_path = record.file_path_a if is_a else record.file_path_b
        sheet_name = (record.sheet_a if is_a else record.sheet_b) or record.sheet
        scope = record.highlight_scope_a if is_a else record.highlight_scope_b
        addresses = record.row_cells_a if is_a else record.row_cells_b
        if scope != "row": addresses = ((record.cell_address_a if is_a else record.cell_address_b) or record.cell_address,)
        if not workbook_path or not sheet_name: continue
        sheet_map = grouped.setdefault(workbook_path, {})
        sheet_map.setdefault(sheet_name, set()).update(address for address in addresses if address)
    fill = PatternFill(fill_type="solid", start_color=excel_color_from_hex(color_hex), end_color=excel_color_from_hex(color_hex))
    changed_cells = 0
    for workbook_path, sheet_map in grouped.items():
        workbook = load_workbook(workbook_path)
        try:
            for sheet_name, addresses in sheet_map.items():
                if sheet_name not in workbook.sheetnames: continue
                worksheet = workbook[sheet_name]
                for address in sorted(addresses):
                    worksheet[address].fill = fill
                    changed_cells += 1
            workbook.save(workbook_path)
        finally:
            workbook.close()
    return changed_cells, len(grouped)


def _diff_output_dir() -> Path:
    output_dir = get_app_paths().output_dir / "excel_diff"; output_dir.mkdir(parents=True, exist_ok=True); return output_dir


def _diff_cache_dir() -> Path:
    cache_dir = _diff_output_dir() / "cache"; cache_dir.mkdir(parents=True, exist_ok=True); return cache_dir


def run_compare_to_cache(path_a: str, path_b: str, *, compare_mode: str = POSITION_MODE, reference_field: str = "", compare_fields: Sequence[str] = (), include_unmatched: bool = False, ignore_case: bool = False, trim_whitespace: bool = True, progress_callback: Callable[[str], None] | None = None, preview_limit: int = DIFF_PREVIEW_LIMIT) -> Dict[str, object]:
    if compare_mode not in {POSITION_MODE, FIELD_MATCH_MODE}: raise ValueError("不支持的比对模式。")
    if compare_mode == FIELD_MATCH_MODE:
        reference_field = str(reference_field or "").strip()
        compare_fields = [str(field).strip() for field in compare_fields if str(field).strip()]
        if not reference_field or not compare_fields: raise ValueError("按字段匹配需要选择参考字段和至少一个比对字段。")
        if reference_field in compare_fields: raise ValueError("参考字段不能同时作为比对字段。")
    progress = progress_callback or (lambda _message: None)
    pairs, file_mode_label, files_in_a, files_in_b = _resolve_pairs(path_a, path_b)
    result_id = uuid4().hex
    cache_path = _diff_cache_dir() / f"{result_id}.jsonl"
    preview_records: List[Dict[str, object]] = []; diff_count = 0
    stats = {"valid_reference_rows": 0, "paired_reference_rows": 0, "unmatched_reference_rows": 0, "skipped_field_count": 0}
    progress("正在比较两个 Excel 文件" if len(pairs) == 1 else "正在扫描并比较 Excel 文件")
    with cache_path.open("w", encoding="utf-8") as handle:
        for index, (paired_a, paired_b, name) in enumerate(pairs, start=1):
            progress(f"正在比较 {name} ({index}/{len(pairs)})")
            records = iter_compare_excel_records(paired_a, paired_b, ignore_case=ignore_case, trim_whitespace=trim_whitespace) if compare_mode == POSITION_MODE else iter_compare_field_records(paired_a, paired_b, reference_field, compare_fields, include_unmatched, stats)
            for record in records:
                item = record.to_dict(); handle.write(json.dumps(item, ensure_ascii=False) + "\n"); diff_count += 1
                if len(preview_records) < max(1, int(preview_limit or DIFF_PREVIEW_LIMIT)): preview_records.append(item)
    meta = CompareMeta(
        mode_label="按位置比对" if compare_mode == POSITION_MODE else "按字段匹配",
        files_in_a=files_in_a, files_in_b=files_in_b, matched_pairs=len(pairs), diff_count=diff_count, **stats,
    )
    return {"result_id": result_id, "cache_file": str(cache_path), "preview_records": preview_records, "preview_limit": max(1, int(preview_limit or DIFF_PREVIEW_LIMIT)), "preview_truncated": diff_count > len(preview_records), "meta": meta.to_dict(), "total_count": diff_count, "file_mode_label": file_mode_label}


def iter_cached_diff_records(cache_file: str, *, query: str = "") -> Iterator[DiffRecord]:
    lowered = str(query or "").strip().lower(); path = Path(cache_file)
    if not path.exists(): return
    with path.open("r", encoding="utf-8") as handle:
        for line in handle:
            text = str(line or "").strip()
            if not text: continue
            record = DiffRecord.from_dict(json.loads(text))
            if not lowered or lowered in record.search_blob: yield record


def read_cached_diff_preview(cache_file: str, *, query: str = "", limit: int = DIFF_PREVIEW_LIMIT) -> Dict[str, object]:
    preview: List[Dict[str, object]] = []; matched_count = 0; safe_limit = max(1, int(limit or DIFF_PREVIEW_LIMIT))
    for record in iter_cached_diff_records(cache_file, query=query):
        matched_count += 1
        if len(preview) < safe_limit: preview.append(record.to_dict())
    return {"records": preview, "matched_count": matched_count, "preview_limit": safe_limit, "preview_truncated": matched_count > len(preview)}


def export_diff_records(records: Iterable[DiffRecord], output_path: str = "") -> Dict[str, object]:
    final_path = Path(str(output_path).strip()) if str(output_path).strip() else _diff_output_dir() / f"{DIFF_RESULT_PREFIX}{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}.xlsx"
    if final_path.suffix.lower() not in SUPPORTED_EXTENSIONS: final_path = final_path.with_suffix(".xlsx")
    final_path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook(); worksheet = workbook.active; worksheet.title = "差异列表"
    headers = ["差异类型", "文件A", "文件B", "参考字段", "参考值", "比对字段", "A Sheet", "B Sheet", "A 单元格", "B 单元格", "A内容", "B内容"]
    worksheet.append(headers); max_lengths = [len(item) for item in headers]; diff_count = 0
    for row in records:
        values = [row.diff_kind, row.filename_a, row.filename_b, row.reference_field, row.reference_value, row.compare_field, row.sheet_a or row.sheet, row.sheet_b or row.sheet, row.cell_address_a or row.cell_address, row.cell_address_b or row.cell_address, row.value_a, row.value_b]
        worksheet.append(values); diff_count += 1
        for index, value in enumerate(values): max_lengths[index] = max(max_lengths[index], min(len(str(value or "")), 80))
    for column_index, max_len in enumerate(max_lengths, start=1): worksheet.column_dimensions[get_column_letter(column_index)].width = max(12, min(max_len + 2, 100))
    workbook.save(final_path); workbook.close()
    return {"output_file": str(final_path), "output_dir": str(final_path.parent), "diff_count": diff_count}


def export_cached_diff_records(cache_file: str, output_path: str = "", *, query: str = "") -> Dict[str, object]:
    return export_diff_records(iter_cached_diff_records(cache_file, query=query), output_path)


def apply_highlight_from_cache(cache_file: str, target: str, color_hex: str, *, query: str = "") -> tuple[int, int]:
    return apply_highlight_to_records(iter_cached_diff_records(cache_file, query=query), target, color_hex)
