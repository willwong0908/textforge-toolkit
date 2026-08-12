"""Excel diff helpers for the Yeehe toolkit."""

from __future__ import annotations

import json
import os
import posixpath
import sqlite3
import tempfile
import zipfile
from collections import defaultdict
from copy import deepcopy
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from itertools import zip_longest
from typing import Callable, Dict, Iterable, Iterator, List, Sequence, Tuple
from uuid import uuid4
from xml.etree import ElementTree as ET
from xml.sax import handler, make_parser
from xml.sax.saxutils import XMLGenerator

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

from .excel_streaming import header_map_from_values, open_streaming_workbook
from .storage import get_app_paths


SUPPORTED_EXTENSIONS = {".xlsx", ".xlsm"}
DIFF_RESULT_PREFIX = "excel_diff_result_"
DIFF_PREVIEW_LIMIT = 1000
CACHE_PAGE_SIZE = 200
POSITION_MODE = "position"
FIELD_MATCH_MODE = "field_match"
_MAIN_XML_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_REL_XML_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PACKAGE_REL_XML_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


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
    compare_mode_label: str = ""

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
            "compare_mode_label": self.compare_mode_label,
        }


def normalize_path(value: str) -> str:
    text = str(value or "").strip().strip("\"'")
    return str(Path(text).expanduser()) if text else ""


def format_cell_value(value: object) -> str:
    if value is None or (isinstance(value, float) and value != value):
        return ""
    return str(value)


def compare_text_values(val_a: object, val_b: object, *, ignore_case: bool = False, trim_whitespace: bool = False) -> bool:
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
    first_row = next(sheet.iter_rows(min_row=1, max_row=1, values_only=True), ())
    return {header: index + 1 for header, index in header_map_from_values(first_row).items()}


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


def iter_compare_excel_records(
    file_a: str,
    file_b: str,
    *,
    ignore_case: bool = False,
    trim_whitespace: bool = True,
    progress_callback: Callable[[str], None] | None = None,
) -> Iterator[DiffRecord]:
    """Compare matching sheets without materializing either workbook."""
    progress = progress_callback or (lambda _message: None)
    with open_streaming_workbook(file_a) as workbook_a, open_streaming_workbook(file_b) as workbook_b:
        for sheet_name in sorted(set(workbook_a.sheetnames) & set(workbook_b.sheetnames)):
            sheet_a, sheet_b = workbook_a[sheet_name], workbook_b[sheet_name]
            rows_a, rows_b = sheet_a.iter_rows(values_only=True), sheet_b.iter_rows(values_only=True)
            estimated_rows = max(int(sheet_a.max_row or 0), int(sheet_b.max_row or 0))
            for row_number, (values_a, values_b) in enumerate(zip_longest(rows_a, rows_b, fillvalue=()), start=1):
                if row_number == 1 or row_number % 1000 == 0:
                    suffix = f" / {estimated_rows}" if estimated_rows else ""
                    progress(f"正在按位置比对 {sheet_name}（第 {row_number}{suffix} 行）")
                for column_number, (value_a, value_b) in enumerate(zip_longest(values_a, values_b, fillvalue=None), start=1):
                    if compare_text_values(value_a, value_b, ignore_case=ignore_case, trim_whitespace=trim_whitespace):
                        yield _position_record(file_a, file_b, sheet_name, row_number, column_number, value_a, value_b)


def compare_excel_files(file_a: str, file_b: str, *, ignore_case: bool = False, trim_whitespace: bool = False) -> List[DiffRecord]:
    return list(iter_compare_excel_records(file_a, file_b, ignore_case=ignore_case, trim_whitespace=trim_whitespace))


def _field_diff_record(file_a: str, file_b: str, row_a: sqlite3.Row | None, row_b: sqlite3.Row | None, reference_field: str, compare_field: str, *, unmatched: bool = False) -> DiffRecord:
    active_row = row_a or row_b
    ref_value = str(active_row["reference_value"]) if active_row else ""
    values_a = json.loads(row_a["values_json"]) if row_a else {}
    values_b = json.loads(row_b["values_json"]) if row_b else {}
    address_a = str(row_a["reference_address"]) if unmatched and row_a else (str(values_a.get(f"{compare_field}:address", "")) if row_a else "")
    address_b = str(row_b["reference_address"]) if unmatched and row_b else (str(values_b.get(f"{compare_field}:address", "")) if row_b else "")
    return DiffRecord(
        filename_a=Path(file_a).name, filename_b=Path(file_b).name,
        sheet=str(active_row["sheet_name"]) if active_row else "",
        cell_address=address_a or address_b,
        value_a="" if unmatched else str(values_a.get(compare_field, "")),
        value_b="" if unmatched else str(values_b.get(compare_field, "")),
        file_path_a=str(Path(file_a).resolve()), file_path_b=str(Path(file_b).resolve()),
        sheet_a=str(row_a["sheet_name"]) if row_a else "", sheet_b=str(row_b["sheet_name"]) if row_b else "",
        cell_address_a=address_a, cell_address_b=address_b,
        row_a=int(row_a["row_number"]) if row_a else 0, row_b=int(row_b["row_number"]) if row_b else 0,
        reference_field=reference_field, reference_value=ref_value, compare_field=compare_field,
        diff_kind="unmatched_reference" if unmatched else "field_value",
        highlight_scope_a="row" if unmatched and row_a else "cell",
        highlight_scope_b="row" if unmatched and row_b else "cell",
    )


def _create_field_index(connection: sqlite3.Connection) -> None:
    connection.executescript(
        """
        CREATE TABLE rows (
            side TEXT NOT NULL,
            reference_value TEXT NOT NULL,
            occurrence_index INTEGER NOT NULL,
            sheet_name TEXT NOT NULL,
            row_number INTEGER NOT NULL,
            reference_address TEXT NOT NULL,
            values_json TEXT NOT NULL,
            PRIMARY KEY (side, reference_value, occurrence_index)
        );
        CREATE INDEX rows_by_side_reference ON rows(side, reference_value, occurrence_index);
        """
    )


def _index_field_rows(
    connection: sqlite3.Connection,
    file_path: str,
    side: str,
    reference_field: str,
    compare_fields: Sequence[str],
    stats: Dict[str, int],
    progress_callback: Callable[[str], None],
) -> None:
    occurrence_counts: Dict[str, int] = defaultdict(int)
    insert_rows: list[tuple[object, ...]] = []
    with open_streaming_workbook(file_path) as workbook:
        for sheet in workbook.worksheets:
            first_row = next(sheet.iter_rows(min_row=1, max_row=1, values_only=True), ())
            columns = header_map_from_values(first_row)
            reference_index = columns.get(reference_field)
            if reference_index is None:
                continue
            for row_number, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
                if row_number == 2 or row_number % 1000 == 0:
                    progress_callback(f"正在建立 {side} 的字段索引：{sheet.title} 第 {row_number} 行")
                reference_value = format_cell_value(row[reference_index] if reference_index < len(row) else None)
                if not reference_value:
                    continue
                values: Dict[str, str] = {}
                for field in compare_fields:
                    field_index = columns.get(field)
                    if field_index is None:
                        continue
                    values[field] = format_cell_value(row[field_index] if field_index < len(row) else None)
                    values[f"{field}:address"] = f"{get_column_letter(field_index + 1)}{row_number}"
                occurrence_counts[reference_value] += 1
                insert_rows.append(
                    (
                        side,
                        reference_value,
                        occurrence_counts[reference_value],
                        sheet.title,
                        row_number,
                        f"{get_column_letter(reference_index + 1)}{row_number}",
                        json.dumps(values, ensure_ascii=False),
                    )
                )
                stats["valid_reference_rows"] += 1
                if len(insert_rows) >= 1000:
                    connection.executemany("INSERT INTO rows VALUES (?, ?, ?, ?, ?, ?, ?)", insert_rows)
                    connection.commit()
                    insert_rows.clear()
    if insert_rows:
        connection.executemany("INSERT INTO rows VALUES (?, ?, ?, ?, ?, ?, ?)", insert_rows)
        connection.commit()


def iter_compare_field_records(
    file_a: str,
    file_b: str,
    reference_field: str,
    compare_fields: Sequence[str],
    include_unmatched: bool,
    stats: Dict[str, int],
    progress_callback: Callable[[str], None] | None = None,
) -> Iterator[DiffRecord]:
    progress = progress_callback or (lambda _message: None)
    with tempfile.NamedTemporaryFile(prefix="yeehe_diff_field_", suffix=".sqlite3", delete=False) as temporary:
        db_path = temporary.name
    connection = sqlite3.connect(db_path)
    connection.row_factory = sqlite3.Row
    try:
        _create_field_index(connection)
        _index_field_rows(connection, file_a, "A", reference_field, compare_fields, stats, progress)
        _index_field_rows(connection, file_b, "B", reference_field, compare_fields, stats, progress)
        progress("正在按字段配对并生成差异")
        query = """
            SELECT a.*, b.sheet_name AS b_sheet_name, b.row_number AS b_row_number,
                   b.reference_address AS b_reference_address, b.values_json AS b_values_json
            FROM rows a JOIN rows b
              ON a.reference_value = b.reference_value AND a.occurrence_index = b.occurrence_index
            WHERE a.side = 'A' AND b.side = 'B'
            ORDER BY a.rowid
        """
        for row_a in connection.execute(query):
            row_b = {
                "reference_value": row_a["reference_value"], "sheet_name": row_a["b_sheet_name"],
                "row_number": row_a["b_row_number"], "reference_address": row_a["b_reference_address"],
                "values_json": row_a["b_values_json"],
            }
            stats["paired_reference_rows"] += 1
            values_a = json.loads(row_a["values_json"])
            values_b = json.loads(row_b["values_json"])
            for field in compare_fields:
                if field not in values_a or field not in values_b:
                    stats["skipped_field_count"] += 1
                    continue
                if compare_text_values(values_a[field], values_b[field], trim_whitespace=False):
                    yield _field_diff_record(file_a, file_b, row_a, row_b, reference_field, field)

        unmatched_query = """
            SELECT a.* FROM rows a LEFT JOIN rows b
              ON a.reference_value = b.reference_value AND a.occurrence_index = b.occurrence_index AND b.side = ?
            WHERE a.side = ? AND b.reference_value IS NULL ORDER BY a.rowid
        """
        for side, other_side in (("A", "B"), ("B", "A")):
            for row in connection.execute(unmatched_query, (other_side, side)):
                stats["unmatched_reference_rows"] += 1
                if include_unmatched:
                    yield _field_diff_record(file_a, file_b, row if side == "A" else None, row if side == "B" else None, reference_field, reference_field, unmatched=True)
    finally:
        connection.close()
        Path(db_path).unlink(missing_ok=True)


def compare_paths(path_a: str, path_b: str, *, ignore_case: bool = False, trim_whitespace: bool = False, progress_callback: Callable[[str], None] | None = None) -> tuple[List[DiffRecord], CompareMeta]:
    result = run_compare_to_cache(path_a, path_b, ignore_case=ignore_case, trim_whitespace=trim_whitespace, progress_callback=progress_callback)
    records = list(iter_cached_diff_records(str(result["cache_file"])))
    meta_data = result["meta"]
    return records, CompareMeta(**meta_data)


def excel_color_from_hex(color_hex: str) -> str:
    cleaned = str(color_hex or "").replace("#", "").strip()
    if len(cleaned) == 6: return f"FF{cleaned.upper()}"
    if len(cleaned) == 8: return cleaned.upper()
    raise ValueError(f"无法识别的颜色值: {color_hex}")


def _highlight_index_path() -> Path:
    cache_dir = _diff_cache_dir()
    return cache_dir / f"highlight_{uuid4().hex}.sqlite3"


def _create_highlight_index(connection: sqlite3.Connection) -> None:
    connection.executescript(
        """
        CREATE TABLE targets (
            file_path TEXT NOT NULL,
            sheet_name TEXT NOT NULL,
            cell_address TEXT NOT NULL,
            PRIMARY KEY (file_path, sheet_name, cell_address)
        );
        CREATE TABLE row_targets (
            file_path TEXT NOT NULL,
            sheet_name TEXT NOT NULL,
            row_number INTEGER NOT NULL,
            PRIMARY KEY (file_path, sheet_name, row_number)
        );
        CREATE INDEX targets_by_file_sheet ON targets(file_path, sheet_name, cell_address);
        CREATE INDEX rows_by_file_sheet ON row_targets(file_path, sheet_name, row_number);
        """
    )


def _build_highlight_index(records: Iterable[DiffRecord], target: str) -> tuple[Path, int]:
    if target not in {"A", "B"}:
        raise ValueError("target 必须是 A 或 B")
    db_path = _highlight_index_path()
    connection = sqlite3.connect(db_path)
    cell_rows: list[tuple[str, str, str]] = []
    row_rows: list[tuple[str, str, int]] = []
    try:
        _create_highlight_index(connection)
        is_a = target == "A"
        for record in records:
            file_path = record.file_path_a if is_a else record.file_path_b
            sheet_name = (record.sheet_a if is_a else record.sheet_b) or record.sheet
            scope = record.highlight_scope_a if is_a else record.highlight_scope_b
            address = (record.cell_address_a if is_a else record.cell_address_b) or record.cell_address
            row_number = record.row_a if is_a else record.row_b
            if not file_path or not sheet_name:
                continue
            if scope == "row" and row_number > 0:
                row_rows.append((file_path, sheet_name, row_number))
            elif address:
                cell_rows.append((file_path, sheet_name, address))
            if len(cell_rows) + len(row_rows) >= 1000:
                if cell_rows:
                    connection.executemany("INSERT OR IGNORE INTO targets VALUES (?, ?, ?)", cell_rows)
                    cell_rows.clear()
                if row_rows:
                    connection.executemany("INSERT OR IGNORE INTO row_targets VALUES (?, ?, ?)", row_rows)
                    row_rows.clear()
                connection.commit()
        if cell_rows:
            connection.executemany("INSERT OR IGNORE INTO targets VALUES (?, ?, ?)", cell_rows)
        if row_rows:
            connection.executemany("INSERT OR IGNORE INTO row_targets VALUES (?, ?, ?)", row_rows)
        connection.commit()
        count = connection.execute("SELECT COUNT(*) FROM targets").fetchone()[0]
        count += connection.execute("SELECT COUNT(*) FROM row_targets").fetchone()[0]
        return db_path, int(count)
    finally:
        connection.close()


def _worksheet_part_map(archive: zipfile.ZipFile) -> dict[str, str]:
    workbook_root = ET.fromstring(archive.read("xl/workbook.xml"))
    rel_root = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    rel_targets = {
        relation.attrib.get("Id", ""): relation.attrib.get("Target", "")
        for relation in rel_root.findall(f"{{{_PACKAGE_REL_XML_NS}}}Relationship")
    }
    result: dict[str, str] = {}
    for sheet in workbook_root.findall(f".//{{{_MAIN_XML_NS}}}sheet"):
        sheet_name = sheet.attrib.get("name", "")
        relation_id = sheet.attrib.get(f"{{{_REL_XML_NS}}}id", "")
        target = rel_targets.get(relation_id, "")
        if sheet_name and target:
            normalized_target = target.lstrip("/")
            result[sheet_name] = (
                normalized_target
                if normalized_target.startswith("xl/")
                else posixpath.normpath(posixpath.join("xl", normalized_target))
            )
    return result


def _append_highlight_style(styles_xml: bytes, color_hex: str) -> tuple[bytes, dict[str, str]]:
    root = ET.fromstring(styles_xml)
    ns = {"x": _MAIN_XML_NS}
    fills = root.find("x:fills", ns)
    cell_xfs = root.find("x:cellXfs", ns)
    if fills is None or cell_xfs is None:
        raise ValueError("Excel 样式表不完整，无法写入标记。")
    fill = ET.Element(f"{{{_MAIN_XML_NS}}}fill")
    pattern = ET.SubElement(fill, f"{{{_MAIN_XML_NS}}}patternFill", {"patternType": "solid"})
    ET.SubElement(pattern, f"{{{_MAIN_XML_NS}}}fgColor", {"rgb": excel_color_from_hex(color_hex)})
    ET.SubElement(pattern, f"{{{_MAIN_XML_NS}}}bgColor", {"indexed": "64"})
    fill_id = len(fills)
    fills.append(fill)
    fills.attrib["count"] = str(len(fills))
    style_map: dict[str, str] = {}
    original_xfs = list(cell_xfs)
    for old_id, old_xf in enumerate(original_xfs):
        new_xf = deepcopy(old_xf)
        new_xf.attrib["fillId"] = str(fill_id)
        new_xf.attrib["applyFill"] = "1"
        new_id = len(cell_xfs)
        cell_xfs.append(new_xf)
        style_map[str(old_id)] = str(new_id)
    cell_xfs.attrib["count"] = str(len(cell_xfs))
    return ET.tostring(root, encoding="utf-8", xml_declaration=True), style_map


class _WorksheetHighlightHandler(handler.ContentHandler):
    def __init__(self, writer: XMLGenerator, addresses: set[str], rows: set[int], style_map: dict[str, str]) -> None:
        super().__init__()
        self.writer = writer
        self.addresses = addresses
        self.rows = rows
        self.style_map = style_map
        self._cell_depth = 0
        self._active_cell = False
        self.changed_count = 0

    def startDocument(self) -> None:
        self.writer.startDocument()

    def endDocument(self) -> None:
        self.writer.endDocument()

    def startElement(self, name: str, attrs) -> None:
        if name == "c":
            self._cell_depth += 1
            address = str(attrs.get("r", ""))
            row_text = "".join(ch for ch in address if ch.isdigit())
            self._active_cell = address in self.addresses or (row_text.isdigit() and int(row_text) in self.rows)
            copied = dict(attrs)
            if self._active_cell:
                copied["s"] = self.style_map.get(str(copied.get("s", "0")), self.style_map.get("0", "0"))
                self.changed_count += 1
            self.writer.startElement(name, copied)
            return
        self.writer.startElement(name, attrs)

    def endElement(self, name: str) -> None:
        self.writer.endElement(name)
        if name == "c":
            self._cell_depth = max(0, self._cell_depth - 1)
            self._active_cell = False

    def characters(self, content: str) -> None:
        self.writer.characters(content)


def _rewrite_sheet_xml(source, output, addresses: set[str], rows: set[int], style_map: dict[str, str]) -> int:
    """Copy one worksheet XML entry while changing only marked cells."""
    generator = XMLGenerator(output, encoding="utf-8")
    parser = make_parser()
    parser.setFeature(handler.feature_namespaces, False)
    content_handler = _WorksheetHighlightHandler(generator, addresses, rows, style_map)
    parser.setContentHandler(content_handler)
    while chunk := source.read(1024 * 1024):
        parser.feed(chunk)
    parser.close()
    return int(content_handler.changed_count)


def _targets_for_workbook(index_path: Path, workbook_path: str) -> tuple[dict[str, set[str]], dict[str, set[int]]]:
    connection = sqlite3.connect(index_path)
    try:
        cell_targets: dict[str, set[str]] = defaultdict(set)
        row_targets: dict[str, set[int]] = defaultdict(set)
        for sheet_name, address in connection.execute("SELECT sheet_name, cell_address FROM targets WHERE file_path = ?", (workbook_path,)):
            cell_targets[str(sheet_name)].add(str(address))
        for sheet_name, row_number in connection.execute("SELECT sheet_name, row_number FROM row_targets WHERE file_path = ?", (workbook_path,)):
            row_targets[str(sheet_name)].add(int(row_number))
        return cell_targets, row_targets
    finally:
        connection.close()


def _stream_highlight_workbook(workbook_path: str, index_path: Path, color_hex: str) -> int:
    cell_targets, row_targets = _targets_for_workbook(index_path, workbook_path)
    if not cell_targets and not row_targets:
        return 0
    # Expand an unmatched-row request to actual non-empty cells with the same
    # read-only iterator used by comparison. Blank cells are never highlighted.
    for sheet_name, rows in row_targets.items():
        if not rows:
            continue
        sheet_addresses = cell_targets.setdefault(sheet_name, set())
        with open_streaming_workbook(workbook_path, data_only=True) as streaming_workbook:
            sheet = streaming_workbook[sheet_name] if sheet_name in streaming_workbook.sheetnames else None
            if sheet is None:
                continue
            for row_number, values in enumerate(sheet.iter_rows(values_only=True), start=1):
                if row_number not in rows:
                    continue
                for column_number, value in enumerate(values, start=1):
                    if value is not None and str(value) != "":
                        sheet_addresses.add(f"{get_column_letter(column_number)}{row_number}")
    row_targets = {}
    source_path = Path(workbook_path)
    temporary_path = source_path.with_name(f".{source_path.stem}.highlight-{uuid4().hex}{source_path.suffix}")
    changed = 0
    with zipfile.ZipFile(source_path, "r") as source_archive:
        sheet_parts = _worksheet_part_map(source_archive)
        styles_xml, style_map = _append_highlight_style(source_archive.read("xl/styles.xml"), color_hex)
        with zipfile.ZipFile(temporary_path, "w", compression=zipfile.ZIP_DEFLATED, allowZip64=True) as output_archive:
            for info in source_archive.infolist():
                if info.filename == "xl/styles.xml":
                    output_archive.writestr(info, styles_xml)
                    continue
                matched_sheet = next((sheet_name for sheet_name, part_name in sheet_parts.items() if info.filename == part_name), "")
                if matched_sheet:
                    addresses = cell_targets.get(matched_sheet, set())
                    rows = row_targets.get(matched_sheet, set())
                    if addresses or rows:
                        with source_archive.open(info, "r") as source_entry, output_archive.open(info, "w") as output_entry:
                            changed += _rewrite_sheet_xml(source_entry, output_entry, addresses, rows, style_map)
                        continue
                with source_archive.open(info, "r") as source_entry, output_archive.open(info, "w") as output_entry:
                    while chunk := source_entry.read(1024 * 1024):
                        output_entry.write(chunk)
    with zipfile.ZipFile(temporary_path, "r") as verification:
        if verification.testzip() is not None:
            raise ValueError("标记后的 Excel 校验失败，原文件未修改。")
    os.replace(temporary_path, source_path)
    return changed


def apply_highlight_to_records(records: Iterable[DiffRecord], target: str, color_hex: str) -> tuple[int, int]:
    index_path, _ = _build_highlight_index(records, target)
    connection = sqlite3.connect(index_path)
    try:
        paths = [str(row[0]) for row in connection.execute("SELECT DISTINCT file_path FROM targets UNION SELECT DISTINCT file_path FROM row_targets")]
    finally:
        connection.close()
    try:
        changed_cells = sum(_stream_highlight_workbook(path, index_path, color_hex) for path in paths)
        return changed_cells, len(paths)
    finally:
        index_path.unlink(missing_ok=True)


def _diff_output_dir() -> Path:
    output_dir = get_app_paths().output_dir / "excel_diff"; output_dir.mkdir(parents=True, exist_ok=True); return output_dir


def _diff_cache_dir() -> Path:
    cache_dir = _diff_output_dir() / "cache"; cache_dir.mkdir(parents=True, exist_ok=True); return cache_dir


def _cache_meta_path(cache_file: str | Path) -> Path:
    return Path(f"{cache_file}.meta.json")


def _compact_cache_item(record: DiffRecord, pair_index: int) -> dict[str, object]:
    """Remove pair-level strings repeated on every large-cache line."""
    if record.diff_kind == "cell":
        return {"p": pair_index, "s": record.sheet, "c": record.cell_address, "a": record.value_a, "b": record.value_b}
    return {
        "p": pair_index, "sa": record.sheet_a, "sb": record.sheet_b,
        "ca": record.cell_address_a, "cb": record.cell_address_b,
        "ra": record.row_a, "rb": record.row_b,
        "rf": record.reference_field, "rv": record.reference_value, "cf": record.compare_field,
        "a": record.value_a, "b": record.value_b, "k": record.diff_kind,
        "ha": record.highlight_scope_a, "hb": record.highlight_scope_b,
    }


def _expand_cache_item(item: dict[str, object], manifest: dict[str, object]) -> DiffRecord:
    pairs = manifest.get("pairs") if isinstance(manifest.get("pairs"), list) else []
    pair_index = int(item.get("p", 0) or 0)
    pair = pairs[pair_index] if 0 <= pair_index < len(pairs) and isinstance(pairs[pair_index], dict) else {}
    filename_a, filename_b = str(pair.get("filename_a", "") or ""), str(pair.get("filename_b", "") or "")
    file_path_a, file_path_b = str(pair.get("file_path_a", "") or ""), str(pair.get("file_path_b", "") or "")
    if "s" in item:
        sheet, cell = str(item.get("s", "") or ""), str(item.get("c", "") or "")
        return DiffRecord(filename_a, filename_b, sheet, cell, str(item.get("a", "") or ""), str(item.get("b", "") or ""), file_path_a, file_path_b, sheet, sheet, cell, cell)
    return DiffRecord(
        filename_a, filename_b,
        str(item.get("sa", "") or "") or str(item.get("sb", "") or ""),
        str(item.get("ca", "") or "") or str(item.get("cb", "") or ""),
        str(item.get("a", "") or ""), str(item.get("b", "") or ""), file_path_a, file_path_b,
        sheet_a=str(item.get("sa", "") or ""), sheet_b=str(item.get("sb", "") or ""),
        cell_address_a=str(item.get("ca", "") or ""), cell_address_b=str(item.get("cb", "") or ""),
        row_a=int(item.get("ra", 0) or 0), row_b=int(item.get("rb", 0) or 0),
        reference_field=str(item.get("rf", "") or ""), reference_value=str(item.get("rv", "") or ""),
        compare_field=str(item.get("cf", "") or ""), diff_kind=str(item.get("k", "field_value") or "field_value"),
        highlight_scope_a=str(item.get("ha", "cell") or "cell"), highlight_scope_b=str(item.get("hb", "cell") or "cell"),
    )


def run_compare_to_cache(path_a: str, path_b: str, *, compare_mode: str = POSITION_MODE, reference_field: str = "", compare_fields: Sequence[str] = (), include_unmatched: bool = False, ignore_case: bool = False, trim_whitespace: bool = False, progress_callback: Callable[[str], None] | None = None, preview_limit: int = CACHE_PAGE_SIZE) -> Dict[str, object]:
    if compare_mode not in {POSITION_MODE, FIELD_MATCH_MODE}: raise ValueError("不支持的比对模式。")
    if compare_mode == FIELD_MATCH_MODE:
        reference_field = str(reference_field or "").strip()
        compare_fields = [str(field).strip() for field in compare_fields if str(field).strip()]
        if not reference_field or not compare_fields: raise ValueError("按字段匹配需要选择参考字段和至少一个比对字段。")
        if reference_field in compare_fields: raise ValueError("参考字段不能同时作为比对字段。")
    progress = progress_callback or (lambda _message: None)
    pairs, file_mode_label, files_in_a, files_in_b = _resolve_pairs(path_a, path_b)
    if not pairs:
        raise ValueError("未找到可配对的 Excel 文件，请确认两个路径有效，目录模式下文件名需要一致。")
    result_id = uuid4().hex
    cache_path = _diff_cache_dir() / f"{result_id}.jsonl"
    preview_records: List[Dict[str, object]] = []; diff_count = 0
    manifest: dict[str, object] = {"format": 2, "pairs": [], "page_size": CACHE_PAGE_SIZE, "page_offsets": []}
    stats = {"valid_reference_rows": 0, "paired_reference_rows": 0, "unmatched_reference_rows": 0, "skipped_field_count": 0}
    progress("正在比较两个 Excel 文件" if len(pairs) == 1 else "正在扫描并比较 Excel 文件")
    with cache_path.open("w", encoding="utf-8") as handle:
        for index, (paired_a, paired_b, name) in enumerate(pairs, start=1):
            pair_index = index - 1
            manifest["pairs"].append(
                {
                    "filename_a": Path(paired_a).name,
                    "filename_b": Path(paired_b).name,
                    "file_path_a": str(Path(paired_a).resolve()),
                    "file_path_b": str(Path(paired_b).resolve()),
                }
            )
            progress(f"正在比较 {name} ({index}/{len(pairs)})")
            records = (
                iter_compare_excel_records(
                    paired_a,
                    paired_b,
                    ignore_case=ignore_case,
                    trim_whitespace=trim_whitespace,
                    progress_callback=progress,
                )
                if compare_mode == POSITION_MODE
                else iter_compare_field_records(
                    paired_a,
                    paired_b,
                    reference_field,
                    compare_fields,
                    include_unmatched,
                    stats,
                    progress,
                )
            )
            for record in records:
                if diff_count % CACHE_PAGE_SIZE == 0:
                    manifest["page_offsets"].append(handle.tell())
                item = _compact_cache_item(record, pair_index)
                handle.write(json.dumps(item, ensure_ascii=False, separators=(",", ":")) + "\n")
                diff_count += 1
                if len(preview_records) < max(1, int(preview_limit or DIFF_PREVIEW_LIMIT)):
                    preview_records.append(record.to_dict())
    manifest["total_count"] = diff_count
    _cache_meta_path(cache_path).write_text(json.dumps(manifest, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    meta = CompareMeta(
        mode_label=file_mode_label,
        files_in_a=files_in_a,
        files_in_b=files_in_b,
        matched_pairs=len(pairs),
        diff_count=diff_count,
        compare_mode_label="按位置比对" if compare_mode == POSITION_MODE else "按字段匹配",
        **stats,
    )
    return {"result_id": result_id, "cache_file": str(cache_path), "preview_records": preview_records, "preview_limit": max(1, int(preview_limit or DIFF_PREVIEW_LIMIT)), "preview_truncated": diff_count > len(preview_records), "meta": meta.to_dict(), "total_count": diff_count, "file_mode_label": file_mode_label}


def iter_cached_diff_records(cache_file: str, *, query: str = "") -> Iterator[DiffRecord]:
    lowered = str(query or "").strip().lower(); path = Path(cache_file)
    if not path.exists(): return
    meta_path = _cache_meta_path(path)
    manifest: dict[str, object] = {}
    if meta_path.exists():
        try:
            manifest = json.loads(meta_path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            manifest = {}
    with path.open("r", encoding="utf-8") as handle:
        for line in handle:
            text = str(line or "").strip()
            if not text: continue
            item = json.loads(text)
            record = _expand_cache_item(item, manifest) if manifest.get("format") == 2 else DiffRecord.from_dict(item)
            if not lowered or lowered in record.search_blob: yield record


def read_cached_diff_preview(
    cache_file: str,
    *,
    query: str = "",
    limit: int = DIFF_PREVIEW_LIMIT,
    offset: int = 0,
) -> Dict[str, object]:
    preview: List[Dict[str, object]] = []
    matched_count = 0
    safe_limit = max(1, int(limit or DIFF_PREVIEW_LIMIT))
    safe_offset = max(0, int(offset or 0))
    path = Path(cache_file)
    meta_path = _cache_meta_path(path)
    manifest: dict[str, object] = {}
    if meta_path.exists():
        try:
            manifest = json.loads(meta_path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            manifest = {}
    if not query and manifest.get("format") == 2 and path.exists():
        page_size = int(manifest.get("page_size", CACHE_PAGE_SIZE) or CACHE_PAGE_SIZE)
        page_offsets = manifest.get("page_offsets") if isinstance(manifest.get("page_offsets"), list) else []
        page_index = safe_offset // page_size
        skip_count = safe_offset % page_size
        records: list[dict[str, object]] = []
        with path.open("r", encoding="utf-8") as handle:
            if page_index < len(page_offsets):
                handle.seek(int(page_offsets[page_index]))
            for _ in range(skip_count):
                if not handle.readline():
                    break
            for _ in range(safe_limit):
                line = handle.readline()
                if not line:
                    break
                records.append(_expand_cache_item(json.loads(line), manifest).to_dict())
        matched_count = int(manifest.get("total_count", 0) or 0)
        return {
            "records": records,
            "matched_count": matched_count,
            "preview_limit": safe_limit,
            "offset": safe_offset,
            "preview_truncated": safe_offset + len(records) < matched_count,
        }
    for record in iter_cached_diff_records(cache_file, query=query):
        matched_count += 1
        if safe_offset <= matched_count - 1 < safe_offset + safe_limit:
            preview.append(record.to_dict())
    return {
        "records": preview,
        "matched_count": matched_count,
        "preview_limit": safe_limit,
        "offset": safe_offset,
        "preview_truncated": safe_offset + len(preview) < matched_count,
    }


def export_diff_records(records: Iterable[DiffRecord], output_path: str = "") -> Dict[str, object]:
    final_path = Path(str(output_path).strip()) if str(output_path).strip() else _diff_output_dir() / f"{DIFF_RESULT_PREFIX}{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}.xlsx"
    if final_path.suffix.lower() not in SUPPORTED_EXTENSIONS: final_path = final_path.with_suffix(".xlsx")
    final_path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook(write_only=True)
    worksheet = workbook.create_sheet("差异列表")
    headers = ["文件A", "文件B", "Sheet", "单元格", "A内容", "B内容", "差异类型", "参考字段", "参考值", "比对字段", "A Sheet", "B Sheet", "A 单元格", "B 单元格"]
    worksheet.append(headers)
    diff_count = 0
    for row in records:
        values = [
            row.filename_a,
            row.filename_b,
            row.sheet,
            row.cell_address,
            row.value_a,
            row.value_b,
            row.diff_kind,
            row.reference_field,
            row.reference_value,
            row.compare_field,
            row.sheet_a or row.sheet,
            row.sheet_b or row.sheet,
            row.cell_address_a or row.cell_address,
            row.cell_address_b or row.cell_address,
        ]
        worksheet.append(values); diff_count += 1
    # Write-only output avoids retaining every difference row in memory.  A
    # fixed width keeps the result readable without reopening it in normal mode.
    workbook.save(final_path); workbook.close()
    return {"output_file": str(final_path), "output_dir": str(final_path.parent), "diff_count": diff_count}


def export_cached_diff_records(cache_file: str, output_path: str = "", *, query: str = "") -> Dict[str, object]:
    return export_diff_records(iter_cached_diff_records(cache_file, query=query), output_path)


def apply_highlight_from_cache(cache_file: str, target: str, color_hex: str, *, query: str = "") -> tuple[int, int]:
    return apply_highlight_to_records(iter_cached_diff_records(cache_file, query=query), target, color_hex)
