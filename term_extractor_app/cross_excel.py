"""Cross-Excel search and merge helpers."""

from __future__ import annotations

from copy import copy
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Dict, Iterable, List, Sequence

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.cell import WriteOnlyCell

from .core import _open_excel_file
from .excel_streaming import header_map_from_values, is_streamable_excel, open_streaming_workbook
from .storage import get_app_paths


SUPPORTED_EXTENSIONS = {".xlsx", ".xls", ".xlsm"}
MERGE_RESULT_PREFIXES = ("合并结果_", "merge_result_")


@dataclass
class CrossExcelSearchMatch:
    file_name: str
    sheet_name: str
    row_index: int
    row_values: List[str]
    matched_columns: List[int]

    def to_dict(self) -> Dict[str, object]:
        return {
            "file_name": self.file_name,
            "sheet_name": self.sheet_name,
            "row_index": self.row_index,
            "row_values": list(self.row_values),
            "matched_columns": list(self.matched_columns),
        }


def _is_valid_excel_file(path: Path) -> bool:
    if path.suffix.lower() not in SUPPORTED_EXTENSIONS:
        return False
    if path.name.startswith("~$"):
        return False
    lowered_name = path.name.lower()
    if any(lowered_name.startswith(prefix.lower()) for prefix in MERGE_RESULT_PREFIXES):
        return False
    return True


def get_all_excel_files(folder_path: str) -> List[Path]:
    root = Path(folder_path)
    if not root.exists() or not root.is_dir():
        raise ValueError("文件夹路径无效。")
    files = [path for path in root.glob("**/*") if path.is_file() and _is_valid_excel_file(path)]
    return sorted(files, key=lambda item: str(item).lower())


def collect_all_headers(excel_files: Sequence[Path]) -> tuple[List[str], Dict[str, Dict[str, List[str]]]]:
    all_headers = set()
    file_sheet_headers: Dict[str, Dict[str, List[str]]] = {}

    for file_path in excel_files:
        file_sheet_headers[str(file_path)] = {}
        if is_streamable_excel(file_path):
            with open_streaming_workbook(file_path) as workbook:
                for sheet in workbook.worksheets:
                    first_row = next(sheet.iter_rows(min_row=1, max_row=1, values_only=True), ())
                    headers = list(header_map_from_values(first_row).keys())
                    file_sheet_headers[str(file_path)][sheet.title] = headers
                    all_headers.update(headers)
        else:
            with _open_excel_file(str(file_path)) as excel_file:
                for sheet_name in excel_file.sheet_names:
                    df = pd.read_excel(excel_file, sheet_name=sheet_name, nrows=0, keep_default_na=False)
                    headers = [str(column) for column in df.columns]
                    file_sheet_headers[str(file_path)][sheet_name] = headers
                    all_headers.update(headers)

    return sorted(all_headers), file_sheet_headers


def scan_cross_excel_folder(folder_path: str) -> Dict[str, object]:
    excel_files = get_all_excel_files(folder_path)
    headers, file_sheet_headers = collect_all_headers(excel_files)
    return {
        "folder_path": folder_path,
        "file_count": len(excel_files),
        "headers": headers,
        "files": [path.name for path in excel_files],
        "file_sheet_headers": file_sheet_headers,
    }


def _trim_row_values(values: Iterable[object]) -> List[str]:
    normalized = ["" if value is None else str(value) for value in values]
    last_non_empty = -1
    for index, value in enumerate(normalized):
        if str(value).strip():
            last_non_empty = index
    if last_non_empty < 0:
        return []
    return normalized[: last_non_empty + 1]


def search_excel_rows(folder_path: str, query: str, limit: int = 300) -> Dict[str, object]:
    search_text = str(query or "").strip()
    if not search_text:
        raise ValueError("请输入要搜索的内容。")

    excel_files = get_all_excel_files(folder_path)
    lowered = search_text.casefold()
    matches: List[CrossExcelSearchMatch] = []
    scanned_rows = 0

    for file_path in excel_files:
        if is_streamable_excel(file_path):
            with open_streaming_workbook(file_path) as workbook:
                source_sheets = ((sheet.title, sheet.iter_rows(values_only=True)) for sheet in workbook.worksheets)
                for sheet_name, source_rows in source_sheets:
                    for row_offset, row_values in enumerate(source_rows, start=1):
                        scanned_rows += 1
                        trimmed = _trim_row_values(row_values)
                        if not trimmed:
                            continue
                        matched_columns = [index for index, cell_value in enumerate(trimmed) if lowered in str(cell_value or "").casefold()]
                        if not matched_columns:
                            continue
                        matches.append(CrossExcelSearchMatch(file_path.name, sheet_name, row_offset, trimmed, matched_columns))
                        if len(matches) >= max(1, int(limit or 300)):
                            return {"query": search_text, "file_count": len(excel_files), "scanned_rows": scanned_rows, "truncated": True, "items": [item.to_dict() for item in matches]}
        else:
            with _open_excel_file(str(file_path)) as excel_file:
                for sheet_name in excel_file.sheet_names:
                    df = pd.read_excel(excel_file, sheet_name=sheet_name, keep_default_na=False, header=None)
                    for row_offset, row_values in enumerate(df.itertuples(index=False, name=None), start=1):
                        scanned_rows += 1
                        trimmed = _trim_row_values(row_values)
                        if not trimmed:
                            continue
                        matched_columns = [
                            index
                            for index, cell_value in enumerate(trimmed)
                            if lowered in str(cell_value or "").casefold()
                        ]
                        if not matched_columns:
                            continue
                        matches.append(
                            CrossExcelSearchMatch(
                                file_name=file_path.name,
                                sheet_name=sheet_name,
                                row_index=row_offset,
                                row_values=trimmed,
                                matched_columns=matched_columns,
                            )
                        )
                        if len(matches) >= max(1, int(limit or 300)):
                            return {
                                "query": search_text,
                                "file_count": len(excel_files),
                                "scanned_rows": scanned_rows,
                                "truncated": True,
                                "items": [item.to_dict() for item in matches],
                            }

    return {
        "query": search_text,
        "file_count": len(excel_files),
        "scanned_rows": scanned_rows,
        "truncated": False,
        "items": [item.to_dict() for item in matches],
    }


def copy_cell_style(source_cell, target_cell) -> None:
    if source_cell is None or not source_cell.has_style:
        return
    if source_cell.fill:
        target_cell.fill = copy(source_cell.fill)
    if source_cell.font:
        target_cell.font = copy(source_cell.font)
    if source_cell.alignment:
        target_cell.alignment = copy(source_cell.alignment)
    if source_cell.border:
        target_cell.border = copy(source_cell.border)
    if source_cell.number_format:
        target_cell.number_format = source_cell.number_format


def _cross_excel_output_dir() -> Path:
    output_dir = get_app_paths().output_dir / "cross_excel_merge"
    output_dir.mkdir(parents=True, exist_ok=True)
    return output_dir


def merge_excel_files_by_headers(
    folder_path: str,
    selected_headers: Sequence[str],
    apply_format: bool = True,
) -> Dict[str, object]:
    headers = [str(item).strip() for item in selected_headers if str(item).strip()]
    if not headers:
        raise ValueError("请至少选择一个表头。")

    excel_files = get_all_excel_files(folder_path)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    output_path = _cross_excel_output_dir() / f"合并结果_{timestamp}.xlsx"
    workbook = Workbook(write_only=True)
    worksheet = workbook.create_sheet("合并数据")
    output_headers = ["来源文件", "来源Sheet", *headers]
    worksheet.append(output_headers)
    row_count = 0

    for file_path in excel_files:
        if is_streamable_excel(file_path):
            with open_streaming_workbook(file_path, data_only=False) as source_workbook:
                for source_sheet in source_workbook.worksheets:
                    first_row = next(source_sheet.iter_rows(min_row=1, max_row=1), ())
                    header_map = header_map_from_values([cell.value for cell in first_row])
                    if not any(header in header_map for header in headers):
                        continue
                    for source_row in source_sheet.iter_rows(min_row=2):
                        result_row = [file_path.name, source_sheet.title]
                        for header in headers:
                            source_index = header_map.get(header)
                            source_cell = source_row[source_index] if source_index is not None and source_index < len(source_row) else None
                            if apply_format and source_cell is not None:
                                target_cell = WriteOnlyCell(worksheet, value=source_cell.value)
                                copy_cell_style(source_cell, target_cell)
                                result_row.append(target_cell)
                            else:
                                result_row.append(source_cell.value if source_cell is not None else None)
                        worksheet.append(result_row)
                        row_count += 1
            continue

        # .xls cannot use openpyxl's read-only reader.  Preserve its existing
        # compatibility behavior while keeping modern Excel files streaming.
        with _open_excel_file(str(file_path)) as source_workbook:
            for sheet_name in source_workbook.sheet_names:
                frame = pd.read_excel(source_workbook, sheet_name=sheet_name, keep_default_na=False)
                frame.columns = [str(column) for column in frame.columns]
                if not any(header in frame.columns for header in headers):
                    continue
                for row in frame.itertuples(index=False, name=None):
                    values_by_header = dict(zip(frame.columns, row))
                    worksheet.append([file_path.name, sheet_name, *[values_by_header.get(header, "") for header in headers]])
                    row_count += 1

    if not row_count:
        workbook.close()
        raise ValueError("没有找到包含所选表头的数据。")
    workbook.save(output_path)

    return {
        "output_file": str(output_path),
        "row_count": row_count,
        "column_count": len(output_headers),
        "selected_headers": headers,
        "apply_format": bool(apply_format),
        "output_dir": str(output_path.parent),
    }
