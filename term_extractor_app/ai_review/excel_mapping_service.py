from __future__ import annotations

import uuid
from typing import Any

from .database import dumps_json, get_connection, loads_json, utc_now


def list_excel_mapping_presets() -> list[dict[str, Any]]:
    with get_connection() as conn:
        rows = conn.execute(
            """
            SELECT * FROM excel_mapping_presets
            ORDER BY updated_at DESC
            """
        ).fetchall()
    return [_preset_to_dict(row) for row in rows]


def get_excel_mapping_preset(preset_id: str) -> dict[str, Any]:
    with get_connection() as conn:
        row = conn.execute("SELECT * FROM excel_mapping_presets WHERE id = ?", (preset_id,)).fetchone()
    if not row:
        raise ValueError("预设不存在")
    return _preset_to_dict(row)


def save_excel_mapping_preset(name: str, mapping: dict[str, Any], preset_id: str | None = None) -> dict[str, Any]:
    preset_name = _normalize_preset_name(name)
    normalized_mapping = _normalize_mapping(mapping)
    now = utc_now()
    preset_id = preset_id or uuid.uuid4().hex
    with get_connection() as conn:
        same_name = conn.execute(
            "SELECT id FROM excel_mapping_presets WHERE lower(trim(name)) = lower(trim(?))",
            (preset_name,),
        ).fetchone()
        if same_name and same_name["id"] != preset_id:
            raise ValueError("已存在同名映射模板")
        row = conn.execute("SELECT id FROM excel_mapping_presets WHERE id = ?", (preset_id,)).fetchone()
        if row:
            conn.execute(
                """
                UPDATE excel_mapping_presets
                SET name = ?, mapping_json = ?, updated_at = ?
                WHERE id = ?
                """,
                (preset_name, dumps_json(normalized_mapping), now, preset_id),
            )
        else:
            conn.execute(
                """
                INSERT INTO excel_mapping_presets (id, name, mapping_json, created_at, updated_at)
                VALUES (?, ?, ?, ?, ?)
                """,
                (preset_id, preset_name, dumps_json(normalized_mapping), now, now),
            )
    return get_excel_mapping_preset(preset_id)


def delete_excel_mapping_preset(preset_id: str) -> None:
    with get_connection() as conn:
        row = conn.execute("SELECT id FROM excel_mapping_presets WHERE id = ?", (preset_id,)).fetchone()
        if not row:
            raise ValueError("映射预设不存在")
        conn.execute("DELETE FROM excel_mapping_presets WHERE id = ?", (preset_id,))


def _preset_to_dict(row: Any) -> dict[str, Any]:
    return {
        "id": row["id"],
        "name": row["name"],
        "mapping": loads_json(row["mapping_json"], {}),
        "created_at": row["created_at"],
        "updated_at": row["updated_at"],
    }


def _normalize_preset_name(name: str) -> str:
    preset_name = str(name or "").strip()
    if not preset_name:
        raise ValueError("请输入模板名称")
    return preset_name


def _normalize_mapping(mapping: dict[str, Any]) -> dict[str, Any]:
    if not isinstance(mapping, dict):
        raise ValueError("映射模板内容无效")

    sheets = mapping.get("sheets")
    if not isinstance(sheets, list) or not sheets:
        raise ValueError("请至少配置一组原文列和译文列")

    normalized_sheets: list[dict[str, Any]] = []
    configured_count = 0
    for sheet_index, sheet in enumerate(sheets, start=1):
        if not isinstance(sheet, dict):
            raise ValueError(f"第 {sheet_index} 个工作表配置无效")
        mappings = sheet.get("mappings")
        if not isinstance(mappings, list):
            raise ValueError(f"第 {sheet_index} 个工作表配置无效")

        normalized_mappings: list[dict[str, Any]] = []
        source_columns: set[int] = set()
        target_columns: set[int] = set()
        for item in mappings:
            if not isinstance(item, dict):
                raise ValueError(f"第 {sheet_index} 个工作表存在无效列配置")
            source_column = _as_column_index(item.get("source_column"))
            target_column = _as_column_index(item.get("target_column"))
            if source_column is None or target_column is None:
                raise ValueError(f"第 {sheet_index} 个工作表存在未完成的原文或译文列配置")
            if source_column == target_column:
                raise ValueError(f"第 {sheet_index} 个工作表的原文列和译文列不能相同")
            if source_column in source_columns:
                raise ValueError(f"第 {sheet_index} 个工作表的原文列不能重复")
            if target_column in target_columns:
                raise ValueError(f"第 {sheet_index} 个工作表的译文列不能重复")

            info_columns = _normalize_info_columns(item.get("info_columns"), source_column, target_column, sheet_index)
            normalized_mappings.append(
                {
                    "source_column": source_column,
                    "target_column": target_column,
                    "info_columns": info_columns,
                }
            )
            source_columns.add(source_column)
            target_columns.add(target_column)
            configured_count += 1

        normalized_sheets.append(
            {
                # Keep the snapshot for people reviewing a template. Matching always uses array order.
                "sheet_name": str(sheet.get("sheet_name") or ""),
                "mappings": normalized_mappings,
            }
        )

    if configured_count == 0:
        raise ValueError("请至少配置一组原文列和译文列")
    return {
        "schema_version": 2,
        "source_language": str(mapping.get("source_language") or "").strip(),
        "target_language": str(mapping.get("target_language") or "").strip(),
        "sheets": normalized_sheets,
    }


def _as_column_index(value: Any) -> int | None:
    try:
        index = int(value)
    except (TypeError, ValueError):
        return None
    return index if index >= 0 else None


def _normalize_info_columns(
    value: Any,
    source_column: int,
    target_column: int,
    sheet_index: int,
) -> list[dict[str, Any]]:
    if not isinstance(value, list):
        return []
    normalized: list[dict[str, Any]] = []
    seen: set[int] = set()
    for item in value:
        if not isinstance(item, dict):
            continue
        column = _as_column_index(item.get("column"))
        if column is None:
            raise ValueError(f"第 {sheet_index} 个工作表存在无效信息列")
        if column in {source_column, target_column}:
            raise ValueError(f"第 {sheet_index} 个工作表的信息列不能与原文列或译文列相同")
        if column in seen:
            continue
        seen.add(column)
        normalized.append({"column": column, "category": str(item.get("category") or "").strip()})
    return normalized
