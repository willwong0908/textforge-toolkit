from __future__ import annotations

import os
import uuid
from pathlib import Path
from typing import Any

from .config import UPLOADS_DIR, ensure_directories
from .database import dumps_json, get_connection, loads_json, utc_now
from ..open_utils import open_folder


def save_upload_file(filename: str, data: bytes) -> Path:
    ensure_directories()
    suffix = Path(filename).suffix.lower()
    stored_name = f"{uuid.uuid4().hex}{suffix}"
    path = UPLOADS_DIR / stored_name
    path.write_bytes(data)
    return path


def create_batch(
    *,
    original_filename: str,
    stored_path: Path,
    file_type: str,
    status: str,
    metadata: dict[str, Any] | None = None,
) -> str:
    batch_id = uuid.uuid4().hex
    now = utc_now()
    with get_connection() as conn:
        conn.execute(
            """
            INSERT INTO file_batches (
                id, original_filename, stored_path, file_type, status,
                item_count, metadata_json, created_at, updated_at
            )
            VALUES (?, ?, ?, ?, ?, 0, ?, ?, ?)
            """,
            (
                batch_id,
                original_filename,
                str(stored_path),
                file_type,
                status,
                dumps_json(metadata or {}),
                now,
                now,
            ),
        )
    return batch_id


def replace_batch_items(
    *,
    batch_id: str,
    items: list[dict[str, Any]],
    source_column: str | None = None,
    target_column: str | None = None,
    status: str = "ready",
    metadata_update: dict[str, Any] | None = None,
) -> None:
    now = utc_now()
    with get_connection() as conn:
        batch = conn.execute("SELECT metadata_json FROM file_batches WHERE id = ?", (batch_id,)).fetchone()
        if not batch:
            raise ValueError("读取批次不存在")
        metadata = loads_json(batch["metadata_json"], {})
        metadata.update(metadata_update or {})

        conn.execute("DELETE FROM file_items WHERE batch_id = ?", (batch_id,))
        conn.executemany(
            """
            INSERT INTO file_items (
                id, batch_id, source_file, sheet_name, segment_id, row_number,
                source_text, target_text, info_json, source_column, target_column,
                target_language, location_json, references_json,
                status_note, item_order, created_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            [
                (
                    uuid.uuid4().hex,
                    batch_id,
                    item["source_file"],
                    item.get("sheet_name"),
                    item.get("segment_id"),
                    item.get("row_number"),
                    item.get("source_text", ""),
                    item.get("target_text", ""),
                    dumps_json(item.get("info", [])),
                    item.get("source_column"),
                    item.get("target_column"),
                    item.get("target_language", ""),
                    dumps_json(item.get("location", {})),
                    dumps_json(item.get("references", [])),
                    item.get("status_note", ""),
                    item["item_order"],
                    now,
                )
                for item in items
            ],
        )
        conn.execute(
            """
            UPDATE file_batches
            SET status = ?, source_column = ?, target_column = ?, item_count = ?,
                metadata_json = ?, updated_at = ?
            WHERE id = ?
            """,
            (
                status,
                source_column,
                target_column,
                len(items),
                dumps_json(metadata),
                now,
                batch_id,
            ),
        )


def get_batch(batch_id: str) -> dict[str, Any] | None:
    with get_connection() as conn:
        row = conn.execute("SELECT * FROM file_batches WHERE id = ?", (batch_id,)).fetchone()
    return _batch_to_dict(row) if row else None


def get_latest_ready_batch() -> dict[str, Any] | None:
    with get_connection() as conn:
        row = conn.execute(
            """
            SELECT * FROM file_batches
            WHERE status = 'ready'
            ORDER BY updated_at DESC
            LIMIT 1
            """
        ).fetchone()
    return _batch_to_dict(row) if row else None


def get_batch_items(batch_id: str, limit: int | None = None) -> list[dict[str, Any]]:
    sql = """
        SELECT * FROM file_items
        WHERE batch_id = ?
        ORDER BY item_order ASC
    """
    params: tuple[Any, ...] = (batch_id,)
    if limit is not None:
        sql += " LIMIT ?"
        params = (batch_id, limit)
    with get_connection() as conn:
        rows = conn.execute(sql, params).fetchall()
    return [_item_to_dict(row) for row in rows]


def clear_read_cache() -> None:
    with get_connection() as conn:
        rows = conn.execute("SELECT stored_path FROM file_batches").fetchall()
        conn.execute("DELETE FROM file_items")
        conn.execute("DELETE FROM file_batches")
    for row in rows:
        path = Path(row["stored_path"])
        if path.exists() and path.is_file():
            try:
                path.unlink()
            except OSError:
                pass
    if UPLOADS_DIR.exists():
        for child in UPLOADS_DIR.iterdir():
            if child.is_file():
                try:
                    child.unlink()
                except OSError:
                    pass


def open_directory(path: str | os.PathLike[str]) -> None:
    open_folder(path)


def _batch_to_dict(row: Any) -> dict[str, Any]:
    return {
        "id": row["id"],
        "original_filename": row["original_filename"],
        "stored_path": row["stored_path"],
        "file_type": row["file_type"],
        "status": row["status"],
        "source_column": row["source_column"],
        "target_column": row["target_column"],
        "item_count": row["item_count"],
        "metadata": loads_json(row["metadata_json"], {}),
        "created_at": row["created_at"],
        "updated_at": row["updated_at"],
    }


def _item_to_dict(row: Any) -> dict[str, Any]:
    return {
        "id": row["id"],
        "batch_id": row["batch_id"],
        "source_file": row["source_file"],
        "sheet_name": row["sheet_name"],
        "segment_id": row["segment_id"],
        "row_number": row["row_number"],
        "source_text": row["source_text"],
        "target_text": row["target_text"],
        "info": loads_json(row["info_json"], []) if "info_json" in row.keys() else [],
        "source_column": row["source_column"] if "source_column" in row.keys() else None,
        "target_column": row["target_column"] if "target_column" in row.keys() else None,
        "target_language": row["target_language"] if "target_language" in row.keys() else "",
        "location": loads_json(row["location_json"], {}) if "location_json" in row.keys() else {},
        "references": loads_json(row["references_json"], []) if "references_json" in row.keys() else [],
        "status_note": row["status_note"],
        "item_order": row["item_order"],
    }
