from __future__ import annotations

import json
import sqlite3
from contextlib import contextmanager
from datetime import datetime
from typing import Any, Iterator

from .config import DB_PATH, ensure_directories


def utc_now() -> str:
    return datetime.utcnow().isoformat(timespec="seconds")


@contextmanager
def get_connection() -> Iterator[sqlite3.Connection]:
    ensure_directories()
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    try:
        yield conn
        conn.commit()
    finally:
        conn.close()


def init_db() -> None:
    with get_connection() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS file_batches (
                id TEXT PRIMARY KEY,
                original_filename TEXT NOT NULL,
                stored_path TEXT NOT NULL,
                file_type TEXT NOT NULL,
                status TEXT NOT NULL,
                source_column TEXT,
                target_column TEXT,
                item_count INTEGER NOT NULL DEFAULT 0,
                metadata_json TEXT NOT NULL DEFAULT '{}',
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS file_items (
                id TEXT PRIMARY KEY,
                batch_id TEXT NOT NULL,
                source_file TEXT NOT NULL,
                sheet_name TEXT,
                segment_id TEXT,
                row_number INTEGER,
                source_text TEXT,
                target_text TEXT,
                info_json TEXT NOT NULL DEFAULT '[]',
                source_column TEXT,
                target_column TEXT,
                status_note TEXT,
                item_order INTEGER NOT NULL,
                created_at TEXT NOT NULL,
                FOREIGN KEY(batch_id) REFERENCES file_batches(id)
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS ai_result_cache (
                cache_key TEXT PRIMARY KEY,
                source_text TEXT NOT NULL,
                target_text TEXT NOT NULL,
                model TEXT NOT NULL,
                prompt_signature TEXT NOT NULL,
                directional_signature TEXT NOT NULL,
                result_json TEXT NOT NULL,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value_json TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS prompt_templates (
                id TEXT PRIMARY KEY,
                name TEXT NOT NULL,
                system_prompt TEXT NOT NULL,
                user_prompt TEXT NOT NULL,
                is_default INTEGER NOT NULL DEFAULT 0,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS directional_templates (
                id TEXT PRIMARY KEY,
                name TEXT NOT NULL,
                items_json TEXT NOT NULL,
                is_default INTEGER NOT NULL DEFAULT 0,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS forbidden_templates (
                id TEXT PRIMARY KEY,
                name TEXT NOT NULL,
                words_text TEXT NOT NULL,
                is_default INTEGER NOT NULL DEFAULT 0,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS excel_mapping_presets (
                id TEXT PRIMARY KEY,
                name TEXT NOT NULL,
                mapping_json TEXT NOT NULL,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS review_tasks (
                id TEXT PRIMARY KEY,
                batch_id TEXT NOT NULL,
                status TEXT NOT NULL,
                total_count INTEGER NOT NULL DEFAULT 0,
                cached_count INTEGER NOT NULL DEFAULT 0,
                requested_count INTEGER NOT NULL DEFAULT 0,
                completed_count INTEGER NOT NULL DEFAULT 0,
                failed_count INTEGER NOT NULL DEFAULT 0,
                output_path TEXT,
                config_json TEXT NOT NULL DEFAULT '{}',
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS review_task_logs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                task_id TEXT NOT NULL,
                level TEXT NOT NULL,
                message TEXT NOT NULL,
                created_at TEXT NOT NULL
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS review_results (
                id TEXT PRIMARY KEY,
                task_id TEXT NOT NULL,
                item_id TEXT NOT NULL,
                cache_key TEXT,
                status TEXT NOT NULL,
                has_issue INTEGER,
                issue_type TEXT,
                issue TEXT,
                suggestion TEXT,
                directional_checks_json TEXT NOT NULL DEFAULT '{}',
                error_message TEXT,
                raw_result_json TEXT NOT NULL DEFAULT '{}',
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS forbidden_results (
                id TEXT PRIMARY KEY,
                task_id TEXT NOT NULL,
                item_id TEXT NOT NULL,
                matched_words TEXT NOT NULL,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS review_followup_messages (
                id TEXT PRIMARY KEY,
                task_id TEXT NOT NULL,
                result_id TEXT NOT NULL,
                role TEXT NOT NULL,
                content TEXT NOT NULL,
                created_at TEXT NOT NULL
            )
            """
        )
        columns = {
            row["name"]
            for row in conn.execute("PRAGMA table_info(review_results)").fetchall()
        }
        if "directional_checks_json" not in columns:
            conn.execute(
                "ALTER TABLE review_results ADD COLUMN directional_checks_json TEXT NOT NULL DEFAULT '{}'"
            )
        item_columns = {
            row["name"]
            for row in conn.execute("PRAGMA table_info(file_items)").fetchall()
        }
        if "info_json" not in item_columns:
            conn.execute("ALTER TABLE file_items ADD COLUMN info_json TEXT NOT NULL DEFAULT '[]'")
        if "source_column" not in item_columns:
            conn.execute("ALTER TABLE file_items ADD COLUMN source_column TEXT")
        if "target_column" not in item_columns:
            conn.execute("ALTER TABLE file_items ADD COLUMN target_column TEXT")
        conn.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_review_followup_messages_result
            ON review_followup_messages(task_id, result_id, created_at)
            """
        )


def dumps_json(data: Any) -> str:
    return json.dumps(data, ensure_ascii=False)


def loads_json(text: str | None, default: Any) -> Any:
    if not text:
        return default
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        return default
