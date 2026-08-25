from __future__ import annotations

import json
import sqlite3
from contextlib import contextmanager
from datetime import UTC, datetime
from typing import Any, Iterator

from .config import DB_PATH, ensure_directories


def utc_now() -> str:
    return datetime.now(UTC).replace(tzinfo=None).isoformat(timespec="seconds")


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
            CREATE TABLE IF NOT EXISTS review_request_states (
                id TEXT PRIMARY KEY,
                task_id TEXT NOT NULL,
                package_index INTEGER NOT NULL,
                package_total INTEGER NOT NULL,
                item_count INTEGER NOT NULL,
                first_item_id TEXT NOT NULL DEFAULT '',
                first_source_file TEXT NOT NULL DEFAULT '',
                first_sheet_name TEXT NOT NULL DEFAULT '',
                first_row_number INTEGER,
                status TEXT NOT NULL DEFAULT 'queued',
                attempt_count INTEGER NOT NULL DEFAULT 0,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                UNIQUE(task_id, package_index)
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
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS review_sessions (
                id TEXT PRIMARY KEY,
                title TEXT NOT NULL,
                title_custom INTEGER NOT NULL DEFAULT 0,
                status TEXT NOT NULL DEFAULT 'draft',
                prompt_template_id TEXT,
                source_language TEXT NOT NULL DEFAULT 'auto',
                target_languages_json TEXT NOT NULL DEFAULT '["auto"]',
                auto_start INTEGER NOT NULL DEFAULT 0,
                context_summary TEXT NOT NULL DEFAULT '',
                error_message TEXT NOT NULL DEFAULT '',
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS review_attachments (
                id TEXT PRIMARY KEY,
                session_id TEXT NOT NULL,
                original_filename TEXT NOT NULL,
                stored_path TEXT NOT NULL,
                file_type TEXT NOT NULL DEFAULT '',
                file_hash TEXT NOT NULL,
                size_bytes INTEGER NOT NULL DEFAULT 0,
                status TEXT NOT NULL DEFAULT 'uploaded',
                manifest_json TEXT NOT NULL DEFAULT '{}',
                mapping_mode TEXT NOT NULL DEFAULT 'ai',
                mapping_preset_id TEXT,
                sent_at TEXT,
                error_message TEXT NOT NULL DEFAULT '',
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                FOREIGN KEY(session_id) REFERENCES review_sessions(id)
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS review_messages (
                id TEXT PRIMARY KEY,
                session_id TEXT NOT NULL,
                role TEXT NOT NULL,
                kind TEXT NOT NULL DEFAULT 'message',
                content TEXT NOT NULL DEFAULT '',
                payload_json TEXT NOT NULL DEFAULT '{}',
                created_at TEXT NOT NULL,
                FOREIGN KEY(session_id) REFERENCES review_sessions(id)
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS workspace_runs (
                id TEXT PRIMARY KEY,
                session_id TEXT NOT NULL,
                status TEXT NOT NULL,
                model TEXT NOT NULL DEFAULT '',
                input_signature TEXT NOT NULL DEFAULT '',
                attachment_ids_json TEXT NOT NULL DEFAULT '[]',
                plan_json TEXT NOT NULL DEFAULT '{}',
                error_message TEXT NOT NULL DEFAULT '',
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                FOREIGN KEY(session_id) REFERENCES review_sessions(id)
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS workspace_questions (
                id TEXT PRIMARY KEY,
                session_id TEXT NOT NULL,
                run_id TEXT NOT NULL,
                prompt TEXT NOT NULL,
                recommended_label TEXT NOT NULL DEFAULT '',
                recommended_value TEXT NOT NULL DEFAULT '',
                answer TEXT NOT NULL DEFAULT '',
                status TEXT NOT NULL DEFAULT 'pending',
                created_at TEXT NOT NULL,
                answered_at TEXT,
                FOREIGN KEY(session_id) REFERENCES review_sessions(id)
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS review_session_tasks (
                session_id TEXT NOT NULL,
                target_language TEXT NOT NULL,
                task_id TEXT NOT NULL,
                created_at TEXT NOT NULL,
                PRIMARY KEY(session_id, target_language, task_id)
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS extraction_profiles (
                signature TEXT PRIMARY KEY,
                reader_name TEXT NOT NULL,
                plan_json TEXT NOT NULL,
                confirmed_count INTEGER NOT NULL DEFAULT 1,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS review_session_events (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                session_id TEXT NOT NULL,
                event_type TEXT NOT NULL,
                payload_json TEXT NOT NULL DEFAULT '{}',
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
        if "target_language" not in item_columns:
            conn.execute("ALTER TABLE file_items ADD COLUMN target_language TEXT NOT NULL DEFAULT ''")
        if "location_json" not in item_columns:
            conn.execute("ALTER TABLE file_items ADD COLUMN location_json TEXT NOT NULL DEFAULT '{}'")
        if "references_json" not in item_columns:
            conn.execute("ALTER TABLE file_items ADD COLUMN references_json TEXT NOT NULL DEFAULT '[]'")
        prompt_columns = {
            row["name"]
            for row in conn.execute("PRAGMA table_info(prompt_templates)").fetchall()
        }
        if "forbidden_words_text" not in prompt_columns:
            conn.execute("ALTER TABLE prompt_templates ADD COLUMN forbidden_words_text TEXT NOT NULL DEFAULT ''")
        task_columns = {
            row["name"]
            for row in conn.execute("PRAGMA table_info(review_tasks)").fetchall()
        }
        if "session_id" not in task_columns:
            conn.execute("ALTER TABLE review_tasks ADD COLUMN session_id TEXT")
        if "target_language" not in task_columns:
            conn.execute("ALTER TABLE review_tasks ADD COLUMN target_language TEXT NOT NULL DEFAULT ''")
        session_columns = {
            row["name"]
            for row in conn.execute("PRAGMA table_info(review_sessions)").fetchall()
        }
        if "title_custom" not in session_columns:
            conn.execute("ALTER TABLE review_sessions ADD COLUMN title_custom INTEGER NOT NULL DEFAULT 0")
        attachment_columns = {
            row["name"]
            for row in conn.execute("PRAGMA table_info(review_attachments)").fetchall()
        }
        if "mapping_mode" not in attachment_columns:
            conn.execute("ALTER TABLE review_attachments ADD COLUMN mapping_mode TEXT NOT NULL DEFAULT 'ai'")
        if "mapping_preset_id" not in attachment_columns:
            conn.execute("ALTER TABLE review_attachments ADD COLUMN mapping_preset_id TEXT")
        if "sent_at" not in attachment_columns:
            conn.execute("ALTER TABLE review_attachments ADD COLUMN sent_at TEXT")
            conn.execute(
                """
                UPDATE review_attachments
                SET sent_at = created_at
                WHERE EXISTS (
                    SELECT 1 FROM workspace_runs
                    WHERE workspace_runs.session_id = review_attachments.session_id
                )
                """
            )
        run_columns = {
            row["name"]
            for row in conn.execute("PRAGMA table_info(workspace_runs)").fetchall()
        }
        if "attachment_ids_json" not in run_columns:
            conn.execute("ALTER TABLE workspace_runs ADD COLUMN attachment_ids_json TEXT NOT NULL DEFAULT '[]'")
        conn.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_review_followup_messages_result
            ON review_followup_messages(task_id, result_id, created_at)
            """
        )
        conn.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_review_request_states_task
            ON review_request_states(task_id, package_index)
            """
        )
        conn.execute(
            "CREATE INDEX IF NOT EXISTS idx_review_messages_session ON review_messages(session_id, created_at)"
        )
        conn.execute(
            "CREATE INDEX IF NOT EXISTS idx_review_attachments_session ON review_attachments(session_id, created_at)"
        )
        conn.execute(
            "CREATE INDEX IF NOT EXISTS idx_review_events_session ON review_session_events(session_id, id)"
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
