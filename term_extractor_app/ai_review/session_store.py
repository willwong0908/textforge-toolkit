from __future__ import annotations

import hashlib
import uuid
from pathlib import Path
from typing import Any

from .cache_service import save_upload_file
from .config import UPLOADS_DIR
from .database import dumps_json, get_connection, init_db, loads_json, utc_now


FINAL_SESSION_STATES = {"completed", "partial", "failed"}


def create_session(
    *,
    title: str = "新审校",
    prompt_template_id: str | None = None,
    term_base_id: str | None = None,
    memoq_term_base_ids: list[str] | None = None,
    source_language: str = "auto",
    target_languages: list[str] | None = None,
    auto_start: bool = False,
) -> dict[str, Any]:
    init_db()
    session_id = uuid.uuid4().hex
    now = utc_now()
    targets = _normalize_languages(target_languages or ["auto"])
    with get_connection() as conn:
        conn.execute(
            """
            INSERT INTO review_sessions (
                id, title, title_custom, status, prompt_template_id, term_base_id, memoq_term_base_ids_json, source_language,
                target_languages_json, auto_start, created_at, updated_at
            ) VALUES (?, ?, ?, 'draft', ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                session_id,
                str(title or "新审校").strip()[:120] or "新审校",
                0 if str(title or "").strip() in {"", "新审校"} else 1,
                prompt_template_id,
                term_base_id,
                dumps_json([str(x) for x in (memoq_term_base_ids or []) if str(x).strip()]),
                str(source_language or "auto").strip() or "auto",
                dumps_json(targets),
                1 if auto_start else 0,
                now,
                now,
            ),
        )
    add_message(session_id, "assistant", "message", "请添加文件或直接输入待审校文本。")
    emit_event(session_id, "session.created", {"session_id": session_id})
    return get_session(session_id) or {}


def list_sessions(limit: int = 100) -> list[dict[str, Any]]:
    init_db()
    with get_connection() as conn:
        rows = conn.execute(
            "SELECT * FROM review_sessions ORDER BY updated_at DESC LIMIT ?",
            (max(1, min(int(limit), 500)),),
        ).fetchall()
    return [_session_to_dict(row) for row in rows]


def get_session(session_id: str) -> dict[str, Any] | None:
    init_db()
    with get_connection() as conn:
        row = conn.execute("SELECT * FROM review_sessions WHERE id = ?", (session_id,)).fetchone()
    return _session_to_dict(row) if row else None


def get_session_snapshot(session_id: str) -> dict[str, Any] | None:
    init_db()
    with get_connection() as conn:
        # A chat refresh reads several related tables. Start one read transaction
        # so it cannot combine a newly-ready Workspace run with the message list
        # captured just before that run published its report.
        conn.execute("BEGIN")
        session_row = conn.execute("SELECT * FROM review_sessions WHERE id = ?", (session_id,)).fetchone()
        if not session_row:
            return None
        session = _session_to_dict(session_row)
        attachments = conn.execute(
            "SELECT * FROM review_attachments WHERE session_id = ? ORDER BY created_at",
            (session_id,),
        ).fetchall()
        messages = conn.execute(
            "SELECT * FROM review_messages WHERE session_id = ? ORDER BY created_at, rowid",
            (session_id,),
        ).fetchall()
        questions = conn.execute(
            "SELECT * FROM workspace_questions WHERE session_id = ? ORDER BY created_at",
            (session_id,),
        ).fetchall()
        runs = conn.execute(
            "SELECT * FROM workspace_runs WHERE session_id = ? ORDER BY created_at DESC, rowid DESC",
            (session_id,),
        ).fetchall()
        task_links = conn.execute(
            "SELECT target_language, task_id FROM review_session_tasks WHERE session_id = ? ORDER BY created_at",
            (session_id,),
        ).fetchall()
    return {
        "session": session,
        "attachments": [_attachment_to_dict(row) for row in attachments],
        "messages": [_message_to_dict(row) for row in messages],
        "questions": [_question_to_dict(row) for row in questions],
        "workspace_runs": [_run_to_dict(row) for row in runs],
        "tasks": [dict(row) for row in task_links],
    }


def update_session(session_id: str, **fields: Any) -> dict[str, Any]:
    allowed = {
        "title", "title_custom", "status", "prompt_template_id", "term_base_id", "memoq_term_base_ids_json", "source_language", "target_languages_json",
        "auto_start", "context_summary", "error_message",
    }
    values = {key: value for key, value in fields.items() if key in allowed}
    if "target_languages_json" in values and isinstance(values["target_languages_json"], list):
        values["target_languages_json"] = dumps_json(_normalize_languages(values["target_languages_json"]))
    if "memoq_term_base_ids_json" in values and isinstance(values["memoq_term_base_ids_json"], list):
        values["memoq_term_base_ids_json"] = dumps_json([str(x) for x in values["memoq_term_base_ids_json"] if str(x).strip()])
    if "auto_start" in values:
        values["auto_start"] = 1 if values["auto_start"] else 0
    if "title_custom" in values:
        values["title_custom"] = 1 if values["title_custom"] else 0
    values["updated_at"] = utc_now()
    assignments = ", ".join(f"{key} = ?" for key in values)
    with get_connection() as conn:
        if not conn.execute("SELECT id FROM review_sessions WHERE id = ?", (session_id,)).fetchone():
            raise ValueError("审校会话不存在")
        conn.execute(
            f"UPDATE review_sessions SET {assignments} WHERE id = ?",
            [*values.values(), session_id],
        )
    return get_session(session_id) or {}


def delete_session(session_id: str) -> None:
    snapshot = get_session_snapshot(session_id)
    if snapshot is None:
        raise ValueError("审校会话不存在")
    upload_root = UPLOADS_DIR.resolve()
    paths: list[Path] = []
    for attachment in snapshot["attachments"]:
        path = Path(str(attachment.get("stored_path") or ""))
        try:
            resolved = path.resolve()
        except OSError:
            continue
        if resolved.parent == upload_root:
            paths.append(resolved)
    task_ids = [str(item["task_id"]) for item in snapshot["tasks"]]
    with get_connection() as conn:
        for task_id in task_ids:
            conn.execute("DELETE FROM review_followup_messages WHERE task_id = ?", (task_id,))
            conn.execute("DELETE FROM forbidden_results WHERE task_id = ?", (task_id,))
            conn.execute("DELETE FROM review_results WHERE task_id = ?", (task_id,))
            conn.execute("DELETE FROM review_task_logs WHERE task_id = ?", (task_id,))
            conn.execute("DELETE FROM review_request_states WHERE task_id = ?", (task_id,))
            conn.execute("DELETE FROM review_tasks WHERE id = ?", (task_id,))
        for table in (
            "review_session_events", "review_session_tasks", "workspace_questions", "workspace_runs",
            "review_messages", "review_attachments",
        ):
            conn.execute(f"DELETE FROM {table} WHERE session_id = ?", (session_id,))
        conn.execute("DELETE FROM review_sessions WHERE id = ?", (session_id,))
    for path in paths:
        try:
            path.unlink(missing_ok=True)
        except OSError:
            pass


def release_completed_session_cache(session_id: str) -> dict[str, int]:
    """Drop regenerable payloads after a completed review without losing its results.

    A conversation keeps its result rows, details, and generated Excel files.  The
    original uploaded copy, however, is no longer needed after every task has
    produced an output.  Workspace plans and model-result caches also duplicate
    the reviewed text, so retaining them makes the local database grow quickly.
    Structural learning is intentionally stored separately in extraction_profiles
    and is never touched here.
    """
    snapshot = get_session_snapshot(session_id)
    if snapshot is None:
        raise ValueError("审校会话不存在")
    if str(snapshot["session"].get("status") or "") not in {"completed", "partial"}:
        return _empty_cache_release_summary()

    task_ids = [str(item.get("task_id") or "") for item in snapshot.get("tasks", []) if item.get("task_id")]
    if not task_ids:
        return _empty_cache_release_summary()
    placeholders = ",".join("?" for _ in task_ids)
    final_statuses = {"completed", "completed_with_errors", "failed"}
    with get_connection() as conn:
        task_rows = conn.execute(
            f"SELECT id, status, output_path FROM review_tasks WHERE id IN ({placeholders})", task_ids
        ).fetchall()
    if len(task_rows) != len(task_ids) or any(str(row["status"] or "") not in final_statuses for row in task_rows):
        return _empty_cache_release_summary()
    if not any(str(row["output_path"] or "").strip() for row in task_rows):
        # Keep the source copy if output generation itself failed.  It may still
        # be required to repair or rerun this particular conversation.
        return _empty_cache_release_summary()

    upload_root = UPLOADS_DIR.resolve()
    released_paths: list[Path] = []
    released_bytes = 0
    attachment_updates: list[tuple[str, str]] = []
    for attachment in snapshot.get("attachments", []):
        stored_path = Path(str(attachment.get("stored_path") or ""))
        try:
            resolved = stored_path.resolve()
        except OSError:
            continue
        if resolved.parent != upload_root:
            continue
        try:
            released_bytes += resolved.stat().st_size
        except OSError:
            pass
        released_paths.append(resolved)
        attachment_updates.append((dumps_json(_compact_attachment_manifest(attachment.get("manifest") or {})), str(attachment["id"])))

    run_ids = [str(item.get("id") or "") for item in snapshot.get("workspace_runs", []) if item.get("id")]
    compacted_runs = 0
    compacted_messages = 0
    removed_result_cache_entries = 0
    now = utc_now()
    with get_connection() as conn:
        if attachment_updates:
            conn.executemany(
                "UPDATE review_attachments SET stored_path = '', manifest_json = ?, updated_at = ? WHERE id = ?",
                [(manifest_json, now, attachment_id) for manifest_json, attachment_id in attachment_updates],
            )

        if run_ids:
            run_placeholders = ",".join("?" for _ in run_ids)
            runs = conn.execute(
                f"SELECT id, plan_json FROM workspace_runs WHERE id IN ({run_placeholders})", run_ids
            ).fetchall()
            compacted_run_rows = [
                (dumps_json(_compact_workspace_plan(loads_json(row["plan_json"], {}))), now, str(row["id"]))
                for row in runs
            ]
            if compacted_run_rows:
                conn.executemany(
                    "UPDATE workspace_runs SET plan_json = ?, updated_at = ? WHERE id = ?", compacted_run_rows
                )
                compacted_runs = len(compacted_run_rows)

            messages = conn.execute(
                "SELECT id, payload_json FROM review_messages WHERE session_id = ? AND kind = 'workspace_report'",
                (session_id,),
            ).fetchall()
            compacted_message_rows: list[tuple[str, str]] = []
            for row in messages:
                payload = loads_json(row["payload_json"], {})
                if str(payload.get("_workspace_run_id") or "") not in run_ids:
                    continue
                compacted_message_rows.append((dumps_json(_compact_workspace_plan(payload)), str(row["id"])))
            if compacted_message_rows:
                conn.executemany(
                    "UPDATE review_messages SET payload_json = ? WHERE id = ?", compacted_message_rows
                )
                compacted_messages = len(compacted_message_rows)

        cache_rows = conn.execute(
            f"SELECT DISTINCT cache_key FROM review_results WHERE task_id IN ({placeholders}) AND cache_key <> ''",
            task_ids,
        ).fetchall()
        cache_keys = [str(row["cache_key"]) for row in cache_rows if row["cache_key"]]
        if cache_keys:
            cache_placeholders = ",".join("?" for _ in cache_keys)
            cursor = conn.execute(
                f"DELETE FROM ai_result_cache WHERE cache_key IN ({cache_placeholders})", cache_keys
            )
            removed_result_cache_entries = max(0, int(cursor.rowcount or 0))
        # The normalized result columns are what the preview, details and Excel
        # writer consume. raw_result_json is a duplicate of those values.
        conn.execute(
            f"UPDATE review_results SET raw_result_json = '{{}}', updated_at = ? WHERE task_id IN ({placeholders})",
            (now, *task_ids),
        )

    for path in released_paths:
        try:
            path.unlink(missing_ok=True)
        except OSError:
            pass
    _remove_orphaned_private_uploads()
    return {
        "released_attachment_count": len(released_paths),
        "released_bytes": released_bytes,
        "removed_result_cache_entries": removed_result_cache_entries,
        "compacted_workspace_runs": compacted_runs,
        "compacted_workspace_messages": compacted_messages,
    }


def _empty_cache_release_summary() -> dict[str, int]:
    return {
        "released_attachment_count": 0,
        "released_bytes": 0,
        "removed_result_cache_entries": 0,
        "compacted_workspace_runs": 0,
        "compacted_workspace_messages": 0,
    }


def _compact_attachment_manifest(manifest: dict[str, Any]) -> dict[str, Any]:
    """Keep the file shape only; samples contain user content and are regenerable."""
    structure = dict(manifest.get("structure") or {})
    if str(manifest.get("file_type") or "") == "excel":
        structure = {
            "sheets": [
                {
                    "name": sheet.get("name"),
                    "row_count": sheet.get("row_count"),
                    "columns": [
                        {
                            "index": column.get("index"),
                            "letter": column.get("letter"),
                            "header": column.get("header"),
                        }
                        for column in sheet.get("columns", [])
                    ],
                }
                for sheet in structure.get("sheets", [])
            ]
        }
    elif str(manifest.get("file_type") or "") in {"csv", "tsv"}:
        structure = {"headers": list(structure.get("headers") or [])}
    return {
        "reader_name": str(manifest.get("reader_name") or ""),
        "file_type": str(manifest.get("file_type") or ""),
        "filename": str(manifest.get("filename") or ""),
        "file_hash": str(manifest.get("file_hash") or ""),
        "structure": structure,
        "block_count": int(manifest.get("block_count") or 0),
        "warnings": list(manifest.get("warnings") or []),
        "cache_released": True,
    }


def _compact_workspace_plan(plan: dict[str, Any]) -> dict[str, Any]:
    """Replace full review units with the compact preview retained in chat."""
    compacted = {
        "source_language": plan.get("source_language") or "auto",
        "targets": [
            {
                "language": target.get("language") or "auto",
                "unit_count": len(target.get("units") or []) if isinstance(target.get("units"), list) else int(target.get("unit_count") or 0),
            }
            for target in plan.get("targets", [])
            if isinstance(target, dict)
        ],
        "assumptions": list(plan.get("assumptions") or [])[:5],
        "warnings": list(plan.get("warnings") or [])[:5],
        "confidence": plan.get("confidence", 0),
        "needs_input": False,
        "question": "",
        "file_summaries": list(plan.get("file_summaries") or []),
        "cache_hit": bool(plan.get("cache_hit")),
        "cache_released": True,
    }
    for key in ("_thinking_text", "_workspace_run_id"):
        if key in plan:
            compacted[key] = plan[key]
    return compacted


def _remove_orphaned_private_uploads() -> None:
    """Remove only unreferenced UUID upload copies; never touch source paths."""
    if not UPLOADS_DIR.exists():
        return
    upload_root = UPLOADS_DIR.resolve()
    with get_connection() as conn:
        attachment_rows = conn.execute(
            "SELECT stored_path FROM review_attachments WHERE stored_path <> ''"
        ).fetchall()
        legacy_rows = conn.execute("SELECT stored_path FROM file_batches").fetchall()
    claimed: set[Path] = set()
    for row in [*attachment_rows, *legacy_rows]:
        try:
            path = Path(str(row["stored_path"] or "")).resolve()
        except OSError:
            continue
        if path.parent == upload_root:
            claimed.add(path)
    for child in upload_root.iterdir():
        try:
            if child.is_file() and child.resolve() not in claimed:
                child.unlink(missing_ok=True)
        except OSError:
            continue


def add_attachment(
    session_id: str,
    filename: str,
    data: bytes,
) -> dict[str, Any]:
    session = get_session(session_id)
    if not session:
        raise ValueError("审校会话不存在")
    if not data:
        raise ValueError("上传文件为空")
    stored_path = save_upload_file(filename, data)
    attachment_id = uuid.uuid4().hex
    now = utc_now()
    digest = hashlib.sha256(data).hexdigest()
    with get_connection() as conn:
        conn.execute(
            """
            INSERT INTO review_attachments (
                id, session_id, original_filename, stored_path, file_hash,
                size_bytes, status, created_at, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, 'uploaded', ?, ?)
            """,
            (
                attachment_id,
                session_id,
                filename,
                str(stored_path),
                digest,
                len(data),
                now,
                now,
            ),
        )
    update_session(session_id, status="draft")
    if not session.get("title_custom"):
        with get_connection() as conn:
            attachment_count = int(
                conn.execute(
                    "SELECT COUNT(*) FROM review_attachments WHERE session_id = ?",
                    (session_id,),
                ).fetchone()[0]
            )
        if attachment_count == 1:
            update_session(session_id, title=Path(filename).name[:120] or "新审校")
    attachment = get_attachment(attachment_id) or {}
    emit_event(session_id, "attachment.uploaded", {"attachment": _public_attachment(attachment)})
    return attachment


def get_attachment(attachment_id: str) -> dict[str, Any] | None:
    with get_connection() as conn:
        row = conn.execute("SELECT * FROM review_attachments WHERE id = ?", (attachment_id,)).fetchone()
    return _attachment_to_dict(row) if row else None


def update_attachment(attachment_id: str, **fields: Any) -> dict[str, Any]:
    allowed = {
        "file_type", "status", "manifest_json", "mapping_mode", "mapping_preset_id",
        "sent_at", "error_message",
    }
    values = {key: value for key, value in fields.items() if key in allowed}
    if "manifest_json" in values and not isinstance(values["manifest_json"], str):
        values["manifest_json"] = dumps_json(values["manifest_json"])
    values["updated_at"] = utc_now()
    assignments = ", ".join(f"{key} = ?" for key in values)
    with get_connection() as conn:
        conn.execute(
            f"UPDATE review_attachments SET {assignments} WHERE id = ?",
            [*values.values(), attachment_id],
        )
    return get_attachment(attachment_id) or {}


def get_pending_attachments(session_id: str) -> list[dict[str, Any]]:
    with get_connection() as conn:
        rows = conn.execute(
            """
            SELECT * FROM review_attachments
            WHERE session_id = ? AND sent_at IS NULL
            ORDER BY created_at, rowid
            """,
            (session_id,),
        ).fetchall()
    return [_attachment_to_dict(row) for row in rows]


def mark_attachments_sent(session_id: str, attachment_ids: list[str]) -> list[dict[str, Any]]:
    normalized = list(dict.fromkeys(str(value) for value in attachment_ids if value))
    if not normalized:
        return []
    placeholders = ",".join("?" for _ in normalized)
    now = utc_now()
    with get_connection() as conn:
        conn.execute(
            f"UPDATE review_attachments SET sent_at = ?, updated_at = ? "
            f"WHERE session_id = ? AND sent_at IS NULL AND id IN ({placeholders})",
            (now, now, session_id, *normalized),
        )
        rows = conn.execute(
            f"SELECT * FROM review_attachments WHERE session_id = ? AND id IN ({placeholders}) "
            "ORDER BY created_at, rowid",
            (session_id, *normalized),
        ).fetchall()
    return [_attachment_to_dict(row) for row in rows]


def delete_pending_attachment(session_id: str, attachment_id: str) -> None:
    with get_connection() as conn:
        row = conn.execute(
            "SELECT * FROM review_attachments WHERE id = ? AND session_id = ?",
            (attachment_id, session_id),
        ).fetchone()
        if not row:
            raise ValueError("附件不存在")
        if row["sent_at"]:
            raise ValueError("已发送的附件不能从对话中移除")
        conn.execute("DELETE FROM review_attachments WHERE id = ?", (attachment_id,))
    path = Path(str(row["stored_path"] or ""))
    try:
        resolved = path.resolve()
        if resolved.parent == UPLOADS_DIR.resolve():
            resolved.unlink(missing_ok=True)
    except OSError:
        pass
    emit_event(session_id, "attachment.deleted", {"attachment_id": attachment_id})


def add_message(
    session_id: str,
    role: str,
    kind: str,
    content: str,
    payload: dict[str, Any] | None = None,
) -> dict[str, Any]:
    message_id = uuid.uuid4().hex
    now = utc_now()
    with get_connection() as conn:
        conn.execute(
            """
            INSERT INTO review_messages (id, session_id, role, kind, content, payload_json, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (message_id, session_id, role, kind, str(content or ""), dumps_json(payload or {}), now),
        )
        conn.execute("UPDATE review_sessions SET updated_at = ? WHERE id = ?", (now, session_id))
        row = conn.execute("SELECT * FROM review_messages WHERE id = ?", (message_id,)).fetchone()
    message = _message_to_dict(row)
    emit_event(session_id, "message.created", {"message": message})
    return message


def recent_context(session_id: str, limit: int = 24, max_chars: int = 32000) -> list[dict[str, str]]:
    with get_connection() as conn:
        rows = conn.execute(
            """
            SELECT role, content FROM review_messages
            WHERE session_id = ? AND kind IN ('message', 'workspace_report', 'question_answer')
            ORDER BY created_at DESC, rowid DESC LIMIT ?
            """,
            (session_id, max(1, limit)),
        ).fetchall()
    result: list[dict[str, str]] = []
    total = 0
    for row in reversed(rows):
        content = str(row["content"] or "")
        if total + len(content) > max_chars:
            content = content[: max(0, max_chars - total)]
        if not content:
            continue
        result.append({"role": str(row["role"]), "content": content})
        total += len(content)
        if total >= max_chars:
            break
    return result


def create_workspace_run(
    session_id: str, *, model: str, input_signature: str, attachment_ids: list[str] | None = None
) -> dict[str, Any]:
    run_id = uuid.uuid4().hex
    now = utc_now()
    with get_connection() as conn:
        conn.execute(
            """
            UPDATE workspace_questions
            SET status = 'superseded', answered_at = ?
            WHERE session_id = ? AND status = 'pending'
            """,
            (now, session_id),
        )
        conn.execute(
            """
            INSERT INTO workspace_runs (
                id, session_id, status, model, input_signature, attachment_ids_json, created_at, updated_at
            ) VALUES (?, ?, 'inspecting', ?, ?, ?, ?, ?)
            """,
            (run_id, session_id, model, input_signature, dumps_json(attachment_ids or []), now, now),
        )
        row = conn.execute("SELECT * FROM workspace_runs WHERE id = ?", (run_id,)).fetchone()
    return _run_to_dict(row)


def update_workspace_run(run_id: str, *, status: str, plan: dict[str, Any] | None = None, error: str = "") -> None:
    with get_connection() as conn:
        conn.execute(
            "UPDATE workspace_runs SET status = ?, plan_json = ?, error_message = ?, updated_at = ? WHERE id = ?",
            (status, dumps_json(plan or {}), error, utc_now(), run_id),
        )


def create_question(
    session_id: str,
    run_id: str,
    prompt: str,
    recommended_value: str = "",
    *,
    recommended_label: str = "确认推荐方案",
) -> dict[str, Any]:
    question_id = uuid.uuid4().hex
    with get_connection() as conn:
        conn.execute(
            """
            INSERT INTO workspace_questions (
                id, session_id, run_id, prompt, recommended_label, recommended_value, status, created_at
            ) VALUES (?, ?, ?, ?, ?, ?, 'pending', ?)
            """,
            (question_id, session_id, run_id, prompt, recommended_label, recommended_value, utc_now()),
        )
        row = conn.execute("SELECT * FROM workspace_questions WHERE id = ?", (question_id,)).fetchone()
    question = _question_to_dict(row)
    emit_event(session_id, "workspace.question", {"question": question})
    return question


def answer_question(question_id: str, answer: str) -> dict[str, Any]:
    with get_connection() as conn:
        row = conn.execute("SELECT * FROM workspace_questions WHERE id = ?", (question_id,)).fetchone()
        if not row:
            raise ValueError("确认问题不存在")
        if row["status"] != "pending":
            raise ValueError("这个问题已经回答")
        conn.execute(
            "UPDATE workspace_questions SET answer = ?, status = 'answered', answered_at = ? WHERE id = ?",
            (str(answer or "").strip(), utc_now(), question_id),
        )
        updated = conn.execute("SELECT * FROM workspace_questions WHERE id = ?", (question_id,)).fetchone()
    question = _question_to_dict(updated)
    add_message(question["session_id"], "user", "question_answer", question["answer"])
    return question


def link_task(session_id: str, target_language: str, task_id: str) -> None:
    with get_connection() as conn:
        conn.execute(
            """
            INSERT OR IGNORE INTO review_session_tasks (session_id, target_language, task_id, created_at)
            VALUES (?, ?, ?, ?)
            """,
            (session_id, target_language, task_id, utc_now()),
        )
    emit_event(session_id, "review.task_created", {"task_id": task_id, "target_language": target_language})


def emit_event(session_id: str, event_type: str, payload: dict[str, Any] | None = None) -> int:
    with get_connection() as conn:
        cursor = conn.execute(
            "INSERT INTO review_session_events (session_id, event_type, payload_json, created_at) VALUES (?, ?, ?, ?)",
            (session_id, event_type, dumps_json(payload or {}), utc_now()),
        )
        return int(cursor.lastrowid)


def get_events(session_id: str, after_id: int = 0, limit: int = 200) -> list[dict[str, Any]]:
    with get_connection() as conn:
        rows = conn.execute(
            """
            SELECT * FROM review_session_events
            WHERE session_id = ? AND id > ? ORDER BY id LIMIT ?
            """,
            (session_id, max(0, int(after_id)), max(1, min(int(limit), 1000))),
        ).fetchall()
    return [
        {
            "id": row["id"],
            "event_type": row["event_type"],
            "payload": loads_json(row["payload_json"], {}),
            "created_at": row["created_at"],
        }
        for row in rows
    ]


def _normalize_languages(values: list[str]) -> list[str]:
    result: list[str] = []
    for value in values:
        normalized = str(value or "").strip() or "auto"
        if normalized not in result:
            result.append(normalized)
    return result or ["auto"]


def _session_to_dict(row: Any) -> dict[str, Any]:
    return {
        "id": row["id"],
        "title": row["title"],
        "title_custom": bool(row["title_custom"]) if "title_custom" in row.keys() else False,
        "status": row["status"],
        "prompt_template_id": row["prompt_template_id"],
        "term_base_id": row["term_base_id"] if "term_base_id" in row.keys() else None,
        "memoq_term_base_ids": loads_json(row["memoq_term_base_ids_json"], []) if "memoq_term_base_ids_json" in row.keys() else [],
        "source_language": row["source_language"],
        "target_languages": loads_json(row["target_languages_json"], ["auto"]),
        "auto_start": bool(row["auto_start"]),
        "context_summary": row["context_summary"],
        "error_message": row["error_message"],
        "created_at": row["created_at"],
        "updated_at": row["updated_at"],
    }


def _attachment_to_dict(row: Any) -> dict[str, Any]:
    return {
        "id": row["id"],
        "session_id": row["session_id"],
        "original_filename": row["original_filename"],
        "stored_path": row["stored_path"],
        "file_type": row["file_type"],
        "file_hash": row["file_hash"],
        "size_bytes": row["size_bytes"],
        "status": row["status"],
        "manifest": loads_json(row["manifest_json"], {}),
        "mapping_mode": row["mapping_mode"] if "mapping_mode" in row.keys() else "ai",
        "mapping_preset_id": row["mapping_preset_id"] if "mapping_preset_id" in row.keys() else None,
        "sent_at": row["sent_at"] if "sent_at" in row.keys() else None,
        "error_message": row["error_message"],
        "created_at": row["created_at"],
        "updated_at": row["updated_at"],
    }


def _public_attachment(item: dict[str, Any]) -> dict[str, Any]:
    return {key: value for key, value in item.items() if key != "stored_path"}


def _message_to_dict(row: Any) -> dict[str, Any]:
    return {
        "id": row["id"], "session_id": row["session_id"], "role": row["role"],
        "kind": row["kind"], "content": row["content"],
        "payload": loads_json(row["payload_json"], {}), "created_at": row["created_at"],
    }


def _question_to_dict(row: Any) -> dict[str, Any]:
    return {
        "id": row["id"], "session_id": row["session_id"], "run_id": row["run_id"],
        "prompt": row["prompt"], "recommended_label": row["recommended_label"],
        "recommended_value": row["recommended_value"], "answer": row["answer"],
        "status": row["status"], "created_at": row["created_at"], "answered_at": row["answered_at"],
    }


def _run_to_dict(row: Any) -> dict[str, Any]:
    return {
        "id": row["id"], "session_id": row["session_id"], "status": row["status"],
        "model": row["model"], "input_signature": row["input_signature"],
        "attachment_ids": loads_json(row["attachment_ids_json"], []) if "attachment_ids_json" in row.keys() else [],
        "plan": loads_json(row["plan_json"], {}), "error_message": row["error_message"],
        "created_at": row["created_at"], "updated_at": row["updated_at"],
    }
