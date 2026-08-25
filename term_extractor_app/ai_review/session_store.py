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
                id, title, title_custom, status, prompt_template_id, source_language,
                target_languages_json, auto_start, created_at, updated_at
            ) VALUES (?, ?, ?, 'draft', ?, ?, ?, ?, ?, ?)
            """,
            (
                session_id,
                str(title or "新审校").strip()[:120] or "新审校",
                0 if str(title or "").strip() in {"", "新审校"} else 1,
                prompt_template_id,
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
    session = get_session(session_id)
    if not session:
        return None
    with get_connection() as conn:
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
        "title", "title_custom", "status", "prompt_template_id", "source_language", "target_languages_json",
        "auto_start", "context_summary", "error_message",
    }
    values = {key: value for key, value in fields.items() if key in allowed}
    if "target_languages_json" in values and isinstance(values["target_languages_json"], list):
        values["target_languages_json"] = dumps_json(_normalize_languages(values["target_languages_json"]))
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


def add_attachment(session_id: str, filename: str, data: bytes) -> dict[str, Any]:
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
            (attachment_id, session_id, filename, str(stored_path), digest, len(data), now, now),
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


def create_question(session_id: str, run_id: str, prompt: str, recommended_value: str) -> dict[str, Any]:
    question_id = uuid.uuid4().hex
    with get_connection() as conn:
        conn.execute(
            """
            INSERT INTO workspace_questions (
                id, session_id, run_id, prompt, recommended_label, recommended_value, status, created_at
            ) VALUES (?, ?, ?, ?, '确认推荐方案', ?, 'pending', ?)
            """,
            (question_id, session_id, run_id, prompt, recommended_value, utc_now()),
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
