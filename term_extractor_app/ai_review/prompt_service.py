from __future__ import annotations

import uuid
from typing import Any

from .database import get_connection, utc_now


DEFAULT_TEMPLATE_NAME = "\u901a\u7528\u5ba1\u6821"
DEFAULT_SYSTEM_PROMPT = (
    "\u4f60\u662f\u4e13\u4e1a\u7ffb\u8bd1\u5ba1\u6821\uff0c"
    "\u8bf7\u68c0\u67e5\u8bd1\u6587\u662f\u5426\u51c6\u786e\u3001\u81ea\u7136\u3001\u5b8c\u6574\uff0c"
    "\u5e76\u53ea\u8f93\u51fa\u7b26\u5408\u8981\u6c42\u7684 JSON\u3002"
)
DEFAULT_USER_PROMPT = (
    "\u8bf7\u5ba1\u6821\u4ee5\u4e0b JSON \u6761\u76ee\uff0c\u8fd4\u56de\u4e25\u683c JSON\uff0c"
    "\u4e0d\u8981\u8f93\u51fa\u989d\u5916\u89e3\u91ca\u3002\n\n"
    "\u8981\u6c42\uff1a\n"
    "1. \u8fd4\u56de\u683c\u5f0f\u5fc5\u987b\u662f {\"items\":[...]}\u3002\n"
    "2. \u6bcf\u4e2a\u6761\u76ee\u5fc5\u987b\u5305\u542b id\u3001has_issue\u3001issue_type\u3001issue\u3001suggestion\u3002\n"
    "3. \u5982\u679c\u6ca1\u6709\u95ee\u9898\uff0chas_issue \u4e3a false\uff0cissue_type\u3001issue\u3001suggestion \u8fd4\u56de\u7a7a\u5b57\u7b26\u4e32\u3002\n"
    "4. \u5982\u679c\u6709\u95ee\u9898\uff0cissue_type \u7528\u7b80\u6d01\u5206\u7c7b\uff0c"
    "\u4f8b\u5982\u6f0f\u8bd1\u3001\u9519\u8bd1\u3001\u8bed\u6cd5\u95ee\u9898\u3001\u683c\u5f0f\u95ee\u9898\u3001\u8bed\u4e49\u4e0d\u6e05\u3002"
    "\u8fd9\u4e9b\u53ea\u662f\u95ee\u9898\u5f52\u7c7b\u6807\u7b7e\uff0c\u4e0d\u662f\u989d\u5916\u68c0\u67e5\u6e05\u5355\uff0c"
    "\u4e0d\u8981\u56e0\u4e3a\u6807\u7b7e\u793a\u4f8b\u800c\u6269\u5927\u5ba1\u6821\u8303\u56f4\u3002\n"
    "5. \u5982\u679c\u6709\u95ee\u9898\uff0csuggestion \u5fc5\u987b\u53ea\u586b\u5199\u5b8c\u6574\u4fee\u6539\u540e\u7684\u8bd1\u6587\u53e5\u5b50\uff0c\u4e0d\u8981\u89e3\u91ca\uff0c\u4e0d\u8981\u5199\u4fee\u6539\u7406\u7531\uff0c\u4e0d\u8981\u53ea\u5199\u7247\u6bb5\u3002\n"
    "6. \u5982\u679c\u6761\u76ee\u5305\u542b info \u5b57\u6bb5\uff0c\u8bf7\u628a info \u4f5c\u4e3a\u53c2\u8003\u4fe1\u606f\uff1b"
    "info \u4e2d category \u662f\u4fe1\u606f\u7c7b\u522b\uff0cvalue \u662f\u4fe1\u606f\u5185\u5bb9\u3002\n"
    "7. \u5982\u679c\u6761\u76ee\u4e0d\u5305\u542b info \u5b57\u6bb5\uff0c\u4e0d\u8981\u5047\u8bbe\u5b58\u5728\u53c2\u8003\u4fe1\u606f\u3002\n\n"
    "\u5f85\u5ba1\u6821 JSON\uff1a\n"
    "{text}"
)
DEFAULT_USER_PROMPT_MARKER = "\u4e0d\u8981\u56e0\u4e3a\u6807\u7b7e\u793a\u4f8b\u800c\u6269\u5927\u5ba1\u6821\u8303\u56f4"


def ensure_default_prompt_template() -> None:
    with get_connection() as conn:
        row = conn.execute("SELECT id FROM prompt_templates WHERE is_default = 1 LIMIT 1").fetchone()
        if row:
            prompt_row = conn.execute(
                "SELECT user_prompt FROM prompt_templates WHERE id = ?",
                (row["id"],),
            ).fetchone()
            if prompt_row and DEFAULT_USER_PROMPT_MARKER not in str(prompt_row["user_prompt"] or ""):
                conn.execute(
                    """
                    UPDATE prompt_templates
                    SET name = ?, system_prompt = ?, user_prompt = ?, updated_at = ?
                    WHERE id = ?
                    """,
                    (DEFAULT_TEMPLATE_NAME, DEFAULT_SYSTEM_PROMPT, DEFAULT_USER_PROMPT, utc_now(), row["id"]),
                )
            return
        now = utc_now()
        conn.execute(
            """
            INSERT INTO prompt_templates (
                id, name, system_prompt, user_prompt, is_default, created_at, updated_at
            )
            VALUES (?, ?, ?, ?, 1, ?, ?)
            """,
            (uuid.uuid4().hex, DEFAULT_TEMPLATE_NAME, DEFAULT_SYSTEM_PROMPT, DEFAULT_USER_PROMPT, now, now),
        )


def reset_default_prompt_template() -> dict[str, Any]:
    now = utc_now()
    with get_connection() as conn:
        row = conn.execute("SELECT id FROM prompt_templates WHERE is_default = 1 LIMIT 1").fetchone()
        if row:
            template_id = row["id"]
            conn.execute(
                """
                UPDATE prompt_templates
                SET name = ?, system_prompt = ?, user_prompt = ?, updated_at = ?
                WHERE id = ?
                """,
                (DEFAULT_TEMPLATE_NAME, DEFAULT_SYSTEM_PROMPT, DEFAULT_USER_PROMPT, now, template_id),
            )
        else:
            template_id = uuid.uuid4().hex
            conn.execute(
                """
                INSERT INTO prompt_templates (
                    id, name, system_prompt, user_prompt, is_default, created_at, updated_at
                )
                VALUES (?, ?, ?, ?, 1, ?, ?)
                """,
                (template_id, DEFAULT_TEMPLATE_NAME, DEFAULT_SYSTEM_PROMPT, DEFAULT_USER_PROMPT, now, now),
            )
    return get_prompt_template(template_id)


def list_prompt_templates() -> list[dict[str, Any]]:
    ensure_default_prompt_template()
    with get_connection() as conn:
        rows = conn.execute(
            """
            SELECT * FROM prompt_templates
            ORDER BY is_default DESC, updated_at DESC
            """
        ).fetchall()
    return [_template_to_dict(row) for row in rows]


def get_prompt_template(template_id: str | None = None) -> dict[str, Any]:
    ensure_default_prompt_template()
    with get_connection() as conn:
        if template_id:
            row = conn.execute("SELECT * FROM prompt_templates WHERE id = ?", (template_id,)).fetchone()
        else:
            row = conn.execute("SELECT * FROM prompt_templates WHERE is_default = 1 LIMIT 1").fetchone()
    if not row:
        raise ValueError("提示词模板不存在")
    return _template_to_dict(row)


def save_prompt_template(
    *,
    template_id: str | None,
    name: str,
    system_prompt: str,
    user_prompt: str,
    forbidden_words_text: str = "",
) -> dict[str, Any]:
    if "{text}" not in user_prompt:
        raise ValueError("用户提示词必须包含 {text} 占位符")
    now = utc_now()
    with get_connection() as conn:
        if template_id:
            row = conn.execute("SELECT is_default FROM prompt_templates WHERE id = ?", (template_id,)).fetchone()
            if not row:
                raise ValueError("提示词模板不存在")
            conn.execute(
                """
                UPDATE prompt_templates
                SET name = ?, system_prompt = ?, user_prompt = ?, forbidden_words_text = ?, updated_at = ?
                WHERE id = ?
                """,
                (name, system_prompt, user_prompt, str(forbidden_words_text or ""), now, template_id),
            )
            saved_id = template_id
        else:
            saved_id = uuid.uuid4().hex
            conn.execute(
                """
                INSERT INTO prompt_templates (
                    id, name, system_prompt, user_prompt, forbidden_words_text, is_default, created_at, updated_at
                )
                VALUES (?, ?, ?, ?, ?, 0, ?, ?)
                """,
                (saved_id, name, system_prompt, user_prompt, str(forbidden_words_text or ""), now, now),
            )
    return get_prompt_template(saved_id)


def delete_prompt_template(template_id: str) -> None:
    with get_connection() as conn:
        row = conn.execute("SELECT is_default FROM prompt_templates WHERE id = ?", (template_id,)).fetchone()
        if not row:
            raise ValueError("提示词模板不存在")
        if row["is_default"]:
            raise ValueError("默认提示词模板不能删除，请使用恢复默认")
        conn.execute("DELETE FROM prompt_templates WHERE id = ?", (template_id,))


def _template_to_dict(row: Any) -> dict[str, Any]:
    return {
        "id": row["id"],
        "name": row["name"],
        "system_prompt": row["system_prompt"],
        "user_prompt": row["user_prompt"],
        "forbidden_words_text": row["forbidden_words_text"] if "forbidden_words_text" in row.keys() else "",
        "is_default": bool(row["is_default"]),
        "created_at": row["created_at"],
        "updated_at": row["updated_at"],
    }
