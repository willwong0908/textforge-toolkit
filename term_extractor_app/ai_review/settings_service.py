from __future__ import annotations

from typing import Any

from .database import dumps_json, get_connection, loads_json, utc_now


def get_setting(key: str, default: Any = None) -> Any:
    with get_connection() as conn:
        row = conn.execute("SELECT value_json FROM settings WHERE key = ?", (key,)).fetchone()
    if not row:
        return default
    return loads_json(row["value_json"], default)


def set_setting(key: str, value: Any) -> None:
    now = utc_now()
    with get_connection() as conn:
        conn.execute(
            """
            INSERT INTO settings (key, value_json, updated_at)
            VALUES (?, ?, ?)
            ON CONFLICT(key) DO UPDATE SET
                value_json = excluded.value_json,
                updated_at = excluded.updated_at
            """,
            (key, dumps_json(value), now),
        )


def get_ai_settings() -> dict[str, Any]:
    return get_setting(
        "ai.deepseek",
        {
            "api_key": "",
            "selected_model": "",
            "models": [],
            "max_concurrency": 8,
            "max_chars_per_request": 20000,
            "enable_thinking": True,
        },
    )


def save_ai_settings(settings: dict[str, Any]) -> None:
    current = get_ai_settings()
    current.update(settings)
    set_setting("ai.deepseek", current)


def public_ai_settings() -> dict[str, Any]:
    settings = get_ai_settings()
    api_key = settings.get("api_key", "")
    return {
        "provider": "deepseek",
        "has_api_key": bool(api_key),
        "api_key_masked": mask_api_key(api_key),
        "selected_model": settings.get("selected_model", ""),
        "models": settings.get("models", []),
        "max_concurrency": settings.get("max_concurrency", 8),
        "max_chars_per_request": settings.get("max_chars_per_request", 20000),
        "enable_thinking": True,
    }


def mask_api_key(api_key: str) -> str:
    if not api_key:
        return ""
    if len(api_key) <= 10:
        return "*" * len(api_key)
    return f"{api_key[:6]}...{api_key[-4:]}"
