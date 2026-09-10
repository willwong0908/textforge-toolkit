from __future__ import annotations

import asyncio
import hashlib
import json
import logging
import traceback
import threading
import uuid
from typing import Any

from .database import dumps_json, get_connection, init_db, loads_json, utc_now
from .directional_service import enabled_review_types, get_directional_template
from .forbidden_service import check_forbidden_words, get_forbidden_template, parse_forbidden_words
from .output_service import generate_review_excel
from .prompt_service import get_prompt_template
from .shared_provider import SharedProviderError, followup_chat, get_shared_ai_settings
from .session_store import emit_event
from .term_base_service import get_term_base, match_terms
from .memoq_service import lookup_terms, MemoQError
from ..models import LLMRequest, LLMResponse
from ..logging_utils import LOGGER_NAME, configure_file_logger
from ..providers import ProviderRegistry
from ..recovery_supervisor import (
    TerminalFailureContext,
    is_ai_correctable_response_failure,
    request_terminal_recovery_sync,
    retry_llm_with_corrective_prompt,
)
from ..scheduler import AdaptiveConcurrencyController, AsyncRequestScheduler
from ..telemetry import infer_model_tier, track_event

DEFAULT_DIRECTIONAL_SYSTEM_PROMPT = """你是专业翻译审校员。你的任务是按照用户指定的定向审校类型，检查译文相对于原文是否存在对应问题。

你必须遵守以下规则：
1. 只检查用户指定的审校类型，不要主动增加其他类型。
2. 每个条目都必须返回。
3. 每个审校类型都必须返回。
4. 如果某个类型没有问题，返回空字符串。
5. 如果某个类型有问题，用简洁中文说明问题。
6. 如果任何类型有问题，suggestion 必须只填写完整修改后的译文句子，不要解释，不要写修改理由，不要只写片段。
7. 如果没有任何问题，suggestion 必须返回空字符串。
8. 不要输出 Markdown。
9. 不要输出解释性文字。
10. 只返回严格 JSON。"""

DEFAULT_DIRECTIONAL_USER_PROMPT = """请按照 review_types 中指定的审校类型，审校 items 中的译文。

输入 JSON：
{text}

返回格式必须严格如下：
{
  "items": [
    {
      "id": "条目 ID",
      "suggestion": "如果任一审校类型有问题，填写完整修改后的译文句子；如果没有问题，返回空字符串",
      "checks": {
        "审校类型名称": "如果有问题，写问题说明；如果没有问题，返回空字符串"
      }
    }
  ]
}

要求：
1. 返回的 items 数量必须和输入 items 数量一致。
2. 返回的 id 必须和输入 id 一致。
3. checks 中的 key 必须和 review_types 的 key 完全一致。
4. 没有问题时，该类型的值必须是空字符串 ""。
5. 有问题时，checks 中只写一句简洁问题说明。
6. 如果任一 checks 不为空，suggestion 必须只填写完整修改后的译文句子，不要解释，不要写修改理由，不要只写片段。
7. 如果所有 checks 都为空，suggestion 必须返回空字符串。
8. 不要返回 review_types 中不存在的类型。
9. 不要省略任何已启用的审校类型。"""

TERM_REVIEW_PLACEHOLDER = "{term_review}"
TERM_REVIEW_INSTRUCTION = (
    "术语表参考规则：如果请求 JSON 包含 term_pairs，它是当前批次所有条目命中的去重术语集合。"
    "审校每个 item 时，只核对该 item 原文中实际出现的 source，不要把其他条目命中的术语套用到当前条目。"
    "source 是在原文中命中的源术语，targets 是推荐的目标术语，entry_note 是可选的语境备注；三者都只作为审校参考，不是强制替换规则。"
    "必须结合当前句子的实际含义、词性、语法、搭配、语气和上下文，判断该术语是否表达同一概念，以及推荐译法是否适合当前用法。"
    "如果现有译文在当前语境中含义准确、表达自然，不得仅因没有逐字采用 targets、使用了合理变形或采用了更合适的上下文表达而判定为术语问题。"
    "不得为了命中术语而改变原句语义、把名词机械改成动词（或反之），也不得生成不符合语法或搭配习惯的建议。"
    "只有在同一概念被明确错译、漏译、与既定专名明显不一致，或确实违背 entry_note 的适用语境时，才判定为术语问题并给出自然、完整的修改后译文；存在歧义时不要强行报错。"
)


class ReviewTaskError(Exception):
    pass


def create_review_task(
    batch_id: str,
    prompt_template_id: str | None = None,
    source_language: str = "",
    target_language: str = "",
    mode: str = "normal",
    directional_template_id: str | None = None,
    enable_ai_review: bool = True,
    enable_forbidden_check: bool = False,
    forbidden_template_id: str | None = None,
    session_id: str | None = None,
    term_base_id: str | None = None,
    memoq_term_base_ids: list[str] | None = None,
) -> str:
    if mode == "directional":
        raise ReviewTaskError("定向审校已移除，请使用提示词模板配置审校规则")
    items = _get_batch_items(batch_id)
    if not items:
        raise ReviewTaskError("当前没有可审校条目，请先读取文件")
    if not enable_ai_review and not enable_forbidden_check:
        raise ReviewTaskError("请至少启用 AI 审校或禁用词检查")

    settings = get_shared_ai_settings()
    api_key = settings.get("api_key", "")
    model = settings.get("selected_model", "")
    if enable_ai_review:
        if not api_key:
            raise ReviewTaskError("请先加载 DeepSeek API Key")
        if not model:
            raise ReviewTaskError("请先选择 DeepSeek 模型")

    forbidden_template = None
    forbidden_words: list[str] = []
    if enable_forbidden_check:
        forbidden_template = get_forbidden_template(forbidden_template_id)
        forbidden_words = parse_forbidden_words(forbidden_template["words_text"])

    term_base = get_term_base(term_base_id)
    if term_base_id and not term_base:
        raise ReviewTaskError("所选术语表不存在，请重新选择")

    task_id = uuid.uuid4().hex
    now = utc_now()
    if not enable_ai_review:
        config = {
            "mode": "forbidden_only",
            "enable_ai_review": False,
            "enable_forbidden_check": True,
            "forbidden_template_id": forbidden_template["id"] if forbidden_template else "",
            "forbidden_template_name": forbidden_template["name"] if forbidden_template else "",
            "forbidden_words": forbidden_words,
            "source_language": source_language.strip(),
            "target_language": target_language.strip(),
            "max_chars_per_request": int(settings.get("max_chars_per_request") or 3000),
            "max_concurrency": int(settings.get("max_concurrency") or 8),
            "enable_thinking": bool(settings.get("enable_thinking", False)),
        }
    elif mode == "directional":
        directional_template = get_directional_template(directional_template_id)
        review_types = enabled_review_types(directional_template)
        if not review_types:
            raise ReviewTaskError("请至少启用一个定向审校类型")
        config = {
            "mode": "directional",
            "enable_ai_review": True,
            "enable_forbidden_check": enable_forbidden_check,
            "forbidden_template_id": forbidden_template["id"] if forbidden_template else "",
            "forbidden_template_name": forbidden_template["name"] if forbidden_template else "",
            "forbidden_words": forbidden_words,
            "model": model,
            "directional_template_id": directional_template["id"],
            "directional_template_name": directional_template["name"],
            "review_types": review_types,
            "system_prompt": DEFAULT_DIRECTIONAL_SYSTEM_PROMPT,
            "user_prompt": DEFAULT_DIRECTIONAL_USER_PROMPT,
            "source_language": source_language.strip(),
            "target_language": target_language.strip(),
            "max_chars_per_request": int(settings.get("max_chars_per_request") or 3000),
            "max_concurrency": int(settings.get("max_concurrency") or 8),
            "enable_thinking": bool(settings.get("enable_thinking", False)),
        }
    else:
        prompt = get_prompt_template(prompt_template_id)
        prompt_forbidden_words = parse_forbidden_words(str(prompt.get("forbidden_words_text") or ""))
        if prompt_forbidden_words:
            enable_forbidden_check = True
            forbidden_words = prompt_forbidden_words
        config = {
            "mode": "normal",
            "enable_ai_review": True,
            "enable_forbidden_check": enable_forbidden_check,
            "forbidden_template_id": forbidden_template["id"] if forbidden_template else "",
            "forbidden_template_name": forbidden_template["name"] if forbidden_template else "",
            "forbidden_words": forbidden_words,
            "model": model,
            "prompt_template_id": prompt["id"],
            "prompt_template_name": prompt["name"],
            "system_prompt": prompt["system_prompt"],
            "user_prompt": prompt["user_prompt"],
            "source_language": source_language.strip(),
            "target_language": target_language.strip(),
            "max_chars_per_request": int(settings.get("max_chars_per_request") or 3000),
            "max_concurrency": int(settings.get("max_concurrency") or 8),
            "max_items_per_request": int(settings.get("max_items_per_request") or 80),
            "enable_thinking": bool(settings.get("enable_thinking", False)),
            "reasoning_effort": str(settings.get("reasoning_effort") or "low"),
            "session_id": str(session_id or ""),
            "term_base_id": str(term_base_id or ""),
            "term_base_name": str((term_base or {}).get("filename") or ""),
            "term_base_hash": str((term_base or {}).get("file_hash") or ""),
            "memoq_term_base_ids": [str(x) for x in (memoq_term_base_ids or []) if str(x).strip()],
            "memoq_term_pairs": [],
        }
    with get_connection() as conn:
        conn.execute(
            """
            INSERT INTO review_tasks (
                id, batch_id, session_id, target_language, status, total_count, config_json, created_at, updated_at
            )
            VALUES (?, ?, ?, ?, 'pending', ?, ?, ?, ?)
            """,
            (task_id, batch_id, session_id, target_language.strip(), len(items), dumps_json(config), now, now),
        )
    _add_log(task_id, "info", f"任务已创建，共 {len(items)} 条")
    _track_ai_review_start(config)
    thread = threading.Thread(target=_run_review_task, args=(task_id,), daemon=True)
    thread.start()
    return task_id


def get_review_task(task_id: str) -> dict[str, Any] | None:
    with get_connection() as conn:
        row = conn.execute("SELECT * FROM review_tasks WHERE id = ?", (task_id,)).fetchone()
    if not row:
        return None
    task = _task_to_dict(row)
    request_states = _get_review_request_states(task_id)
    task["request_states"] = request_states
    task["request_status_counts"] = {
        status: sum(1 for item in request_states if item["status"] == status)
        for status in ("queued", "submitted", "thinking", "output", "retrying", "completed", "failed")
    }
    return task


def recover_interrupted_review_tasks() -> int:
    """Close tasks whose daemon worker disappeared during a previous app process."""
    with get_connection() as conn:
        rows = conn.execute(
            "SELECT id FROM review_tasks WHERE status IN ('pending', 'running', 'recovering')"
        ).fetchall()
        task_ids = [str(row["id"]) for row in rows]
        if task_ids:
            placeholders = ",".join("?" for _ in task_ids)
            now = utc_now()
            conn.execute(
                f"UPDATE review_tasks SET status = 'failed', updated_at = ? WHERE id IN ({placeholders})",
                (now, *task_ids),
            )
            conn.execute(
                f"""
                UPDATE review_request_states
                SET status = 'failed', updated_at = ?
                WHERE task_id IN ({placeholders}) AND status NOT IN ('completed', 'failed')
                """,
                (now, *task_ids),
            )
    for task_id in task_ids:
        _add_log(task_id, "error", "应用上次退出时任务仍未结束，已标记为中断；请重新发起审校。")
    return len(task_ids)


def get_review_logs(task_id: str, after_id: int = 0) -> list[dict[str, Any]]:
    with get_connection() as conn:
        rows = conn.execute(
            """
            SELECT * FROM review_task_logs
            WHERE task_id = ? AND id > ?
            ORDER BY id ASC
            """,
            (task_id, after_id),
        ).fetchall()
    return [
        {
            "id": row["id"],
            "task_id": row["task_id"],
            "level": row["level"],
            "message": row["message"],
            "created_at": row["created_at"],
        }
        for row in rows
    ]


def get_review_results(task_id: str, limit: int | None = 20) -> list[dict[str, Any]]:
    limit_clause = "" if limit is None else "LIMIT ?"
    params: tuple[Any, ...] = (task_id,) if limit is None else (task_id, limit)
    with get_connection() as conn:
        rows = conn.execute(
            f"""
            SELECT r.*, i.source_text, i.target_text, i.source_file, i.sheet_name,
                   i.segment_id, i.row_number,
                   COALESCE(f.matched_words, '') AS matched_words
            FROM review_results r
            JOIN file_items i ON i.id = r.item_id
            LEFT JOIN forbidden_results f ON f.task_id = r.task_id AND f.item_id = r.item_id
            WHERE r.task_id = ?
            ORDER BY i.item_order ASC
            {limit_clause}
            """,
            params,
        ).fetchall()
    return [_result_to_dict(row) for row in rows]


def get_review_issue_results(task_id: str) -> list[dict[str, Any]]:
    task = get_review_task(task_id)
    if not task:
        return []
    config = task.get("config", {})
    return [
        item
        for item in get_review_results(task_id, limit=None)
        if _review_result_has_issue(item, config)
    ]


def get_review_followup_messages(task_id: str, result_id: str) -> dict[str, Any]:
    init_db()
    item = get_review_result_detail(task_id, result_id)
    if item is None:
        raise ReviewTaskError("问题条目不存在")
    with get_connection() as conn:
        rows = conn.execute(
            """
            SELECT id, role, content, created_at
            FROM review_followup_messages
            WHERE task_id = ? AND result_id = ?
            ORDER BY created_at ASC
            """,
            (task_id, result_id),
        ).fetchall()
    return {
        "item": item,
        "messages": [
            {
                "id": row["id"],
                "role": row["role"],
                "content": row["content"],
                "created_at": row["created_at"],
            }
            for row in rows
        ],
    }


def send_review_followup_message(task_id: str, result_id: str, message: str) -> dict[str, Any]:
    init_db()
    user_message = str(message or "").strip()
    if not user_message:
        raise ReviewTaskError("请输入追问内容")
    if len(user_message) > 4000:
        raise ReviewTaskError("追问内容过长，请缩短后再发送")

    task = get_review_task(task_id)
    if not task:
        raise ReviewTaskError("审校任务不存在")
    item = get_review_result_detail(task_id, result_id)
    if item is None:
        raise ReviewTaskError("问题条目不存在")

    config = task.get("config", {})
    history = get_review_followup_messages(task_id, result_id)["messages"]
    limited_history = history[-12:]
    messages = _build_followup_messages(item, config, limited_history, user_message)
    try:
        reply = followup_chat(
            task_id=task_id,
            messages=messages,
            enable_thinking=bool(config.get("enable_thinking", False)),
        ).strip()
    except SharedProviderError as exc:
        raise ReviewTaskError(str(exc)) from exc
    except Exception as exc:
        _add_log(task_id, "error", "追问请求失败：{0}\n{1}".format(exc, traceback.format_exc()))
        raise ReviewTaskError("追问请求失败：{0}".format(exc)) from exc
    if not reply:
        raise ReviewTaskError("模型没有返回内容")

    now = utc_now()
    user_id = uuid.uuid4().hex
    assistant_id = uuid.uuid4().hex
    with get_connection() as conn:
        conn.execute(
            """
            INSERT INTO review_followup_messages (id, task_id, result_id, role, content, created_at)
            VALUES (?, ?, ?, 'user', ?, ?)
            """,
            (user_id, task_id, result_id, user_message, now),
        )
        conn.execute(
            """
            INSERT INTO review_followup_messages (id, task_id, result_id, role, content, created_at)
            VALUES (?, ?, ?, 'assistant', ?, ?)
            """,
            (assistant_id, task_id, result_id, reply, utc_now()),
        )
    return get_review_followup_messages(task_id, result_id)


def get_review_result_detail(task_id: str, result_id: str) -> dict[str, Any] | None:
    with get_connection() as conn:
        row = conn.execute(
            """
            SELECT r.*, i.source_text, i.target_text, i.source_file, i.sheet_name,
                   i.segment_id, i.row_number,
                   COALESCE(f.matched_words, '') AS matched_words
            FROM review_results r
            JOIN file_items i ON i.id = r.item_id
            LEFT JOIN forbidden_results f ON f.task_id = r.task_id AND f.item_id = r.item_id
            WHERE r.task_id = ? AND r.id = ?
            """,
            (task_id, result_id),
        ).fetchone()
    return _result_to_dict(row) if row else None


def _review_result_has_issue(item: dict[str, Any], config: dict[str, Any]) -> bool:
    status = str(item.get("status") or "")
    if status == "failed":
        return True
    if str(config.get("mode") or "") == "directional":
        checks = item.get("checks", {})
        return any(str(value or "").strip() for value in checks.values()) or bool(str(item.get("suggestion") or "").strip())
    if str(config.get("mode") or "") == "forbidden_only":
        return bool(str(item.get("matched_words") or "").strip())
    return bool(item.get("has_issue")) or bool(str(item.get("matched_words") or "").strip())


def _build_followup_messages(
    item: dict[str, Any],
    config: dict[str, Any],
    history: list[dict[str, Any]],
    user_message: str,
) -> list[dict[str, str]]:
    context = _format_followup_context(item, config)
    messages = [
        {
            "role": "system",
            "content": (
                "你是翻译审校结果的追问助手。请只围绕给定审校条目回答，"
                "帮助用户理解问题、判断建议是否合理，或给出更好的改法。"
                "回答要简洁、明确，不要编造原文和译文之外的信息。"
            ),
        },
        {
            "role": "user",
            "content": "这是当前审校条目的固定上下文：\n" + context,
        },
    ]
    for message in history:
        role = str(message.get("role") or "").strip()
        if role not in {"user", "assistant"}:
            continue
        content = str(message.get("content") or "").strip()
        if content:
            messages.append({"role": role, "content": content[:6000]})
    messages.append({"role": "user", "content": user_message})
    return messages


def _format_followup_context(item: dict[str, Any], config: dict[str, Any]) -> str:
    lines = [
        f"审校模式：{config.get('mode', 'normal')}",
        f"原文：{item.get('source_text') or ''}",
        f"译文：{item.get('target_text') or ''}",
        f"修改建议：{item.get('suggestion') or ''}",
    ]
    issue_type = str(item.get("issue_type") or "").strip()
    issue = str(item.get("error_message") or item.get("issue") or "").strip()
    if issue_type:
        lines.append(f"问题类型：{issue_type}")
    if issue:
        lines.append(f"问题说明：{issue}")
    checks = item.get("checks", {})
    if isinstance(checks, dict):
        check_lines = [f"{key}：{value}" for key, value in checks.items() if str(value or "").strip()]
        if check_lines:
            lines.append("定向检查：\n" + "\n".join(check_lines))
    matched_words = str(item.get("matched_words") or "").strip()
    if matched_words:
        lines.append(f"禁用词命中：{matched_words}")
    location = []
    if item.get("sheet_name"):
        location.append(str(item.get("sheet_name")))
    if item.get("row_number"):
        location.append(f"第 {item.get('row_number')} 行")
    if location:
        lines.append("位置：" + " / ".join(location))
    return "\n".join(lines)


def _run_review_task(task_id: str) -> None:
    for run_index in range(2):
        try:
            _run_review_task_impl(task_id)
            return
        except Exception as exc:  # pragma: no cover - runtime safety net
            message = f"任务异常终止：{type(exc).__name__}: {exc}"
            _add_log(task_id, "error", message)
            task = get_review_task(task_id)
            if run_index == 0 and task and _attempt_review_task_terminal_recovery(task_id, task, exc):
                continue
            _mark_unfinished_review_requests_failed(task_id)
            if task:
                _update_task(task_id, status="failed")
                _track_ai_review_finish(task.get("config") or {}, success=False)
            return


def _attempt_review_task_terminal_recovery(task_id: str, task: dict[str, Any], exc: Exception) -> bool:
    request_states = list(task.get("request_states") or [])
    completed_units = sum(
        int(item.get("item_count") or 0) for item in request_states if item.get("status") == "completed"
    )
    total_units = sum(int(item.get("item_count") or 0) for item in request_states)
    recent_logs = _get_recent_review_log_messages(task_id, limit=6)
    _update_task(task_id, status="recovering")
    _add_log(task_id, "warning", "任务即将进入失败状态，正在启动最后一次 Recovery Agent 兜底。")
    decision = request_terminal_recovery_sync(
        TerminalFailureContext(
            workflow="ai_review",
            task_id=task_id,
            phase="task_runtime",
            error_type=type(exc).__name__,
            message=str(exc),
            has_checkpoint=True,
            completed_units=completed_units,
            remaining_units=max(0, total_units - completed_units),
            metadata={
                "target_language": str(task.get("target_language") or ""),
                "recent_events": [str(item or "")[-500:] for item in recent_logs],
            },
        ),
        allowed_actions=("retry_task", "abort"),
    )
    if decision.error:
        _add_log(task_id, "error", f"Recovery Agent 未能形成可执行方案：{decision.error}")
        return False
    _add_log(task_id, "warning", f"Recovery Agent 决策：{decision.action}；{decision.reason or '未提供说明'}")
    if decision.action != "retry_task":
        return False
    _add_log(task_id, "warning", "Recovery Agent 已批准从未完成条目继续；该动作最多执行一次。")
    return True


def _run_review_task_impl(task_id: str) -> None:
    task = get_review_task(task_id)
    if not task:
        return
    config = task["config"]
    items = _get_batch_items(task["batch_id"])
    enable_ai_review = bool(config.get("enable_ai_review", True))
    max_chars = max(1, int(config["max_chars_per_request"]))
    settings = get_shared_ai_settings()
    api_key = settings.get("api_key", "")

    _update_task(task_id, status="running")
    requested_count = 0
    failed_count = 0
    completed_count = 0

    if enable_ai_review:
        model = config["model"]
        memoq_ids = [str(x) for x in (config.get("memoq_term_base_ids") or []) if str(x).strip()]
        if memoq_ids:
            try:
                pairs = lookup_terms(
                    memoq_ids,
                    str(config.get("source_language") or "auto"),
                    str(config.get("target_language") or "auto"),
                    [str(item.get("source_text") or "") for item in items if str(item.get("source_text") or "").strip()],
                )
                unique: list[dict[str, Any]] = []
                seen: set[str] = set()
                for pair in pairs:
                    key = _term_pair_key(pair)
                    if key not in seen:
                        seen.add(key)
                        unique.append(pair)
                config["memoq_term_pairs"] = unique
                _add_log(task_id, "info", f"memoQ 术语库已匹配 {len(unique)} 个术语对，将按批次提供给模型。")
            except MemoQError as exc:
                _add_log(task_id, "warning", f"memoQ 术语提取失败，继续执行但不应用远程术语库：{exc}")
        term_base_id = str(config.get("term_base_id") or "")
        term_match_count = 0
        term_language_warning: dict[str, Any] | None = None
        if term_base_id:
            for item in items:
                term_pairs, match_meta = match_terms(
                    term_base_id,
                    str(config.get("source_language") or ""),
                    str(config.get("target_language") or ""),
                    str(item.get("source_text") or ""),
                )
                if term_pairs:
                    item["term_pairs"] = term_pairs
                    term_match_count += len(term_pairs)
                if match_meta.get("status") == "language_unmapped":
                    term_language_warning = match_meta
            if term_language_warning:
                _add_log(
                    task_id,
                    "warning",
                    "术语表未找到当前源语种或目标语种列，本任务将继续执行但不应用术语表。",
                )
            else:
                _add_log(
                    task_id,
                    "info",
                    f"术语表 {config.get('term_base_name') or term_base_id} 已匹配 {term_match_count} 个术语条目。",
                )
        prompt_signature = _prompt_signature(
            config["system_prompt"],
            config["user_prompt"],
            config.get("source_language", ""),
            config.get("target_language", ""),
            str(config.get("term_base_hash") or ""),
        )
        directional_signature = _directional_signature(config)
        enable_thinking = bool(config.get("enable_thinking", False))
        existing_completed_ids = _get_completed_review_item_ids(task_id)
        request_items = []
        cached_pairs: list[tuple[dict[str, Any], dict[str, Any]]] = []
        for item in items:
            if item["id"] in existing_completed_ids:
                continue
            cache_key = _cache_key(
                item["source_text"],
                item["target_text"],
                _item_info(item),
                _item_term_pairs(item),
                model,
                prompt_signature,
                directional_signature,
                enable_thinking,
            )
            item_with_cache = {**item, "cache_key": cache_key}
            cached = _get_cached_result(cache_key)
            if cached is None:
                request_items.append(item_with_cache)
            else:
                cached_pairs.append((item_with_cache, cached))

        if cached_pairs:
            _save_review_results_bulk(task_id, cached_pairs, status="cached")
        cached_count = len(cached_pairs)
        initial_completed = len(existing_completed_ids) + cached_count
        _update_task(task_id, cached_count=cached_count, completed_count=initial_completed)

        packages = _build_packages(
            request_items,
            max_chars,
            max_items=max(1, int(config.get("max_items_per_request") or 80)),
        )
        _replace_review_request_states(task_id, packages)
        requested_count, completed_count, failed_count = asyncio.run(
            _run_review_packages(
                task_id=task_id,
                api_key=api_key,
                model=model,
                config=config,
                packages=packages,
                initial_completed=initial_completed,
            )
        )
    else:
        _add_log(task_id, "info", "未启用 AI 审校，跳过 AI 请求")
        for item in items:
            _save_review_result(task_id, item["id"], "", "skipped_ai", {})
        completed_count = len(items)
        _update_task(task_id, completed_count=completed_count)

    if config.get("enable_forbidden_check"):
        _run_forbidden_check(task_id, items, config.get("forbidden_words", []))

    try:
        output_path = generate_review_excel(task_id)
        _update_task(task_id, output_path=str(output_path))
        _add_log(task_id, "info", f"结果已自动保存：{output_path}")
    except Exception as exc:
        _add_log(task_id, "error", f"结果 Excel 输出失败：{exc}")
        _update_task(task_id, status="completed_with_errors")
        _track_ai_review_finish(config, success=False)
        return

    final_status = "completed" if failed_count == 0 else "completed_with_errors"
    _update_task(task_id, status=final_status)
    _track_ai_review_finish(config, success=True)
    _add_log(task_id, "info", f"任务完成：成功 {completed_count - failed_count} 条，失败 {failed_count} 条")


def _run_forbidden_check(task_id: str, items: list[dict[str, Any]], words: list[str]) -> None:
    _add_log(task_id, "info", "开始禁用词检查")
    now = utc_now()
    hit_count = 0
    with get_connection() as conn:
        conn.execute("DELETE FROM forbidden_results WHERE task_id = ?", (task_id,))
        for item in items:
            matched = check_forbidden_words(item["target_text"], words)
            if matched:
                hit_count += 1
            conn.execute(
                """
                INSERT INTO forbidden_results (id, task_id, item_id, matched_words, created_at, updated_at)
                VALUES (?, ?, ?, ?, ?, ?)
                """,
                (uuid.uuid4().hex, task_id, item["id"], "; ".join(matched), now, now),
            )
    _add_log(task_id, "info", f"禁用词检查完成，命中 {hit_count} 条")


async def _run_review_packages(
    task_id: str,
    api_key: str,
    model: str,
    config: dict[str, Any],
    packages: list[list[dict[str, Any]]],
    initial_completed: int = 0,
) -> tuple[int, int, int]:
    """Run independent review packages concurrently with adaptive backpressure."""
    if not packages:
        return 0, max(0, int(initial_completed)), 0

    provider_name, provider_settings = _load_review_provider(api_key, model)
    adapter = ProviderRegistry.create_adapter(provider_name, provider_settings)

    def on_stream_phase(package_index: int, phase: str) -> None:
        _set_review_request_status(task_id, package_index, phase)
        phase_label = {"thinking": "思考", "output": "输出"}.get(str(phase))
        if phase_label:
            _add_log(task_id, "info", f"第 {package_index}/{len(packages)} 包进入{phase_label}阶段。")

    requests: list[LLMRequest] = []
    for index, package in enumerate(packages, start=1):
        payload = _build_request_payload(package, config)
        package_term_pairs = list(payload.get("term_pairs") or [])
        user_prompt = _build_user_prompt(
            config,
            json.dumps(payload, ensure_ascii=False),
            any(_item_info(item) for item in package),
            bool(package_term_pairs),
        )
        metadata = {
            "enable_thinking": bool(config.get("enable_thinking", False)),
            "reasoning_effort": str(config.get("reasoning_effort") or "low"),
            "package_index": index,
            "package": package,
        }
        if "deepseek" in f"{provider_name} {provider_settings.base_url}".lower():
            metadata["stream_phase_callback"] = lambda phase, package_index=index: on_stream_phase(
                package_index, phase
            )
        requests.append(
            LLMRequest(
                task_id=f"ai_review_package_{index}",
                task_type="candidate_review_batch",
                prompt=user_prompt,
                messages=[
                    {"role": "system", "content": config["system_prompt"]},
                    {"role": "user", "content": user_prompt},
                ],
                metadata=metadata,
            )
        )

    requested_count = 0
    completed_count = max(0, int(initial_completed))
    failed_count = 0
    counters_lock = threading.Lock()
    terminal_recovery_queue: list[tuple[LLMRequest, LLMResponse, Any]] = []

    def validate_response(request: LLMRequest, response: LLMResponse) -> LLMResponse:
        if not response.success:
            return response
        try:
            parsed = _parse_json_object(response.content)
            package = list(request.metadata.get("package") or [])
            _validate_response_items(parsed, package, config)
        except (ReviewTaskError, json.JSONDecodeError) as exc:
            return LLMResponse(
                task_id=response.task_id,
                task_type=response.task_type,
                content=response.content,
                provider=response.provider,
                model=response.model,
                latency_ms=response.latency_ms,
                attempts=response.attempts,
                success=False,
                error=str(exc),
                error_type="response_validation",
                retryable=True,
                response_metadata=response.response_metadata,
            )
        return response

    def on_result(request: LLMRequest, response: LLMResponse, snapshot) -> None:
        nonlocal completed_count, failed_count
        package = list(request.metadata.get("package") or [])
        package_index = int(request.metadata.get("package_index") or 0)
        can_use_terminal_ai = (
            is_ai_correctable_response_failure(response.error_type)
            and len(terminal_recovery_queue) < 8
        )
        if (
            not response.success
            and not bool(request.metadata.get("terminal_recovery_attempted"))
            and can_use_terminal_ai
        ):
            request.metadata["terminal_recovery_attempted"] = True
            terminal_recovery_queue.append((request, response, snapshot))
            _set_review_request_status(task_id, package_index, "retrying", max(1, response.attempts) + 1)
            _add_log(
                task_id,
                "warning",
                f"第 {package_index}/{len(packages)} 包常规重试已耗尽，进入最后一次 AI 兜底纠错。",
            )
            return
        _set_review_request_status(task_id, package_index, "completed" if response.success else "failed", response.attempts)
        _add_log(task_id, "info", f"第 {package_index}/{len(packages)} 包完成，包含 {len(package)} 条，并发 {snapshot.current_concurrency}")
        if response.success:
            parsed_items = _parse_json_object(response.content).get("items", [])
            result_by_id = {str(item.get("id")): item for item in parsed_items if isinstance(item, dict)}
            saved_results = []
            for item in package:
                if config.get("mode") == "directional":
                    result = _normalize_directional_result(
                        result_by_id.get(item["id"]), item["id"], config.get("review_types", [])
                    )
                else:
                    result = _normalize_result(result_by_id.get(item["id"]), item["id"])
                saved_results.append((item, result))
            _save_review_results_bulk(task_id, saved_results)
            _cache_review_results_bulk(
                saved_results,
                model=model,
                prompt_signature=_prompt_signature(
                    config["system_prompt"],
                    config["user_prompt"],
                    str(config.get("source_language") or ""),
                    str(config.get("target_language") or ""),
                    str(config.get("term_base_hash") or ""),
                ),
                directional_signature=_directional_signature(config),
            )
            _add_log(task_id, "info", f"第 {package_index} 包校验通过：{len(parsed_items)} 条")
        else:
            message = str(response.error or "请求失败")
            _add_log(task_id, "error", f"第 {package_index} 包失败，已重试 {max(0, response.attempts - 1)} 次：{message}")
            _save_review_errors_bulk(task_id, package, message)

        with counters_lock:
            completed_count += len(package)
            if not response.success:
                failed_count += len(package)
            _update_task(
                task_id,
                requested_count=requested_count,
                completed_count=completed_count,
                failed_count=failed_count,
            )
        session_id = str(config.get("session_id") or "")
        if session_id:
            emit_event(
                session_id,
                "review.package_completed",
                {
                    "task_id": task_id,
                    "package_index": package_index,
                    "package_total": len(packages),
                    "item_count": len(package),
                    "success": bool(response.success),
                    "completed_count": completed_count,
                    "failed_count": failed_count,
                },
            )

    def on_request_started(request: LLMRequest) -> None:
        nonlocal requested_count
        package = list(request.metadata.get("package") or [])
        package_index = int(request.metadata.get("package_index") or 0)
        with counters_lock:
            requested_count += len(package)
            _update_task(task_id, requested_count=requested_count)
        _set_review_request_status(task_id, package_index, "submitted", 1)

    def on_retry(request: LLMRequest, response: LLMResponse, next_attempt: int, backoff: float) -> None:
        package_index = int(request.metadata.get("package_index") or 0)
        _set_review_request_status(task_id, package_index, "retrying", next_attempt)
        reason = str(response.error_type or "request_failed")
        detail = str(response.error or "").strip()
        _add_log(
            task_id,
            "warning",
            f"第 {package_index}/{len(packages)} 包触发重试；原因 {reason}；"
            f"{('详情 ' + detail + '；') if detail else ''}"
            f"{backoff:.1f} 秒后进行第 {next_attempt} 次请求。",
        )

    original_send_prompt = adapter.send_prompt

    async def logged_send_prompt(request: LLMRequest, attempt: int = 1) -> LLMResponse:
        package_index = int(request.metadata.get("package_index") or 0)
        if attempt == 1:
            _add_log(task_id, "info", f"已发送第 {package_index}/{len(packages)} 包，正在等待模型响应。")
        else:
            _set_review_request_status(task_id, package_index, "submitted", attempt)
            _add_log(task_id, "warning", f"第 {package_index}/{len(packages)} 包响应无效或请求失败，正在进行第 {attempt} 次请求。")
        _add_log(
            task_id,
            "debug",
            _format_ai_request_log(
                package_index=package_index,
                package_total=len(packages),
                attempt=attempt,
                model=model,
                request=request,
            ),
        )
        try:
            response = await original_send_prompt(request, attempt=attempt)
        except Exception as exc:
            _add_log(
                task_id,
                "debug",
                f"AI 请求异常｜包 {package_index}/{len(packages)}｜第 {attempt} 次\n"
                f"{type(exc).__name__}: {exc}",
            )
            raise
        _add_log(
            task_id,
            "debug",
            _format_ai_response_log(
                package_index=package_index,
                package_total=len(packages),
                attempt=attempt,
                response=response,
            ),
        )
        return response

    adapter.send_prompt = logged_send_prompt

    configured_max = max(1, int(config.get("max_concurrency") or provider_settings.max_concurrency or 1))
    provider_max = max(1, int(provider_settings.max_concurrency or configured_max))
    controller = AdaptiveConcurrencyController(
        mode="自动",
        user_max=configured_max,
        provider_max=provider_max,
        local_max=32,
        start_concurrency=min(4, configured_max, provider_max),
    )
    scheduler = AsyncRequestScheduler(
        adapter=adapter,
        controller=controller,
        max_retries=2,
        stop_requested=lambda: False,
        response_validator=validate_response,
        attempt_timeout_seconds=max(10, int(provider_settings.timeout_seconds or 90)),
        on_retry=on_retry,
    )
    try:
        _add_log(
            task_id,
            "info",
            f"开始 AI 审校，共 {len(packages)} 包；模型 {model}；自动并发从 "
            f"{controller.current_concurrency} 路启动，最高 {controller.effective_max} 路；"
            f"单次请求硬超时 {max(10, int(provider_settings.timeout_seconds or 90))} 秒；最多请求 3 次。",
        )
        await scheduler.run(requests, on_result=on_result, on_request_started=on_request_started)
        for request, failed_response, snapshot in terminal_recovery_queue:
            package_index = int(request.metadata.get("package_index") or 0)
            recovered_response = await retry_llm_with_corrective_prompt(
                adapter=adapter,
                request=request,
                failed_response=failed_response,
                response_validator=validate_response,
                timeout_seconds=min(60, max(10, int(provider_settings.timeout_seconds or 90))),
            )
            if recovered_response.success:
                _add_log(task_id, "warning", f"第 {package_index}/{len(packages)} 包已通过最终 AI 兜底纠错。")
            else:
                _add_log(
                    task_id,
                    "error",
                    f"第 {package_index}/{len(packages)} 包最终 AI 兜底仍失败："
                    f"{recovered_response.error or '响应无效'}",
                )
            on_result(request, recovered_response, snapshot)
    finally:
        await adapter.close()
    return requested_count, completed_count, failed_count


def _load_review_provider(api_key: str, model: str):
    from .shared_provider import _load_provider_settings

    provider_name, provider_settings = _load_provider_settings(api_key_override=api_key)
    provider_settings.model = model
    return provider_name, provider_settings


def _build_packages(
    items: list[dict[str, Any]], max_chars: int, *, max_items: int = 80
) -> list[list[dict[str, Any]]]:
    packages: list[list[dict[str, Any]]] = []
    current: list[dict[str, Any]] = []
    current_chars = 0
    current_term_keys: set[str] = set()

    for item in items:
        item_chars = _item_character_cost(item)
        item_pairs = _item_term_pairs(item)
        new_term_chars, new_term_keys = _new_term_pair_cost(item_pairs, current_term_keys)
        item_chars += new_term_chars
        if item_chars > max_chars:
            if current:
                packages.append(current)
                current = []
                current_chars = 0
                current_term_keys = set()
            packages.append([item])
            continue
        if current and (current_chars + item_chars > max_chars or len(current) >= max_items):
            packages.append(current)
            current = []
            current_chars = 0
            current_term_keys = set()
            new_term_chars, new_term_keys = _new_term_pair_cost(item_pairs, current_term_keys)
            item_chars = _item_character_cost(item) + new_term_chars
        current.append(item)
        current_chars += item_chars
        current_term_keys.update(new_term_keys)

    if current:
        packages.append(current)
    return packages


def _item_character_cost(item: dict[str, Any]) -> int:
    """Count item-local review content; package-level terms are counted separately."""
    total = len(str(item.get("source_text") or "")) + len(str(item.get("target_text") or ""))
    total += sum(
        len(str(info.get("category") or "")) + len(str(info.get("value") or ""))
        for info in _item_info(item)
    )
    return total


def _new_term_pair_cost(
    pairs: list[dict[str, Any]], existing_keys: set[str]
) -> tuple[int, set[str]]:
    total = 0
    new_keys: set[str] = set()
    for pair in pairs:
        key = _term_pair_key(pair)
        if key in existing_keys or key in new_keys:
            continue
        new_keys.add(key)
        total += _term_pair_character_cost(pair)
    return total, new_keys


def _term_pair_character_cost(pair: dict[str, Any]) -> int:
    total = len(str(pair.get("source") or "")) + len(str(pair.get("entry_note") or ""))
    total += sum(len(str(value or "")) for value in pair.get("targets") or [])
    return total


def _parse_json_object(content: str) -> dict[str, Any]:
    text = content.strip()
    if text.startswith("```"):
        text = text.strip("`")
        if text.startswith("json"):
            text = text[4:].strip()
    start = text.find("{")
    end = text.rfind("}")
    if start >= 0 and end >= start:
        text = text[start : end + 1]
    return json.loads(text)


def _validate_response_items(
    parsed: dict[str, Any], package: list[dict[str, Any]], config: dict[str, Any]
) -> None:
    raw_items = parsed.get("items")
    if not isinstance(raw_items, list):
        raise ReviewTaskError("AI 返回 JSON 中缺少 items 数组")
    expected_ids = [str(item["id"]) for item in package]
    returned_ids: list[str] = []
    for index, item in enumerate(raw_items, 1):
        if not isinstance(item, dict):
            raise ReviewTaskError(f"AI 返回第 {index} 个条目不是对象")
        item_id = str(item.get("id") or "")
        if not item_id:
            raise ReviewTaskError(f"AI 返回第 {index} 个条目缺少 id")
        returned_ids.append(item_id)
        if config.get("mode") == "directional":
            if not isinstance(item.get("checks"), dict):
                raise ReviewTaskError(f"AI 返回条目 {item_id} 缺少 checks 对象")
            if not isinstance(item.get("suggestion", ""), str):
                raise ReviewTaskError(f"AI 返回条目 {item_id} 的 suggestion 不是字符串")
        else:
            required = {"has_issue", "issue_type", "issue", "suggestion"}
            missing = required.difference(item)
            if missing:
                raise ReviewTaskError(f"AI 返回条目 {item_id} 缺少字段：{', '.join(sorted(missing))}")
            if not isinstance(item["has_issue"], bool):
                raise ReviewTaskError(f"AI 返回条目 {item_id} 的 has_issue 必须是布尔值")
            for key in ("issue_type", "issue", "suggestion"):
                if not isinstance(item[key], str):
                    raise ReviewTaskError(f"AI 返回条目 {item_id} 的 {key} 必须是字符串")
    if len(returned_ids) != len(expected_ids):
        raise ReviewTaskError(f"AI 返回 {len(returned_ids)} 条，预期 {len(expected_ids)} 条")
    if len(set(returned_ids)) != len(returned_ids):
        raise ReviewTaskError("AI 返回存在重复 id")
    if set(returned_ids) != set(expected_ids):
        missing = sorted(set(expected_ids) - set(returned_ids))
        extra = sorted(set(returned_ids) - set(expected_ids))
        raise ReviewTaskError(f"AI 返回 id 不匹配；缺少 {missing[:5]}，多出 {extra[:5]}")


def _save_review_results_bulk(
    task_id: str,
    pairs: list[tuple[dict[str, Any], dict[str, Any]]],
    *,
    status: str = "completed",
) -> None:
    if not pairs:
        return
    now = utc_now()
    with get_connection() as conn:
        conn.executemany(
            "DELETE FROM review_results WHERE task_id = ? AND item_id = ? AND status = 'failed'",
            [(task_id, item["id"]) for item, _result in pairs],
        )
        conn.executemany(
            """
            INSERT INTO review_results (
                id, task_id, item_id, cache_key, status, has_issue, issue_type,
                issue, suggestion, directional_checks_json, error_message,
                raw_result_json, created_at, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, '', ?, ?, ?)
            """,
            [
                (
                    uuid.uuid4().hex,
                    task_id,
                    item["id"],
                    item.get("cache_key", ""),
                    status,
                    1 if result.get("has_issue") else 0,
                    result.get("issue_type", ""),
                    result.get("issue", ""),
                    result.get("suggestion", ""),
                    dumps_json(result.get("checks", {})),
                    dumps_json(result),
                    now,
                    now,
                )
                for item, result in pairs
            ],
        )


def _save_review_errors_bulk(task_id: str, package: list[dict[str, Any]], message: str) -> None:
    if not package:
        return
    now = utc_now()
    with get_connection() as conn:
        conn.executemany(
            "DELETE FROM review_results WHERE task_id = ? AND item_id = ? AND status = 'failed'",
            [(task_id, item["id"]) for item in package],
        )
        conn.executemany(
            """
            INSERT INTO review_results (
                id, task_id, item_id, cache_key, status, error_message, created_at, updated_at
            ) VALUES (?, ?, ?, ?, 'failed', ?, ?, ?)
            """,
            [
                (uuid.uuid4().hex, task_id, item["id"], item.get("cache_key", ""), message, now, now)
                for item in package
            ],
        )


def _cache_review_results_bulk(
    pairs: list[tuple[dict[str, Any], dict[str, Any]]],
    *,
    model: str,
    prompt_signature: str,
    directional_signature: str,
) -> None:
    if not pairs:
        return
    now = utc_now()
    with get_connection() as conn:
        conn.executemany(
            """
            INSERT INTO ai_result_cache (
                cache_key, source_text, target_text, model, prompt_signature,
                directional_signature, result_json, created_at, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(cache_key) DO UPDATE SET
                result_json = excluded.result_json,
                updated_at = excluded.updated_at
            """,
            [
                (
                    item["cache_key"], item.get("source_text", ""), item.get("target_text", ""),
                    model, prompt_signature, directional_signature, dumps_json(result), now, now,
                )
                for item, result in pairs
            ],
        )


def _normalize_result(result: Any, item_id: str) -> dict[str, Any]:
    if not isinstance(result, dict):
        return {
            "id": item_id,
            "has_issue": False,
            "issue_type": "",
            "issue": "",
            "suggestion": "",
        }
    return {
        "id": str(result.get("id") or item_id),
        "has_issue": bool(result.get("has_issue")),
        "issue_type": str(result.get("issue_type") or ""),
        "issue": str(result.get("issue") or ""),
        "suggestion": str(result.get("suggestion") or ""),
    }


def _normalize_directional_result(result: Any, item_id: str, review_types: list[dict[str, str]]) -> dict[str, Any]:
    checks = {}
    raw_checks = result.get("checks", {}) if isinstance(result, dict) else {}
    if not isinstance(raw_checks, dict):
        raw_checks = {}
    for review_type in review_types:
        key = review_type["key"]
        checks[key] = str(raw_checks.get(key) or "")
    suggestion = str(result.get("suggestion") or "") if isinstance(result, dict) else ""
    return {"id": item_id, "suggestion": suggestion, "checks": checks}


def _get_batch_items(batch_id: str) -> list[dict[str, Any]]:
    with get_connection() as conn:
        rows = conn.execute(
            """
            SELECT * FROM file_items
            WHERE batch_id = ?
            ORDER BY item_order ASC
            """,
            (batch_id,),
        ).fetchall()
    return [dict(row) for row in rows]


def _item_info(item: dict[str, Any]) -> list[dict[str, str]]:
    data = loads_json(item.get("info_json"), []) if item.get("info_json") else []
    if not isinstance(data, list):
        return []
    normalized = []
    for info in data:
        if not isinstance(info, dict):
            continue
        value = str(info.get("value") or "").strip()
        if value:
            normalized.append(
                {
                    "category": str(info.get("category") or "").strip(),
                    "value": value,
                }
            )
    return normalized


def _build_request_payload(
    package: list[dict[str, Any]], config: dict[str, Any]
) -> dict[str, Any]:
    payload: dict[str, Any] = {"items": [_payload_item(item) for item in package]}
    term_pairs = _package_term_pairs(package)
    for pair in config.get("memoq_term_pairs") or []:
        if isinstance(pair, dict) and _term_pair_key(pair) not in {_term_pair_key(x) for x in term_pairs}:
            term_pairs.append(pair)
    if term_pairs:
        payload["term_pairs"] = term_pairs
    if config.get("mode") == "directional":
        payload["review_types"] = config.get("review_types", [])
    source_language = str(config.get("source_language") or "").strip()
    target_language = str(config.get("target_language") or "").strip()
    if source_language or target_language:
        payload["language"] = {"source": source_language, "target": target_language}
    return payload


def _payload_item(item: dict[str, Any]) -> dict[str, Any]:
    payload = {"id": item["id"], "source": item["source_text"], "target": item["target_text"]}
    info = _item_info(item)
    if info:
        payload["info"] = info
    return payload


def _item_term_pairs(item: dict[str, Any]) -> list[dict[str, Any]]:
    pairs = item.get("term_pairs") or []
    return [pair for pair in pairs if isinstance(pair, dict)] if isinstance(pairs, list) else []


def _package_term_pairs(package: list[dict[str, Any]]) -> list[dict[str, Any]]:
    pairs: list[dict[str, Any]] = []
    seen: set[str] = set()
    for item in package:
        for pair in _item_term_pairs(item):
            key = _term_pair_key(pair)
            if key in seen:
                continue
            seen.add(key)
            pairs.append(pair)
    return pairs


def _term_pair_key(pair: dict[str, Any]) -> str:
    return json.dumps(pair, ensure_ascii=False, sort_keys=True, separators=(",", ":"), default=str)


def _build_user_prompt(
    config: dict[str, Any], text: str, has_info: bool, has_term_pairs: bool = False
) -> str:
    prefixes = []
    source_language = str(config.get("source_language") or "").strip()
    target_language = str(config.get("target_language") or "").strip()
    if source_language.lower() in {"none", "auto"}:
        source_language = ""
    if target_language.lower() == "auto":
        target_language = ""
    if not source_language and target_language:
        prefixes.append(f"以下是{target_language}译文；本次没有原文，请仅检查译文本身的语言、格式、术语和表达问题。")
    elif source_language or target_language:
        prefixes.append(f"以下是 {source_language or '未指定语种'} 原文和 {target_language or '未指定语种'} 译文。")
    if has_info:
        prefixes.append(
            "如果 item 包含 info 字段，请把 info 作为参考信息。"
            "info 中 category 是信息类别，value 是信息内容；没有 info 的条目不要假设存在参考信息。"
        )
    term_instruction = TERM_REVIEW_INSTRUCTION if has_term_pairs else ""
    template = str(config["user_prompt"])
    had_placeholder = TERM_REVIEW_PLACEHOLDER in template
    template = template.replace(TERM_REVIEW_PLACEHOLDER, term_instruction)
    prompt = template.replace("{text}", text)
    if term_instruction and not had_placeholder:
        prefixes.append(term_instruction)
    return "\n".join([*prefixes, prompt]) if prefixes else prompt


def _format_ai_request_log(
    *,
    package_index: int,
    package_total: int,
    attempt: int,
    model: str,
    request: LLMRequest,
) -> str:
    messages = list(request.messages or [])
    if not messages:
        messages = [{"role": "user", "content": request.prompt}]
    sections = [
        f"AI 请求｜包 {package_index}/{package_total}｜第 {attempt} 次｜模型 {model}"
    ]
    for message in messages:
        role = str(message.get("role") or "unknown")
        content = str(message.get("content") or "")
        sections.extend((f"[{role}]", content))
    return "\n".join(sections)


def _format_ai_response_log(
    *,
    package_index: int,
    package_total: int,
    attempt: int,
    response: LLMResponse,
) -> str:
    sections = [
        f"AI 返回｜包 {package_index}/{package_total}｜第 {attempt} 次｜"
        f"成功 {str(bool(response.success)).lower()}｜耗时 {int(response.latency_ms or 0)} ms",
        "[content]",
        str(response.content or ""),
    ]
    if response.error_type or response.error:
        sections.extend(
            (
                "[error]",
                f"type={response.error_type or 'unknown'}\n{response.error or ''}",
            )
        )
    if response.response_metadata:
        sections.extend(
            (
                "[metadata]",
                json.dumps(response.response_metadata, ensure_ascii=False, indent=2, default=str),
            )
        )
    return "\n".join(sections)


def _cache_key(
    source: str,
    target: str,
    info: list[dict[str, str]],
    term_pairs: list[dict[str, Any]],
    model: str,
    prompt_signature: str,
    directional_signature: str,
    enable_thinking: bool,
) -> str:
    raw = dumps_json(
        {
            "source": source,
            "target": target,
            "info": info,
            "term_pairs": term_pairs,
            "model": model,
            "prompt_signature": prompt_signature,
            "directional_signature": directional_signature,
            "enable_thinking": enable_thinking,
        }
    )
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


def _prompt_signature(
    system_prompt: str,
    user_prompt: str,
    source_language: str,
    target_language: str,
    term_base_hash: str = "",
) -> str:
    raw = dumps_json(
        {
            "system_prompt": system_prompt,
            "user_prompt": user_prompt,
            "source_language": source_language,
            "target_language": target_language,
            "term_base_hash": term_base_hash,
        }
    )
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


def _directional_signature(config: dict[str, Any]) -> str:
    raw = dumps_json(
        {
            "mode": config.get("mode", "normal"),
            "review_types": config.get("review_types", []),
        }
    )
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


def _get_cached_result(cache_key: str) -> dict[str, Any] | None:
    with get_connection() as conn:
        row = conn.execute(
            "SELECT result_json FROM ai_result_cache WHERE cache_key = ?",
            (cache_key,),
        ).fetchone()
    if not row:
        return None
    return loads_json(row["result_json"], None)


def _set_cached_result(
    cache_key: str,
    source_text: str,
    target_text: str,
    model: str,
    prompt_signature: str,
    directional_signature: str,
    result: dict[str, Any],
) -> None:
    now = utc_now()
    with get_connection() as conn:
        conn.execute(
            """
            INSERT INTO ai_result_cache (
                cache_key, source_text, target_text, model, prompt_signature,
                directional_signature, result_json, created_at, updated_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(cache_key) DO UPDATE SET
                result_json = excluded.result_json,
                updated_at = excluded.updated_at
            """,
            (
                cache_key,
                source_text,
                target_text,
                model,
                prompt_signature,
                directional_signature,
                dumps_json(result),
                now,
                now,
            ),
        )


def _save_review_result(task_id: str, item_id: str, cache_key: str, status: str, result: dict[str, Any]) -> None:
    now = utc_now()
    with get_connection() as conn:
        conn.execute(
            """
            INSERT INTO review_results (
                id, task_id, item_id, cache_key, status, has_issue, issue_type,
                issue, suggestion, directional_checks_json, error_message, raw_result_json, created_at, updated_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, '', ?, ?, ?)
            """,
            (
                uuid.uuid4().hex,
                task_id,
                item_id,
                cache_key,
                status,
                1 if result.get("has_issue") else 0,
                result.get("issue_type", ""),
                result.get("issue", ""),
                result.get("suggestion", ""),
                dumps_json(result.get("checks", {})),
                dumps_json(result),
                now,
                now,
            ),
        )


def _save_review_error(task_id: str, item_id: str, cache_key: str, message: str) -> None:
    now = utc_now()
    with get_connection() as conn:
        conn.execute(
            """
            INSERT INTO review_results (
                id, task_id, item_id, cache_key, status, error_message, created_at, updated_at
            )
            VALUES (?, ?, ?, ?, 'failed', ?, ?, ?)
            """,
            (uuid.uuid4().hex, task_id, item_id, cache_key, message, now, now),
        )


def _update_task(task_id: str, **fields: Any) -> None:
    if not fields:
        return
    fields["updated_at"] = utc_now()
    assignments = ", ".join(f"{key} = ?" for key in fields)
    values = list(fields.values())
    values.append(task_id)
    with get_connection() as conn:
        conn.execute(f"UPDATE review_tasks SET {assignments} WHERE id = ?", values)


def _get_completed_review_item_ids(task_id: str) -> set[str]:
    with get_connection() as conn:
        rows = conn.execute(
            """
            SELECT DISTINCT item_id
            FROM review_results
            WHERE task_id = ? AND status IN ('completed', 'cached', 'skipped_ai')
            """,
            (task_id,),
        ).fetchall()
    return {str(row["item_id"] or "") for row in rows if str(row["item_id"] or "").strip()}


def _get_recent_review_log_messages(task_id: str, limit: int = 6) -> list[str]:
    with get_connection() as conn:
        rows = conn.execute(
            """
            SELECT message
            FROM review_task_logs
            WHERE task_id = ?
            ORDER BY id DESC
            LIMIT ?
            """,
            (task_id, max(1, int(limit))),
        ).fetchall()
    return [str(row["message"] or "") for row in reversed(rows)]


def _replace_review_request_states(task_id: str, packages: list[list[dict[str, Any]]]) -> None:
    now = utc_now()
    with get_connection() as conn:
        conn.execute("DELETE FROM review_request_states WHERE task_id = ?", (task_id,))
        for index, package in enumerate(packages, start=1):
            first = dict(package[0] if package else {})
            conn.execute(
                """
                INSERT INTO review_request_states (
                    id, task_id, package_index, package_total, item_count, first_item_id,
                    first_source_file, first_sheet_name, first_row_number, status, created_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'queued', ?, ?)
                """,
                (
                    uuid.uuid4().hex, task_id, index, len(packages), len(package), str(first.get("id") or ""),
                    str(first.get("source_file") or ""), str(first.get("sheet_name") or ""), first.get("row_number"), now, now,
                ),
            )


def _set_review_request_status(task_id: str, package_index: int, status: str, attempt_count: int | None = None) -> None:
    if status not in {"queued", "submitted", "thinking", "output", "retrying", "completed", "failed"}:
        return
    assignments = ["status = ?", "updated_at = ?"]
    values: list[Any] = [status, utc_now()]
    if attempt_count is not None:
        assignments.append("attempt_count = ?")
        values.append(max(0, int(attempt_count)))
    values.extend([task_id, package_index])
    with get_connection() as conn:
        conn.execute(
            f"UPDATE review_request_states SET {', '.join(assignments)} WHERE task_id = ? AND package_index = ?",
            values,
        )
        row = conn.execute(
            "SELECT * FROM review_request_states WHERE task_id = ? AND package_index = ?", (task_id, package_index)
        ).fetchone()
        task = conn.execute("SELECT session_id, target_language FROM review_tasks WHERE id = ?", (task_id,)).fetchone()
    if row and task and task["session_id"]:
        emit_event(
            str(task["session_id"]),
            "review.request_status",
            {"task_id": task_id, "target_language": str(task["target_language"] or ""), "request": _request_state_to_dict(row)},
        )


def _get_review_request_states(task_id: str) -> list[dict[str, Any]]:
    with get_connection() as conn:
        rows = conn.execute(
            "SELECT * FROM review_request_states WHERE task_id = ? ORDER BY package_index", (task_id,)
        ).fetchall()
    return [_request_state_to_dict(row) for row in rows]


def _request_state_to_dict(row: Any) -> dict[str, Any]:
    status = str(row["status"] or "queued")
    label_map = {
        "queued": "排队中", "submitted": "已提交 · 等待首段", "thinking": "思考中",
        "output": "输出中", "retrying": "等待重试", "completed": "已完成", "failed": "失败",
    }
    location = " / ".join(
        item for item in (str(row["first_source_file"] or ""), str(row["first_sheet_name"] or "")) if item
    )
    if row["first_row_number"]:
        location = f"{location}{' · ' if location else ''}第 {row['first_row_number']} 行"
    return {
        "id": row["id"], "package_index": int(row["package_index"]), "package_total": int(row["package_total"]),
        "item_count": int(row["item_count"]), "status": status, "status_label": label_map.get(status, status),
        "attempt_count": int(row["attempt_count"] or 0), "location": location, "updated_at": row["updated_at"],
    }


def _add_log(task_id: str, level: str, message: str) -> None:
    timestamp = utc_now()
    with get_connection() as conn:
        conn.execute(
            """
            INSERT INTO review_task_logs (task_id, level, message, created_at)
            VALUES (?, ?, ?, ?)
            """,
            (task_id, level, message, timestamp),
        )
    logger = logging.getLogger(LOGGER_NAME)
    if not logger.handlers:
        logger = configure_file_logger(with_console=True)
    log_method = getattr(logger, str(level or "info").lower(), logger.info)
    log_method("[AI_REVIEW][%s] %s", task_id, message)


def _mark_unfinished_review_requests_failed(task_id: str) -> None:
    with get_connection() as conn:
        conn.execute(
            """
            UPDATE review_request_states
            SET status = 'failed', updated_at = ?
            WHERE task_id = ? AND status NOT IN ('completed', 'failed')
            """,
            (utc_now(), task_id),
        )


def _track_ai_review_start(config: dict[str, Any]) -> None:
    track_event("task_start.ai_review")
    if bool(config.get("enable_forbidden_check")):
        track_event("task_option.blocked_terms_enabled")
    if not bool(config.get("enable_ai_review", True)):
        return
    if str(config.get("mode", "normal")) == "directional":
        track_event("task_mode.directional_review")
    else:
        track_event("task_mode.general_review")
    track_event("task_start.ai_tool")
    if bool(config.get("enable_thinking", False)):
        track_event("model_mode.thinking_enabled")
    tier = infer_model_tier(str(config.get("model", "")))
    if tier == "flash":
        track_event("model_tier.flash")
    elif tier == "pro":
        track_event("model_tier.pro")


def _track_ai_review_finish(config: dict[str, Any], *, success: bool) -> None:
    if success:
        track_event("task_success.ai_review")
    else:
        track_event("task_fail.ai_review")
    if not bool(config.get("enable_ai_review", True)):
        return
    if success:
        track_event("task_success.ai_tool")
    else:
        track_event("task_fail.ai_tool")


def _task_to_dict(row: Any) -> dict[str, Any]:
    total_count = int(row["total_count"] or 0)
    completed_count = int(row["completed_count"] or 0)
    failed_count = int(row["failed_count"] or 0)
    requested_count = int(row["requested_count"] or 0)
    status = str(row["status"] or "")
    progress_percent = round((completed_count / total_count) * 100) if total_count > 0 else 0
    status_label_map = {
        "pending": "等待开始",
        "running": "审校中",
        "recovering": "AI 终止兜底",
        "completed": "已完成",
        "completed_with_errors": "已完成，有失败",
        "failed": "失败",
    }
    status_label = status_label_map.get(status, status or "未开始")
    if status in {"completed", "completed_with_errors"}:
        message = f"已处理 {completed_count}/{total_count}"
    elif status == "running":
        message = f"正在审校 {completed_count}/{total_count}"
    elif status == "recovering":
        message = "常规恢复已耗尽，正在进行最后一次 AI 兜底"
    elif status == "failed":
        message = "审校失败"
    else:
        message = "等待开始"
    return {
        "id": row["id"],
        "batch_id": row["batch_id"],
        "status": status,
        "status_label": status_label,
        "message": message,
        "total_count": total_count,
        "cached_count": row["cached_count"],
        "requested_count": requested_count,
        "completed_count": completed_count,
        "failed_count": failed_count,
        "progress_current": completed_count,
        "progress_total": total_count,
        "progress_percent": max(0, min(100, progress_percent)),
        "output_path": row["output_path"],
        "config": loads_json(row["config_json"], {}),
        "created_at": row["created_at"],
        "updated_at": row["updated_at"],
    }


def _result_to_dict(row: Any) -> dict[str, Any]:
    return {
        "id": row["id"],
        "item_id": row["item_id"],
        "status": row["status"],
        "has_issue": bool(row["has_issue"]) if row["has_issue"] is not None else None,
        "issue_type": row["issue_type"],
        "issue": row["issue"],
        "suggestion": row["suggestion"],
        "checks": loads_json(row["directional_checks_json"], {}),
        "error_message": row["error_message"],
        "source_text": row["source_text"],
        "target_text": row["target_text"],
        "source_file": row["source_file"],
        "sheet_name": row["sheet_name"],
        "segment_id": row["segment_id"],
        "row_number": row["row_number"],
        "matched_words": row["matched_words"] if "matched_words" in row.keys() else "",
    }
