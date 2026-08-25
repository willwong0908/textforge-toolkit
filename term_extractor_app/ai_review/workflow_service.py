from __future__ import annotations

import threading
import time
from pathlib import Path
from typing import Any

from .cache_service import create_batch, replace_batch_items
from .database import get_connection, loads_json
from .prompt_service import get_prompt_template
from .review_service import create_review_task, get_review_results, get_review_task
from .session_store import add_message, emit_event, get_session_snapshot, link_task, update_session
from .workspace_service import save_confirmed_profile
from ..telemetry import track_event


def start_session_review(session_id: str, run_id: str | None = None) -> list[dict[str, str]]:
    snapshot = get_session_snapshot(session_id)
    if snapshot is None:
        raise ValueError("审校会话不存在")
    session = snapshot["session"]
    run = _get_workspace_run(session_id, run_id)
    if run is None:
        raise ValueError("请先完成文件识别")
    if run["status"] not in {"ready", "needs_input"}:
        raise ValueError("当前识别结果尚未就绪")
    plan = run["plan"]
    targets = [target for target in plan.get("targets", []) if target.get("units")]
    if not targets:
        raise ValueError("识别结果中没有可审校条目")
    if snapshot["tasks"]:
        active = [link for link in snapshot["tasks"] if (get_review_task(link["task_id"]) or {}).get("status") not in {"completed", "completed_with_errors", "failed"}]
        if active:
            raise ValueError("当前会话已有审校任务在运行")

    template = get_prompt_template(session.get("prompt_template_id"))
    created: list[dict[str, str]] = []
    for target in targets:
        language = str(target.get("language") or "auto")
        units = list(target.get("units") or [])
        batch_id = create_batch(
            original_filename=f"{session['title']}_{language}",
            stored_path=Path("workspace://") / session_id,
            file_type="workspace",
            status="uploaded",
            metadata={"session_id": session_id, "target_language": language, "workspace_run_id": run["id"]},
        )
        items = []
        for order, unit in enumerate(units, 1):
            pointer = str(unit.get("pointer") or "")
            metadata = dict(unit.get("metadata") or {})
            references = [item for item in unit.get("references", []) if isinstance(item, dict)]
            info = [
                {"category": str(item.get("category") or "参考"), "value": str(item.get("value") or "")}
                for item in references
                if str(item.get("value") or "").strip()
            ]
            items.append(
                {
                    "source_file": str(unit.get("source_file") or "直接输入"),
                    "sheet_name": str(metadata.get("scope") or ""),
                    "segment_id": pointer,
                    "row_number": metadata.get("row"),
                    "source_text": str(unit.get("source_text") or ""),
                    "target_text": str(unit.get("target_text") or ""),
                    "target_language": language,
                    "status_note": "原文为空" if not str(unit.get("source_text") or "").strip() else "",
                    "item_order": order,
                    "info": info,
                    "references": references,
                    "location": {"pointer": pointer, **metadata},
                }
            )
        replace_batch_items(
            batch_id=batch_id,
            items=items,
            metadata_update={"preview_ready": True, "session_id": session_id, "target_language": language},
        )
        task_id = create_review_task(
            batch_id=batch_id,
            prompt_template_id=template["id"],
            source_language=str(plan.get("source_language") or session.get("source_language") or "auto"),
            target_language=language,
            mode="normal",
            enable_ai_review=True,
            enable_forbidden_check=bool(str(template.get("forbidden_words_text") or "").strip()),
            session_id=session_id,
        )
        link_task(session_id, language, task_id)
        created.append({"target_language": language, "task_id": task_id})

    save_confirmed_profile(run["id"])
    update_session(session_id, status="reviewing")
    add_message(
        session_id,
        "assistant",
        "review_started",
        "已按目标语言分别启动审校：" + "、".join(item["target_language"] for item in created) + "。",
        {"tasks": created},
    )
    emit_event(session_id, "review.started", {"tasks": created})
    if len(created) > 1:
        track_event("task_mode.ai_review_multi_target")
    monitor = threading.Thread(target=_monitor_session_tasks, args=(session_id, created), daemon=True)
    monitor.start()
    return created


def get_session_task_results(session_id: str) -> list[dict[str, Any]]:
    snapshot = get_session_snapshot(session_id)
    if snapshot is None:
        raise ValueError("审校会话不存在")
    result: list[dict[str, Any]] = []
    for link in snapshot["tasks"]:
        task = get_review_task(link["task_id"])
        if not task:
            continue
        result.append(
            {
                "target_language": link["target_language"],
                "task": task,
                "results": get_review_results(link["task_id"], limit=20),
            }
        )
    return result


def _get_workspace_run(session_id: str, run_id: str | None) -> dict[str, Any] | None:
    with get_connection() as conn:
        if run_id:
            row = conn.execute(
                "SELECT * FROM workspace_runs WHERE id = ? AND session_id = ?",
                (run_id, session_id),
            ).fetchone()
        else:
            row = conn.execute(
                "SELECT * FROM workspace_runs WHERE session_id = ? ORDER BY created_at DESC LIMIT 1",
                (session_id,),
            ).fetchone()
    if not row:
        return None
    return {
        "id": row["id"],
        "session_id": row["session_id"],
        "status": row["status"],
        "plan": loads_json(row["plan_json"], {}),
    }


def _monitor_session_tasks(session_id: str, links: list[dict[str, str]]) -> None:
    previous: dict[str, tuple[str, int, int]] = {}
    while True:
        tasks = [get_review_task(link["task_id"]) for link in links]
        tasks = [task for task in tasks if task]
        if not tasks:
            update_session(session_id, status="failed", error_message="审校任务不存在")
            return
        for task in tasks:
            key = (str(task["status"]), int(task["completed_count"]), int(task["failed_count"]))
            if previous.get(task["id"]) != key:
                previous[task["id"]] = key
                emit_event(
                    session_id,
                    "review.progress",
                    {
                        "task_id": task["id"],
                        "status": task["status"],
                        "completed_count": task["completed_count"],
                        "total_count": task["total_count"],
                        "failed_count": task["failed_count"],
                        "output_path": task.get("output_path") or "",
                    },
                )
        final = {"completed", "completed_with_errors", "failed"}
        if all(str(task["status"]) in final for task in tasks):
            failed = [task for task in tasks if task["status"] == "failed"]
            partial = [task for task in tasks if task["status"] == "completed_with_errors"]
            status = "failed" if len(failed) == len(tasks) else "partial" if failed or partial else "completed"
            update_session(session_id, status=status)
            outputs = [str(task.get("output_path") or "") for task in tasks if task.get("output_path")]
            add_message(
                session_id,
                "assistant",
                "review_completed",
                "审校已完成。" if status == "completed" else "审校已结束，部分目标语言存在失败项。",
                {"status": status, "outputs": outputs},
            )
            emit_event(session_id, "review.completed", {"status": status, "outputs": outputs})
            return
        time.sleep(0.5)
