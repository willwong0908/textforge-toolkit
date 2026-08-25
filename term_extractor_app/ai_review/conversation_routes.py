from __future__ import annotations

import asyncio
import json
from typing import Any

from fastapi import APIRouter, File, HTTPException, Query, UploadFile
from fastapi.responses import StreamingResponse
from pydantic import BaseModel, Field

from .session_store import (
    add_attachment,
    answer_question,
    create_session,
    delete_session,
    get_events,
    get_session,
    get_session_snapshot,
    list_sessions,
    update_session,
)
from .workflow_service import get_session_task_results, start_session_review
from .workspace_service import submit_workspace_inspection
from ..telemetry import track_event


router = APIRouter(prefix="/api/ai-review/conversations", tags=["ai-review-conversations"])


class SessionCreatePayload(BaseModel):
    title: str = "新审校"
    prompt_template_id: str | None = None
    source_language: str = "auto"
    target_languages: list[str] = Field(default_factory=lambda: ["auto"])
    auto_start: bool = False


class SessionMessagePayload(BaseModel):
    text: str = ""
    prompt_template_id: str | None = None
    source_language: str = "auto"
    target_languages: list[str] = Field(default_factory=lambda: ["auto"])
    auto_start: bool | None = None


class SessionDeletePayload(BaseModel):
    confirm: bool = False


class WorkspaceDecisionPayload(BaseModel):
    run_id: str | None = None
    question_id: str | None = None
    action: str = "confirm"
    answer: str = ""


@router.get("")
def conversations() -> dict[str, Any]:
    return {"sessions": list_sessions()}


@router.post("")
def create_conversation(payload: SessionCreatePayload) -> dict[str, Any]:
    return {"session": create_session(**payload.model_dump())}


@router.get("/{session_id}")
def conversation(session_id: str) -> dict[str, Any]:
    snapshot = get_session_snapshot(session_id)
    if snapshot is None:
        raise HTTPException(status_code=404, detail="审校会话不存在")
    snapshot["task_results"] = get_session_task_results(session_id)
    return snapshot


@router.delete("/{session_id}")
def remove_conversation(session_id: str, payload: SessionDeletePayload) -> dict[str, Any]:
    if not payload.confirm:
        raise HTTPException(status_code=400, detail="删除会话前必须二次确认")
    try:
        delete_session(session_id)
    except ValueError as exc:
        raise HTTPException(status_code=404, detail=str(exc)) from exc
    return {"ok": True}


@router.post("/{session_id}/attachments")
async def upload_attachments(session_id: str, files: list[UploadFile] = File(...)) -> dict[str, Any]:
    if not get_session(session_id):
        raise HTTPException(status_code=404, detail="审校会话不存在")
    if not files:
        raise HTTPException(status_code=400, detail="请选择文件")
    if len(files) > 50:
        raise HTTPException(status_code=400, detail="一次最多上传 50 个文件")
    result = []
    for file in files:
        data = await file.read()
        try:
            result.append(add_attachment(session_id, file.filename or "unknown", data))
        except ValueError as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc
    return {"attachments": result}


@router.post("/{session_id}/messages")
def send_message(session_id: str, payload: SessionMessagePayload) -> dict[str, Any]:
    session = get_session(session_id)
    if session is None:
        raise HTTPException(status_code=404, detail="审校会话不存在")
    update_fields: dict[str, Any] = {
        "prompt_template_id": payload.prompt_template_id,
        "source_language": payload.source_language or "auto",
        "target_languages_json": payload.target_languages or ["auto"],
    }
    if payload.auto_start is not None:
        update_fields["auto_start"] = payload.auto_start
    update_session(session_id, **update_fields)
    try:
        run_id = submit_workspace_inspection(session_id, payload.text)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    track_event("task_start.ai_review_workspace")
    return {"ok": True, "run_id": run_id}


@router.post("/{session_id}/decision")
def submit_decision(session_id: str, payload: WorkspaceDecisionPayload) -> dict[str, Any]:
    if not get_session(session_id):
        raise HTTPException(status_code=404, detail="审校会话不存在")
    try:
        if payload.question_id:
            answer_question(payload.question_id, payload.answer or payload.action)
        if payload.action == "confirm":
            tasks = start_session_review(session_id, payload.run_id)
            return {"ok": True, "tasks": tasks}
        run_id = submit_workspace_inspection(session_id, "", adjustment_text=payload.answer)
        return {"ok": True, "run_id": run_id}
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


@router.get("/{session_id}/results")
def conversation_results(session_id: str) -> dict[str, Any]:
    try:
        return {"tasks": get_session_task_results(session_id)}
    except ValueError as exc:
        raise HTTPException(status_code=404, detail=str(exc)) from exc


@router.get("/{session_id}/events")
async def conversation_events(session_id: str, after_id: int = Query(0, ge=0)) -> StreamingResponse:
    if not get_session(session_id):
        raise HTTPException(status_code=404, detail="审校会话不存在")

    async def stream():
        cursor = after_id
        idle_ticks = 0
        while True:
            events = get_events(session_id, cursor)
            if events:
                idle_ticks = 0
                for event in events:
                    cursor = max(cursor, int(event["id"]))
                    payload = json.dumps(event, ensure_ascii=False)
                    yield f"id: {cursor}\nevent: {event['event_type']}\ndata: {payload}\n\n"
            else:
                idle_ticks += 1
                if idle_ticks % 15 == 0:
                    yield ": keep-alive\n\n"
            await asyncio.sleep(0.4)

    return StreamingResponse(
        stream(),
        media_type="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )
