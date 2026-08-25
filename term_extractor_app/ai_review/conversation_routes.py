from __future__ import annotations

import asyncio
import json
from pathlib import Path
from typing import Any

from fastapi import APIRouter, File, Header, HTTPException, Query, UploadFile
from fastapi.responses import StreamingResponse
from pydantic import BaseModel, Field

from .session_store import (
    add_attachment,
    add_message,
    answer_question,
    create_session,
    delete_session,
    delete_pending_attachment,
    get_attachment,
    get_events,
    get_pending_attachments,
    get_session,
    get_session_snapshot,
    list_sessions,
    mark_attachments_sent,
    update_session,
    update_attachment,
)
from .excel_mapping_service import get_excel_mapping_preset
from .readers import ReaderError, build_default_registry
from .workflow_service import get_session_task_results, start_session_review
from .workspace_service import submit_workspace_inspection
from ..telemetry import track_event
from ..open_utils import open_path


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


class SessionRenamePayload(BaseModel):
    title: str


class AttachmentMappingPayload(BaseModel):
    mode: str = "ai"
    preset_id: str | None = None


class LocalAttachmentPayload(BaseModel):
    file_path: str


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


@router.patch("/{session_id}")
def rename_conversation(session_id: str, payload: SessionRenamePayload) -> dict[str, Any]:
    title = payload.title.strip()
    if not title:
        raise HTTPException(status_code=400, detail="会话名称不能为空")
    try:
        session = update_session(session_id, title=title[:120], title_custom=True)
    except ValueError as exc:
        raise HTTPException(status_code=404, detail=str(exc)) from exc
    return {"ok": True, "session": session}


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


@router.post("/{session_id}/attachments/local")
def attach_local_file(session_id: str, payload: LocalAttachmentPayload) -> dict[str, Any]:
    if not get_session(session_id):
        raise HTTPException(status_code=404, detail="审校会话不存在")
    path = Path(str(payload.file_path or "").strip())
    if not path.exists() or not path.is_file():
        raise HTTPException(status_code=404, detail="找不到原始文件，无法添加。")
    try:
        attachment = add_attachment(
            session_id,
            path.name,
            path.read_bytes(),
            original_path=str(path),
        )
    except OSError as exc:
        raise HTTPException(status_code=400, detail=f"读取原始文件失败：{exc}") from exc
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    return {"attachment": attachment}


@router.patch("/{session_id}/attachments/{attachment_id}")
def configure_attachment_mapping(
    session_id: str, attachment_id: str, payload: AttachmentMappingPayload
) -> dict[str, Any]:
    attachment = get_attachment(attachment_id)
    if not attachment or attachment["session_id"] != session_id:
        raise HTTPException(status_code=404, detail="附件不存在")
    if attachment.get("sent_at"):
        raise HTTPException(status_code=400, detail="已发送附件不能修改导入方式")
    mode = payload.mode.strip().lower()
    if mode not in {"ai", "preset"}:
        raise HTTPException(status_code=400, detail="导入方式无效")
    preset_id = None
    if mode == "preset":
        if attachment["original_filename"].lower().endswith((".xlsx", ".xlsm")) is False:
            raise HTTPException(status_code=400, detail="映射模板仅适用于 Excel 文件")
        try:
            preset = get_excel_mapping_preset(str(payload.preset_id or ""))
        except ValueError as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc
        preset_id = preset["id"]
    updated = update_attachment(
        attachment_id,
        mapping_mode=mode,
        mapping_preset_id=preset_id,
    )
    return {"ok": True, "attachment": updated}


@router.get("/{session_id}/attachments/{attachment_id}/structure")
def inspect_attachment_structure(session_id: str, attachment_id: str) -> dict[str, Any]:
    attachment = get_attachment(attachment_id)
    if not attachment or attachment["session_id"] != session_id:
        raise HTTPException(status_code=404, detail="附件不存在")
    if attachment.get("sent_at"):
        raise HTTPException(status_code=400, detail="已发送附件不能修改导入方式")
    if not str(attachment.get("original_filename") or "").lower().endswith((".xlsx", ".xlsm")):
        raise HTTPException(status_code=400, detail="只有 Excel 文件支持手动映射")
    try:
        document = build_default_registry().read(
            Path(str(attachment.get("stored_path") or "")),
            str(attachment.get("original_filename") or ""),
        )
    except ReaderError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    manifest = document.to_manifest()
    updated = update_attachment(
        attachment_id,
        file_type=document.file_type,
        status="ready",
        manifest_json=manifest,
        error_message="",
    )
    return {"attachment": updated, "manifest": manifest}


@router.delete("/{session_id}/attachments/{attachment_id}")
def remove_pending_attachment(session_id: str, attachment_id: str) -> dict[str, Any]:
    try:
        delete_pending_attachment(session_id, attachment_id)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    return {"ok": True}


@router.post("/{session_id}/attachments/{attachment_id}/open-original")
def open_conversation_attachment_original(session_id: str, attachment_id: str) -> dict[str, Any]:
    attachment = get_attachment(attachment_id)
    if not attachment or attachment["session_id"] != session_id:
        raise HTTPException(status_code=404, detail="附件不存在")
    original_path_text = str(attachment.get("original_path") or "").strip()
    if not original_path_text:
        raise HTTPException(status_code=404, detail="未记录原始文件路径，无法溯源打开。")
    original_path = Path(original_path_text)
    if not original_path.exists() or not original_path.is_file():
        raise HTTPException(status_code=404, detail=f"找不到原始文件：{original_path}")
    try:
        open_path(original_path)
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"打开原始文件失败：{exc}") from exc
    return {"ok": True, "file_path": str(original_path)}


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
    pending = get_pending_attachments(session_id)
    attachment_ids = [str(item["id"]) for item in pending]
    if not payload.text.strip() and not attachment_ids:
        raise HTTPException(status_code=400, detail="请输入待审校文本或添加文件")
    try:
        sent = mark_attachments_sent(session_id, attachment_ids)
        content = payload.text.strip() or f"已发送 {len(sent)} 个文件"
        add_message(
            session_id,
            "user",
            "attachment_message" if sent else "message",
            content,
            {"attachments": [{key: value for key, value in item.items() if key != "stored_path"} for item in sent]},
        )
        run_id = submit_workspace_inspection(
            session_id,
            payload.text,
            attachment_ids=attachment_ids,
            record_user_message=False,
        )
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    track_event("task_start.ai_review_workspace")
    return {"ok": True, "run_id": run_id}


@router.post("/{session_id}/decision")
def submit_decision(session_id: str, payload: WorkspaceDecisionPayload) -> dict[str, Any]:
    if not get_session(session_id):
        raise HTTPException(status_code=404, detail="审校会话不存在")
    try:
        snapshot = get_session_snapshot(session_id) or {}
        source_run = next(
            (item for item in snapshot.get("workspace_runs", []) if item.get("id") == payload.run_id),
            None,
        )
        if payload.action == "confirm":
            targets = list((source_run or {}).get("plan", {}).get("targets", []) or [])
            if not any(item.get("units") for item in targets):
                raise ValueError("当前没有可确认的审校方案，请补充目标语种和译文内容后重新识别")
        if payload.question_id:
            answer_question(payload.question_id, payload.answer or payload.action)
        if payload.action == "confirm":
            tasks = start_session_review(session_id, payload.run_id)
            return {"ok": True, "tasks": tasks}
        run_id = submit_workspace_inspection(
            session_id,
            "",
            adjustment_text=payload.answer,
            attachment_ids=list((source_run or {}).get("attachment_ids") or []),
            parent_run_id=str(payload.run_id or ""),
        )
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
async def conversation_events(
    session_id: str,
    after_id: int = Query(0, ge=0),
    last_event_id: str | None = Header(default=None, alias="Last-Event-ID"),
) -> StreamingResponse:
    if not get_session(session_id):
        raise HTTPException(status_code=404, detail="审校会话不存在")

    async def stream():
        try:
            header_cursor = int(last_event_id or 0)
        except ValueError:
            header_cursor = 0
        cursor = max(after_id, header_cursor)
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
