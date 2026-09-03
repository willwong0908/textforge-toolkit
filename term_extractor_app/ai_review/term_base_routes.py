from __future__ import annotations

from typing import Any

from fastapi import APIRouter, File, HTTPException, UploadFile

from .term_base_service import TermBaseError, delete_term_base, list_term_bases, upload_term_base
from .memoq_service import MemoQError, clear_credentials, credential_status, list_termbases, save_credentials


router = APIRouter(prefix="/api/ai-review/term-bases", tags=["ai-review-term-bases"])


def _public_term_base(item: dict[str, Any]) -> dict[str, Any]:
    return {key: value for key, value in item.items() if key != "stored_path"}


@router.get("")
def term_bases() -> dict[str, Any]:
    return {"term_bases": [_public_term_base(item) for item in list_term_bases()]}


@router.post("")
async def add_term_base(file: UploadFile = File(...)) -> dict[str, Any]:
    data = await file.read()
    try:
        term_base = upload_term_base(file.filename or "", data)
    except TermBaseError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    return {"term_base": _public_term_base(term_base)}


@router.delete("/{term_base_id}")
def remove_term_base(term_base_id: str) -> dict[str, Any]:
    try:
        delete_term_base(term_base_id)
    except TermBaseError as exc:
        raise HTTPException(status_code=404, detail=str(exc)) from exc
    return {"ok": True}


@router.get("/memoq/status")
def memoq_status() -> dict[str, Any]:
    return {"memoq": credential_status()}


@router.post("/memoq/bind")
def bind_memoq(payload: dict[str, Any]) -> dict[str, Any]:
    try:
        save_credentials(str(payload.get("username") or ""), str(payload.get("password") or ""), str(payload.get("base_url") or ""))
        # Validate credentials immediately so a typo is not persisted as a false success.
        from .memoq_service import MemoQClient, _load_credentials
        credentials = _load_credentials() or {}
        client = MemoQClient(credentials)
        try:
            client.login()
        finally:
            client.close()
    except MemoQError as exc:
        clear_credentials()
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    return {"memoq": credential_status()}


@router.delete("/memoq/bind")
def unbind_memoq() -> dict[str, Any]:
    clear_credentials()
    return {"memoq": credential_status()}


@router.get("/memoq/termbases")
def memoq_termbases(q: str = "") -> dict[str, Any]:
    try:
        return {"termbases": list_termbases(q)}
    except MemoQError as exc:
        raise HTTPException(status_code=409, detail=str(exc)) from exc
