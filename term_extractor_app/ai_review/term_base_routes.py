from __future__ import annotations

from typing import Any

from fastapi import APIRouter, File, HTTPException, UploadFile

from .term_base_service import TermBaseError, delete_term_base, list_term_bases, upload_term_base


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
