from __future__ import annotations

from dataclasses import asdict, dataclass, field
from typing import Any


@dataclass(slots=True)
class ReaderBlock:
    pointer: str
    text: str
    label: str = ""
    language_hint: str = ""
    metadata: dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(slots=True)
class ReaderDocument:
    reader_name: str
    file_type: str
    filename: str
    file_hash: str
    # A reader can describe structure, but it cannot reliably decide a file's
    # business role. Workspace makes that decision from row-aligned evidence.
    role_hint: str = "candidate"
    structure: dict[str, Any] = field(default_factory=dict)
    blocks: list[ReaderBlock] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    def to_manifest(self, sample_limit: int = 20) -> dict[str, Any]:
        manifest = {
            "reader_name": self.reader_name,
            "file_type": self.file_type,
            "filename": self.filename,
            "file_hash": self.file_hash,
            "role_hint": self.role_hint,
            "structure": self.structure,
            "samples": [block.to_dict() for block in self.blocks[:sample_limit]],
            "block_count": len(self.blocks),
            "warnings": list(self.warnings),
        }
        if self.file_type in {"excel", "csv", "tsv"}:
            manifest["table_evidence"] = _table_evidence(self.blocks)
        return manifest


def _table_evidence(blocks: list[ReaderBlock]) -> list[dict[str, Any]]:
    """Keep row-aligned, bounded examples so Workspace can compare table columns."""
    grouped: dict[str, dict[int, list[ReaderBlock]]] = {}
    for block in blocks:
        metadata = block.metadata or {}
        row = metadata.get("row")
        column_index = metadata.get("column_index")
        if not isinstance(row, int) or not isinstance(column_index, int):
            continue
        scope = str(metadata.get("sheet") or "table")
        grouped.setdefault(scope, {}).setdefault(row, []).append(block)

    evidence: list[dict[str, Any]] = []
    for scope in sorted(grouped)[:3]:
        rows = grouped[scope]
        ordered_rows = sorted(rows.items())
        multi_cell_rows = [(row, cells) for row, cells in ordered_rows if len(cells) > 1]
        selected_rows = (multi_cell_rows or ordered_rows)[:3]
        first_by_column: dict[int, ReaderBlock] = {}
        columns: dict[int, dict[str, Any]] = {}
        for _row, cells in ordered_rows:
            for block in cells:
                metadata = block.metadata or {}
                column_index = int(metadata["column_index"])
                columns.setdefault(
                    column_index,
                    {
                        "column_index": column_index,
                        "column": str(metadata.get("column") or ""),
                        "header": str(metadata.get("header") or block.label or ""),
                    },
                )
                first_by_column.setdefault(column_index, block)

        def cell(block: ReaderBlock) -> dict[str, Any]:
            metadata = block.metadata or {}
            text = " ".join(str(block.text or "").split())
            return {
                "column_index": int(metadata["column_index"]),
                "column": str(metadata.get("column") or ""),
                "header": str(metadata.get("header") or block.label or ""),
                "pointer": block.pointer,
                "text": text[:160] + ("…" if len(text) > 160 else ""),
            }

        evidence.append(
            {
                "scope": scope,
                "columns": [columns[index] for index in sorted(columns)[:16]],
                "row_examples": [
                    {"row": row, "cells": [cell(block) for block in sorted(cells, key=lambda item: int((item.metadata or {}).get("column_index", 0)))[:16]]}
                    for row, cells in selected_rows
                ],
                "column_examples": [cell(first_by_column[index]) for index in sorted(first_by_column)[:16]],
            }
        )
    return evidence


@dataclass(slots=True)
class ReviewUnit:
    id: str
    session_id: str
    target_language: str
    target_text: str
    source_text: str = ""
    source_file: str = ""
    pointer: str = ""
    references: list[dict[str, str]] = field(default_factory=list)
    metadata: dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(slots=True)
class TargetPlan:
    language: str
    units: list[ReviewUnit] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {"language": self.language, "units": [unit.to_dict() for unit in self.units]}


@dataclass(slots=True)
class ExtractionPlan:
    source_language: str = "auto"
    targets: list[TargetPlan] = field(default_factory=list)
    content_files: list[str] = field(default_factory=list)
    reference_files: list[str] = field(default_factory=list)
    relationships: list[dict[str, str]] = field(default_factory=list)
    assumptions: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    confidence: float = 0.0
    needs_input: bool = False
    question: str = ""

    def to_dict(self) -> dict[str, Any]:
        return {
            "source_language": self.source_language,
            "targets": [target.to_dict() for target in self.targets],
            "content_files": list(self.content_files),
            "reference_files": list(self.reference_files),
            "relationships": list(self.relationships),
            "assumptions": list(self.assumptions),
            "warnings": list(self.warnings),
            "confidence": self.confidence,
            "needs_input": self.needs_input,
            "question": self.question,
        }
