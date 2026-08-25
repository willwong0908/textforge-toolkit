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
    role_hint: str = "content"
    structure: dict[str, Any] = field(default_factory=dict)
    blocks: list[ReaderBlock] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    def to_manifest(self, sample_limit: int = 20) -> dict[str, Any]:
        return {
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
