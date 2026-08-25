from __future__ import annotations

import hashlib
import json
import re
import threading
import time
import uuid
from collections import defaultdict
from pathlib import Path
from typing import Any

from .database import dumps_json, get_connection, loads_json, utc_now
from .excel_mapping_service import get_excel_mapping_preset
from .readers import ReaderDependencyError, ReaderDocument, ReaderError, build_default_registry
from .session_store import (
    add_message,
    create_question,
    create_workspace_run,
    emit_event,
    get_session_snapshot,
    update_attachment,
    update_session,
    update_workspace_run,
)
from .shared_provider import SharedProviderError, get_shared_ai_settings, workspace_chat
from .types import ExtractionPlan, ReaderBlock, ReviewUnit, TargetPlan
from ..telemetry import track_event


WORKSPACE_SYSTEM_PROMPT = """你是翻译审校工具的 Workspace Agent。你只负责理解文件结构和文件关系，不执行审校。
根据结构清单判断正文、参考资料、原文位置和每种目标语言的译文位置。优先采用已有稳定指针，不要编造不存在的内容。
约束优先级：用户本轮最新的自然语言说明 > 用户明确选择的具体语言/none > 界面中的 auto > 文件名和样例推测。auto 仅表示“请你自行判断”，绝不是用户坚持存在原文或某一源语言的约束；用户先前或界面中是 auto、随后说明“不用原文/全文是译文”时，应结合全文语义和文件结构判断，并可返回 source_language 为 none。不要按单个关键词机械判断，例如“缺失源文”可能是在描述问题而不是要求无原文审校。
用户选择源语言为 none 时，所有映射都必须把 source_column_index/source_pointer 设为空，并把 source_language 返回为 none，不得补造原文。当你判断没有原文、只需要审校译文时，也必须显式返回 source_language 为 none，并让所有映射的 source_column_index/source_pointer 为空；不能因为没有原文就要求用户补充原文。
对于 DOCX、PPTX、TXT、PDF、JSON、XML 等非表格正文，target_pointer 可以使用 "*" 表示该附件所有可读文本位置，也可以使用一个文档级指针前缀（例如 word/document.xml）；后端会展开为实际位置。DOCX/PPTX 必须保持 Reader 的段落/文本位置为独立审校单元，绝不能把整篇文档拼成一个审校条目。用户明确无原文时，应优先以这个方式生成可执行的全文译文审校方案。
只有完全无法形成可执行提取方案时 needs_input 才能为 true。只返回严格 JSON，不要输出 Markdown。"""

WORKSPACE_USER_PROMPT = """请检查以下本地解析器生成的结构清单，并返回：
{{
  "content_files": ["attachment_id"],
  "reference_files": ["attachment_id"],
  "source_language": "语言或 auto",
  "targets": [{{"language":"语言", "mappings":[{{"attachment_id":"...", "scope":"工作表名、table 或 document", "source_column_index":"表格原文列的零基索引，可为空", "target_column_index":"表格译文列的零基索引", "target_pointer":"非表格格式使用精确位置、文档级前缀或 *（全文）", "source_pointer":"非表格格式可为空"}}]}}],
  "relationships": [{{"from":"attachment_id", "to":"attachment_id", "kind":"reference|translation"}}],
  "assumptions": [], "warnings": [], "confidence": 0.0,
  "needs_input": false, "question": ""
}}

界面源语言设置（auto 表示交给你判断，不是硬约束）：{source_language}
界面目标语言设置：{target_languages}
用户本轮最新说明（优先于 auto，必须结合整句话理解）：{adjustment_text}
本轮待调整的上一版识别方案（只有用户从“其他：自行输入”补充时才提供；普通新消息为空）：
{adjustment_context_json}
结构清单：
{manifest_json}"""

def submit_workspace_inspection(
    session_id: str,
    direct_text: str = "",
    *,
    adjustment_text: str = "",
    attachment_ids: list[str] | None = None,
    record_user_message: bool = True,
    parent_run_id: str = "",
) -> str:
    snapshot = get_session_snapshot(session_id)
    if snapshot is None:
        raise ValueError("审校会话不存在")
    if direct_text.strip() and record_user_message:
        add_message(session_id, "user", "message", direct_text.strip())
    effective_attachment_ids = (
        [str(value) for value in attachment_ids]
        if attachment_ids is not None
        else [str(item["id"]) for item in snapshot["attachments"]]
    )
    if not direct_text.strip() and not effective_attachment_ids:
        raise ValueError("请添加文件或输入待审校文本")
    update_session(session_id, status="inspecting", error_message="")
    settings = get_shared_ai_settings()
    signature = _input_signature(snapshot, direct_text, adjustment_text, effective_attachment_ids)
    run = create_workspace_run(
        session_id,
        model=str(settings.get("selected_model") or ""),
        input_signature=signature,
        attachment_ids=effective_attachment_ids,
    )
    emit_event(session_id, "workspace.started", {"run_id": run["id"]})
    thread = threading.Thread(
        target=_run_inspection,
        args=(session_id, run["id"], direct_text, adjustment_text, effective_attachment_ids, parent_run_id),
        daemon=True,
    )
    thread.start()
    return run["id"]


def inspect_workspace_sync(
    session_id: str,
    direct_text: str = "",
    *,
    adjustment_text: str = "",
    attachment_ids: list[str] | None = None,
    parent_run_id: str = "",
) -> dict[str, Any]:
    snapshot = get_session_snapshot(session_id)
    if snapshot is None:
        raise ValueError("审校会话不存在")
    effective_attachment_ids = (
        [str(value) for value in attachment_ids]
        if attachment_ids is not None
        else [str(item["id"]) for item in snapshot["attachments"]]
    )
    settings = get_shared_ai_settings()
    run = create_workspace_run(
        session_id,
        model=str(settings.get("selected_model") or ""),
        input_signature=_input_signature(snapshot, direct_text, adjustment_text, effective_attachment_ids),
        attachment_ids=effective_attachment_ids,
    )
    return _inspect(session_id, run["id"], direct_text, adjustment_text, effective_attachment_ids, parent_run_id)


def _run_inspection(
    session_id: str,
    run_id: str,
    direct_text: str,
    adjustment_text: str,
    attachment_ids: list[str],
    parent_run_id: str,
) -> None:
    try:
        _inspect(session_id, run_id, direct_text, adjustment_text, attachment_ids, parent_run_id)
    except Exception as exc:
        message = str(exc) or repr(exc)
        update_workspace_run(run_id, status="failed", error=message)
        update_session(session_id, status="failed", error_message=message)
        add_message(session_id, "assistant", "error", f"文件识别失败：{message}")
        emit_event(session_id, "workspace.failed", {"run_id": run_id, "message": message})
        track_event("task_fail.ai_review_workspace")


def _inspect(
    session_id: str,
    run_id: str,
    direct_text: str,
    adjustment_text: str = "",
    attachment_ids: list[str] | None = None,
    parent_run_id: str = "",
) -> dict[str, Any]:
    snapshot = get_session_snapshot(session_id)
    if snapshot is None:
        raise ValueError("审校会话不存在")
    session = snapshot["session"]
    documents: dict[str, ReaderDocument] = {}
    registry = build_default_registry()
    dependency_errors: list[dict[str, str]] = []
    allowed_attachment_ids = set(attachment_ids) if attachment_ids is not None else None
    selected_attachments = [
        attachment for attachment in snapshot["attachments"]
        if allowed_attachment_ids is None or str(attachment["id"]) in allowed_attachment_ids
    ]
    attachment_settings = {str(item["id"]): item for item in selected_attachments}
    for attachment in selected_attachments:
        attachment_id = str(attachment["id"])
        try:
            document = registry.read(Path(attachment["stored_path"]), attachment["original_filename"])
            documents[attachment_id] = document
            update_attachment(
                attachment_id,
                file_type=document.file_type,
                status="ready",
                manifest_json=document.to_manifest(),
                error_message="",
            )
            emit_event(
                session_id,
                "attachment.inspected",
                {"attachment_id": attachment_id, "manifest": document.to_manifest(sample_limit=5)},
            )
        except ReaderDependencyError as exc:
            update_attachment(attachment_id, status="needs_reader", error_message=str(exc))
            dependency_errors.append(
                {
                    "attachment_id": attachment_id,
                    "extension": Path(attachment["original_filename"]).suffix.lower(),
                    "package_name": exc.package_name,
                    "source_url": exc.source_url,
                    "message": str(exc),
                }
            )
        except ReaderError as exc:
            update_attachment(attachment_id, status="needs_reader", error_message=str(exc))
            dependency_errors.append(
                {
                    "attachment_id": attachment_id,
                    "extension": Path(attachment["original_filename"]).suffix.lower(),
                    "package_name": "",
                    "source_url": "",
                    "message": str(exc),
                }
            )

    if dependency_errors:
        plan = ExtractionPlan(
            source_language=session["source_language"],
            warnings=[item["message"] for item in dependency_errors],
            needs_input=True,
            question="部分文件格式当前不支持，请移除或转换为内置格式后重试。",
        )
        plan_data = plan.to_dict()
        plan_data["reader_requirements"] = dependency_errors
        plan_data["file_summaries"] = _build_file_summaries(plan_data, documents)
        update_workspace_run(run_id, status="needs_input", plan=plan_data)
        update_session(session_id, status="needs_input")
        create_question(session_id, run_id, plan.question, recommended_label="")
        add_message(session_id, "assistant", "workspace_report", _format_report(plan_data), plan_data)
        track_event("task_question.ai_review_workspace")
        return plan_data

    adjustment = adjustment_text.strip()
    content_documents = [document for document in documents.values() if document.role_hint != "reference"]
    xliff_local_only = (
        bool(content_documents)
        and not direct_text.strip()
        and not adjustment
        and all(document.file_type == "xliff" for document in content_documents)
    )
    candidate = _apply_natural_language_adjustment(session, documents, adjustment) if adjustment else None
    candidate = candidate or _build_deterministic_plan(session, documents, direct_text)
    preset_plan = (
        _build_preset_plan(session, documents, attachment_settings)
        if not adjustment and not direct_text.strip()
        else None
    )
    cached_plan = (
        _get_cached_profile(run_id, documents, session)
        if not adjustment and not direct_text.strip() and not xliff_local_only and preset_plan is None
        else None
    )
    cache_hit = cached_plan is not None
    if xliff_local_only:
        plan = candidate
        plan.assumptions.append("XLIFF 自带原文、译文及语言结构，本次使用本地确定性读取。")
    elif preset_plan is not None:
        plan = preset_plan
    elif cached_plan is not None:
        plan = cached_plan
    else:
        refined, thinking_text = _refine_with_model(
            session,
            documents,
            direct_text=direct_text,
            adjustment_text=adjustment,
            run_id=run_id,
            parent_run_id=parent_run_id,
        )
        if refined is not None:
            plan = refined
            if adjustment:
                plan.assumptions.append(f"Workspace Agent 已重新应用用户调整：{_short_sample(adjustment, 140)}")
        else:
            plan = candidate
            plan.confidence = min(plan.confidence, 0.49)
            if adjustment:
                plan.needs_input = True
                plan.question = "Workspace Agent 没有成功应用这条调整。请确认模型配置后重试，或明确到文件名、工作表以及原文/译文所在列。"
                plan.warnings.append("本轮只保留本地候选映射供核对，未把它当作调整后的正式方案。")
            else:
                plan.needs_input = True
                plan.question = "Workspace Agent 未能完成文件关系和正文位置判断。请检查模型配置后重试，或补充文件名、工作表及原文/译文所在列。"
                plan.warnings.append("下方内容只是 Reader 生成的本地候选映射，尚未经过 Workspace Agent 确认。")
    if direct_text.strip():
        _normalize_direct_text_plan(plan, session, documents)
    _attach_references(plan, documents)
    plan_data = plan.to_dict()
    plan_data["cache_hit"] = cache_hit
    plan_data["file_summaries"] = _build_file_summaries(plan_data, documents)
    state = "needs_input" if plan.needs_input else "ready"
    update_workspace_run(run_id, status=state, plan=plan_data)
    update_session(session_id, status=state)
    message_payload = dict(plan_data)
    if "thinking_text" in locals() and thinking_text.strip():
        message_payload["_thinking_text"] = thinking_text[-30000:]
    message_payload["_workspace_run_id"] = run_id
    add_message(session_id, "assistant", "workspace_report", _format_report(plan_data), message_payload)
    emit_event(session_id, f"workspace.{state}", {"run_id": run_id, "plan": plan_data})
    if plan.needs_input:
        create_question(
            session_id,
            run_id,
            plan.question or "无法确定待审校内容，请补充说明。",
            recommended_label="",
        )
        track_event("task_question.ai_review_workspace")
    else:
        track_event("task_success.ai_review_workspace")
        if bool(session.get("auto_start")) and not cache_hit:
            from .workflow_service import start_session_review

            start_session_review(session_id, run_id)
        else:
            create_question(
                session_id,
                run_id,
                (
                    "已读取完全匹配的本地结构缓存，请确认后开始审校；选择其他将重新交给 Workspace Agent 识别。"
                    if cache_hit
                    else "识别方案已准备好，请确认后开始审校；如有需要，也可以输入补充说明后重新识别。"
                ),
                "确认推荐方案",
            )
    return plan_data


def _build_deterministic_plan(
    session: dict[str, Any], documents: dict[str, ReaderDocument], direct_text: str
) -> ExtractionPlan:
    content_ids = [attachment_id for attachment_id, doc in documents.items() if doc.role_hint != "reference"]
    reference_ids = [attachment_id for attachment_id, doc in documents.items() if doc.role_hint == "reference"]
    targets: dict[str, list[ReviewUnit]] = defaultdict(list)
    assumptions: list[str] = []
    warnings: list[str] = []

    if direct_text.strip():
        for index, chunk in enumerate(_split_direct_text(direct_text), 1):
            detected = detect_language(chunk)
            language = _select_target_language(detected, session["target_languages"])
            targets[language].append(
                ReviewUnit(
                    id=uuid.uuid4().hex,
                    session_id=session["id"],
                    target_language=language,
                    target_text=chunk,
                    pointer=f"direct:{index}",
                )
            )
        assumptions.append("直接输入文本按段落和文字脚本自动识别目标语言。")

    for attachment_id in content_ids:
        document = documents[attachment_id]
        if document.file_type == "xliff":
            for block in document.blocks:
                language = str(block.metadata.get("target_language") or block.language_hint or "auto")
                language = _select_target_language(language, session["target_languages"])
                targets[language].append(
                    ReviewUnit(
                        id=uuid.uuid4().hex,
                        session_id=session["id"],
                        target_language=language,
                        source_text=str(block.metadata.get("source_text") or ""),
                        target_text=block.text,
                        source_file=document.filename,
                        pointer=block.pointer,
                        metadata={"attachment_id": attachment_id, **block.metadata},
                    )
                )
            continue
        if document.file_type in {"excel", "csv", "tsv"}:
            mapped, mapping_assumptions = _map_tabular_document(session, attachment_id, document)
            for language, units in mapped.items():
                targets[language].extend(units)
            assumptions.extend(mapping_assumptions)
            continue
        for block in document.blocks:
            detected = block.language_hint or detect_language(block.text)
            language = _select_target_language(detected, session["target_languages"])
            targets[language].append(
                ReviewUnit(
                    id=uuid.uuid4().hex,
                    session_id=session["id"],
                    target_language=language,
                    target_text=block.text,
                    source_file=document.filename,
                    pointer=block.pointer,
                    metadata={"attachment_id": attachment_id, **block.metadata},
                )
            )
        assumptions.append(f"{document.filename} 未发现稳定双语结构，按无源文文本处理。")

    target_plans = [TargetPlan(language=language, units=units) for language, units in targets.items() if units]
    total_units = sum(len(item.units) for item in target_plans)
    content_documents = [document for document in documents.values() if document.role_hint != "reference"]
    if target_plans and content_documents and all(document.file_type == "xliff" for document in content_documents):
        confidence = 0.98
    elif any("未发现标准表头" in value or "只有一列" in value or "未发现稳定双语结构" in value for value in assumptions):
        confidence = 0.62
    elif target_plans and content_documents and all(document.file_type in {"excel", "csv", "tsv", "xliff"} for document in content_documents):
        confidence = 0.94
    elif direct_text.strip() and not content_documents:
        confidence = 0.90
    else:
        confidence = 0.72
    if not target_plans:
        confidence = 0.0
        warnings.append("没有识别到可审校译文。")
    return ExtractionPlan(
        source_language=session["source_language"],
        targets=target_plans,
        content_files=content_ids,
        reference_files=reference_ids,
        relationships=[
            {"from": reference_id, "to": content_id, "kind": "reference"}
            for reference_id in reference_ids for content_id in content_ids
        ],
        assumptions=list(dict.fromkeys(assumptions)),
        warnings=warnings,
        confidence=confidence,
        needs_input=total_units == 0,
        question="无法识别待审校译文，请说明文件关系或正文位置。" if total_units == 0 else "",
    )


def _normalize_direct_text_plan(
    plan: ExtractionPlan, session: dict[str, Any], documents: dict[str, ReaderDocument]
) -> None:
    """Prevent an Agent's placeholder `auto` from becoming a real review language.

    Direct input is always review content.  It has no source counterpart unless a
    source was explicitly supplied elsewhere, so its detected script must become
    the target language before a task is created.
    """
    target_languages = list(session.get("target_languages") or ["auto"])
    regrouped: dict[str, list[ReviewUnit]] = defaultdict(list)
    order: list[str] = []
    direct_unit_count = 0
    for target in plan.targets:
        for unit in target.units:
            if str(unit.pointer or "").startswith("direct:"):
                direct_unit_count += 1
                detected = detect_language(unit.target_text)
                language = _select_target_language(detected, target_languages)
                unit.target_language = language
                unit.source_text = ""
                unit.source_file = "直接输入"
            else:
                language = str(unit.target_language or target.language or "auto")
            if language not in regrouped:
                order.append(language)
            regrouped[language].append(unit)
    if direct_unit_count:
        plan.targets = [TargetPlan(language=language, units=regrouped[language]) for language in order]
        if not documents and str(session.get("source_language") or "auto").lower() in {"auto", "none"}:
            plan.source_language = "none"
        plan.assumptions.append("直接输入内容已按文字脚本确定为待审校译文；没有原文时仅检查译文本身。")

def _build_preset_plan(
    session: dict[str, Any],
    documents: dict[str, ReaderDocument],
    attachment_settings: dict[str, dict[str, Any]],
) -> ExtractionPlan | None:
    content_ids = [attachment_id for attachment_id, doc in documents.items() if doc.role_hint != "reference"]
    reference_ids = [attachment_id for attachment_id, doc in documents.items() if doc.role_hint == "reference"]
    if not content_ids:
        return None
    targets: dict[str, list[ReviewUnit]] = defaultdict(list)
    preset_names: list[str] = []
    preset_source_languages: list[str] = []
    for attachment_id in content_ids:
        document = documents[attachment_id]
        if document.file_type == "xliff":
            local = _build_deterministic_plan(session, {attachment_id: document}, "")
            for target in local.targets:
                targets[target.language].extend(target.units)
            continue
        attachment = attachment_settings.get(attachment_id) or {}
        if attachment.get("mapping_mode") != "preset" or document.file_type != "excel":
            return None
        try:
            preset = get_excel_mapping_preset(str(attachment.get("mapping_preset_id") or ""))
        except ValueError:
            return None
        mapping = dict(preset.get("mapping") or {})
        preset_names.append(str(preset.get("name") or "未命名模板"))
        if str(mapping.get("source_language") or "").strip():
            preset_source_languages.append(str(mapping["source_language"]).strip())
        sheet_names = [str(item.get("name") or "") for item in document.structure.get("sheets", [])]
        configured_language = str(mapping.get("target_language") or "auto")
        for sheet_index, sheet_mapping in enumerate(mapping.get("sheets") or []):
            scope = sheet_names[sheet_index] if sheet_index < len(sheet_names) else str(sheet_mapping.get("sheet_name") or "")
            for column_mapping in sheet_mapping.get("mappings") or []:
                mapped = _map_explicit_tabular_columns(
                    session,
                    attachment_id,
                    document,
                    source_column=(
                        None
                        if str(session.get("source_language") or "").lower() == "none"
                        else int(column_mapping["source_column"])
                    ),
                    target_columns=[(configured_language, int(column_mapping["target_column"]))],
                    scopes={scope} if scope else set(),
                )
                for language, units in mapped.items():
                    for unit in units:
                        unit.metadata["mapping_preset_id"] = preset["id"]
                    targets[language].extend(units)
    if not targets:
        return None
    selected_source = str(session.get("source_language") or "auto")
    if selected_source == "auto" and preset_source_languages:
        selected_source = preset_source_languages[0]
    return ExtractionPlan(
        source_language=selected_source,
        targets=[TargetPlan(language=language, units=units) for language, units in targets.items()],
        content_files=content_ids,
        reference_files=reference_ids,
        relationships=[
            {"from": reference_id, "to": content_id, "kind": "reference"}
            for reference_id in reference_ids for content_id in content_ids
        ],
        assumptions=["已按用户选择的导入映射模板读取：" + "、".join(preset_names)],
        confidence=1.0,
    )


def _map_tabular_document(
    session: dict[str, Any], attachment_id: str, document: ReaderDocument
) -> tuple[dict[str, list[ReviewUnit]], list[str]]:
    by_scope: dict[str, dict[int, dict[int, Any]]] = defaultdict(lambda: defaultdict(dict))
    labels: dict[tuple[str, int], str] = {}
    for block in document.blocks:
        scope = str(block.metadata.get("sheet") or "table")
        row = int(block.metadata.get("row") or 0)
        column = int(block.metadata.get("column_index") or 0)
        by_scope[scope][row][column] = block
        labels[(scope, column)] = str(block.metadata.get("header") or block.label or "")

    result: dict[str, list[ReviewUnit]] = defaultdict(list)
    assumptions: list[str] = []
    for scope, rows in by_scope.items():
        column_indexes = sorted({column for row in rows.values() for column in row})
        source_columns = [column for column in column_indexes if _header_role(labels.get((scope, column), "")) == "source"]
        target_columns = [column for column in column_indexes if _header_role(labels.get((scope, column), "")) == "target"]
        if not target_columns and len(column_indexes) >= 2:
            source_columns = source_columns or [column_indexes[0]]
            target_columns = [column for column in column_indexes if column not in source_columns]
            assumptions.append(f"{document.filename} / {scope} 未发现标准表头，按首列原文、其余列译文处理。")
        elif not target_columns and len(column_indexes) == 1:
            target_columns = column_indexes
            assumptions.append(f"{document.filename} / {scope} 只有一列，按无源文译文处理。")
        source_column = source_columns[0] if source_columns else None
        for target_column in target_columns:
            header = labels.get((scope, target_column), "")
            language = _language_from_label(header) or "auto"
            language = _select_target_language(language, session["target_languages"])
            for row_number, row in sorted(rows.items()):
                target_block = row.get(target_column)
                if target_block is None or not target_block.text.strip():
                    continue
                source_block = row.get(source_column) if source_column is not None else None
                result[language].append(
                    ReviewUnit(
                        id=uuid.uuid4().hex,
                        session_id=session["id"],
                        target_language=language,
                        source_text=source_block.text if source_block else "",
                        target_text=target_block.text,
                        source_file=document.filename,
                        pointer=target_block.pointer,
                        metadata={
                            "attachment_id": attachment_id,
                            "scope": scope,
                            "row": row_number,
                            "source_header": labels.get((scope, source_column), "") if source_column is not None else "",
                            "target_header": header,
                            "source_pointer": source_block.pointer if source_block else "",
                            "target_pointer": target_block.pointer,
                        },
                    )
                )
    return result, assumptions


def _apply_natural_language_adjustment(
    session: dict[str, Any], documents: dict[str, ReaderDocument], adjustment: str
) -> ExtractionPlan | None:
    if not adjustment:
        return None
    tabular_documents = {
        attachment_id: document
        for attachment_id, document in documents.items()
        if document.file_type in {"excel", "csv", "tsv"} and document.role_hint != "reference"
    }
    if not tabular_documents:
        return None
    lowered = adjustment.lower()
    named_ids = [
        attachment_id for attachment_id, document in tabular_documents.items()
        if document.filename.lower() in lowered or Path(document.filename).stem.lower() in lowered
    ]
    selected_documents = {
        attachment_id: tabular_documents[attachment_id]
        for attachment_id in (named_ids or list(tabular_documents))
    }
    source_column = (
        None
        if str(session.get("source_language") or "").strip().lower() == "none"
        else _find_adjustment_column(adjustment, ("原文", "源文", "source"))
    )

    language_labels = [
        "简体中文", "繁体中文", "英语", "英文", "日语", "韩语", "法语", "德语", "西班牙语",
        "葡萄牙语", "意大利语", "俄语", "阿拉伯语", "泰语", "越南语", "印尼语", "土耳其语",
        "波兰语", "荷兰语", "瑞典语", "挪威语", "丹麦语", "芬兰语", "捷克语", "匈牙利语",
        "罗马尼亚语", "希腊语", "希伯来语", "乌克兰语",
    ]
    target_columns: list[tuple[str, int]] = []
    for label in language_labels:
        column = _find_adjustment_column(adjustment, (label,))
        if column is not None:
            language = "英语" if label == "英文" else label
            target_columns.append((language, column))
    if not target_columns:
        generic_target = _find_adjustment_column(adjustment, ("译文", "目标文本", "target"))
        configured_targets = [value for value in session.get("target_languages") or [] if value != "auto"]
        if generic_target is not None and len(configured_targets) == 1:
            target_columns.append((configured_targets[0], generic_target))
    if not target_columns:
        return None

    targets: dict[str, list[ReviewUnit]] = defaultdict(list)
    content_files: list[str] = []
    for attachment_id, document in selected_documents.items():
        content_files.append(attachment_id)
        scopes = [
            str(item.get("name") or "") for item in (document.structure.get("sheets") or [])
            if str(item.get("name") or "") and str(item.get("name") or "").lower() in lowered
        ]
        mapped = _map_explicit_tabular_columns(
            session,
            attachment_id,
            document,
            source_column=source_column,
            target_columns=target_columns,
            scopes=set(scopes),
        )
        for language, units in mapped.items():
            targets[language].extend(units)
    target_plans = [TargetPlan(language=language, units=units) for language, units in targets.items() if units]
    if not target_plans:
        return None
    source_label = "无源文" if source_column is None else f"第 {_column_label(source_column)} 列"
    target_label = "、".join(f"{language}第 {_column_label(column)} 列" for language, column in target_columns)
    return ExtractionPlan(
        source_language=session.get("source_language") or "auto",
        targets=target_plans,
        content_files=content_files,
        reference_files=[attachment_id for attachment_id in documents if attachment_id not in content_files and documents[attachment_id].role_hint == "reference"],
        assumptions=[f"已应用用户调整：原文 {source_label}；{target_label}。"],
        confidence=0.99,
    )


def _find_adjustment_column(text: str, labels: tuple[str, ...]) -> int | None:
    label_pattern = "|".join(re.escape(label) for label in labels)
    match = re.search(
        rf"(?:{label_pattern})[^\n，,。；;]{{0,14}}?(?:在|为|是|:|：)?\s*([A-Za-z]+|\d+)\s*列",
        text,
        re.IGNORECASE,
    )
    if not match:
        return None
    token = match.group(1)
    if token.isdigit():
        return max(0, int(token) - 1)
    value = 0
    for char in token.upper():
        value = value * 26 + ord(char) - ord("A") + 1
    return value - 1


def _column_label(index: int) -> str:
    value = index + 1
    result = ""
    while value:
        value, remainder = divmod(value - 1, 26)
        result = chr(ord("A") + remainder) + result
    return result


def _map_explicit_tabular_columns(
    session: dict[str, Any],
    attachment_id: str,
    document: ReaderDocument,
    *,
    source_column: int | None,
    target_columns: list[tuple[str, int]],
    scopes: set[str],
) -> dict[str, list[ReviewUnit]]:
    rows_by_scope: dict[str, dict[int, dict[int, Any]]] = defaultdict(lambda: defaultdict(dict))
    labels: dict[tuple[str, int], str] = {}
    for block in document.blocks:
        scope = str(block.metadata.get("sheet") or "table")
        if scopes and scope not in scopes:
            continue
        row = int(block.metadata.get("row") or 0)
        column = int(block.metadata.get("column_index") or 0)
        rows_by_scope[scope][row][column] = block
        labels[(scope, column)] = str(block.metadata.get("header") or block.label or "")
    result: dict[str, list[ReviewUnit]] = defaultdict(list)
    for scope, rows in rows_by_scope.items():
        for configured_language, target_column in target_columns:
            language = _select_target_language(configured_language, session.get("target_languages") or [configured_language])
            for row_number, row in sorted(rows.items()):
                target_block = row.get(target_column)
                if target_block is None or not target_block.text.strip():
                    continue
                source_block = row.get(source_column) if source_column is not None else None
                result[language].append(
                    ReviewUnit(
                        id=uuid.uuid4().hex,
                        session_id=session["id"],
                        target_language=language,
                        source_text=source_block.text if source_block else "",
                        target_text=target_block.text,
                        source_file=document.filename,
                        pointer=target_block.pointer,
                        metadata={
                            "attachment_id": attachment_id,
                            "scope": scope,
                            "row": row_number,
                            "source_header": labels.get((scope, source_column), "") if source_column is not None else "",
                            "target_header": labels.get((scope, target_column), ""),
                            "source_column_index": source_column,
                            "target_column_index": target_column,
                            "source_pointer": source_block.pointer if source_block else "",
                            "target_pointer": target_block.pointer,
                            "adjusted_by_user": True,
                        },
                    )
                )
    return result


def _refine_with_model(
    session: dict[str, Any],
    documents: dict[str, ReaderDocument],
    *,
    direct_text: str = "",
    adjustment_text: str = "",
    run_id: str = "",
    parent_run_id: str = "",
) -> tuple[ExtractionPlan | None, str]:
    settings = get_shared_ai_settings()
    if not str(settings.get("api_key") or "").strip() or not str(settings.get("selected_model") or "").strip():
        return None, ""
    analysis_documents = dict(documents)
    if direct_text.strip():
        direct_blocks = [
            ReaderBlock(pointer=f"direct:{index}", text=value)
            for index, value in enumerate(_split_direct_text(direct_text), 1)
        ]
        analysis_documents["__direct_text__"] = ReaderDocument(
            reader_name="direct_text",
            file_type="direct_text",
            filename="直接输入",
            file_hash=hashlib.sha256(direct_text.encode("utf-8")).hexdigest(),
            structure={"paragraph_count": len(direct_blocks)},
            blocks=direct_blocks,
        )
    manifests = {
        attachment_id: document.to_manifest(sample_limit=12)
        for attachment_id, document in analysis_documents.items()
    }
    user_prompt = WORKSPACE_USER_PROMPT.format(
        source_language=session["source_language"],
        target_languages=json.dumps(session["target_languages"], ensure_ascii=False),
        adjustment_text=adjustment_text or "无",
        adjustment_context_json=json.dumps(
            _workspace_adjustment_context(session["id"], parent_run_id) if adjustment_text.strip() else {},
            ensure_ascii=False,
        ),
        manifest_json=json.dumps(manifests, ensure_ascii=False),
    )
    try:
        pending_chunks: list[str] = []
        thinking_chunks: list[str] = []
        pending_phase = "content"
        received_chars = 0
        last_emit = time.monotonic()
        emitted_any = False

        def flush_pending() -> None:
            nonlocal last_emit, emitted_any
            if not pending_chunks:
                return
            emit_event(
                session["id"],
                "workspace.delta",
                {
                    "run_id": run_id,
                    "delta": "".join(pending_chunks),
                    "phase": pending_phase,
                    "received_chars": received_chars,
                },
            )
            pending_chunks.clear()
            last_emit = time.monotonic()
            emitted_any = True

        def emit_delta(delta: str, phase: str = "content") -> None:
            nonlocal received_chars, pending_phase
            value = str(delta or "")
            if phase == "reasoning" and value:
                thinking_chunks.append(value)
            if pending_chunks and phase != pending_phase:
                flush_pending()
            pending_phase = phase
            if value:
                pending_chunks.append(value)
                received_chars += len(value)
            now = time.monotonic()
            if pending_chunks and (not emitted_any or sum(map(len, pending_chunks)) >= 64 or now - last_emit >= 0.12):
                flush_pending()

        emit_event(session["id"], "workspace.output_started", {"run_id": run_id})
        raw = workspace_chat(
            [{"role": "system", "content": WORKSPACE_SYSTEM_PROMPT}, {"role": "user", "content": user_prompt}],
            on_delta=emit_delta,
        )
        flush_pending()
        parsed = _parse_json(raw)
        plan = _plan_from_agent(session["id"], parsed, analysis_documents)
        if str(session.get("source_language") or "").lower() == "none":
            plan.source_language = "none"
            for target in plan.targets:
                for unit in target.units:
                    unit.source_text = ""
                    unit.metadata["source_pointer"] = ""
                    unit.metadata["source_column_index"] = None
        return plan, "".join(thinking_chunks)
    except (SharedProviderError, ValueError, json.JSONDecodeError):
        return None, "".join(thinking_chunks) if 'thinking_chunks' in locals() else ""


def _workspace_adjustment_context(session_id: str, parent_run_id: str) -> dict[str, Any]:
    if not parent_run_id:
        return {}
    with get_connection() as conn:
        row = conn.execute(
            "SELECT plan_json FROM workspace_runs WHERE id = ? AND session_id = ?",
            (parent_run_id, session_id),
        ).fetchone()
    return loads_json(row["plan_json"], {}) if row else {}


def _plan_from_agent(
    session_id: str, data: dict[str, Any], documents: dict[str, ReaderDocument]
) -> ExtractionPlan:
    agent_declares_no_source = str(data.get("source_language") or "").strip().lower() == "none"
    block_maps = {
        attachment_id: {block.pointer: block for block in document.blocks}
        for attachment_id, document in documents.items()
    }
    targets: list[TargetPlan] = []
    for raw_target in data.get("targets", []):
        language = str(raw_target.get("language") or "auto")
        units: list[ReviewUnit] = []
        for mapping in raw_target.get("mappings", []):
            attachment_id = str(mapping.get("attachment_id") or "")
            document = documents.get(attachment_id)
            if document is None:
                continue
            if document.file_type in {"excel", "csv", "tsv"} and mapping.get("target_column_index") is not None:
                try:
                    target_column = int(mapping["target_column_index"])
                    source_column = (
                        int(mapping["source_column_index"])
                        if mapping.get("source_column_index") is not None and str(mapping.get("source_column_index")) != ""
                        else None
                    )
                except (TypeError, ValueError):
                    continue
                if agent_declares_no_source:
                    source_column = None
                scope = str(mapping.get("scope") or "")
                mapped = _map_explicit_tabular_columns(
                    {"id": session_id, "target_languages": [language]},
                    attachment_id,
                    document,
                    source_column=source_column,
                    target_columns=[(language, target_column)],
                    scopes={scope} if scope else set(),
                )
                units.extend(mapped.get(language) or [])
                continue
            target_pointer = str(mapping.get("target_pointer") or "")
            source_pointer = "" if agent_declares_no_source else str(mapping.get("source_pointer") or "")
            source_block = block_maps.get(attachment_id, {}).get(source_pointer)
            target_blocks = _resolve_non_tabular_target_blocks(
                document,
                target_pointer=target_pointer,
                scope=str(mapping.get("scope") or ""),
            )
            for target_block in target_blocks:
                units.append(
                    ReviewUnit(
                        id=uuid.uuid4().hex,
                        session_id=session_id,
                        target_language=language,
                        source_text=(
                            ""
                            if agent_declares_no_source
                            else source_block.text if source_block else str(target_block.metadata.get("source_text") or "")
                        ),
                        target_text=target_block.text,
                        source_file=document.filename,
                        pointer=target_block.pointer,
                        metadata={
                            "attachment_id": attachment_id,
                            "scope": str(mapping.get("scope") or ""),
                            "source_pointer": source_pointer,
                            "target_pointer": target_block.pointer,
                        },
                    )
                )
        if units:
            targets.append(TargetPlan(language=language, units=units))
    return ExtractionPlan(
        source_language=str(data.get("source_language") or "auto"),
        targets=targets,
        content_files=[value for value in data.get("content_files", []) if value in documents],
        reference_files=[value for value in data.get("reference_files", []) if value in documents],
        relationships=[item for item in data.get("relationships", []) if isinstance(item, dict)],
        assumptions=[str(value) for value in data.get("assumptions", [])],
        warnings=[str(value) for value in data.get("warnings", [])],
        confidence=max(0.0, min(float(data.get("confidence") or 0.0), 1.0)),
        needs_input=bool(data.get("needs_input")) or not targets,
        question=str(data.get("question") or ""),
    )


def _resolve_non_tabular_target_blocks(
    document: ReaderDocument, *, target_pointer: str, scope: str = ""
) -> list[ReaderBlock]:
    """Resolve exact, document-prefix and explicit whole-document mappings safely."""
    pointer = str(target_pointer or "").strip()
    normalized_scope = str(scope or "").strip().lower()
    if pointer in {"*", "all", "全文"} or (not pointer and normalized_scope in {"document", "docx", "pptx", "text", "all"}):
        return list(document.blocks)
    exact = [block for block in document.blocks if block.pointer == pointer]
    if exact:
        return exact
    prefix = pointer.rstrip("/")
    if prefix:
        return [block for block in document.blocks if block.pointer.startswith(prefix + "/")]
    return []


def _attach_references(plan: ExtractionPlan, documents: dict[str, ReaderDocument]) -> None:
    reference_blocks = [
        {"category": doc.filename, "value": block.text, "pointer": block.pointer}
        for attachment_id, doc in documents.items()
        if attachment_id in plan.reference_files
        for block in doc.blocks
    ]
    if not reference_blocks:
        return
    for target in plan.targets:
        for unit in target.units:
            unit_terms = _search_terms(unit.source_text + " " + unit.target_text)
            scored: list[tuple[int, dict[str, str]]] = []
            for reference in reference_blocks:
                score = len(unit_terms & _search_terms(reference["value"]))
                if score:
                    scored.append((score, reference))
            unit.references = [item for _, item in sorted(scored, key=lambda pair: pair[0], reverse=True)[:5]]


def _get_cached_profile(
    run_id: str, documents: dict[str, ReaderDocument], session: dict[str, Any]
) -> ExtractionPlan | None:
    if not documents or any(
        document.file_type not in {"excel", "csv", "tsv"}
        for document in documents.values() if document.role_hint != "reference"
    ):
        return None
    signature = _structure_signature(
        documents,
        str(session.get("source_language") or "auto"),
        list(session.get("target_languages") or ["auto"]),
    )
    with get_connection() as conn:
        row = conn.execute("SELECT plan_json FROM extraction_profiles WHERE signature = ?", (signature,)).fetchone()
        run = conn.execute("SELECT session_id FROM workspace_runs WHERE id = ?", (run_id,)).fetchone()
    if not row or not run:
        return None
    profile = loads_json(row["plan_json"], {})
    if profile.get("profile_version") != 2:
        return None
    attachment_ids = list(documents)
    targets: dict[str, list[ReviewUnit]] = defaultdict(list)
    try:
        for mapping in profile.get("mappings", []):
            attachment_index = int(mapping.get("attachment_index", -1))
            if attachment_index < 0 or attachment_index >= len(attachment_ids):
                return None
            attachment_id = attachment_ids[attachment_index]
            language = str(mapping.get("language") or "auto")
            mapped = _map_explicit_tabular_columns(
                {"id": str(run["session_id"]), "target_languages": [language]},
                attachment_id,
                documents[attachment_id],
                source_column=(
                    int(mapping["source_column_index"])
                    if mapping.get("source_column_index") is not None
                    else None
                ),
                target_columns=[(language, int(mapping["target_column_index"]))],
                scopes={str(mapping.get("scope") or "")} if mapping.get("scope") else set(),
            )
            targets[language].extend(mapped.get(language) or [])
    except (TypeError, ValueError, KeyError):
        return None
    if not targets:
        return None
    content_files = [
        attachment_ids[index]
        for index in profile.get("content_indices", [])
        if isinstance(index, int) and 0 <= index < len(attachment_ids)
    ]
    reference_files = [
        attachment_ids[index]
        for index in profile.get("reference_indices", [])
        if isinstance(index, int) and 0 <= index < len(attachment_ids)
    ]
    return ExtractionPlan(
        source_language=str(profile.get("source_language") or session.get("source_language") or "auto"),
        targets=[TargetPlan(language=language, units=units) for language, units in targets.items()],
        content_files=content_files,
        reference_files=reference_files,
        relationships=[
            {"from": reference_id, "to": content_id, "kind": "reference"}
            for reference_id in reference_files for content_id in content_files
        ],
        assumptions=["语言选择和表头结构完全一致，已读取确认过的本地结构缓存。"],
        confidence=0.99,
    )


def save_confirmed_profile(run_id: str) -> None:
    with get_connection() as conn:
        run = conn.execute("SELECT * FROM workspace_runs WHERE id = ?", (run_id,)).fetchone()
        if not run:
            raise ValueError("Workspace 运行不存在")
        attachment_ids = loads_json(run["attachment_ids_json"], []) if "attachment_ids_json" in run.keys() else []
        if not attachment_ids:
            return
        placeholders = ",".join("?" for _ in attachment_ids)
        attachments = conn.execute(
            f"SELECT id, manifest_json FROM review_attachments WHERE session_id = ? AND status = 'ready' "
            f"AND id IN ({placeholders}) ORDER BY created_at, rowid",
            (run["session_id"], *attachment_ids),
        ).fetchall()
        manifests = [loads_json(row["manifest_json"], {}) for row in attachments]
        if not manifests or any(
            item.get("file_type") not in {"excel", "csv", "tsv"}
            for item in manifests if item.get("role_hint") != "reference"
        ):
            return
        attachment_indexes = {str(row["id"]): index for index, row in enumerate(attachments)}
        raw_plan = loads_json(run["plan_json"], {})
        session_row = conn.execute(
            "SELECT source_language, target_languages_json FROM review_sessions WHERE id = ?",
            (run["session_id"],),
        ).fetchone()
        if not session_row:
            return
        profile_plan = {
            "profile_version": 2,
            "source_language": raw_plan.get("source_language", "auto"),
            "content_indices": [attachment_indexes[value] for value in raw_plan.get("content_files", []) if value in attachment_indexes],
            "reference_indices": [attachment_indexes[value] for value in raw_plan.get("reference_files", []) if value in attachment_indexes],
            "mappings": [],
        }
        seen_mappings: set[tuple[Any, ...]] = set()
        for target in raw_plan.get("targets", []):
            for unit in target.get("units", []):
                metadata = dict(unit.get("metadata") or {})
                attachment_id = str(metadata.get("attachment_id") or "")
                if attachment_id not in attachment_indexes or metadata.get("target_column_index") is None:
                    continue
                key = (
                    attachment_indexes[attachment_id],
                    str(target.get("language") or "auto"),
                    str(metadata.get("scope") or ""),
                    metadata.get("source_column_index"),
                    metadata.get("target_column_index"),
                )
                if key in seen_mappings:
                    continue
                seen_mappings.add(key)
                profile_plan["mappings"].append(
                    {
                        "attachment_index": key[0],
                        "language": key[1],
                        "scope": key[2],
                        "source_column_index": key[3],
                        "target_column_index": key[4],
                    }
                )
        if not profile_plan["mappings"]:
            return
        signature = hashlib.sha256(
            dumps_json(
                {
                    "source_language": str(session_row["source_language"] or "auto"),
                    "target_languages": loads_json(session_row["target_languages_json"], ["auto"]),
                    "documents": [_manifest_shape(item) for item in manifests],
                }
            ).encode("utf-8")
        ).hexdigest()
        now = utc_now()
        conn.execute(
            """
            INSERT INTO extraction_profiles (signature, reader_name, plan_json, confirmed_count, created_at, updated_at)
            VALUES (?, 'workspace', ?, 1, ?, ?)
            ON CONFLICT(signature) DO UPDATE SET
                plan_json = excluded.plan_json,
                confirmed_count = extraction_profiles.confirmed_count + 1,
                updated_at = excluded.updated_at
            """,
            (signature, dumps_json(profile_plan), now, now),
        )


def _input_signature(
    snapshot: dict[str, Any],
    direct_text: str,
    adjustment_text: str = "",
    attachment_ids: list[str] | None = None,
) -> str:
    allowed = set(attachment_ids) if attachment_ids is not None else None
    raw = {
        "attachments": [
            item["file_hash"] for item in snapshot["attachments"]
            if allowed is None or str(item["id"]) in allowed
        ],
        "text": hashlib.sha256(direct_text.encode("utf-8")).hexdigest() if direct_text else "",
        "adjustment": hashlib.sha256(adjustment_text.encode("utf-8")).hexdigest() if adjustment_text else "",
        "source": snapshot["session"]["source_language"],
        "targets": snapshot["session"]["target_languages"],
    }
    return hashlib.sha256(dumps_json(raw).encode("utf-8")).hexdigest()


def _structure_signature(
    documents: dict[str, ReaderDocument], source_language: str, target_languages: list[str]
) -> str:
    shapes = [_manifest_shape(document.to_manifest(sample_limit=0)) for document in documents.values()]
    return hashlib.sha256(
        dumps_json(
            {
                "source_language": source_language,
                "target_languages": target_languages,
                "documents": shapes,
            }
        ).encode("utf-8")
    ).hexdigest()


def _manifest_shape(manifest: dict[str, Any]) -> dict[str, Any]:
    structure = dict(manifest.get("structure") or {})
    if manifest.get("file_type") == "excel":
        header_structure = {
            "sheets": [
                {
                    "name": sheet.get("name"),
                    "columns": [
                        {"index": column.get("index"), "header": column.get("header")}
                        for column in sheet.get("columns", [])
                    ],
                }
                for sheet in structure.get("sheets", [])
            ]
        }
    elif manifest.get("file_type") in {"csv", "tsv"}:
        header_structure = {"headers": list(structure.get("headers") or [])}
    else:
        header_structure = structure
    return {
        "reader_name": manifest.get("reader_name"),
        "file_type": manifest.get("file_type"),
        "role_hint": manifest.get("role_hint"),
        "structure": header_structure,
    }


def _header_role(label: str) -> str:
    normalized = re.sub(r"[\s_\-]+", "", str(label or "").lower())
    source_markers = ("source", "original", "src", "原文", "源文", "待翻译", "英文")
    target_markers = ("target", "translation", "translated", "tgt", "译文", "翻译", "本地化")
    if any(marker in normalized for marker in source_markers):
        return "source"
    if any(marker in normalized for marker in target_markers) or _language_from_label(label):
        return "target"
    return "unknown"


def _language_from_label(label: str) -> str:
    normalized = str(label or "").lower()
    mappings = {
        "简体中文": ("zh-cn", "zh_cn", "zh-hans", "简中", "简体"),
        "繁体中文": ("zh-tw", "zh_tw", "zh-hant", "繁中", "繁体"),
        "英语": ("english", "en-us", "en_us", "en-gb", "en_gb", "英语", "英文"),
        "日语": ("japanese", "ja-jp", "ja_jp", "日语", "日文", "日本語"),
        "韩语": ("korean", "ko-kr", "ko_kr", "韩语", "韩文", "한국어"),
        "法语": ("french", "fr-fr", "fr_fr", "法语"),
        "德语": ("german", "de-de", "de_de", "德语"),
        "西班牙语": ("spanish", "es-es", "es_es", "西班牙语"),
        "俄语": ("russian", "ru-ru", "ru_ru", "俄语"),
    }
    for language, markers in mappings.items():
        if any(marker in normalized for marker in markers):
            return language
    return ""


def detect_language(text: str) -> str:
    counts = {
        "日语": len(re.findall(r"[\u3040-\u30ff]", text)),
        "韩语": len(re.findall(r"[\uac00-\ud7af]", text)),
        "中文": len(re.findall(r"[\u4e00-\u9fff]", text)),
        "俄语": len(re.findall(r"[\u0400-\u04ff]", text)),
        "英语": len(re.findall(r"[A-Za-z]", text)),
    }
    language, count = max(counts.items(), key=lambda item: item[1])
    return language if count else "auto"


def _select_target_language(detected: str, selected: list[str]) -> str:
    chosen = [value for value in selected if value != "auto"]
    if not chosen:
        return detected or "auto"
    normalized = str(detected or "").lower()
    for value in chosen:
        if value.lower() in normalized or normalized in value.lower():
            return value
    return chosen[0]


def _split_direct_text(text: str) -> list[str]:
    return [part.strip() for part in re.split(r"\n\s*\n|\n", text) if part.strip()]


def _search_terms(text: str) -> set[str]:
    words = set(re.findall(r"[A-Za-z][A-Za-z0-9_\-]{2,}|[\u4e00-\u9fff]{2,}", text.lower()))
    return {word for word in words if len(word) <= 80}


def _parse_json(content: str) -> dict[str, Any]:
    text = content.strip()
    start = text.find("{")
    end = text.rfind("}")
    if start < 0 or end < start:
        raise ValueError("Workspace Agent 未返回 JSON")
    result = json.loads(text[start : end + 1])
    if not isinstance(result, dict):
        raise ValueError("Workspace Agent 返回结构无效")
    return result


def _format_report(plan: dict[str, Any]) -> str:
    target_lines = [
        f"{item.get('language') or '自动识别'}：{len(item.get('units') or [])} 条"
        for item in plan.get("targets", [])
    ]
    lines = ["已完成文件识别。"]
    if target_lines:
        lines.append("审校任务：" + "；".join(target_lines) + "。")
    for file_index, file_info in enumerate(plan.get("file_summaries") or [], 1):
        filename = str(file_info.get("filename") or "未命名文件")
        role = "参考资料" if file_info.get("role") == "reference" else "正文"
        lines.append("")
        lines.append(f"文件 {file_index}｜{filename}｜{role}")
        mappings = file_info.get("mappings") or []
        if not mappings:
            lines.append(f"  位置：{file_info.get('structure_label') or '未形成正文映射'}")
            continue
        for mapping in mappings:
            lines.append(
                f"  {mapping.get('language') or '自动识别'} · {mapping.get('count') or 0} 条"
                f"｜原文：{mapping.get('source_location') or '无源文'}"
                f"｜译文：{mapping.get('target_location') or '未识别'}"
            )
            for sample_index, sample in enumerate(mapping.get("samples") or [], 1):
                source = str(sample.get("source") or "")
                target = str(sample.get("target") or "")
                position = str(sample.get("position") or "")
                if source:
                    lines.append(f"    样例 {sample_index} [{position}] 原文：{source}")
                    lines.append(f"                         译文：{target}")
                else:
                    lines.append(f"    样例 {sample_index} [{position}] 译文：{target}")
    direct_summary = _build_direct_summary(plan)
    if direct_summary:
        lines.extend(["", *direct_summary])
    if plan.get("assumptions"):
        lines.extend(["", "识别依据：" + "；".join(str(value) for value in plan["assumptions"][:4])])
    if plan.get("warnings"):
        lines.append("注意：" + "；".join(str(value) for value in plan["warnings"][:4]))
    return "\n".join(lines)


def _build_file_summaries(
    plan: dict[str, Any], documents: dict[str, ReaderDocument]
) -> list[dict[str, Any]]:
    summaries: list[dict[str, Any]] = []
    reference_ids = set(plan.get("reference_files") or [])
    targets = plan.get("targets") or []
    for attachment_id, document in documents.items():
        mappings: list[dict[str, Any]] = []
        for target in targets:
            units = [
                unit for unit in target.get("units") or []
                if str((unit.get("metadata") or {}).get("attachment_id") or "") == attachment_id
                or str(unit.get("source_file") or "") == document.filename
            ]
            if not units:
                continue
            source_pointers = [str((unit.get("metadata") or {}).get("source_pointer") or "") for unit in units]
            target_pointers = [
                str((unit.get("metadata") or {}).get("target_pointer") or unit.get("pointer") or "")
                for unit in units
            ]
            source_header = next(
                (str((unit.get("metadata") or {}).get("source_header") or "") for unit in units
                 if (unit.get("metadata") or {}).get("source_header")),
                "",
            )
            target_header = next(
                (str((unit.get("metadata") or {}).get("target_header") or "") for unit in units
                 if (unit.get("metadata") or {}).get("target_header")),
                "",
            )
            samples = []
            for unit in units[:3]:
                metadata = unit.get("metadata") or {}
                source_pointer = str(metadata.get("source_pointer") or "")
                target_pointer = str(metadata.get("target_pointer") or unit.get("pointer") or "")
                samples.append(
                    {
                        "position": _sample_position(source_pointer, target_pointer),
                        "source": _short_sample(str(unit.get("source_text") or "")),
                        "target": _short_sample(str(unit.get("target_text") or "")),
                    }
                )
            mappings.append(
                {
                    "language": target.get("language") or "auto",
                    "count": len(units),
                    "source_location": _compact_locations(source_pointers, source_header),
                    "target_location": _compact_locations(target_pointers, target_header),
                    "samples": samples,
                }
            )
        role = "reference" if attachment_id in reference_ids or document.role_hint == "reference" else "content"
        summaries.append(
            {
                "attachment_id": attachment_id,
                "filename": document.filename,
                "role": role,
                "file_type": document.file_type,
                "structure_label": _document_structure_label(document),
                "mappings": mappings,
            }
        )
    return summaries


def _build_direct_summary(plan: dict[str, Any]) -> list[str]:
    lines: list[str] = []
    for target in plan.get("targets") or []:
        units = [unit for unit in target.get("units") or [] if str(unit.get("pointer") or "").startswith("direct:")]
        if not units:
            continue
        lines.append(f"直接输入｜{target.get('language') or '自动识别'} · {len(units)} 条｜无源文")
        for index, unit in enumerate(units[:3], 1):
            lines.append(
                f"  样例 {index} [{unit.get('pointer') or ''}] 译文：{_short_sample(str(unit.get('target_text') or ''))}"
            )
    return lines


def _compact_locations(pointers: list[str], header: str = "") -> str:
    values = [value for value in pointers if value]
    if not values:
        return ""
    excel_cells = [re.fullmatch(r"sheet:(.+?)/cell:([A-Z]+)(\d+)", value) for value in values]
    if all(excel_cells):
        matches = [match for match in excel_cells if match]
        sheets = {match.group(1) for match in matches}
        columns = {match.group(2) for match in matches}
        rows = [int(match.group(3)) for match in matches]
        if len(sheets) == 1 and len(columns) == 1:
            label = f'工作表“{next(iter(sheets))}” · {next(iter(columns))}列'
            if header:
                label += f'（{header}）'
            return f"{label} · 第 {min(rows)}–{max(rows)} 行"
    csv_cells = [re.fullmatch(r"row:(\d+)/column:(\d+)", value) for value in values]
    if all(csv_cells):
        matches = [match for match in csv_cells if match]
        columns = {int(match.group(2)) for match in matches}
        rows = [int(match.group(1)) for match in matches]
        if len(columns) == 1:
            label = f"第 {next(iter(columns))} 列"
            if header:
                label += f"（{header}）"
            return f"{label} · 第 {min(rows)}–{max(rows)} 行"
    if len(values) == 1:
        return values[0]
    return f"{values[0]} → {values[-1]}"


def _sample_position(source_pointer: str, target_pointer: str) -> str:
    if source_pointer and target_pointer:
        return f"{source_pointer} → {target_pointer}"
    return target_pointer or source_pointer or "位置未知"


def _short_sample(text: str, limit: int = 90) -> str:
    compact = " ".join(text.split())
    return compact if len(compact) <= limit else compact[: limit - 1] + "…"


def _document_structure_label(document: ReaderDocument) -> str:
    structure = document.structure or {}
    if document.file_type == "excel":
        sheets = structure.get("sheets") or []
        return "工作表：" + "、".join(str(item.get("name") or "") for item in sheets[:6])
    if document.file_type == "pdf":
        return f"PDF · {structure.get('page_count') or 0} 页"
    return f"{document.file_type.upper()} · {len(document.blocks)} 个文本位置"
