from __future__ import annotations

import tempfile
import json
import time
import threading
import unittest
from io import BytesIO
from pathlib import Path
from unittest.mock import patch

from fastapi import FastAPI
from fastapi.testclient import TestClient
from openpyxl import Workbook

from term_extractor_app.ai_review import cache_service, config, database, session_store, workspace_service
from term_extractor_app.ai_review.conversation_routes import router


workspace_prompts: list[str] = []


def _workspace_response(messages: list[dict[str, str]], on_delta=None) -> str:
    prompt = messages[-1]["content"]
    workspace_prompts.append(prompt)
    if "用户直接输入（其中可能混有操作说明；只提取真正待审校的正文）" in prompt:
        direct_text = prompt.split("\n---\n", 1)[1].rsplit("\n---", 1)[0].strip()
        cleaned = direct_text.split("：", 1)[-1].strip() if "：" in direct_text else direct_text
        response = json.dumps({
            "source_language": "none",
            "targets": [{"language": "英语", "mappings": [{
                "attachment_id": "__direct_text__", "target_pointer": "direct:1", "target_text": cleaned,
            }]}],
            "assumptions": ["已分离操作说明与正文。"], "warnings": [], "confidence": 0.97,
            "needs_input": False, "question": "",
        }, ensure_ascii=False)
        if on_delta:
            on_delta("正在整理直接输入。", "reasoning")
            on_delta(response, "content")
        return response
    manifests = json.loads(prompt.split("结构清单：\n", 1)[1])
    attachment_id = next(iter(manifests))
    if "没有原文" in prompt:
        targets = [{"language": "简体中文", "mappings": [{"attachment_id": attachment_id, "scope": "table", "source_column_index": None, "target_column_index": 1}]}]
    else:
        targets = [
            {"language": "简体中文", "mappings": [{"attachment_id": attachment_id, "scope": "table", "source_column_index": 0, "target_column_index": 1}]},
            {"language": "日语", "mappings": [{"attachment_id": attachment_id, "scope": "table", "source_column_index": 0, "target_column_index": 2}]},
        ]
    response = json.dumps({
        "source_language": "英语", "targets": targets, "assumptions": ["Workspace Agent 已确认映射。"],
        "warnings": [], "confidence": 0.97, "needs_input": False, "question": "",
    }, ensure_ascii=False)
    if on_delta:
        on_delta("正在判断文件与列之间的关系。", "reasoning")
        on_delta(response, "content")
    return response


class ConversationApiTests(unittest.TestCase):
    def setUp(self) -> None:
        self.initial_threads = set(threading.enumerate())
        workspace_prompts.clear()
        self.temp = tempfile.TemporaryDirectory()
        root = Path(self.temp.name)
        self.db_path = root / "data" / "app.sqlite3"
        self.db_path.parent.mkdir(parents=True)
        self.uploads = root / "uploads"
        self.uploads.mkdir()
        self.patches = [
            patch.object(database, "DB_PATH", self.db_path),
            patch.object(config, "DB_PATH", self.db_path),
            patch.object(cache_service, "UPLOADS_DIR", self.uploads),
            patch.object(session_store, "UPLOADS_DIR", self.uploads),
            patch.object(workspace_service, "get_shared_ai_settings", return_value={"selected_model": "configured-model", "api_key": "test-key"}),
            patch.object(workspace_service, "workspace_chat", side_effect=_workspace_response),
        ]
        for item in self.patches:
            item.start()
        database.init_db()
        app = FastAPI()
        app.include_router(router)
        self.client = TestClient(app)

    def tearDown(self) -> None:
        self.client.close()
        for thread in set(threading.enumerate()) - self.initial_threads:
            if getattr(thread, '_target', None) is workspace_service._run_inspection:
                thread.join(timeout=10)
                self.assertFalse(thread.is_alive(), 'Workspace worker must stop before removing its test database')
        for item in reversed(self.patches):
            item.stop()
        self.temp.cleanup()

    def test_create_upload_inspect_and_delete_conversation(self) -> None:
        created = self.client.post(
            "/api/ai-review/conversations",
            json={"target_languages": ["简体中文", "日语"]},
        )
        self.assertEqual(created.status_code, 200)
        session_id = created.json()["session"]["id"]
        upload = self.client.post(
            f"/api/ai-review/conversations/{session_id}/attachments",
            files={"files": ("translations.csv", "Source,简体中文,日本語\nHello,你好,こんにちは\n", "text/csv")},
        )
        self.assertEqual(upload.status_code, 200)
        uploaded_id = upload.json()["attachments"][0]["id"]
        self.assertEqual(
            self.client.get(f"/api/ai-review/conversations/{session_id}").json()["session"]["title"],
            "translations.csv",
        )
        renamed = self.client.patch(
            f"/api/ai-review/conversations/{session_id}", json={"title": "每日测验审校"}
        )
        self.assertEqual(renamed.status_code, 200)
        self.assertTrue(renamed.json()["session"]["title_custom"])
        sent = self.client.post(
            f"/api/ai-review/conversations/{session_id}/messages",
            json={
                "text": "",
                "source_language": "英语",
                "target_languages": ["简体中文", "日语"],
                "auto_start": False,
            },
        )
        self.assertEqual(sent.status_code, 200)
        snapshot = None
        for _ in range(40):
            response = self.client.get(f"/api/ai-review/conversations/{session_id}")
            snapshot = response.json()
            if snapshot["session"]["status"] in {"ready", "needs_input", "failed"}:
                break
            time.sleep(0.05)
        self.assertEqual(snapshot["session"]["status"], "ready")
        self.assertEqual(snapshot["session"]["title"], "每日测验审校")
        self.assertIsNotNone(snapshot["attachments"][0]["sent_at"])
        attachment_message = next(item for item in snapshot["messages"] if item["kind"] == "attachment_message")
        self.assertEqual(attachment_message["payload"]["attachments"][0]["id"], uploaded_id)
        event_types = {item["event_type"] for item in session_store.get_events(session_id)}
        self.assertIn("workspace.output_started", event_types)
        self.assertIn("workspace.delta", event_types)
        workspace_message = next(item for item in snapshot["messages"] if item["kind"] == "workspace_report")
        self.assertEqual(workspace_message["payload"]["_thinking_text"], "正在判断文件与列之间的关系。")
        self.assertNotIn("_thinking_text", snapshot["workspace_runs"][0]["plan"])
        plan = snapshot["workspace_runs"][0]["plan"]
        self.assertEqual({item["language"] for item in plan["targets"]}, {"简体中文", "日语"})
        self.assertEqual(len(snapshot["questions"]), 1)
        self.assertEqual(snapshot["questions"][0]["recommended_label"], "确认推荐方案")
        adjusted = self.client.post(
            f"/api/ai-review/conversations/{session_id}/decision",
            json={
                "run_id": snapshot["workspace_runs"][0]["id"],
                "question_id": snapshot["questions"][0]["id"],
                "action": "other",
                "answer": "translations.csv 没有原文，简体中文译文在 B 列",
            },
        )
        self.assertEqual(adjusted.status_code, 200)
        for _ in range(40):
            snapshot = self.client.get(f"/api/ai-review/conversations/{session_id}").json()
            if (
                len(snapshot["workspace_runs"]) >= 2
                and snapshot["workspace_runs"][0]["status"] == "ready"
                and any(item["status"] == "pending" for item in snapshot["questions"])
            ):
                break
            time.sleep(0.05)
        adjusted_plan = snapshot["workspace_runs"][0]["plan"]
        self.assertEqual([item["language"] for item in adjusted_plan["targets"]], ["简体中文"])
        self.assertEqual(adjusted_plan["targets"][0]["units"][0]["target_text"], "你好")
        self.assertEqual(adjusted_plan["targets"][0]["units"][0]["source_text"], "")
        pending_questions = [item for item in snapshot["questions"] if item["status"] == "pending"]
        superseded_questions = [item for item in snapshot["questions"] if item["status"] == "superseded"]
        self.assertEqual(len(pending_questions), 1)
        self.assertEqual(len(superseded_questions), 0)
        delete = self.client.request(
            "DELETE",
            f"/api/ai-review/conversations/{session_id}",
            json={"confirm": True},
        )
        self.assertEqual(delete.status_code, 200)
        self.assertEqual(self.client.get(f"/api/ai-review/conversations/{session_id}").status_code, 404)

    def test_local_attachment_keeps_a_private_review_copy_without_source_path(self) -> None:
        selected_path = Path(self.temp.name) / "selected.csv"
        selected_path.write_text("Source,Target\nHello,你好\n", encoding="utf-8")
        session_id = self.client.post("/api/ai-review/conversations", json={}).json()["session"]["id"]
        attached = self.client.post(
            f"/api/ai-review/conversations/{session_id}/attachments/local",
            json={"file_path": str(selected_path)},
        )
        self.assertEqual(attached.status_code, 200)
        attachment = attached.json()["attachment"]
        self.assertNotIn("original_path", attachment)
        self.assertTrue(Path(attachment["stored_path"]).is_file())

    def test_delete_requires_confirmation(self) -> None:
        session_id = self.client.post("/api/ai-review/conversations", json={}).json()["session"]["id"]
        response = self.client.request(
            "DELETE",
            f"/api/ai-review/conversations/{session_id}",
            json={"confirm": False},
        )
        self.assertEqual(response.status_code, 400)

    def test_auto_start_is_saved_per_session_immediately(self) -> None:
        first_session = self.client.post("/api/ai-review/conversations", json={}).json()["session"]["id"]
        second_session = self.client.post("/api/ai-review/conversations", json={}).json()["session"]["id"]
        updated = self.client.patch(
            f"/api/ai-review/conversations/{first_session}",
            json={"auto_start": True},
        )
        self.assertEqual(updated.status_code, 200)
        self.assertTrue(updated.json()["session"]["auto_start"])
        first_snapshot = self.client.get(f"/api/ai-review/conversations/{first_session}").json()
        second_snapshot = self.client.get(f"/api/ai-review/conversations/{second_session}").json()
        self.assertTrue(first_snapshot["session"]["auto_start"])
        self.assertFalse(second_snapshot["session"]["auto_start"])

    def test_composer_settings_are_saved_per_session_immediately(self) -> None:
        first_session = self.client.post("/api/ai-review/conversations", json={}).json()["session"]["id"]
        second_session = self.client.post("/api/ai-review/conversations", json={}).json()["session"]["id"]
        updated = self.client.patch(
            f"/api/ai-review/conversations/{first_session}",
            json={
                "prompt_template_id": "template-session-one",
                "source_language": "日语",
                "target_languages": ["简体中文", "英语"],
            },
        )
        self.assertEqual(updated.status_code, 200)
        first_snapshot = self.client.get(f"/api/ai-review/conversations/{first_session}").json()
        second_snapshot = self.client.get(f"/api/ai-review/conversations/{second_session}").json()
        self.assertEqual(first_snapshot["session"]["prompt_template_id"], "template-session-one")
        self.assertEqual(first_snapshot["session"]["source_language"], "日语")
        self.assertEqual(first_snapshot["session"]["target_languages"], ["简体中文", "英语"])
        self.assertIsNone(second_snapshot["session"]["prompt_template_id"])
        self.assertEqual(second_snapshot["session"]["source_language"], "auto")
        self.assertEqual(second_snapshot["session"]["target_languages"], ["auto"])

    def test_starting_new_workspace_run_supersedes_old_pending_question(self) -> None:
        session_id = self.client.post("/api/ai-review/conversations", json={}).json()["session"]["id"]
        first_run = session_store.create_workspace_run(
            session_id, model="configured-model", input_signature="first"
        )
        first_question = session_store.create_question(
            session_id, first_run["id"], "第一次确认", "确认推荐方案"
        )
        second_run = session_store.create_workspace_run(
            session_id, model="configured-model", input_signature="second"
        )
        session_store.create_question(session_id, second_run["id"], "第二次确认", "确认推荐方案")

        questions = self.client.get(f"/api/ai-review/conversations/{session_id}").json()["questions"]
        first = next(item for item in questions if item["id"] == first_question["id"])
        self.assertEqual(first["status"], "superseded")
        self.assertEqual(len([item for item in questions if item["status"] == "pending"]), 1)

    def test_needs_input_question_has_no_false_confirmation(self) -> None:
        session_id = self.client.post("/api/ai-review/conversations", json={}).json()["session"]["id"]
        run = session_store.create_workspace_run(session_id, model="configured-model", input_signature="needs-input")
        question = session_store.create_question(
            session_id, run["id"], "请补充译文位置", recommended_label=""
        )
        self.assertEqual(question["recommended_label"], "")
        self.assertEqual(question["recommended_value"], "")

    def test_direct_text_followup_reuses_original_payload_and_returns_a_new_workspace_reply(self) -> None:
        session_id = self.client.post("/api/ai-review/conversations", json={}).json()["session"]["id"]
        session_store.add_message(session_id, "user", "message", "审校英文：Hi, Biby, I love you. Would you love me?")
        first_run = session_store.create_workspace_run(session_id, model="configured-model", input_signature="direct-input")
        session_store.update_workspace_run(
            first_run["id"],
            status="needs_input",
            plan={"source_language": "auto", "targets": [], "needs_input": True, "question": "请说明文本用途。"},
        )
        question = session_store.create_question(session_id, first_run["id"], "请说明文本用途。", recommended_label="")
        response = self.client.post(
            f"/api/ai-review/conversations/{session_id}/decision",
            json={"run_id": first_run["id"], "question_id": question["id"], "action": "other", "answer": "直接审校"},
        )
        self.assertEqual(response.status_code, 200)
        for _ in range(80):
            snapshot = self.client.get(f"/api/ai-review/conversations/{session_id}").json()
            if len(snapshot["workspace_runs"]) >= 2 and snapshot["workspace_runs"][0]["status"] == "ready":
                break
            time.sleep(0.05)
        self.assertEqual(snapshot["workspace_runs"][0]["status"], "ready")
        plan = snapshot["workspace_runs"][0]["plan"]
        self.assertEqual(plan["source_language"], "none")
        self.assertEqual(plan["targets"][0]["units"][0]["target_text"], "Hi, Biby, I love you. Would you love me?")
        self.assertTrue(any(item["kind"] == "workspace_report" for item in snapshot["messages"]))

    def test_pending_attachment_can_be_removed_and_source_can_be_none(self) -> None:
        session_id = self.client.post("/api/ai-review/conversations", json={}).json()["session"]["id"]
        upload = self.client.post(
            f"/api/ai-review/conversations/{session_id}/attachments",
            files={"files": ("remove.csv", "Source,Target\nHello,你好\n", "text/csv")},
        ).json()
        attachment_id = upload["attachments"][0]["id"]
        removed = self.client.delete(
            f"/api/ai-review/conversations/{session_id}/attachments/{attachment_id}"
        )
        self.assertEqual(removed.status_code, 200)
        self.assertEqual(
            self.client.get(f"/api/ai-review/conversations/{session_id}").json()["attachments"], []
        )

        self.client.post(
            f"/api/ai-review/conversations/{session_id}/attachments",
            files={"files": ("translations.csv", "Source,简体中文,日本語\nHello,你好,こんにちは\n", "text/csv")},
        )
        sent = self.client.post(
            f"/api/ai-review/conversations/{session_id}/messages",
            json={"source_language": "none", "target_languages": ["简体中文"]},
        )
        self.assertEqual(sent.status_code, 200)
        snapshot = None
        for _ in range(40):
            snapshot = self.client.get(f"/api/ai-review/conversations/{session_id}").json()
            if snapshot["session"]["status"] in {"ready", "needs_input", "failed"}:
                break
            time.sleep(0.05)
        plan = snapshot["workspace_runs"][0]["plan"]
        self.assertEqual(plan["source_language"], "none")
        self.assertTrue(
            all(not unit["source_text"] for target in plan["targets"] for unit in target["units"])
        )

    def test_pending_excel_structure_can_be_read_before_mapping(self) -> None:
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "翻译"
        sheet.append(["原文", "译文", "备注"])
        sheet.append(["Hello", "你好", "首页"])
        stream = BytesIO()
        workbook.save(stream)
        session_id = self.client.post("/api/ai-review/conversations", json={}).json()["session"]["id"]
        upload = self.client.post(
            f"/api/ai-review/conversations/{session_id}/attachments",
            files={"files": ("mapping.xlsx", stream.getvalue(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
        ).json()
        attachment_id = upload["attachments"][0]["id"]
        response = self.client.get(
            f"/api/ai-review/conversations/{session_id}/attachments/{attachment_id}/structure"
        )
        self.assertEqual(response.status_code, 200)
        structure = response.json()["manifest"]["structure"]
        self.assertEqual(structure["sheets"][0]["name"], "翻译")
        self.assertEqual(
            [item["header"] for item in structure["sheets"][0]["columns"]],
            ["原文", "译文", "备注"],
        )

    def test_new_workspace_message_does_not_include_session_history(self) -> None:
        session_id = self.client.post("/api/ai-review/conversations", json={}).json()["session"]["id"]
        session_store.add_message(session_id, "user", "message", "NOISE_SENTINEL")
        self.client.post(
            f"/api/ai-review/conversations/{session_id}/attachments",
            files={"files": ("independent.csv", "Source,Target\nHello,你好\n", "text/csv")},
        )
        response = self.client.post(
            f"/api/ai-review/conversations/{session_id}/messages",
            json={"source_language": "英语", "target_languages": ["简体中文"]},
        )
        self.assertEqual(response.status_code, 200)
        for _ in range(80):
            snapshot = self.client.get(f"/api/ai-review/conversations/{session_id}").json()
            if workspace_prompts and snapshot["session"]["status"] in {"ready", "needs_input", "failed"}:
                break
            time.sleep(0.05)
        self.assertTrue(workspace_prompts)
        self.assertNotIn("NOISE_SENTINEL", workspace_prompts[-1])
        self.assertIn("普通新消息为空", workspace_prompts[-1])


if __name__ == "__main__":
    unittest.main()
