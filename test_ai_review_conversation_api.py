from __future__ import annotations

import tempfile
import json
import time
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
        "content_files": [attachment_id], "reference_files": [], "source_language": "英语",
        "targets": targets, "relationships": [], "assumptions": ["Workspace Agent 已确认映射。"],
        "warnings": [], "confidence": 0.97, "needs_input": False, "question": "",
    }, ensure_ascii=False)
    if on_delta:
        on_delta("正在判断文件与列之间的关系。", "reasoning")
        on_delta(response, "content")
    return response


class ConversationApiTests(unittest.TestCase):
    def setUp(self) -> None:
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
        opened = self.client.get(
            f"/api/ai-review/conversations/{session_id}/attachments/{uploaded_id}/file"
        )
        self.assertEqual(opened.status_code, 200)
        self.assertIn(b"Source", opened.content)
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
            if len(snapshot["workspace_runs"]) >= 2 and snapshot["workspace_runs"][0]["status"] == "ready":
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

    def test_delete_requires_confirmation(self) -> None:
        session_id = self.client.post("/api/ai-review/conversations", json={}).json()["session"]["id"]
        response = self.client.request(
            "DELETE",
            f"/api/ai-review/conversations/{session_id}",
            json={"confirm": False},
        )
        self.assertEqual(response.status_code, 400)

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
