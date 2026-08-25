from __future__ import annotations

import tempfile
import json
import time
import unittest
from pathlib import Path
from unittest.mock import patch

from fastapi import FastAPI
from fastapi.testclient import TestClient

from term_extractor_app.ai_review import cache_service, config, database, session_store, workspace_service
from term_extractor_app.ai_review.conversation_routes import router


def _workspace_response(messages: list[dict[str, str]]) -> str:
    prompt = messages[-1]["content"]
    manifests = json.loads(prompt.split("结构清单：\n", 1)[1])
    attachment_id = next(iter(manifests))
    if "没有原文" in prompt:
        targets = [{"language": "简体中文", "mappings": [{"attachment_id": attachment_id, "scope": "table", "source_column_index": None, "target_column_index": 1}]}]
    else:
        targets = [
            {"language": "简体中文", "mappings": [{"attachment_id": attachment_id, "scope": "table", "source_column_index": 0, "target_column_index": 1}]},
            {"language": "日语", "mappings": [{"attachment_id": attachment_id, "scope": "table", "source_column_index": 0, "target_column_index": 2}]},
        ]
    return json.dumps({
        "content_files": [attachment_id], "reference_files": [], "source_language": "英语",
        "targets": targets, "relationships": [], "assumptions": ["Workspace Agent 已确认映射。"],
        "warnings": [], "confidence": 0.97, "needs_input": False, "question": "",
    }, ensure_ascii=False)


class ConversationApiTests(unittest.TestCase):
    def setUp(self) -> None:
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


if __name__ == "__main__":
    unittest.main()
