from __future__ import annotations

import tempfile
import unittest
from io import BytesIO
from pathlib import Path
from unittest.mock import patch

from fastapi import FastAPI
from fastapi.testclient import TestClient
from openpyxl import Workbook

from term_extractor_app.ai_review import config, database, term_base_service
from term_extractor_app.ai_review.conversation_routes import router as conversation_router
from term_extractor_app.ai_review.term_base_routes import router as term_base_router


def _term_table_bytes(*, target: str = "哎呀讨厌", include_note: bool = True) -> bytes:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Sheet1"
    metadata_headers = ["Entry_ID", "Entry_Subject"] + (["Entry_Note"] if include_note else [])
    sheet.append(metadata_headers + [
        "English", "English", "Japanese", "Korean", "Chinese_PRC", "Chinese_PRC",
    ])
    note_values = ["マラード"] if include_note else []
    sheet.append([700, "", *note_values, "oh my", "goodness", "あらやだ", "어머", target, "哎呀，真讨厌"])
    note_values = ["人名"] if include_note else []
    sheet.append([701, "character", *note_values, "Luke", "", "ルーク", "루크", "路克", "卢克"])
    output = BytesIO()
    workbook.save(output)
    workbook.close()
    return output.getvalue()


class ReviewTermBaseTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        root = Path(self.temp.name)
        self.db_path = root / "data" / "app.sqlite3"
        self.db_path.parent.mkdir(parents=True)
        self.term_bases_dir = root / "term_bases"
        self.term_bases_dir.mkdir(parents=True)
        self.patches = [
            patch.object(database, "DB_PATH", self.db_path),
            patch.object(config, "DB_PATH", self.db_path),
            patch.object(config, "TERM_BASES_DIR", self.term_bases_dir),
            patch.object(term_base_service, "TERM_BASES_DIR", self.term_bases_dir),
        ]
        for item in self.patches:
            item.start()
        database.init_db()

    def tearDown(self) -> None:
        for item in reversed(self.patches):
            item.stop()
        self.temp.cleanup()

    def test_duplicate_language_columns_and_entry_note_are_preserved(self) -> None:
        term_base = term_base_service.upload_term_base("project_tb.xlsx", _term_table_bytes())
        self.assertEqual(term_base["entry_count"], 2)
        self.assertTrue(term_base["has_entry_note"])
        self.assertEqual(term_base["languages"].count("English"), 1)

        pairs, metadata = term_base_service.match_terms(
            term_base["id"], "日语", "简体中文", "ええっ、あらやだ！", limit=50
        )
        self.assertEqual(metadata["source_column"], "Japanese")
        self.assertEqual(metadata["target_column"], "Chinese_PRC")
        self.assertEqual(pairs, [{
            "source": "あらやだ",
            "targets": ["哎呀讨厌", "哎呀，真讨厌"],
            "entry_note": "マラード",
        }])

    def test_same_filename_overwrites_in_place_and_session_selection_is_independent(self) -> None:
        app = FastAPI()
        app.include_router(conversation_router)
        app.include_router(term_base_router)
        with TestClient(app) as client:
            invalid = client.post("/api/ai-review/conversations", json={"term_base_id": "missing"})
            self.assertEqual(invalid.status_code, 400)
            uploaded = client.post(
                "/api/ai-review/term-bases",
                files={"file": ("project_tb.xlsx", _term_table_bytes(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
            )
            self.assertEqual(uploaded.status_code, 200)
            term_base_id = uploaded.json()["term_base"]["id"]
            first = client.post("/api/ai-review/conversations", json={}).json()["session"]["id"]
            second = client.post("/api/ai-review/conversations", json={}).json()["session"]["id"]
            selected = client.patch(
                f"/api/ai-review/conversations/{first}", json={"term_base_id": term_base_id}
            )
            self.assertEqual(selected.status_code, 200)
            self.assertEqual(selected.json()["session"]["term_base_id"], term_base_id)
            self.assertIsNone(client.get(f"/api/ai-review/conversations/{second}").json()["session"]["term_base_id"])

            overwritten = client.post(
                "/api/ai-review/term-bases",
                files={"file": ("project_tb.xlsx", _term_table_bytes(target="哎呀"), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
            )
            self.assertEqual(overwritten.json()["term_base"]["id"], term_base_id)
            self.assertEqual(
                client.get(f"/api/ai-review/conversations/{first}").json()["session"]["term_base_id"],
                term_base_id,
            )

            deleted = client.delete(f"/api/ai-review/term-bases/{term_base_id}")
            self.assertEqual(deleted.status_code, 200)
            self.assertIsNone(client.get(f"/api/ai-review/conversations/{first}").json()["session"]["term_base_id"])

    def test_entry_note_column_is_optional(self) -> None:
        term_base = term_base_service.upload_term_base("without_note.xlsx", _term_table_bytes(include_note=False))
        self.assertFalse(term_base["has_entry_note"])
        pairs, _ = term_base_service.match_terms(term_base["id"], "日语", "简体中文", "あらやだ")
        self.assertEqual(pairs[0]["source"], "あらやだ")
        self.assertNotIn("entry_note", pairs[0])


if __name__ == "__main__":
    unittest.main()
