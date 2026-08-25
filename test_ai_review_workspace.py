from __future__ import annotations

import tempfile
import json
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook

from term_extractor_app.ai_review import cache_service, config, database, session_store
from term_extractor_app.ai_review.readers import ReaderError, build_default_registry
from term_extractor_app.ai_review.excel_mapping_service import save_excel_mapping_preset
from term_extractor_app.ai_review.shared_provider import _stream_reasoning_content
from term_extractor_app.ai_review.workspace_service import (
    _structure_signature,
    inspect_workspace_sync,
    save_confirmed_profile,
)


def _workspace_response(messages: list[dict[str, str]], on_delta=None) -> str:
    prompt = messages[-1]["content"]
    manifests = json.loads(prompt.split("结构清单：\n", 1)[1])
    content_files = list(manifests)
    targets: list[dict[str, object]] = []
    for attachment_id, manifest in manifests.items():
        if manifest["file_type"] == "direct_text":
            for sample in manifest["samples"]:
                language = "简体中文" if any("\u4e00" <= char <= "\u9fff" for char in sample["text"]) else "英语"
                targets.append({"language": language, "mappings": [{"attachment_id": attachment_id, "target_pointer": sample["pointer"]}]})
            continue
        filename = manifest["filename"]
        if "没有原文" in prompt:
            target_column = 1 if "B 列" in prompt else 2
            mappings = [{"attachment_id": attachment_id, "scope": "table", "source_column_index": None, "target_column_index": target_column}]
            targets.append({"language": "简体中文", "mappings": mappings})
        elif filename == "translations.csv":
            targets.extend([
                {"language": "简体中文", "mappings": [{"attachment_id": attachment_id, "scope": "table", "source_column_index": 0, "target_column_index": 1}]},
                {"language": "日语", "mappings": [{"attachment_id": attachment_id, "scope": "table", "source_column_index": 0, "target_column_index": 2}]},
            ])
    response = json.dumps({
        "content_files": content_files,
        "reference_files": [],
        "source_language": "auto",
        "targets": targets,
        "relationships": [],
        "assumptions": ["Workspace Agent 已根据结构清单确认映射。"],
        "warnings": [],
        "confidence": 0.97,
        "needs_input": False,
        "question": "",
    }, ensure_ascii=False)
    if on_delta:
        for start in range(0, len(response), 24):
            on_delta(response[start : start + 24])
    return response


class ReaderRegistryTests(unittest.TestCase):
    def test_stream_reasoning_delta_is_exposed(self) -> None:
        payload = {"choices": [{"delta": {"reasoning_content": "正在判断工作表"}}]}
        self.assertEqual(_stream_reasoning_content(payload), "正在判断工作表")

    def test_reads_excel_csv_json_and_docx(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            workbook = Workbook()
            sheet = workbook.active
            sheet.append(["Source", "简体中文", "日本語"])
            sheet.append(["Hello", "你好", "こんにちは"])
            excel_path = root / "sample.xlsx"
            workbook.save(excel_path)

            csv_path = root / "sample.csv"
            csv_path.write_text("Source,Target\nHello,你好\n", encoding="utf-8")
            json_path = root / "sample.json"
            json_path.write_text('{"source":"Hello","target":"你好"}', encoding="utf-8")
            docx_path = root / "sample.docx"
            with zipfile.ZipFile(docx_path, "w") as archive:
                archive.writestr(
                    "word/document.xml",
                    '<w:document xmlns:w="urn:test"><w:body><w:p><w:r><w:t>你好</w:t></w:r></w:p></w:body></w:document>',
                )

            registry = build_default_registry()
            self.assertEqual(registry.read(excel_path).file_type, "excel")
            self.assertEqual(registry.read(csv_path).file_type, "csv")
            self.assertEqual(registry.read(json_path).blocks[0].pointer, "$/source")
            self.assertEqual(registry.read(docx_path).blocks[0].text, "你好")

    def test_rejects_unknown_extension(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "sample.unknown"
            path.write_text("data", encoding="utf-8")
            with self.assertRaises(ReaderError):
                build_default_registry().read(path)


class WorkspaceSessionTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        root = Path(self.temp.name)
        self.db_path = root / "data" / "app.sqlite3"
        self.db_path.parent.mkdir(parents=True)
        self.uploads = root / "uploads"
        self.uploads.mkdir(parents=True)
        self.patches = [
            patch.object(database, "DB_PATH", self.db_path),
            patch.object(config, "DB_PATH", self.db_path),
            patch.object(cache_service, "UPLOADS_DIR", self.uploads),
            patch.object(session_store, "UPLOADS_DIR", self.uploads),
        ]
        for item in self.patches:
            item.start()
        database.init_db()

    def tearDown(self) -> None:
        for item in reversed(self.patches):
            item.stop()
        self.temp.cleanup()

    def test_session_persists_and_workspace_creates_separate_target_plans(self) -> None:
        session = session_store.create_session(target_languages=["简体中文", "日语"])
        csv_data = "Source,简体中文,日本語\nHello,你好,こんにちは\nBye,再见,さようなら\n".encode("utf-8")
        session_store.add_attachment(session["id"], "translations.csv", csv_data)
        with patch(
            "term_extractor_app.ai_review.workspace_service.get_shared_ai_settings",
            return_value={"selected_model": "configured-model", "api_key": "test-key"},
        ), patch(
            "term_extractor_app.ai_review.workspace_service.workspace_chat",
            side_effect=_workspace_response,
        ) as workspace_mock:
            plan = inspect_workspace_sync(session["id"])
        workspace_mock.assert_called_once()
        targets = {item["language"]: item["units"] for item in plan["targets"]}
        self.assertEqual(len(targets["简体中文"]), 2)
        self.assertEqual(len(targets["日语"]), 2)
        self.assertEqual(targets["简体中文"][0]["source_text"], "Hello")
        file_summary = plan["file_summaries"][0]
        self.assertEqual(file_summary["filename"], "translations.csv")
        chinese_mapping = next(item for item in file_summary["mappings"] if item["language"] == "简体中文")
        self.assertIn("第 1 列", chinese_mapping["source_location"])
        self.assertIn("第 2 列", chinese_mapping["target_location"])
        self.assertEqual(chinese_mapping["samples"][0]["source"], "Hello")
        self.assertEqual(chinese_mapping["samples"][0]["target"], "你好")
        snapshot = session_store.get_session_snapshot(session["id"])
        self.assertEqual(snapshot["session"]["status"], "ready")
        self.assertEqual(snapshot["workspace_runs"][0]["model"], "configured-model")
        report = next(message["content"] for message in snapshot["messages"] if message["kind"] == "workspace_report")
        self.assertIn("translations.csv", report)
        self.assertIn("样例 1", report)
        self.assertIn("第 1 列", report)

    def test_natural_language_adjustment_rebuilds_full_column_mapping(self) -> None:
        session = session_store.create_session(target_languages=["简体中文"])
        csv_data = "OldSource,OldTarget,FinalTarget\nS1,T1,F1\nS2,T2,F2\n".encode("utf-8")
        session_store.add_attachment(session["id"], "adjust.csv", csv_data)
        with patch(
            "term_extractor_app.ai_review.workspace_service.get_shared_ai_settings",
            return_value={"selected_model": "configured-model", "api_key": "test-key"},
        ), patch(
            "term_extractor_app.ai_review.workspace_service.workspace_chat",
            side_effect=_workspace_response,
        ):
            plan = inspect_workspace_sync(
                session["id"],
                adjustment_text="adjust.csv 没有原文，简体中文译文在 C 列",
            )
        units = plan["targets"][0]["units"]
        self.assertEqual([unit["target_text"] for unit in units], ["F1", "F2"])
        self.assertTrue(all(unit["source_text"] == "" for unit in units))
        self.assertTrue(all(unit["metadata"]["adjusted_by_user"] for unit in units))
        self.assertFalse(plan["needs_input"])

    def test_direct_text_can_be_reviewed_without_source(self) -> None:
        session = session_store.create_session(target_languages=["auto"])
        with patch(
            "term_extractor_app.ai_review.workspace_service.get_shared_ai_settings",
            return_value={"selected_model": "configured-model", "api_key": "test-key"},
        ), patch(
            "term_extractor_app.ai_review.workspace_service.workspace_chat",
            side_effect=_workspace_response,
        ):
            plan = inspect_workspace_sync(session["id"], "你好世界\n\nHello world")
        units = [unit for target in plan["targets"] for unit in target["units"]]
        self.assertEqual(len(units), 2)
        self.assertTrue(all(unit["source_text"] == "" for unit in units))

    def test_xliff_is_the_only_local_only_fast_path(self) -> None:
        session = session_store.create_session(target_languages=["简体中文"])
        xliff = b'''<?xml version="1.0" encoding="UTF-8"?>
<xliff version="1.2"><file source-language="en" target-language="zh-CN"><body>
<trans-unit id="1"><source>Hello</source><target>\xe4\xbd\xa0\xe5\xa5\xbd</target></trans-unit>
</body></file></xliff>'''
        session_store.add_attachment(session["id"], "sample.xlf", xliff)
        with patch(
            "term_extractor_app.ai_review.workspace_service.get_shared_ai_settings",
            return_value={"selected_model": "configured-model", "api_key": "test-key"},
        ), patch(
            "term_extractor_app.ai_review.workspace_service.workspace_chat",
        ) as workspace_mock:
            plan = inspect_workspace_sync(session["id"])
        workspace_mock.assert_not_called()
        self.assertEqual(plan["targets"][0]["units"][0]["source_text"], "Hello")
        self.assertEqual(plan["confidence"], 0.98)

    def test_confirmed_table_structure_cache_requires_exact_languages_and_headers(self) -> None:
        csv_data = "Source,简体中文,日本語\nHello,你好,こんにちは\n".encode("utf-8")
        first = session_store.create_session(source_language="英语", target_languages=["简体中文", "日语"])
        session_store.add_attachment(first["id"], "translations.csv", csv_data)
        with patch(
            "term_extractor_app.ai_review.workspace_service.get_shared_ai_settings",
            return_value={"selected_model": "configured-model", "api_key": "test-key"},
        ), patch(
            "term_extractor_app.ai_review.workspace_service.workspace_chat",
            side_effect=_workspace_response,
        ):
            inspect_workspace_sync(first["id"])
        first_run = session_store.get_session_snapshot(first["id"])["workspace_runs"][0]
        self.assertTrue(first_run["attachment_ids"])
        self.assertIn("target_column_index", first_run["plan"]["targets"][0]["units"][0]["metadata"])
        save_confirmed_profile(first_run["id"])

        second = session_store.create_session(source_language="英语", target_languages=["简体中文", "日语"])
        second_attachment = session_store.add_attachment(second["id"], "translations.csv", csv_data)
        second_document = build_default_registry().read(
            Path(second_attachment["stored_path"]), second_attachment["original_filename"]
        )
        expected_signature = _structure_signature(
            {second_attachment["id"]: second_document}, "英语", ["简体中文", "日语"]
        )
        with database.get_connection() as conn:
            cached_row = conn.execute(
                "SELECT signature FROM extraction_profiles WHERE signature = ?", (expected_signature,)
            ).fetchone()
            all_signatures = [row["signature"] for row in conn.execute("SELECT signature FROM extraction_profiles")]
        self.assertEqual(all_signatures, [expected_signature])
        with patch(
            "term_extractor_app.ai_review.workspace_service.get_shared_ai_settings",
            return_value={"selected_model": "configured-model", "api_key": "test-key"},
        ), patch(
            "term_extractor_app.ai_review.workspace_service.workspace_chat",
        ) as workspace_mock:
            cached = inspect_workspace_sync(second["id"])
        workspace_mock.assert_not_called()
        self.assertTrue(cached["cache_hit"])
        self.assertIn("本地结构缓存", cached["assumptions"][0])

        changed = session_store.create_session(source_language="英语", target_languages=["简体中文"])
        session_store.add_attachment(changed["id"], "translations.csv", csv_data)
        with patch(
            "term_extractor_app.ai_review.workspace_service.get_shared_ai_settings",
            return_value={"selected_model": "configured-model", "api_key": "test-key"},
        ), patch(
            "term_extractor_app.ai_review.workspace_service.workspace_chat",
            side_effect=_workspace_response,
        ) as workspace_mock:
            inspect_workspace_sync(changed["id"])
        workspace_mock.assert_called_once()

    def test_selected_excel_mapping_preset_bypasses_workspace(self) -> None:
        workbook = Workbook()
        workbook.active.title = "Text"
        workbook.active.append(["Source", "Target"])
        workbook.active.append(["Hello", "你好"])
        path = Path(self.temp.name) / "mapped.xlsx"
        workbook.save(path)
        preset = save_excel_mapping_preset(
            "双语模板",
            {
                "source_language": "英语",
                "target_language": "简体中文",
                "sheets": [{"sheet_name": "Text", "mappings": [{"source_column": 0, "target_column": 1}]}],
            },
        )
        session = session_store.create_session(source_language="英语", target_languages=["简体中文"])
        attachment = session_store.add_attachment(session["id"], "mapped.xlsx", path.read_bytes())
        session_store.update_attachment(
            attachment["id"], mapping_mode="preset", mapping_preset_id=preset["id"]
        )
        with patch(
            "term_extractor_app.ai_review.workspace_service.get_shared_ai_settings",
            return_value={"selected_model": "configured-model", "api_key": "test-key"},
        ), patch(
            "term_extractor_app.ai_review.workspace_service.workspace_chat",
        ) as workspace_mock:
            plan = inspect_workspace_sync(session["id"])
        workspace_mock.assert_not_called()
        self.assertEqual(plan["targets"][0]["units"][0]["source_text"], "Hello")
        self.assertEqual(plan["targets"][0]["units"][0]["target_text"], "你好")

    def test_delete_session_removes_private_attachments(self) -> None:
        session = session_store.create_session()
        attachment = session_store.add_attachment(session["id"], "sample.txt", b"hello")
        path = Path(attachment["stored_path"])
        self.assertTrue(path.exists())
        session_store.delete_session(session["id"])
        self.assertFalse(path.exists())
        self.assertIsNone(session_store.get_session(session["id"]))


if __name__ == "__main__":
    unittest.main()
