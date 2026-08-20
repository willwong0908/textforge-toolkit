from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from term_extractor_app.ai_review import config, database
from term_extractor_app.ai_review import excel_mapping_service as presets


class ExcelMappingPresetServiceTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.original_data_dir = config.DATA_DIR
        self.original_db_path = database.DB_PATH
        config.DATA_DIR = Path(self.temp_dir.name)
        database.DB_PATH = config.DATA_DIR / "app.sqlite3"
        database.init_db()

    def tearDown(self) -> None:
        config.DATA_DIR = self.original_data_dir
        database.DB_PATH = self.original_db_path
        self.temp_dir.cleanup()

    def test_save_overwrite_and_delete_mapping_preset(self) -> None:
        mapping = {
            "source_language": "English",
            "target_language": "Chinese",
            "sheets": [
                {
                    "sheet_name": "Original Sheet",
                    "mappings": [
                        {
                            "source_column": 0,
                            "target_column": 1,
                            "info_columns": [{"column": 2, "category": "context"}],
                        }
                    ],
                }
            ],
        }

        saved = presets.save_excel_mapping_preset("  Game Text  ", mapping)
        self.assertEqual(2, saved["mapping"]["schema_version"])
        self.assertEqual("Game Text", saved["name"])
        self.assertEqual("context", saved["mapping"]["sheets"][0]["mappings"][0]["info_columns"][0]["category"])

        with self.assertRaisesRegex(ValueError, "同名"):
            presets.save_excel_mapping_preset("game text", mapping)

        updated = presets.save_excel_mapping_preset("game text", mapping, saved["id"])
        self.assertEqual(saved["id"], updated["id"])
        self.assertEqual(1, len(presets.list_excel_mapping_presets()))

        presets.delete_excel_mapping_preset(saved["id"])
        self.assertEqual([], presets.list_excel_mapping_presets())

    def test_rejects_empty_name_and_invalid_mapping(self) -> None:
        with self.assertRaisesRegex(ValueError, "模板名称"):
            presets.save_excel_mapping_preset("   ", {"sheets": []})
        with self.assertRaisesRegex(ValueError, "原文列和译文列"):
            presets.save_excel_mapping_preset("Invalid", {"sheets": []})
        with self.assertRaisesRegex(ValueError, "不能相同"):
            presets.save_excel_mapping_preset(
                "Invalid",
                {
                    "sheets": [
                        {"mappings": [{"source_column": 0, "target_column": 0, "info_columns": []}]}
                    ]
                },
            )


if __name__ == "__main__":
    unittest.main()
