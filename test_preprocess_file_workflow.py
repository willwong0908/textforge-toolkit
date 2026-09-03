from pathlib import Path

from openpyxl import Workbook
from fastapi.testclient import TestClient

from term_extractor_app.core import read_source_records
from term_extractor_app.storage import build_default_settings
from term_extractor_app.web_app import create_app


def _write_excel(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Main"
    sheet.append(["Source", "Context"])
    sheet.append(["Alpha", "First"])
    workbook.save(path)


def _write_xliff(path: Path) -> None:
    path.write_text(
        """<?xml version="1.0" encoding="UTF-8"?>
<xliff version="1.2" xmlns="urn:oasis:names:tc:xliff:document:1.2">
  <file source-language="en" datatype="plaintext" original="qa">
    <body><trans-unit id="1"><source>Gamma</source></trans-unit></body>
  </file>
</xliff>
""",
        encoding="utf-8",
    )


def test_explicit_mixed_files_do_not_require_a_scanned_folder(tmp_path: Path) -> None:
    excel_path = tmp_path / "source.xlsx"
    csv_path = tmp_path / "source.csv"
    xliff_path = tmp_path / "source.xlf"
    _write_excel(excel_path)
    csv_path.write_text("Source,Context\nBeta,Second\n", encoding="utf-8")
    _write_xliff(xliff_path)

    input_files = [str(excel_path), str(csv_path), str(xliff_path)]
    mappings = {
        str(excel_path): {"Main": ["Source"]},
        str(csv_path): {"CSV": ["Source"]},
        str(xliff_path): {},
    }

    records, processed_files = read_source_records(
        folder_path=str(tmp_path / "folder-does-not-exist"),
        file_type="mixed",
        header_name=[],
        input_files=input_files,
        file_mappings=mappings,
    )

    assert processed_files == ["source.xlsx", "source.csv", "source.xlf"]
    assert [(record.source_type, record.text) for record in records] == [
        ("excel", "Alpha"),
        ("csv", "Beta"),
        ("xliff", "Gamma"),
    ]


def test_explicit_missing_file_is_reported_instead_of_silently_skipped(tmp_path: Path) -> None:
    missing_path = tmp_path / "missing.xlsx"

    try:
        read_source_records(
            folder_path="",
            file_type="excel",
            header_name=["Source"],
            input_files=[str(missing_path)],
            file_mappings={str(missing_path): {"Main": ["Source"]}},
        )
    except ValueError as exc:
        assert "输入文件不存在" in str(exc)
        assert "missing.xlsx" in str(exc)
    else:
        raise AssertionError("missing explicit files must fail before extraction")


def test_start_api_passes_explicit_mapping_to_task_without_valid_folder(tmp_path: Path) -> None:
    excel_path = tmp_path / "source.xlsx"
    _write_excel(excel_path)

    class RecordingFacade:
        def __init__(self) -> None:
            self.task_input = None

        def load_settings(self):
            return build_default_settings()

        def start(self, task_input, resume=False, settings=None) -> None:
            self.task_input = task_input

    facade = RecordingFacade()
    client = TestClient(create_app(facade))
    response = client.post(
        "/api/tasks/start",
        json={
            "folder_path": str(tmp_path / "folder-does-not-exist"),
            "header_name": ["Source"],
            "source_language": "eng",
            "input_files": [str(excel_path)],
            "file_mappings": {str(excel_path): {"Main": ["Source"]}},
        },
    )

    assert response.status_code == 200
    assert facade.task_input is not None
    assert facade.task_input.input_files == [str(excel_path)]
    assert facade.task_input.file_mappings == {str(excel_path): {"Main": ["Source"]}}
