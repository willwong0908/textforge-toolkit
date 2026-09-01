from __future__ import annotations

from pathlib import Path

from ..storage import get_app_paths, get_app_root


APP_DIR = get_app_root()
APP_PATHS = get_app_paths()
BUNDLE_DIR = APP_DIR
BASE_DIR = APP_DIR
DATA_DIR = APP_PATHS.output_dir / "ai_review" / "data"
UPLOADS_DIR = APP_PATHS.output_dir / "ai_review" / "uploads"
OUTPUTS_DIR = APP_PATHS.output_dir / "ai_review" / "outputs"
TERM_BASES_DIR = APP_PATHS.output_dir / "ai_review" / "term_bases"
STATIC_DIR = APP_DIR / "static"
DB_PATH = DATA_DIR / "app.sqlite3"

HOST = "127.0.0.1"
PORT = 8765

SUPPORTED_EXCEL_EXTENSIONS = {".xlsx", ".xlsm"}
SUPPORTED_XLIFF_EXTENSIONS = {".xlf", ".xliff"}


def ensure_directories() -> None:
    for directory in (DATA_DIR, UPLOADS_DIR, OUTPUTS_DIR, TERM_BASES_DIR):
        directory.mkdir(parents=True, exist_ok=True)
