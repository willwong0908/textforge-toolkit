"""Logging helpers."""

import logging
import sys
from pathlib import Path
from typing import Optional

from .storage import AppPaths, get_app_paths

LOGGER_NAME = "term_extractor_app"


def configure_file_logger(paths: Optional[AppPaths] = None, *, with_console: bool = True) -> logging.Logger:
    paths = paths or get_app_paths()
    logger = logging.getLogger(LOGGER_NAME)
    logger.setLevel(logging.DEBUG)
    logger.propagate = False

    target_file = str(paths.log_file)
    has_target = any(
        isinstance(handler, logging.FileHandler) and getattr(handler, "baseFilename", "") == target_file
        for handler in logger.handlers
    )
    if not has_target:
        paths.output_dir.mkdir(parents=True, exist_ok=True)
        file_handler = logging.FileHandler(paths.log_file, encoding="utf-8")
        file_handler.setFormatter(logging.Formatter("%(asctime)s - %(levelname)s - %(message)s"))
        logger.addHandler(file_handler)

    console_handlers = [
        handler
        for handler in logger.handlers
        if isinstance(handler, logging.StreamHandler) and not isinstance(handler, logging.FileHandler)
    ]
    if with_console:
        if not console_handlers:
            console_handler = logging.StreamHandler(sys.stdout)
            console_handler.setFormatter(logging.Formatter("%(asctime)s - %(levelname)s - %(message)s"))
            logger.addHandler(console_handler)
    else:
        for handler in console_handlers:
            handler.close()
            logger.removeHandler(handler)
    return logger


def reset_file_logger(paths: Optional[AppPaths] = None, *, with_console: bool = True) -> logging.Logger:
    paths = paths or get_app_paths()
    logger = logging.getLogger(LOGGER_NAME)
    target_file = str(paths.log_file)

    retained_handlers = []
    for handler in list(logger.handlers):
        is_target_file = isinstance(handler, logging.FileHandler) and getattr(handler, "baseFilename", "") == target_file
        if is_target_file:
            try:
                handler.flush()
            except Exception:
                pass
            handler.close()
            logger.removeHandler(handler)
        else:
            retained_handlers.append(handler)

    paths.output_dir.mkdir(parents=True, exist_ok=True)
    Path(paths.log_file).write_text("", encoding="utf-8")
    return configure_file_logger(paths, with_console=with_console)
