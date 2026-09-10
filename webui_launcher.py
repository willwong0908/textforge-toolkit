from __future__ import annotations

import contextlib
import ctypes
import json
import logging
import os
import socket
import sys
import threading
import time
import traceback
import urllib.request
import webbrowser

import uvicorn

from term_extractor_app.constants import APP_NAME, APP_VERSION
from term_extractor_app.logging_utils import configure_file_logger
from term_extractor_app.storage import get_app_paths
from term_extractor_app.web_app import app


HOST = "127.0.0.1"
PORT = 8765
URL = f"http://{HOST}:{PORT}/"

LOGGER = configure_file_logger(get_app_paths(), with_console=False)


def _prepare_console() -> None:
    os.environ.setdefault("PYTHONUTF8", "1")
    for stream_name in ("stdout", "stderr"):
        stream = getattr(sys, stream_name, None)
        if hasattr(stream, "reconfigure"):
            with contextlib.suppress(Exception):
                stream.reconfigure(encoding="utf-8", errors="replace")
    if os.name == "nt":
        with contextlib.suppress(Exception):
            ctypes.windll.kernel32.SetConsoleOutputCP(65001)
            ctypes.windll.kernel32.SetConsoleCP(65001)
            ctypes.windll.kernel32.SetConsoleTitleW(f"{APP_NAME} {APP_VERSION}")


def _log_file(level: str, message: str, *args) -> None:
    log_method = getattr(LOGGER, level, LOGGER.info)
    if args:
        log_method("[WEBUI] " + message, *args)
    else:
        log_method("[WEBUI] " + message)


def _log_console(message: str, *, level: str = "info") -> None:
    print(message)
    sys.stdout.flush()
    _log_file(level, message)


def _print_banner() -> None:
    print("\033[96m")
    print("  ██╗   ██╗███████╗███████╗██╗  ██╗███████╗")
    print("  ╚██╗ ██╔╝██╔════╝██╔════╝██║  ██║██╔════╝")
    print("   ╚████╔╝ █████╗  █████╗  ███████║█████╗  ")
    print("    ╚██╔╝  ██╔══╝  ██╔══╝  ██╔══██║██╔══╝  ")
    print("     ██║   ███████╗███████╗██║  ██║███████╗")
    print("     ╚═╝   ╚══════╝╚══════╝╚═╝  ╚═╝╚══════╝")
    print("\033[0m", end="")
    print("  ┌──────────────────────────────────────────────┐")
    print(f"  │ {APP_NAME}  v{APP_VERSION:<27}│")
    print("  │                                              │")
    print(f"  │ 服务地址  {URL:<31}│")
    print("  │ 关闭此窗口将停止当前服务                     │")
    print("  │ 输入 q 回车可主动关闭服务并退出窗口          │")
    print("  └──────────────────────────────────────────────┘")
    print("")
    print("  运行日志")
    print("  ────────────────────────────────────────────────")
    sys.stdout.flush()
    _log_file("info", "Launcher opened. version=%s url=%s", APP_VERSION, URL)


def _is_port_open() -> bool:
    with contextlib.closing(socket.socket(socket.AF_INET, socket.SOCK_STREAM)) as sock:
        sock.settimeout(0.3)
        return sock.connect_ex((HOST, PORT)) == 0


def _is_service_ready() -> bool:
    try:
        request = urllib.request.Request(
            f"{URL}api/status",
            headers={"Accept": "application/json", "User-Agent": "Yeehe-Toolkit-Launcher"},
        )
        with urllib.request.urlopen(request, timeout=1.0) as response:
            if response.status != 200:
                return False
            payload = json.loads(response.read().decode("utf-8"))
        return isinstance(payload, dict) and "is_running" in payload and "stage" in payload
    except Exception:
        return False


def _wait_until_ready(timeout_seconds: float = 20.0) -> bool:
    deadline = time.time() + timeout_seconds
    while time.time() < deadline:
        if _is_service_ready():
            return True
        time.sleep(0.2)
    return False


def _print_idle_hint() -> None:
    _log_console("  输入 q 回车可退出当前窗口。")


def _open_browser_when_ready(existing_service: bool = False) -> None:
    if _wait_until_ready():
        if existing_service:
            _log_console(f"  检测到已有运行中的服务：{URL}")
            _log_console("  已直接打开浏览器。")
            _print_idle_hint()
        else:
            _log_console(f"  已启动：{URL}")
            _log_console("  浏览器已打开。保持本窗口开启即可使用工具。")
        try:
            webbrowser.open(URL)
            _log_file("info", "Browser open requested for %s", URL)
        except Exception:
            LOGGER.exception("[WEBUI] Failed to open browser for %s", URL)
    else:
        _log_console("  启动超时：请查看下方错误信息，或确认端口是否被占用。", level="error")
        _print_idle_hint()


def _listen_for_quit(server: uvicorn.Server | None) -> None:
    while True:
        try:
            command = input().strip().lower()
        except EOFError:
            _log_file("info", "Launcher input stream closed.")
            if server is not None:
                _log_console("  检测到主窗口已关闭，正在停止服务...")
                server.should_exit = True
            return
        except KeyboardInterrupt:
            _log_file("info", "Launcher interrupted by keyboard.")
            if server is not None:
                _log_console("  检测到主窗口已关闭，正在停止服务...")
                server.should_exit = True
            return
        if command in {"q", "quit", "exit"}:
            if server is None:
                _log_console("  当前窗口已退出。")
            else:
                _log_console("  正在关闭服务...")
                LOGGER.info("User requested shutdown with command=%s", command)
                server.should_exit = True
            return


def main() -> int:
    _prepare_console()
    _print_banner()

    if _is_port_open():
        existing_service_ready = _is_service_ready()
        if existing_service_ready:
            _log_file("info", "Existing service detected on port %s", PORT)
            threading.Thread(target=_open_browser_when_ready, kwargs={"existing_service": True}, daemon=True).start()
        else:
            _log_console(f"  启动失败：端口 {PORT} 已被其他程序占用，或旧服务没有正常响应。", level="error")
            _log_console("  请关闭占用该端口的程序后重新启动。本工具不会打开无效页面。", level="error")
            _print_idle_hint()
        _listen_for_quit(None)
        return 0 if existing_service_ready else 2

    logging.getLogger("uvicorn").setLevel(logging.WARNING)
    logging.getLogger("uvicorn.error").setLevel(logging.WARNING)
    logging.getLogger("uvicorn.access").disabled = True

    config = uvicorn.Config(
        app,
        host=HOST,
        port=PORT,
        access_log=False,
        log_level="warning",
    )
    server = uvicorn.Server(config)
    threading.Thread(target=_open_browser_when_ready, daemon=True).start()
    threading.Thread(target=_listen_for_quit, args=(server,), daemon=True).start()

    try:
        _log_file("info", "Starting WebUI server on %s", URL)
        server.run()
    except KeyboardInterrupt:
        _log_file("info", "WebUI server interrupted by keyboard.")
        server.should_exit = True
    except Exception as exc:
        LOGGER.exception("[WEBUI] Launcher crashed: %s", exc)
        _log_console("  启动失败，详细报错已写入 output/log.txt。", level="error")
        traceback.print_exc()
        return 1
    finally:
        _log_file("info", "WebUI server stopped.")
        _log_console("  服务已关闭。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
