from __future__ import annotations

"""memoQ Server termbase integration.

Credentials are encrypted with Windows DPAPI and are never returned by the API.
The service deliberately keeps the old uploaded termbase service independent so
sessions can fall back to it when memoQ is not bound.
"""

import base64
import ctypes
import ctypes.wintypes as wintypes
import json
import hashlib
import threading
import time
from typing import Any

import httpx

from .database import dumps_json, get_connection, loads_json, utc_now

DEFAULT_BASE_URL = "https://memoq.iyeehe.com:8080"
_CREDENTIAL_KEY = "memoq_server_credentials"
_cache_lock = threading.Lock()
_resource_cache: tuple[float, list[dict[str, Any]]] | None = None


class MemoQError(Exception):
    pass


def _protect(raw: bytes) -> bytes:
    if not hasattr(ctypes, "windll"):
        raise MemoQError("memoQ 凭证加密仅支持 Windows")
    crypt32 = ctypes.windll.crypt32
    kernel32 = ctypes.windll.kernel32
    class Blob(ctypes.Structure):
        _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_byte))]
    data = Blob(len(raw), ctypes.cast(ctypes.create_string_buffer(raw), ctypes.POINTER(ctypes.c_byte)))
    out = Blob()
    if not crypt32.CryptProtectData(ctypes.byref(data), "memoQ", None, None, None, 0, ctypes.byref(out)):
        raise MemoQError("无法安全保存 memoQ 凭证")
    try:
        return ctypes.string_at(out.pbData, out.cbData)
    finally:
        kernel32.LocalFree(out.pbData)


def _unprotect(raw: bytes) -> bytes:
    if not hasattr(ctypes, "windll"):
        raise MemoQError("memoQ 凭证加密仅支持 Windows")
    crypt32 = ctypes.windll.crypt32
    kernel32 = ctypes.windll.kernel32
    class Blob(ctypes.Structure):
        _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_byte))]
    buf = ctypes.create_string_buffer(raw)
    data = Blob(len(raw), ctypes.cast(buf, ctypes.POINTER(ctypes.c_byte)))
    out = Blob()
    if not crypt32.CryptUnprotectData(ctypes.byref(data), None, None, None, None, 0, ctypes.byref(out)):
        raise MemoQError("memoQ 凭证已失效，请重新绑定")
    try:
        return ctypes.string_at(out.pbData, out.cbData)
    finally:
        kernel32.LocalFree(out.pbData)


def _load_credentials() -> dict[str, str] | None:
    with get_connection() as conn:
        row = conn.execute("SELECT value_json FROM settings WHERE key = ?", (_CREDENTIAL_KEY,)).fetchone()
    if not row:
        return None
    try:
        encoded = json.loads(row["value_json"])
        raw = _unprotect(base64.b64decode(str(encoded)))
        value = json.loads(raw.decode("utf-8"))
        return {"username": str(value.get("username") or ""), "password": str(value.get("password") or ""), "base_url": str(value.get("base_url") or DEFAULT_BASE_URL)}
    except (ValueError, KeyError, json.JSONDecodeError) as exc:
        raise MemoQError("memoQ 凭证无法读取，请重新绑定") from exc


def save_credentials(username: str, password: str, base_url: str = DEFAULT_BASE_URL) -> None:
    username, password = username.strip(), password.strip()
    if not username or not password:
        raise MemoQError("请输入 memoQ 用户名和密码")
    value = {"username": username, "password": password, "base_url": base_url.strip().rstrip("/") or DEFAULT_BASE_URL}
    encrypted = base64.b64encode(_protect(json.dumps(value, ensure_ascii=False).encode("utf-8"))).decode("ascii")
    now = utc_now()
    with get_connection() as conn:
        conn.execute("INSERT INTO settings(key,value_json,updated_at) VALUES(?,?,?) ON CONFLICT(key) DO UPDATE SET value_json=excluded.value_json, updated_at=excluded.updated_at", (_CREDENTIAL_KEY, json.dumps(encrypted), now))
    _invalidate_cache()


def clear_credentials() -> None:
    with get_connection() as conn:
        conn.execute("DELETE FROM settings WHERE key = ?", (_CREDENTIAL_KEY,))
    _invalidate_cache()


def credential_status() -> dict[str, Any]:
    value = _load_credentials()
    return {"bound": bool(value), "username": value.get("username", "") if value else "", "base_url": value.get("base_url", DEFAULT_BASE_URL) if value else DEFAULT_BASE_URL}


def _invalidate_cache() -> None:
    global _resource_cache
    with _cache_lock:
        _resource_cache = None


class MemoQClient:
    def __init__(self, credentials: dict[str, str]):
        self.credentials = credentials
        self.base_url = credentials.get("base_url") or DEFAULT_BASE_URL
        self.token = ""
        self.client = httpx.Client(timeout=20, verify=True, follow_redirects=True)

    def close(self) -> None:
        self.client.close()

    def login(self) -> None:
        response = self.client.post(self.base_url + "/memoqserverhttpapi/v1/auth/login", json={"username": self.credentials["username"], "password": self.credentials["password"], "LoginMode": "0"})
        if response.status_code >= 400:
            raise MemoQError(f"memoQ 登录失败（HTTP {response.status_code}）")
        data = response.json() if response.content else {}
        self.token = str(data.get("AccessToken") or data.get("accessToken") or data.get("token") or "")
        if not self.token:
            raise MemoQError("memoQ 登录响应未返回访问令牌")

    def request(self, method: str, path: str, **kwargs: Any) -> Any:
        if not self.token:
            self.login()
        headers = dict(kwargs.pop("headers", {}) or {})
        headers["Authorization"] = f"MQS-API {self.token}"
        response = self.client.request(method, self.base_url + "/memoqserverhttpapi/v1" + path, headers=headers, **kwargs)
        if response.status_code in {401, 403}:
            self.login()
            headers["Authorization"] = f"MQS-API {self.token}"
            response = self.client.request(method, self.base_url + "/memoqserverhttpapi/v1" + path, headers=headers, **kwargs)
        if response.status_code >= 400:
            raise MemoQError(f"memoQ 请求失败（HTTP {response.status_code}）")
        return response.json() if response.content else {}

    def termbases(self) -> list[dict[str, Any]]:
        data = self.request("GET", "/tbs")
        raw = data if isinstance(data, list) else data.get("Items") or data.get("items") or data.get("Termbases") or data.get("termbases") or []
        result = []
        for item in raw:
            if not isinstance(item, dict):
                continue
            guid = str(item.get("TBGuid") or item.get("tbGuid") or item.get("Guid") or item.get("Id") or "").strip()
            name = str(item.get("FriendlyName") or item.get("Name") or item.get("name") or guid).strip()
            if not guid or item.get("IsUsedInProject") is False:
                continue
            languages = item.get("Languages") or item.get("languages") or []
            if isinstance(languages, str):
                languages = [languages]
            result.append({"id": guid, "name": name, "project": str(item.get("Project") or item.get("ProjectName") or ""), "languages": [str(x) for x in languages if x],
                           "revision": hashlib.sha256(json.dumps(item, sort_keys=True, ensure_ascii=False).encode('utf-8')).hexdigest()})
        return result

    def lookup(self, termbase_ids: list[str], source_language: str, target_language: str, segments: list[str]) -> list[dict[str, Any]]:
        matches: list[dict[str, Any]] = []
        for termbase_id in termbase_ids:
            for offset in range(0, len(segments), 32):
                data = self.request("POST", f"/tbs/{termbase_id}/lookupterms", json={"SourceLanguage": source_language, "TargetLanguage": target_language, "Segments": [f"<seg>{_xml_escape(x)}</seg>" for x in segments[offset:offset + 32]]})
                _collect_matches(data, matches)
        return matches


def _xml_escape(value: str) -> str:
    return value.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;").replace("'", "&apos;")


def _collect_matches(value: Any, output: list[dict[str, Any]]) -> None:
    if isinstance(value, dict):
        source = value.get("Source") or value.get("source") or value.get("SourceTerm") or value.get("sourceTerm")
        target = value.get("Target") or value.get("target") or value.get("TargetTerm") or value.get("targetTerm")
        if source and target:
            output.append({"source": str(source), "targets": [str(target)], "entry_note": str(value.get("EntryNote") or value.get("entry_note") or value.get("Note") or "")})
        for child in value.values():
            _collect_matches(child, output)
    elif isinstance(value, list):
        for child in value:
            _collect_matches(child, output)


def list_termbases(query: str = "") -> list[dict[str, Any]]:
    global _resource_cache
    now = time.monotonic()
    with _cache_lock:
        cached = _resource_cache
    if cached is None or now - cached[0] > 300:
        credentials = _load_credentials()
        if not credentials:
            raise MemoQError("尚未绑定 memoQ 账号")
        client = MemoQClient(credentials)
        try:
            resources = client.termbases()
        finally:
            client.close()
        with _cache_lock:
            _resource_cache = (now, resources)
    else:
        resources = cached[1]
    needle = query.strip().casefold()
    return [item for item in resources if not needle or needle in f"{item['name']} {item['project']}".casefold()]


def lookup_terms(termbase_ids: list[str], source_language: str, target_language: str, segments: list[str]) -> list[dict[str, Any]]:
    credentials = _load_credentials()
    if not credentials:
        raise MemoQError("尚未绑定 memoQ 账号")
    client = MemoQClient(credentials)
    try:
        return client.lookup([str(x) for x in termbase_ids if str(x).strip()], _language_code(source_language), _language_code(target_language), segments)
    finally:
        client.close()


def lookup_terms_snapshot(termbase_ids: list[str], source_language: str, target_language: str,
                          segments: list[str]) -> tuple[list[dict[str, Any]], str]:
    """Fetch live library metadata and matched content in one authenticated client."""
    credentials = _load_credentials()
    if not credentials:
        raise MemoQError('尚未绑定 memoQ 账号')
    client = MemoQClient(credentials)
    try:
        selected = set(termbase_ids)
        resources = {r['id']: r['revision'] for r in client.termbases() if r['id'] in selected}
        if set(resources) != selected:
            raise MemoQError('部分术语库不存在或无访问权限')
        revision = hashlib.sha256(json.dumps(resources, sort_keys=True).encode()).hexdigest()
        pairs = client.lookup(sorted(selected), _language_code(source_language), _language_code(target_language), segments)
        return pairs, revision
    finally:
        client.close()


def _language_code(value: str) -> str:
    aliases = {
        "自动": "", "自动检测": "", "auto": "", "无": "", "无源文": "",
        "简体中文": "zho-CN", "繁体中文": "zho-TW", "英语": "eng", "日语": "jpn", "韩语": "kor",
        "法语": "fra", "德语": "deu", "西班牙语": "spa", "葡萄牙语": "por", "意大利语": "ita",
        "俄语": "rus", "阿拉伯语": "ara", "泰语": "tha", "越南语": "vie", "印尼语": "ind",
    }
    return aliases.get(str(value or "").strip(), str(value or "").strip())
