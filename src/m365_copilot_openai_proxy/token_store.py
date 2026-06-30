from __future__ import annotations

import base64
import json
import os
import threading
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

SUBSTRATE_AUDIENCE_PREFIX = "https://substrate.office.com/"

# Token storage paths — initialized lazily via init_token_dir() or from TOKEN_DIR env var
_TOKEN_DIR: Path | None = None
_TOKEN_FILE: Path | None = None
_ENV_PATH = Path(".env")


def _get_token_dir() -> Path:
    global _TOKEN_DIR
    if _TOKEN_DIR is None:
        _TOKEN_DIR = Path(os.environ.get("TOKEN_DIR", "/home/app/token"))
    return _TOKEN_DIR


def _get_token_file() -> Path:
    global _TOKEN_FILE
    if _TOKEN_FILE is None:
        _TOKEN_FILE = _get_token_dir() / "token"
    return _TOKEN_FILE


def init_token_dir(token_dir: str) -> None:
    """Initialize token directory from Settings (called once at app startup)."""
    global _TOKEN_DIR, _TOKEN_FILE
    _TOKEN_DIR = Path(token_dir)
    _TOKEN_FILE = _TOKEN_DIR / "token"


def decode_jwt_payload(token: str) -> dict[str, Any]:
    payload = token.split(".")[1]
    payload += "=" * (-len(payload) % 4)
    return json.loads(base64.urlsafe_b64decode(payload))


def is_substrate_token_claims(claims: dict[str, Any]) -> bool:
    return str(claims.get("aud", "")).startswith(SUBSTRATE_AUDIENCE_PREFIX)


class AccessTokenStore:
    def __init__(self, token: str, env_path: Path | str = ".env"):
        self._token = token
        self._env_path = Path(env_path)
        self._mtime_ns = self._read_mtime()
        self._lock = threading.RLock()

    def get(self) -> str:
        with self._lock:
            self._reload_if_changed()
            return self._token

    def status(self) -> dict[str, Any]:
        token = self.get()
        now = time.time()
        try:
            claims = decode_jwt_payload(token)
            if not is_substrate_token_claims(claims):
                return {
                    "valid": False,
                    "error": "Access token is not a substrate.office.com token.",
                    "expires_at": None,
                    "seconds_remaining": 0,
                }
            expires_at = int(claims["exp"])
        except Exception as exc:
            return {
                "valid": False,
                "error": f"Cannot decode access token: {exc}",
                "expires_at": None,
                "seconds_remaining": 0,
            }

        seconds_remaining = max(0, expires_at - int(now))
        return {
            "valid": seconds_remaining > 0,
            "expires_at": datetime.fromtimestamp(expires_at, tz=timezone.utc).isoformat(),
            "seconds_remaining": seconds_remaining,
        }

    def _reload_if_changed(self) -> None:
        mtime_ns = self._read_mtime()
        if mtime_ns is None or mtime_ns == self._mtime_ns:
            return
        token = read_token()
        if token:
            self._token = token
            self._mtime_ns = mtime_ns

    def _read_mtime(self) -> int | None:
        # Prefer isolated token file mtime
        try:
            return _TOKEN_FILE.stat().st_mtime_ns
        except FileNotFoundError:
            pass
        # Fallback: .env file mtime
        try:
            return self._env_path.stat().st_mtime_ns
        except FileNotFoundError:
            return None


def read_token() -> str | None:
    """Read token from isolated token file first, then fall back to .env."""
    # Try isolated token file (TOKEN_DIR volume)
    try:
        token = _get_token_file().read_text(encoding="utf-8").strip()
        if token:
            return token
    except FileNotFoundError:
        pass
    # Fallback: read from .env for backward compatibility
    return _read_env_token(_ENV_PATH)


def write_token(token: str) -> None:
    """Write token to isolated token file on TOKEN_DIR volume."""
    token_dir = _get_token_dir()
    token_file = _get_token_file()
    token_dir.mkdir(parents=True, exist_ok=True)
    token_file.write_text(token, encoding="utf-8")
    try:
        token_file.chmod(0o600)
    except OSError:
        pass


def write_username(username: str) -> None:
    """Write username to the token directory for persistence across restarts."""
    token_dir = _get_token_dir()
    username_file = token_dir / "username"
    token_dir.mkdir(parents=True, exist_ok=True)
    username_file.write_text(username, encoding="utf-8")


def read_username() -> str:
    """Read persisted username from the token directory."""
    try:
        return (_get_token_dir() / "username").read_text(encoding="utf-8").strip()
    except FileNotFoundError:
        return ""


def _read_env_token(path: Path) -> str | None:
    try:
        text = path.read_text(encoding="utf-8")
    except FileNotFoundError:
        return None

    for line in text.splitlines():
        stripped = line.strip()
        if not stripped or stripped.startswith("#"):
            continue
        key, sep, value = stripped.partition("=")
        if sep and key.strip() == "M365_ACCESS_TOKEN":
            return _clean_env_value(value)
    return None


def _clean_env_value(value: str) -> str:
    value = value.strip()
    if len(value) >= 2 and value[0] == value[-1] and value[0] in {"'", '"'}:
        return value[1:-1]
    return value
