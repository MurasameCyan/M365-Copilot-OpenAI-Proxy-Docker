from __future__ import annotations

import json
import threading
import time
import uuid
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any

from .token_store import decode_jwt_payload, is_substrate_token_claims


# Base CDP port for per-account Chromium profiles. Each account gets a unique
# port derived from this base so their debug browsers never collide when the
# refresh scheduler brings one up on demand.
_CDP_PORT_BASE = 9222


@dataclass
class Account:
    """A single M365 Copilot account in the multi-tenant pool.

    Each account owns an isolated Substrate token plus a dedicated Chromium
    profile / CDP port so the refresh scheduler can bring its debug browser up
    on demand (and tear it down afterwards) without colliding with others.
    """

    id: str = field(default_factory=lambda: "acct_" + uuid.uuid4().hex[:12])
    name: str = ""
    token: str = ""
    cdp_port: int = _CDP_PORT_BASE
    # "manual" = token pushed by user (Tampermonkey / paste); "cdp" = auto-captured.
    token_source: str = "manual"
    created_at: float = field(default_factory=time.time)
    updated_at: float = field(default_factory=time.time)

    def token_status(self) -> dict[str, Any]:
        """Decode the JWT and report validity / expiry, mirroring AccessTokenStore.status()."""
        token = self.token
        now = time.time()
        if not token:
            return {"valid": False, "error": "No token", "expires_at": None, "seconds_remaining": 0}
        try:
            claims = decode_jwt_payload(token)
            if not is_substrate_token_claims(claims):
                return {
                    "valid": False,
                    "error": "Token is not a substrate.office.com token.",
                    "expires_at": None,
                    "seconds_remaining": 0,
                }
            expires_at = int(claims["exp"])
        except Exception as exc:  # noqa: BLE001 - report any decode failure to the UI
            return {"valid": False, "error": f"Cannot decode token: {exc}", "expires_at": None, "seconds_remaining": 0}
        seconds_remaining = max(0, expires_at - int(now))
        from datetime import datetime, timezone

        return {
            "valid": seconds_remaining > 0,
            "expires_at": datetime.fromtimestamp(expires_at, tz=timezone.utc).isoformat(),
            "seconds_remaining": seconds_remaining,
        }


class AccountStore:
    """Thread-safe account pool with best-effort JSON persistence."""

    def __init__(self, persist_path: str | Path | None = None):
        self._accounts: dict[str, Account] = {}
        self._lock = threading.RLock()
        self._persist_path = Path(persist_path) if persist_path else None
        if self._persist_path is not None:
            self._load()

    # ------------------------------------------------------------------ IO
    def _load(self) -> None:
        try:
            data = json.loads(self._persist_path.read_text(encoding="utf-8"))
        except (FileNotFoundError, json.JSONDecodeError, OSError):
            return
        if not isinstance(data, dict):
            return
        for acc_id, raw in data.items():
            if not isinstance(raw, dict):
                continue
            try:
                self._accounts[acc_id] = Account(
                    id=raw.get("id", acc_id),
                    name=raw.get("name", ""),
                    token=raw.get("token", ""),
                    cdp_port=int(raw.get("cdp_port", _CDP_PORT_BASE)),
                    token_source=raw.get("token_source", "manual"),
                    created_at=float(raw.get("created_at", time.time())),
                    updated_at=float(raw.get("updated_at", time.time())),
                )
            except (TypeError, ValueError):
                continue

    def _save(self) -> None:
        if self._persist_path is None:
            return
        with self._lock:
            data = {acc_id: asdict(acc) for acc_id, acc in self._accounts.items()}
        try:
            self._persist_path.parent.mkdir(parents=True, exist_ok=True)
            tmp = self._persist_path.with_suffix(".tmp")
            tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
            tmp.replace(self._persist_path)
        except OSError:
            pass  # Persistence is best-effort; never break a request over a disk error

    # -------------------------------------------------------------- queries
    def get(self, acc_id: str) -> Account | None:
        with self._lock:
            return self._accounts.get(acc_id)

    def list(self) -> list[Account]:
        with self._lock:
            return list(self._accounts.values())

    def _next_cdp_port(self) -> int:
        used = {acc.cdp_port for acc in self._accounts.values()}
        port = _CDP_PORT_BASE
        while port in used:
            port += 1
        return port

    # -------------------------------------------------------------- mutations
    def add(self, name: str = "", token: str = "", token_source: str = "manual") -> Account:
        with self._lock:
            acc = Account(
                name=name,
                token=token,
                cdp_port=self._next_cdp_port(),
                token_source=token_source,
            )
            self._accounts[acc.id] = acc
            self._save()
            return acc

    def update_token(self, acc_id: str, token: str, token_source: str | None = None) -> Account | None:
        with self._lock:
            acc = self._accounts.get(acc_id)
            if acc is None:
                return None
            acc.token = token
            if token_source is not None:
                acc.token_source = token_source
            acc.updated_at = time.time()
            self._save()
            return acc

    def rename(self, acc_id: str, name: str) -> Account | None:
        with self._lock:
            acc = self._accounts.get(acc_id)
            if acc is None:
                return None
            acc.name = name
            acc.updated_at = time.time()
            self._save()
            return acc

    def remove(self, acc_id: str) -> bool:
        with self._lock:
            if acc_id in self._accounts:
                del self._accounts[acc_id]
                self._save()
                return True
            return False
