from __future__ import annotations

import json
import secrets
import threading
import time
import uuid
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any


@dataclass
class ApiKey:
    """A single API key bound to one account (scheme B: key -> one account).

    Each key carries its own tone / tool_prompt / system_prompt so different
    users get independent conversation modes and prompt tuning while sharing
    the multi-tenant proxy. Disabled keys are rejected at the auth middleware.
    """

    id: str = field(default_factory=lambda: "key_" + uuid.uuid4().hex[:12])
    key: str = field(default_factory=lambda: "sk-" + secrets.token_urlsafe(32))
    name: str = ""
    account_id: str = ""
    enabled: bool = True
    tone: str = "Magic"
    tool_prompt: str = ""
    system_prompt: str = ""
    created_at: float = field(default_factory=time.time)
    updated_at: float = field(default_factory=time.time)


class KeyStore:
    """Thread-safe API key table with best-effort JSON persistence.

    Lookups by the raw key string are O(1) via an in-memory index that is
    rebuilt on every mutation so the auth middleware stays fast.
    """

    def __init__(self, persist_path: str | Path | None = None):
        self._keys: dict[str, ApiKey] = {}  # id -> ApiKey
        self._by_secret: dict[str, str] = {}  # raw key string -> id
        self._lock = threading.RLock()
        self._persist_path = Path(persist_path) if persist_path else None
        if self._persist_path is not None:
            self._load()

    # ------------------------------------------------------------------ IO
    def _reindex(self) -> None:
        self._by_secret = {k.key: k.id for k in self._keys.values()}

    def _load(self) -> None:
        try:
            data = json.loads(self._persist_path.read_text(encoding="utf-8"))
        except (FileNotFoundError, json.JSONDecodeError, OSError):
            return
        if not isinstance(data, dict):
            return
        for key_id, raw in data.items():
            if not isinstance(raw, dict):
                continue
            try:
                self._keys[key_id] = ApiKey(
                    id=raw.get("id", key_id),
                    key=raw["key"],
                    name=raw.get("name", ""),
                    account_id=raw.get("account_id", ""),
                    enabled=bool(raw.get("enabled", True)),
                    tone=raw.get("tone", "Magic"),
                    tool_prompt=raw.get("tool_prompt", ""),
                    system_prompt=raw.get("system_prompt", ""),
                    created_at=float(raw.get("created_at", time.time())),
                    updated_at=float(raw.get("updated_at", time.time())),
                )
            except (KeyError, TypeError, ValueError):
                continue
        self._reindex()

    def _save(self) -> None:
        if self._persist_path is None:
            return
        with self._lock:
            data = {key_id: asdict(k) for key_id, k in self._keys.items()}
        try:
            self._persist_path.parent.mkdir(parents=True, exist_ok=True)
            tmp = self._persist_path.with_suffix(".tmp")
            tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
            tmp.replace(self._persist_path)
        except OSError:
            pass  # Persistence is best-effort; never break a request over a disk error

    # -------------------------------------------------------------- queries
    def get(self, key_id: str) -> ApiKey | None:
        with self._lock:
            return self._keys.get(key_id)

    def resolve(self, raw_key: str) -> ApiKey | None:
        """Look up an ApiKey by its raw secret string (used by auth middleware)."""
        with self._lock:
            key_id = self._by_secret.get(raw_key)
            return self._keys.get(key_id) if key_id else None

    def list(self) -> list[ApiKey]:
        with self._lock:
            return list(self._keys.values())

    def list_for_account(self, account_id: str) -> list[ApiKey]:
        with self._lock:
            return [k for k in self._keys.values() if k.account_id == account_id]

    # -------------------------------------------------------------- mutations
    def add(self, name: str = "", account_id: str = "", tone: str = "Magic") -> ApiKey:
        with self._lock:
            k = ApiKey(name=name, account_id=account_id, tone=tone)
            self._keys[k.id] = k
            self._reindex()
            self._save()
            return k

    def update(self, key_id: str, **fields: Any) -> ApiKey | None:
        """Update mutable fields (name, account_id, enabled, tone, tool_prompt, system_prompt)."""
        allowed = {"name", "account_id", "enabled", "tone", "tool_prompt", "system_prompt"}
        with self._lock:
            k = self._keys.get(key_id)
            if k is None:
                return None
            for name, value in fields.items():
                if name in allowed:
                    setattr(k, name, value)
            k.updated_at = time.time()
            self._save()
            return k

    def remove(self, key_id: str) -> bool:
        with self._lock:
            if key_id in self._keys:
                del self._keys[key_id]
                self._reindex()
                self._save()
                return True
            return False

    def detach_account(self, account_id: str) -> None:
        """Clear the binding on any keys pointing at a removed account."""
        with self._lock:
            changed = False
            for k in self._keys.values():
                if k.account_id == account_id:
                    k.account_id = ""
                    k.updated_at = time.time()
                    changed = True
            if changed:
                self._save()
