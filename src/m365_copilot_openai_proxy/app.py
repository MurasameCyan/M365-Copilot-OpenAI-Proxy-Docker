from __future__ import annotations

import hashlib
import json
import logging
import os
import re
import secrets
import time
import uuid
from collections.abc import AsyncIterator, Callable
from pathlib import Path

from fastapi import Depends, FastAPI, HTTPException, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse, JSONResponse, RedirectResponse, StreamingResponse, Response

from .config import Settings
from .account_store import Account, AccountStore
from .key_store import ApiKey, KeyStore
from .refresh_scheduler import RefreshScheduler
from .session_store import PersistentSession, PersistentSessionStore
from .substrate_client import SubstrateCopilotClient, SubstrateCopilotError
from .token_store import AccessTokenStore, write_token, write_username, read_username, decode_jwt_payload, is_substrate_token_claims, init_token_dir, write_tone, read_tone, write_tool_prompt, read_tool_prompt, write_system_prompt, read_system_prompt
from .models import AnthropicMessagesRequest, OpenAIChatRequest, OpenAIResponsesRequest
from .translator import translate_anthropic_request, translate_openai_request, translate_responses_request, flatten_content, default_tool_system_prompt

_PERSIST_MODEL_SUFFIX = ":persist"
_SESSION_ID_HEADER = "x-m365-session-id"


def _detect_conversation_session(request: OpenAIChatRequest) -> tuple[str, str]:
    """Auto-detect conversation session from the request messages.

    Returns (session_id, title):
    - session_id: stable hash based on the first user message content
    - title: first ~60 chars of the first user message for display
    When the user starts a new chat in Trae, the first user message changes -> new session.
    Agentic tool-result turns reuse the same first user message -> same session.
    """
    for msg in request.messages:
        if msg.role == "user":
            text = flatten_content(msg.content).strip()
            if text:
                sid = "conv_" + hashlib.sha256(text.encode()).hexdigest()[:12]
                title = text[:60].replace("\n", " ")
                return sid, title
    # Fallback: random session
    return "conv_" + uuid.uuid4().hex[:12], "New conversation"


import re as _re

# Primary: fenced ```tool_call blocks. Fallback: ```json blocks that look like a tool call.
# Note: closing/opening newlines are optional — the model often emits the closing ``` right
# after the JSON (e.g. `}}``` ) with no preceding newline, which would otherwise fail to match.
_TOOL_CALL_RE = _re.compile(
    r"```tool_call\s*(\{.*?\})\s*```",
    _re.DOTALL,
)
_JSON_BLOCK_RE = _re.compile(
    r"```(?:json)?\s*(\{.*?\})\s*```",
    _re.DOTALL,
)


def _coerce_tool_call(obj: dict) -> dict | None:
    """Turn a parsed JSON object into an OpenAI tool_call dict if it looks like one."""
    if not isinstance(obj, dict):
        return None
    # Accept {"name": ..., "arguments": {...}} or common variants
    name = obj.get("name") or obj.get("tool") or obj.get("tool_name") or obj.get("function")
    if not name or not isinstance(name, str):
        return None
    arguments = obj.get("arguments")
    if arguments is None:
        arguments = obj.get("parameters")
    if arguments is None:
        # Treat remaining keys (minus name markers) as the arguments
        arguments = {k: v for k, v in obj.items()
                     if k not in ("name", "tool", "tool_name", "function")}
    if isinstance(arguments, dict):
        arguments = json.dumps(arguments, ensure_ascii=False)
    elif not isinstance(arguments, str):
        arguments = str(arguments)
    return {
        "id": f"call_{uuid.uuid4().hex[:24]}",
        "type": "function",
        "function": {"name": name, "arguments": arguments},
    }


def _extract_tool_calls(text: str) -> list[dict]:
    """Parse tool_call JSON blocks from model text output into OpenAI tool_calls format.

    Tolerant to several formats the M365 Copilot model may emit:
    1. ```tool_call fenced blocks (preferred)
    2. ```json (or bare ```) fenced blocks whose JSON has a "name" key
    """
    calls = []
    matched_spans: list[tuple[int, int]] = []

    # 1. Preferred tool_call blocks
    for m in _TOOL_CALL_RE.finditer(text):
        try:
            obj = json.loads(m.group(1).strip())
        except json.JSONDecodeError:
            continue
        tc = _coerce_tool_call(obj)
        if tc:
            calls.append(tc)
            matched_spans.append(m.span())

    # 2. Fallback: json/plain fenced blocks that look like tool calls
    for m in _JSON_BLOCK_RE.finditer(text):
        # Skip if this span overlaps an already-matched tool_call block
        if any(s <= m.start() < e for s, e in matched_spans):
            continue
        try:
            obj = json.loads(m.group(1).strip())
        except json.JSONDecodeError:
            continue
        tc = _coerce_tool_call(obj)
        if tc:
            calls.append(tc)

    return calls


# Prose fallback: model writes "save as `<path>`" then a fenced code block,
# instead of emitting a tool_call. Synthesize a Write tool_call ONLY when the
# code block's language tag matches the target file's extension — this avoids
# mistaking a usage example (e.g. ```bash python foo.py```) for the file content.
_PROSE_PATH_RE = _re.compile(
    r"`([A-Za-z]:[\\/][^`\n]+?\.[A-Za-z0-9]{1,8}|/[^`\n]+?\.[A-Za-z0-9]{1,8})`"
)
# Capture the language tag (group 1) and the body (group 2).
_PROSE_CODE_RE = _re.compile(r"```([A-Za-z0-9_+#.\-]*)[ \t]*\n(.*?)```", _re.DOTALL)

# Map a file extension to the set of fenced-code-block language tags that count
# as matching content for that extension.
_EXT_LANG = {
    "py": {"python", "py", "python3"},
    "pyw": {"python", "py"},
    "bat": {"bat", "batch", "cmd", "dos", "bat文件"},
    "cmd": {"bat", "batch", "cmd", "dos"},
    "sh": {"bash", "sh", "shell", "zsh"},
    "bash": {"bash", "sh", "shell"},
    "ps1": {"powershell", "ps1", "pwsh", "posh"},
    "js": {"javascript", "js", "node", "jsx"},
    "mjs": {"javascript", "js", "node"},
    "cjs": {"javascript", "js", "node"},
    "ts": {"typescript", "ts", "tsx"},
    "tsx": {"typescript", "tsx", "ts"},
    "jsx": {"javascript", "jsx", "js"},
    "json": {"json", "json5", "jsonc"},
    "html": {"html", "htm", "xhtml"},
    "htm": {"html", "htm"},
    "css": {"css"},
    "scss": {"scss", "sass", "css"},
    "less": {"less", "css"},
    "java": {"java"},
    "kt": {"kotlin", "kt"},
    "c": {"c"},
    "h": {"c", "cpp", "c++"},
    "cpp": {"cpp", "c++", "cxx", "cc"},
    "cc": {"cpp", "c++", "cc"},
    "cs": {"csharp", "cs", "c#"},
    "go": {"go", "golang"},
    "rs": {"rust", "rs"},
    "rb": {"ruby", "rb"},
    "php": {"php"},
    "swift": {"swift"},
    "yml": {"yaml", "yml"},
    "yaml": {"yaml", "yml"},
    "xml": {"xml"},
    "sql": {"sql"},
    "md": {"markdown", "md"},
    "txt": {"text", "txt", "plaintext", ""},
    "toml": {"toml"},
    "ini": {"ini", "cfg", "conf"},
    "cfg": {"ini", "cfg", "conf"},
    "conf": {"ini", "cfg", "conf"},
    "env": {"dotenv", "env", "bash", "sh", ""},
    "dockerfile": {"dockerfile", "docker"},
    "vue": {"vue", "html"},
    "r": {"r"},
    "lua": {"lua"},
    "pl": {"perl", "pl"},
    "scala": {"scala"},
    "dart": {"dart"},
    "gradle": {"gradle", "groovy"},
    "groovy": {"groovy"},
    "makefile": {"makefile", "make"},
}


def _extract_prose_write(text: str, tool_names: set[str]) -> list[dict]:
    """Fallback: synthesize a Write tool_call from a 'save as <path>' + code block prose.

    Strict matching to avoid corrupting files:
    - A Write-like tool must be available.
    - A LOCAL file path (drive letter or absolute unix path, not a URL) with an
      extension must be present.
    - A fenced code block whose language tag matches the file's extension must
      exist. This prevents usage-example blocks (```bash, ```text) from being
      mistaken for the file content and overwriting a correctly written file.
    """
    if not any(n.lower() == "write" for n in tool_names):
        return []

    # Collect candidate local paths (skip URLs).
    file_path = None
    target_ext = None
    for path_m in _PROSE_PATH_RE.finditer(text):
        candidate = path_m.group(1).strip()
        if "://" in candidate or candidate.lower().startswith("http"):
            continue
        ext = candidate.rsplit(".", 1)[-1].lower() if "." in candidate else ""
        if not ext:
            continue
        file_path = candidate
        target_ext = ext
        break
    if not file_path or not target_ext:
        return []

    allowed_langs = _EXT_LANG.get(target_ext)

    # Find a code block whose language matches the target extension.
    best_content = None
    for code_m in _PROSE_CODE_RE.finditer(text):
        lang = (code_m.group(1) or "").strip().lower()
        body = code_m.group(2)
        if allowed_langs is not None:
            if lang in allowed_langs:
                best_content = body
                break
        else:
            # Unknown extension: only accept an exactly-matching language tag.
            if lang == target_ext:
                best_content = body
                break
    if best_content is None:
        return []

    # Trim a single trailing newline that fenced blocks usually carry.
    if best_content.endswith("\n"):
        best_content = best_content[:-1]
    if not best_content.strip():
        return []

    write_name = next((n for n in tool_names if n.lower() == "write"), "Write")
    arguments = json.dumps({"file_path": file_path, "content": best_content}, ensure_ascii=False)
    return [{
        "id": f"call_{uuid.uuid4().hex[:24]}",
        "type": "function",
        "function": {"name": write_name, "arguments": arguments},
    }]


def _strip_tool_call_blocks(text: str) -> str:
    """Remove tool_call code blocks from text, keeping surrounding content."""
    cleaned = _TOOL_CALL_RE.sub("", text)
    # Also strip json/plain blocks that were parsed as tool calls
    def _maybe_strip(m):
        try:
            obj = json.loads(m.group(1).strip())
        except json.JSONDecodeError:
            return m.group(0)
        return "" if _coerce_tool_call(obj) else m.group(0)
    cleaned = _JSON_BLOCK_RE.sub(_maybe_strip, cleaned)
    return cleaned.strip()


# M365 Copilot has a native "generate a file" feature that hosts the file on its
# own object storage (asyncgw/Teams) and returns a download URL, instead of
# emitting our tool_call. From the model's view the task is "done", so prompt
# rules alone can't stop it. We detect this pattern and force a corrective retry.
_FILE_CLAIM_URL_RE = _re.compile(
    r"https?://[^\s`)]+?\.(?:py|js|ts|tsx|jsx|json|txt|md|html?|css|sh|bat|ps1|"
    r"java|kt|c|cpp|cc|h|cs|go|rs|rb|php|swift|ya?ml|xml|sql|ini|toml|cfg)\b",
    _re.IGNORECASE,
)
# Phrases that claim a file was produced (zh + en).
_FILE_CLAIM_PHRASE_RE = _re.compile(
    r"已生成|已创建|已保存|已写入|已经生成|已经创建|生成脚本|生成了|创建了|保存到|"
    r"file (?:created|saved|generated|written)|created the file|saved to|generated the",
    _re.IGNORECASE,
)


def _looks_like_fake_file_claim(text: str) -> bool:
    """True if the model claims to have produced a file but emitted no tool_call.

    Two triggers:
    1. A hosted attachment URL pointing at a code/text file (M365 native file gen).
    2. A "file created/生成" style phrase.
    The caller only invokes this when NO tool_call was parsed from the response.
    """
    if not text:
        return False
    if _FILE_CLAIM_URL_RE.search(text):
        return True
    if _FILE_CLAIM_PHRASE_RE.search(text):
        return True
    return False


_RETRY_INSTRUCTION = (
    "[SYSTEM] Your previous reply did NOT create any file on the host. "
    "You may have used a hosted attachment link or an out-of-band file feature — that does NOT work here; "
    "the host only creates files when you emit a tool_call block. "
    "Re-do the task NOW: output ONLY a fenced ```tool_call block whose JSON is "
    '{"name": "Write", "arguments": {"file_path": "<the exact path the user gave>", "content": "<the FULL file body>"}}. '
    "No prose, no links, no usage examples — just the tool_call block with the complete file content.[/SYSTEM]"
)


def _update_username_from_token(token: str, state) -> None:
    """Extract username from JWT claims and persist it if not already set."""
    if getattr(state, 'username', None) and len(state.username) > 1:
        return  # Already have a valid username, keep it
    try:
        claims = decode_jwt_payload(token)
        name = claims.get("name") or claims.get("upn") or ""
        if isinstance(name, str):
            name = name.strip()
            # If upn is email, take the local part
            if "@" in name and " " not in name:
                name = name.split("@")[0]
        if name and len(name) > 1:
            state.username = name
            write_username(name)
    except Exception:
        pass


def create_app(
    settings: Settings | None = None,
    copilot_client_factory: Callable[[], SubstrateCopilotClient] | None = None,
) -> FastAPI:
    app = FastAPI(title="Ciallo Ms-365 OpenAI Proxy")
    resolved_settings = settings or Settings()
    init_token_dir(resolved_settings.token_dir)
    app.state.settings = resolved_settings
    app.state.token_store = AccessTokenStore(resolved_settings.access_token)
    app.state.session_store = PersistentSessionStore(
        persist_path=Path(resolved_settings.token_dir) / "sessions.json"
    )  # Persist to mounted volume so conversations survive container restarts
    # Multi-tenant stores: account pool (each owns an isolated token + Chromium
    # profile/CDP port) and API key table (each key bound to one account, scheme B).
    app.state.account_store = AccountStore(
        persist_path=Path(resolved_settings.token_dir) / "accounts.json"
    )
    app.state.key_store = KeyStore(
        persist_path=Path(resolved_settings.token_dir) / "keys.json"
    )
    # On-demand token refresh scheduler: brings one account's Chromium up at a
    # time (serial), captures a fresh token via CDP, then tears it down. Keeps
    # peak memory close to single-tenant even with many accounts.
    app.state.refresh_scheduler = RefreshScheduler(
        app.state.account_store,
        profile_root=Path(resolved_settings.token_dir) / "profiles",
    )
    app.state.call_log: list[dict] = []  # API call log for web UI display
    app.state.captured_payloads: list[dict] = []  # Substrate chat payloads captured via get_token.js for mode comparison
    app.state.auto_refresh_enabled = False  # On-demand: only refresh when /v1/ requests come in
    app.state.last_request_time = 0  # 0 means never received any /v1/ request
    app.state.idle_timeout_minutes = resolved_settings.idle_timeout_minutes
    app.state.username = read_username()  # Restore persisted username (set via get_token.js push or CDP extraction)
    app.state.current_tone = read_tone() or "Magic"  # Restore persisted conversation tone (mode), default "Magic" (Auto)
    app.state.tool_prompt = read_tool_prompt()  # Restore persisted user-defined extra tool-call instruction
    app.state.system_prompt = read_system_prompt()  # Restore persisted system-level tool-call instruction override (empty = use default)
    if not resolved_settings.api_key:
        print("WARNING: API_KEY is not set. All /v1/ API endpoints are open without authentication. Set API_KEY in .env to secure your instance.")
    _admin_secret = resolved_settings.admin_password or resolved_settings.api_key
    if not _admin_secret:
        print("WARNING: Neither API_KEY nor ADMIN_PASSWORD is set. Web admin page is open without authentication. Set ADMIN_PASSWORD in .env to secure it.")

    # Generate a random admin session token instead of deterministic hash
    _admin_session_token: str | None = secrets.token_hex(32) if _admin_secret else None

    # Login rate limiting: track failed attempts by client IP
    _login_failures: dict[str, list[float]] = {}
    _LOGIN_RATE_LIMIT = 5       # max failures
    _LOGIN_LOCKOUT_SEC = 60.0   # lockout duration

    app.state.copilot_client_factory = copilot_client_factory or (
        lambda token=None, tone=None, tool_prompt=None: SubstrateCopilotClient(
            token if token is not None else app.state.token_store.get(),
            resolved_settings.time_zone,
            tone if tone is not None else getattr(app.state, 'current_tone', 'Magic'),
            tool_prompt if tool_prompt is not None else getattr(app.state, 'tool_prompt', ''),
        )
    )

    def _is_admin_authenticated(request: Request) -> bool:
        """Check if the request has a valid admin auth cookie."""
        if not _admin_secret:
            return True
        if _admin_session_token is None:
            return False
        cookie_val = request.cookies.get("admin_auth", "")
        return secrets.compare_digest(cookie_val, _admin_session_token)

    def _require_admin(request: Request):
        """Check admin cookie auth; return error response or None."""
        if _admin_secret and not _is_admin_authenticated(request):
            return JSONResponse({"error": {"message": "Admin authentication required", "type": "auth_error"}}, status_code=401)
        return None

    # CORS: use configurable origin whitelist (comma-separated ALLOWED_ORIGINS env var)
    _allowed_origins_raw = os.environ.get("ALLOWED_ORIGINS", "").strip()
    _allowed_origins = [o.strip() for o in _allowed_origins_raw.split(",") if o.strip()] if _allowed_origins_raw else ["*"]
    _cors_is_wildcard = "*" in _allowed_origins

    # CORS middleware
    app.add_middleware(
        CORSMiddleware,
        allow_origins=_allowed_origins,
        allow_methods=["GET", "POST", "OPTIONS"],
        allow_headers=["Content-Type", "Authorization", "x-m365-session-id"],
        max_age=86400,
    )

    # API Key authentication middleware (runs after CORS)
    @app.middleware("http")
    async def api_key_auth(request: Request, call_next):
        # Always handle preflight first
        if request.method == "OPTIONS":
            return await call_next(request)
        # Add CORS headers to all responses from this middleware
        def with_cors(resp):
            if _cors_is_wildcard:
                resp.headers["Access-Control-Allow-Origin"] = "*"
            else:
                origin = request.headers.get("origin", "")
                if origin in _allowed_origins:
                    resp.headers["Access-Control-Allow-Origin"] = origin
            resp.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
            resp.headers["Access-Control-Allow-Headers"] = "Content-Type, Authorization, x-m365-session-id"
            resp.headers["Access-Control-Max-Age"] = "86400"
            return resp
        path = request.url.path
        # Track last request time for idle detection & on-demand refresh
        if path.startswith("/v1/"):
            app.state.last_request_time = time.time()

        # Public paths: admin page (own cookie auth), user page, health, and all
        # /admin/* + /user/* endpoints (each does its own cookie/key check).
        if path in ("/", "/admin", "/favicon.ico", "/healthz") or path.startswith("/admin/") or path.startswith("/user/"):
            return await call_next(request)

        auth = request.headers.get("Authorization", "")
        match = re.match(r"^Bearer\s+(.+)$", auth, re.IGNORECASE)
        raw_key = match.group(1) if match else ""

        # Multi-tenant: resolve the API key -> ApiKey -> bound Account.
        key_obj = app.state.key_store.resolve(raw_key) if raw_key else None
        if key_obj is not None:
            if not key_obj.enabled:
                return with_cors(JSONResponse(
                    status_code=401,
                    content={"error": {"message": "API key is disabled", "type": "auth_error"}},
                ))
            account = app.state.account_store.get(key_obj.account_id) if key_obj.account_id else None
            # Bring the bound account's token up to date on demand (serial, one
            # Chromium at a time). No-op if the token is still valid.
            if account is not None and path.startswith("/v1/"):
                try:
                    await app.state.refresh_scheduler.ensure_fresh(account.id)
                    account = app.state.account_store.get(account.id) or account
                except Exception:
                    pass  # Scheduler failures fall back to whatever token we have
            request.state.api_key_obj = key_obj
            request.state.account = account
            return await call_next(request)

        # Legacy single API_KEY fallback (global admin key, no bound account).
        if resolved_settings.api_key and raw_key == resolved_settings.api_key:
            request.state.api_key_obj = None
            request.state.account = None
            return await call_next(request)

        # No auth configured at all (no keys registered and no legacy key): open.
        if not resolved_settings.api_key and not app.state.key_store.list():
            request.state.api_key_obj = None
            request.state.account = None
            return await call_next(request)

        return with_cors(JSONResponse(
            status_code=401,
            content={"error": {"message": "Invalid API key", "type": "auth_error"}},
        ))

    def get_settings() -> Settings:
        return app.state.settings

    def get_copilot_client(raw_request: Request) -> SubstrateCopilotClient:
        try:
            key_obj = getattr(raw_request.state, "api_key_obj", None)
            account = getattr(raw_request.state, "account", None)
            # Per-key overrides: bound account's token + the key's own tone /
            # tool_prompt. Falls back to global app.state when unbound (legacy key).
            token = account.token if account is not None else None
            tone = key_obj.tone if key_obj is not None else None
            tool_prompt = key_obj.tool_prompt if key_obj is not None else None
            return app.state.copilot_client_factory(token=token, tone=tone, tool_prompt=tool_prompt)
        except TypeError:
            # Test-injected factory may take no arguments.
            return app.state.copilot_client_factory()
        except Exception as exc:
            raise HTTPException(status_code=503, detail=str(exc)) from exc

    # Global exception handler — always return JSON (never HTML error pages)
    @app.exception_handler(Exception)
    async def global_exception_handler(request: Request, exc: Exception):
        return JSONResponse(
            status_code=500,
            content={"error": {"message": str(exc), "type": "internal_error"}},
            headers={"Access-Control-Allow-Origin": "*"},
        )

    @app.exception_handler(HTTPException)
    async def http_exception_handler(request: Request, exc: HTTPException):
        return JSONResponse(
            status_code=exc.status_code,
            content={"error": {"message": exc.detail, "type": "http_error"}},
            headers={"Access-Control-Allow-Origin": "*"},
        )

    def _json_err(status: int, message: str, error_type: str = "error") -> JSONResponse:
        """Return a JSON error response with CORS headers."""
        return JSONResponse(
            status_code=status,
            content={"error": {"message": message, "type": error_type}},
            headers={"Access-Control-Allow-Origin": "*"},
        )

    @app.get("/healthz")
    async def healthz() -> dict:
        return {"status": "ok", "token": app.state.token_store.status()}

    @app.get("/admin/token/status")
    async def token_status(request: Request) -> dict:
        err = _require_admin(request)
        if err: return err
        status = app.state.token_store.status()
        status["auto_refresh"] = app.state.auto_refresh_enabled
        status["username"] = (getattr(app.state, 'username', '') or None) if len(getattr(app.state, 'username', '')) > 1 else None
        return status

    @app.post("/admin/token/auto-refresh-toggle")
    async def toggle_auto_refresh(request: Request) -> dict:
        err = _require_admin(request)
        if err: return err
        app.state.auto_refresh_enabled = not app.state.auto_refresh_enabled
        return {"status": "ok", "auto_refresh": app.state.auto_refresh_enabled}

    @app.post("/admin/token/update")
    async def update_token(request: Request) -> dict:
        err = _require_admin(request)
        if err: return err
        body = await request.json()
        token = body.get("token", "").strip()
        username = body.get("username", "").strip()
        if not token:
            return _json_err(400, "Token is empty")
        # Extract token from full WebSocket URL if needed
        match = re.search(r"access_token=([^&\s]+)", token)
        if match:
            token = match.group(1)
        if not token.startswith("eyJ"):
            return _json_err(400, "Not a valid JWT token")
        # Write to isolated token file
        write_token(token)
        # Update in-memory store
        app.state.token_store._token = token
        app.state.token_store._mtime_ns = None
        if username and len(username) > 1:
            app.state.username = username
            write_username(username)
        else:
            _update_username_from_token(token, app.state)
        return {"status": "ok", "message": "Token updated", "token_status": app.state.token_store.status()}

    @app.post("/admin/token/auto-capture")
    async def auto_capture_token(request: Request) -> dict:
        """Auto-capture token from Chromium CDP running inside the container."""
        err = _require_admin(request)
        if err: return err
        import asyncio
        from .cli import _cdp_extract_token
        cdp_port = 9222
        try:
            token = await _cdp_extract_token(cdp_port, allow_nudge=True)
        except Exception as exc:
            return _json_err(502, f"CDP capture failed: {exc}")
        if not token:
            return _json_err(404, "No substrate token found. Make sure M365 Copilot is open and logged in in Chromium.")
        # Write to token file and update in-memory
        write_token(token)
        app.state.token_store._token = token
        app.state.token_store._mtime_ns = None
        _update_username_from_token(token, app.state)
        return {"status": "ok", "message": "Token auto-captured", "token_status": app.state.token_store.status()}

    @app.post("/admin/cookie/inject")
    async def inject_cookie(request: Request) -> dict:
        err = _require_admin(request)
        if err: return err
        body = await request.json()
        cookies = body.get("cookies", [])
        username = body.get("username", "")
        if username and len(str(username).strip()) > 1:
            app.state.username = str(username).strip()
            write_username(str(username).strip())
        if not cookies:
            return _json_err(400, "No cookies provided")
        import asyncio as _async
        import httpx as _httpx
        import websockets as _ws

        cdp_port = 9222
        try:
            async with _httpx.AsyncClient(timeout=3) as client:
                tabs = (await client.get(f"http://localhost:{cdp_port}/json")).json()
        except Exception as exc:
            return _json_err(502, f"Cannot connect to Chromium CDP: {exc}")

        tab = next((t for t in tabs if t.get("type") == "page" and t.get("url", "").startswith("https://m365.cloud.microsoft/")), None)
        if not tab:
            tab = next((t for t in tabs if t.get("type") == "page"), None)
        if not tab:
            return _json_err(404, "No browser tab found in Chromium")

        injected = 0
        try:
            async with _ws.connect(tab["webSocketDebuggerUrl"]) as ws:
                if "m365.cloud.microsoft" not in tab.get("url", ""):
                    await ws.send(json.dumps({"id": 1, "method": "Page.navigate", "params": {"url": "https://m365.cloud.microsoft/chat"}}))
                    await _async.sleep(3)
                    try:
                        await _async.wait_for(ws.recv(), timeout=2)
                    except (_async.TimeoutError, Exception):
                        pass

                for i, cookie in enumerate(cookies):
                    cookie_params = {
                        "name": cookie.get("name", ""),
                        "value": cookie.get("value", ""),
                        "domain": cookie.get("domain", ".microsoft.com"),
                        "path": cookie.get("path", "/"),
                        "secure": cookie.get("secure", True),
                        "httpOnly": cookie.get("httpOnly", False),
                    }
                    ss = cookie.get("sameSite", "")
                    if ss:
                        ss_cap = ss.capitalize()
                        if ss_cap in ("Strict", "Lax", "None"):
                            cookie_params["sameSite"] = ss_cap
                    # sameSite=None requires secure=true in CDP
                    if cookie_params.get("sameSite") == "None":
                        cookie_params["secure"] = True
                    if cookie.get("expirationDate") or cookie.get("expires"):
                        cookie_params["expires"] = cookie.get("expirationDate") or cookie.get("expires")
                    await ws.send(json.dumps({"id": 100 + i, "method": "Network.setCookie", "params": cookie_params}))
                    try:
                        resp = await _async.wait_for(ws.recv(), timeout=5)
                        result = json.loads(resp)
                        if result.get("result", {}).get("success"):
                            injected += 1
                    except (_async.TimeoutError, Exception):
                        pass

                # Navigate to M365 chat (full load, not just reload)
                await ws.send(json.dumps({"id": 998, "method": "Page.navigate", "params": {"url": "https://m365.cloud.microsoft/chat"}}))
                # Wait for page to load and potentially complete auth redirect
                await _async.sleep(8)
                # Drain any pending CDP messages
                try:
                    while True:
                        await _async.wait_for(ws.recv(), timeout=0.5)
                except (_async.TimeoutError, Exception):
                    pass
        except Exception as exc:
            return _json_err(502, f"CDP cookie injection failed: {exc}")

        return {"status": "ok", "message": f"Injected {injected}/{len(cookies)} cookies. Page navigating to M365...", "injected": injected, "total": len(cookies)}

    @app.get("/admin/chromium/login-status")
    async def chromium_login_status(request: Request) -> dict:
        err = _require_admin(request)
        if err: return err
        import httpx as _httpx
        import websockets as _ws
        import asyncio as _async

        cdp_port = 9222
        # Check CDP availability
        try:
            async with _httpx.AsyncClient(timeout=3) as client:
                tabs = (await client.get(f"http://localhost:{cdp_port}/json")).json()
        except Exception:
            return {"chromium_running": False, "logged_in": False, "url": None, "title": None, "cookies": []}

        # Find M365 tab
        tab = next((t for t in tabs if t.get("type") == "page" and "m365.cloud.microsoft" in t.get("url", "")), None)
        if not tab:
            tab = next((t for t in tabs if t.get("type") == "page"), None)

        if not tab:
            return {"chromium_running": True, "logged_in": False, "url": None, "title": None, "cookies": []}

        # Try to detect login state via CDP
        logged_in = False
        page_title = tab.get("title", "")
        page_url = tab.get("url", "")
        cookie_details = []
        # Extract username: prefer CDP extraction, fallback to app.state.username (set by get_token.js push)
        username = getattr(app.state, 'username', '') or None
        try:
            async with _ws.connect(tab["webSocketDebuggerUrl"]) as ws:
                # Get page cookies for M365 domain
                await ws.send(json.dumps({"id": 1, "method": "Network.getCookies", "params": {"urls": ["https://m365.cloud.microsoft", "https://login.microsoftonline.com", "https://microsoft.com", "https://office.com"]}}))
                resp = await _async.wait_for(ws.recv(), timeout=5)
                result = json.loads(resp)
                cookies = result.get("result", {}).get("cookies", [])
                cookie_details = [{"name": c.get("name", ""), "domain": c.get("domain", ""), "httpOnly": c.get("httpOnly", False), "secure": c.get("secure", False)} for c in cookies]
                # Check for authentication cookies
                auth_cookie_names = {"SignInStateCookie", "ESTSAUTH", "ESTSAUTHPERSISTENT", "brcap", "MUID"}
                found = any(c.get("name", "") in auth_cookie_names for c in cookies)
                # Also check URL — if redirected to login page, not logged in
                if "login.microsoftonline.com" in page_url or "login.windows.net" in page_url:
                    logged_in = False
                elif found or "m365.cloud.microsoft/chat" in page_url:
                    logged_in = True
                else:
                    logged_in = False
                # Extract username from page JS (try multiple sources)
                if logged_in:
                    try:
                        _USER_JS = """(() => {
                            try { const s = sessionStorage.getItem('ms-m365-shell-session-data'); if (s) { const d = JSON.parse(s); if (d && d.userDisplayName) return d.userDisplayName; if (d && d.upn) return d.upn.split('@')[0]; } } catch {}
                            try {
                                const av = document.querySelectorAll('[data-testid="header-person-menu"], [data-testid="persona"], button[aria-label*="Account"], button[aria-label*="Manager"], [role="button"][aria-label*="for "], [role="button"][title*="for "], [role="button"][aria-label*="概要"]');
                                for (const el of av) {
                                    const a = el.getAttribute('aria-label') || el.getAttribute('title') || '';
                                    const m = a.match(/(?:for\\s+|的[帐账]户(?:管理器)?[：:]?\\s*)(.+)/i) || a.match(/^(.+?)(?:\\s*\\(|\\s*-|\\s*的)/);
                                    if (m && m[1] && m[1].trim().length > 1 && m[1].trim().length < 80) return m[1].trim();
                                    if (a && a.length > 1 && a.length < 80 && !/^(home|copilot|apps|chat|create|menu|back|close)$/i.test(a)) return a.trim();
                                }
                            } catch {}
                            try {
                                const els = document.querySelectorAll('[data-testid="header-person-menu"], [data-testid="persona"], [aria-label*="Account"], [aria-label*="Profiles"], .ms-Icon--People, button[title*="Account"], span[id*="person"]');
                                for (const el of els) { const t = el.textContent.trim(); if (t && t.length > 1 && t.length < 80) return t; }
                            } catch {}
                            try {
                                const profile = document.querySelector('div[class*="persona"] span, div[class*="UserProfile"] span, img[alt]'); if (profile) { const a = profile.getAttribute('alt') || profile.textContent; if (a && a.trim() && a.trim().length > 1) return a.trim(); } } catch {}
                            try {
                                const fus = document.querySelectorAll('span.fui-Text, span[class*="fai-bebop"]');
                                const skip = /^(home|copilot|apps|chat|create|new|file|edit|view|insert|format|tools|help|share|send|save|open|close|settings|back|next|previous|more|menu|search|filter|sort|refresh|delete|cancel|ok|yes|no)$/i;
                                for (const el of fus) { const t = el.textContent.trim(); if (t && t.length > 1 && t.length < 80 && !skip.test(t)) return t; }
                            } catch {}
                            return null;
                        })()"""
                        next_id = 2
                        # Drain any pending CDP messages before sending
                        while True:
                            try:
                                await _async.wait_for(ws.recv(), timeout=0.1)
                            except (_async.TimeoutError, Exception):
                                break
                        await ws.send(json.dumps({"id": next_id, "method": "Runtime.evaluate", "params": {"expression": _USER_JS}}))
                        # Wait for the specific response by id
                        deadline = _async.get_event_loop().time() + 3
                        while _async.get_event_loop().time() < deadline:
                            raw_msg = await _async.wait_for(ws.recv(), timeout=2)
                            msg = json.loads(raw_msg)
                            if msg.get("id") == next_id:
                                name_val = msg.get("result", {}).get("result", {}).get("value")
                                if name_val and isinstance(name_val, str) and len(name_val.strip()) > 1:
                                    username = name_val.strip()
                                    app.state.username = username
                                    write_username(username)
                                break
                    except Exception:
                        pass
        except Exception:
            logged_in = "m365.cloud.microsoft/chat" in page_url

        # Fallback to persisted username if CDP extraction returned nothing
        if not username:
            username = getattr(app.state, 'username', '') or None

        return {
            "chromium_running": True,
            "logged_in": logged_in,
            "username": username,
            "url": page_url,
            "title": page_title,
            "cookies": cookie_details,
        }

    @app.post("/admin/chromium/logout")
    async def chromium_logout(request: Request) -> dict:
        """Logout from M365 in Chromium by clearing cookies and navigating to login page."""
        err = _require_admin(request)
        if err: return err
        import httpx as _httpx
        import websockets as _ws
        import asyncio as _async

        cdp_port = 9222
        try:
            async with _httpx.AsyncClient(timeout=3) as client:
                tabs = (await client.get(f"http://localhost:{cdp_port}/json")).json()
        except Exception as exc:
            return _json_err(502, f"Cannot connect to Chromium CDP: {exc}")

        tab = next((t for t in tabs if t.get("type") == "page" and "m365.cloud.microsoft" in t.get("url", "")), None)
        if not tab:
            tab = next((t for t in tabs if t.get("type") == "page"), None)
        if not tab:
            return _json_err(404, "No browser tab found in Chromium")

        try:
            async with _ws.connect(tab["webSocketDebuggerUrl"]) as ws:
                # Clear all cookies for Microsoft domains
                await ws.send(json.dumps({"id": 1, "method": "Network.getCookies", "params": {"urls": ["https://m365.cloud.microsoft", "https://login.microsoftonline.com", "https://microsoft.com", "https://office.com"]}}))
                resp = await _async.wait_for(ws.recv(), timeout=5)
                result = json.loads(resp)
                cookies = result.get("result", {}).get("cookies", [])
                cleared = 0
                for i, c in enumerate(cookies):
                    await ws.send(json.dumps({"id": 100 + i, "method": "Network.deleteCookies", "params": {"name": c.get("name", ""), "domain": c.get("domain", "")}}))
                    try:
                        await _async.wait_for(ws.recv(), timeout=2)
                        cleared += 1
                    except Exception:
                        pass
                # Clear sessionStorage and localStorage
                await ws.send(json.dumps({"id": 500, "method": "Runtime.evaluate", "params": {"expression": "sessionStorage.clear();localStorage.clear();true"}}))
                try:
                    await _async.wait_for(ws.recv(), timeout=3)
                except Exception:
                    pass
                # Navigate to login page
                await ws.send(json.dumps({"id": 501, "method": "Page.navigate", "params": {"url": "https://m365.cloud.microsoft/chat"}}))
                try:
                    await _async.wait_for(ws.recv(), timeout=5)
                except Exception:
                    pass
        except Exception as exc:
            return _json_err(502, f"CDP logout failed: {exc}")

        app.state.username = ""
        write_username("")
        return {"status": "ok", "message": f"Logged out. Cleared {cleared}/{len(cookies)} cookies.", "username": ""}

    @app.post("/admin/login")
    async def admin_login(request: Request) -> Response:
        # Rate limiting: check if client IP is locked out
        client_ip = request.client.host if request.client else "unknown"
        now = time.time()
        failures = _login_failures.get(client_ip, [])
        # Remove expired entries
        failures = [t for t in failures if now - t < _LOGIN_LOCKOUT_SEC]
        _login_failures[client_ip] = failures
        if len(failures) >= _LOGIN_RATE_LIMIT:
            return JSONResponse({"error": {"message": "Too many login attempts, try again later", "type": "auth_error"}}, status_code=429)

        body = await request.json()
        password = body.get("password", "")
        if _admin_secret and secrets.compare_digest(password, _admin_secret):
            resp = JSONResponse({"status": "ok"})
            resp.set_cookie("admin_auth", _admin_session_token, max_age=86400 * 7, httponly=True, samesite="lax", secure=bool(int(os.environ.get("ADMIN_COOKIE_SECURE", "0"))), path="/")
            return resp
        # Record failed attempt
        _login_failures.setdefault(client_ip, []).append(now)
        return JSONResponse({"error": {"message": "Wrong password", "type": "auth_error"}}, status_code=401)

    @app.get("/admin/call-log")
    async def get_call_log(request: Request) -> dict:
        err = _require_admin(request)
        if err: return err
        return {"logs": getattr(app.state, 'call_log', [])}

    @app.post("/admin/capture-payload")
    async def capture_payload(request: Request) -> dict:
        err = _require_admin(request)
        if err: return err
        body = await request.json()
        payloads = body.get("payloads", [])
        if not isinstance(payloads, list):
            return _json_err(400, "payloads must be a list")
        app.state.captured_payloads = payloads[:20]
        return {"status": "ok", "count": len(app.state.captured_payloads)}

    @app.get("/admin/capture-payload")
    async def get_captured_payload(request: Request) -> dict:
        err = _require_admin(request)
        if err: return err
        return {"payloads": getattr(app.state, 'captured_payloads', [])}

    # Conversation tone (mode) options discovered from M365 Copilot's mode picker.
    # The `tone` field in the Substrate chat payload controls which model/mode is used.
    _TONE_OPTIONS = [
        {"value": "Magic", "label": "自动 / Auto", "label_zh": "自动", "label_en": "Auto"},
        {"value": "Chat", "label": "快速答复 / Fast", "label_zh": "快速答复", "label_en": "Fast"},
        {"value": "Reasoning", "label": "深度思考 / Think", "label_zh": "深度思考", "label_en": "Think"},
        {"value": "Gpt_5_5_Chat", "label": "GPT 5.5 快速响应", "label_zh": "GPT 5.5 快速响应", "label_en": "GPT 5.5 Fast"},
        {"value": "Gpt_5_5_Reasoning", "label": "GPT 5.5 深度思考", "label_zh": "GPT 5.5 深度思考", "label_en": "GPT 5.5 Think"},
        {"value": "Gpt_5_2_Chat", "label": "GPT 5.2 快速响应", "label_zh": "GPT 5.2 快速响应", "label_en": "GPT 5.2 Fast"},
        {"value": "Gpt_5_2_Reasoning", "label": "GPT 5.2 深度思考", "label_zh": "GPT 5.2 深度思考", "label_en": "GPT 5.2 Think"},
    ]
    _TONE_VALUES = {o["value"] for o in _TONE_OPTIONS}

    @app.get("/admin/tone")
    async def get_tone(request: Request) -> dict:
        err = _require_admin(request)
        if err: return err
        return {"tone": getattr(app.state, 'current_tone', 'Magic'), "options": _TONE_OPTIONS}

    @app.post("/admin/tone")
    async def set_tone(request: Request) -> dict:
        err = _require_admin(request)
        if err: return err
        body = await request.json()
        tone = (body.get("tone") or "").strip()
        if tone not in _TONE_VALUES:
            return _json_err(400, f"Invalid tone. Allowed: {', '.join(sorted(_TONE_VALUES))}")
        app.state.current_tone = tone
        write_tone(tone)
        return {"status": "ok", "tone": tone}

    @app.get("/admin/tool-prompt")
    async def get_tool_prompt(request: Request) -> dict:
        err = _require_admin(request)
        if err: return err
        return {"tool_prompt": getattr(app.state, 'tool_prompt', '')}

    @app.post("/admin/tool-prompt")
    async def set_tool_prompt(request: Request) -> dict:
        err = _require_admin(request)
        if err: return err
        body = await request.json()
        prompt = body.get("tool_prompt")
        if not isinstance(prompt, str):
            return _json_err(400, "tool_prompt must be a string")
        prompt = prompt[:4000]  # cap length to avoid bloating every request
        app.state.tool_prompt = prompt
        write_tool_prompt(prompt)
        return {"status": "ok", "tool_prompt": prompt}

    @app.get("/admin/system-prompt")
    async def get_system_prompt(request: Request) -> dict:
        err = _require_admin(request)
        if err: return err
        # Return the saved override plus the built-in default (for restore/initial fill).
        return {
            "system_prompt": getattr(app.state, 'system_prompt', ''),
            "default": default_tool_system_prompt(),
        }

    @app.post("/admin/system-prompt")
    async def set_system_prompt(request: Request) -> dict:
        err = _require_admin(request)
        if err: return err
        body = await request.json()
        prompt = body.get("system_prompt")
        if not isinstance(prompt, str):
            return _json_err(400, "system_prompt must be a string")
        prompt = prompt[:8000]  # cap length to avoid bloating every request
        app.state.system_prompt = prompt
        write_system_prompt(prompt)
        return {"status": "ok", "system_prompt": prompt}

    # ============================ Multi-tenant admin API ============================
    def _account_public(acc: Account) -> dict:
        """Serialize an account for the admin UI (never leak the raw token)."""
        return {
            "id": acc.id,
            "name": acc.name,
            "cdp_port": acc.cdp_port,
            "token_source": acc.token_source,
            "has_token": bool(acc.token),
            "token_status": acc.token_status(),
            "key_count": len(app.state.key_store.list_for_account(acc.id)),
            "created_at": acc.created_at,
            "updated_at": acc.updated_at,
        }

    def _key_public(k: ApiKey) -> dict:
        """Serialize an API key for the admin UI (raw key shown so admin can copy)."""
        acc = app.state.account_store.get(k.account_id) if k.account_id else None
        return {
            "id": k.id,
            "key": k.key,
            "name": k.name,
            "account_id": k.account_id,
            "account_name": acc.name if acc is not None else "",
            "enabled": k.enabled,
            "tone": k.tone,
            "tool_prompt": k.tool_prompt,
            "system_prompt": k.system_prompt,
            "created_at": k.created_at,
            "updated_at": k.updated_at,
        }

    @app.get("/admin/accounts")
    async def list_accounts(request: Request) -> dict:
        err = _require_admin(request)
        if err: return err
        return {"accounts": [_account_public(a) for a in app.state.account_store.list()]}

    @app.post("/admin/accounts")
    async def add_account(request: Request) -> dict:
        err = _require_admin(request)
        if err: return err
        body = await request.json()
        name = str(body.get("name", "")).strip()
        token = str(body.get("token", "")).strip()
        if token:
            match = re.search(r"access_token=([^&\s]+)", token)
            token = match.group(1) if match else token
            try:
                claims = decode_jwt_payload(token)
                if not is_substrate_token_claims(claims):
                    return _json_err(400, "Token is not a substrate.office.com token")
            except Exception:
                return _json_err(400, "Not a valid JWT token")
        acc = app.state.account_store.add(name=name, token=token,
                                          token_source="manual" if token else "cdp")
        return {"status": "ok", "account": _account_public(acc)}

    @app.post("/admin/accounts/{acc_id}/token")
    async def update_account_token(acc_id: str, request: Request) -> dict:
        err = _require_admin(request)
        if err: return err
        body = await request.json()
        token = str(body.get("token", "")).strip()
        if not token:
            return _json_err(400, "Token is empty")
        match = re.search(r"access_token=([^&\s]+)", token)
        token = match.group(1) if match else token
        try:
            claims = decode_jwt_payload(token)
            if not is_substrate_token_claims(claims):
                return _json_err(400, "Token is not a substrate.office.com token")
        except Exception:
            return _json_err(400, "Not a valid JWT token")
        acc = app.state.account_store.update_token(acc_id, token, token_source="manual")
        if acc is None:
            return _json_err(404, "Account not found")
        return {"status": "ok", "account": _account_public(acc)}

    @app.post("/admin/accounts/{acc_id}/rename")
    async def rename_account(acc_id: str, request: Request) -> dict:
        err = _require_admin(request)
        if err: return err
        body = await request.json()
        name = str(body.get("name", "")).strip()
        acc = app.state.account_store.rename(acc_id, name)
        if acc is None:
            return _json_err(404, "Account not found")
        return {"status": "ok", "account": _account_public(acc)}

    @app.post("/admin/accounts/{acc_id}/refresh")
    async def refresh_account(acc_id: str, request: Request) -> dict:
        err = _require_admin(request)
        if err: return err
        if app.state.account_store.get(acc_id) is None:
            return _json_err(404, "Account not found")
        try:
            ok = await app.state.refresh_scheduler.ensure_fresh(acc_id, force=True)
        except Exception as exc:
            return _json_err(502, f"Refresh failed: {exc}")
        acc = app.state.account_store.get(acc_id)
        return {"status": "ok", "refreshed": ok, "account": _account_public(acc) if acc else None}

    @app.delete("/admin/accounts/{acc_id}")
    async def remove_account(acc_id: str, request: Request) -> dict:
        err = _require_admin(request)
        if err: return err
        if not app.state.account_store.remove(acc_id):
            return _json_err(404, "Account not found")
        app.state.key_store.detach_account(acc_id)  # unbind keys that pointed here
        return {"status": "ok"}

    @app.get("/admin/keys")
    async def list_keys(request: Request) -> dict:
        err = _require_admin(request)
        if err: return err
        return {"keys": [_key_public(k) for k in app.state.key_store.list()]}

    @app.post("/admin/keys")
    async def add_key(request: Request) -> dict:
        err = _require_admin(request)
        if err: return err
        body = await request.json()
        name = str(body.get("name", "")).strip()
        account_id = str(body.get("account_id", "")).strip()
        tone = str(body.get("tone", "Magic")).strip() or "Magic"
        if tone not in _TONE_VALUES:
            return _json_err(400, f"Invalid tone. Allowed: {', '.join(sorted(_TONE_VALUES))}")
        if account_id and app.state.account_store.get(account_id) is None:
            return _json_err(404, "Bound account not found")
        k = app.state.key_store.add(name=name, account_id=account_id, tone=tone)
        return {"status": "ok", "key": _key_public(k)}

    @app.post("/admin/keys/{key_id}")
    async def update_key(key_id: str, request: Request) -> dict:
        err = _require_admin(request)
        if err: return err
        body = await request.json()
        fields: dict = {}
        if "name" in body:
            fields["name"] = str(body["name"]).strip()
        if "account_id" in body:
            aid = str(body["account_id"]).strip()
            if aid and app.state.account_store.get(aid) is None:
                return _json_err(404, "Bound account not found")
            fields["account_id"] = aid
        if "enabled" in body:
            fields["enabled"] = bool(body["enabled"])
        if "tone" in body:
            tone = str(body["tone"]).strip() or "Magic"
            if tone not in _TONE_VALUES:
                return _json_err(400, f"Invalid tone. Allowed: {', '.join(sorted(_TONE_VALUES))}")
            fields["tone"] = tone
        if "tool_prompt" in body:
            if not isinstance(body["tool_prompt"], str):
                return _json_err(400, "tool_prompt must be a string")
            fields["tool_prompt"] = body["tool_prompt"][:4000]
        if "system_prompt" in body:
            if not isinstance(body["system_prompt"], str):
                return _json_err(400, "system_prompt must be a string")
            fields["system_prompt"] = body["system_prompt"][:8000]
        k = app.state.key_store.update(key_id, **fields)
        if k is None:
            return _json_err(404, "Key not found")
        return {"status": "ok", "key": _key_public(k)}

    @app.delete("/admin/keys/{key_id}")
    async def remove_key(key_id: str, request: Request) -> dict:
        err = _require_admin(request)
        if err: return err
        if not app.state.key_store.remove(key_id):
            return _json_err(404, "Key not found")
        return {"status": "ok"}

    # ============================ User self-service API ============================
    def _resolve_user_key(request: Request) -> ApiKey | None:
        """Resolve the caller's own ApiKey from the Authorization header.

        /user/* paths bypass the auth middleware, so they authenticate here by
        their own API key instead of an admin cookie.
        """
        auth = request.headers.get("Authorization", "")
        m = re.match(r"^Bearer\s+(.+)$", auth, re.IGNORECASE)
        if not m:
            return None
        return app.state.key_store.resolve(m.group(1).strip())

    @app.get("/user/me")
    async def user_me(request: Request) -> dict:
        k = _resolve_user_key(request)
        if k is None:
            return _json_err(401, "Invalid API key", "auth_error")
        acc = app.state.account_store.get(k.account_id) if k.account_id else None
        return {
            "name": k.name,
            "enabled": k.enabled,
            "tone": k.tone,
            "tool_prompt": k.tool_prompt,
            "system_prompt": k.system_prompt,
            "default_system_prompt": default_tool_system_prompt(),
            "account": {
                "id": acc.id,
                "name": acc.name,
                "has_token": bool(acc.token),
                "token_status": acc.token_status(),
            } if acc is not None else None,
            "tone_options": _TONE_OPTIONS,
        }

    @app.post("/user/tone")
    async def user_set_tone(request: Request) -> dict:
        k = _resolve_user_key(request)
        if k is None:
            return _json_err(401, "Invalid API key", "auth_error")
        body = await request.json()
        tone = str(body.get("tone", "")).strip()
        if tone not in _TONE_VALUES:
            return _json_err(400, f"Invalid tone. Allowed: {', '.join(sorted(_TONE_VALUES))}")
        app.state.key_store.update(k.id, tone=tone)
        return {"status": "ok", "tone": tone}

    @app.post("/user/tool-prompt")
    async def user_set_tool_prompt(request: Request) -> dict:
        k = _resolve_user_key(request)
        if k is None:
            return _json_err(401, "Invalid API key", "auth_error")
        body = await request.json()
        prompt = body.get("tool_prompt")
        if not isinstance(prompt, str):
            return _json_err(400, "tool_prompt must be a string")
        app.state.key_store.update(k.id, tool_prompt=prompt[:4000])
        return {"status": "ok", "tool_prompt": prompt[:4000]}

    @app.post("/user/system-prompt")
    async def user_set_system_prompt(request: Request) -> dict:
        k = _resolve_user_key(request)
        if k is None:
            return _json_err(401, "Invalid API key", "auth_error")
        body = await request.json()
        prompt = body.get("system_prompt")
        if not isinstance(prompt, str):
            return _json_err(400, "system_prompt must be a string")
        app.state.key_store.update(k.id, system_prompt=prompt[:8000])
        return {"status": "ok", "system_prompt": prompt[:8000]}

    @app.post("/user/account/token")
    async def user_set_account_token(request: Request) -> dict:
        """Let a user push/update the token for their own bound account.

        If the key has no bound account yet, create one and bind it (self-service
        account provisioning requested for the user UI).
        """
        k = _resolve_user_key(request)
        if k is None:
            return _json_err(401, "Invalid API key", "auth_error")
        body = await request.json()
        token = str(body.get("token", "")).strip()
        if not token:
            return _json_err(400, "Token is empty")
        match = re.search(r"access_token=([^&\s]+)", token)
        token = match.group(1) if match else token
        try:
            claims = decode_jwt_payload(token)
            if not is_substrate_token_claims(claims):
                return _json_err(400, "Token is not a substrate.office.com token")
        except Exception:
            return _json_err(400, "Not a valid JWT token")
        acc_id = k.account_id
        if not acc_id or app.state.account_store.get(acc_id) is None:
            acc = app.state.account_store.add(name=k.name or "user", token=token, token_source="manual")
            app.state.key_store.update(k.id, account_id=acc.id)
        else:
            acc = app.state.account_store.update_token(acc_id, token, token_source="manual")
        return {"status": "ok", "token_status": acc.token_status() if acc else None}

    @app.get("/", response_class=HTMLResponse)
    async def user_page(request: Request) -> str:
        # Root is the user-facing page. Admin console moved to /admin.
        return _USER_HTML

    @app.get("/admin", response_class=HTMLResponse)
    async def admin_page(request: Request) -> str:
        if _admin_secret and not _is_admin_authenticated(request):
            return _LOGIN_HTML
        return _ADMIN_HTML

    @app.get("/favicon.ico")
    async def favicon():
        from starlette.responses import Response
        return Response(status_code=204)

    @app.get("/v1/models")
    async def list_models(settings: Settings = Depends(get_settings)) -> dict:
        return {
            "object": "list",
            "data": [
                {
                    "id": settings.model_alias,
                    "object": "model",
                    "owned_by": "microsoft-365-copilot",
                },
                {
                    "id": f"{settings.model_alias}{_PERSIST_MODEL_SUFFIX}",
                    "object": "model",
                    "owned_by": "microsoft-365-copilot",
                },
            ],
        }

    @app.post("/v1/chat/completions")
    async def chat_completions(
        raw_request: Request,
        request: OpenAIChatRequest,
        settings: Settings = Depends(get_settings),
        client: SubstrateCopilotClient = Depends(get_copilot_client),
    ):
        _log = logging.getLogger("copilot_proxy")
        _log.info("[/v1/chat/completions] stream=%s tools=%d messages=%d model=%s",
                  request.stream, len(request.tools) if request.tools else 0,
                  len(request.messages), request.model)
        if request.tools:
            for t in request.tools:
                _log.info("  tool: %s", t.function.name if t.function else "?")
        # Record call for web UI
        call_record = {
            "time": time.strftime("%H:%M:%S"),
            "stream": request.stream,
            "tools": [t.function.name for t in request.tools] if request.tools else [],
            "messages": len(request.messages),
            "model": request.model,
            "tool_calls_result": None,
        }
        try:
            session = _persistent_session(app, raw_request, request.model, request.user, request)
            # Whenever we reuse a persistent M365 session that already has history
            # (both auto mode and explicit :persist mode), the server remembers the
            # prior turns — so only send the incremental turn instead of resending the
            # whole transcript on every request.
            incremental = (
                session is not None
                and session.turn_count > 0
            )
            # Diagnostics: surface in the web call-log so we can see whether the
            # incremental optimization actually kicks in across turns.
            call_record["incremental"] = incremental
            call_record["turn_count"] = session.turn_count if session is not None else None
            _key_obj = getattr(raw_request.state, "api_key_obj", None)
            _system_override = _key_obj.system_prompt if _key_obj is not None else getattr(app.state, 'system_prompt', '')
            translated = translate_openai_request(request, incremental=incremental, system_override=_system_override)
            if request.stream:
                # Save call record for streaming (tool_calls_result resolved later)
                call_record["streaming"] = True
                app.state.call_log.append(call_record)
                if len(app.state.call_log) > 100:
                    app.state.call_log = app.state.call_log[-100:]
                if request.tools:
                    # When tools are present, buffer the full stream then parse tool_calls
                    return StreamingResponse(
                        _openai_stream_with_tools(
                            settings.model_alias,
                            client,
                            translated.prompt,
                            translated.additional_context,
                            session,
                            call_log=app.state.call_log,
                            call_record=call_record,
                            tool_names={t.function.name for t in request.tools if t.function},
                        ),
                        media_type="text/event-stream",
                    )
                return StreamingResponse(
                    _openai_stream(
                        settings.model_alias,
                        client,
                        translated.prompt,
                        translated.additional_context,
                        session,
                    ),
                    media_type="text/event-stream",
                )
            text = await client.chat(translated.prompt, translated.additional_context, session)
        except ValueError as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc
        except SubstrateCopilotError as exc:
            raise HTTPException(status_code=502, detail=str(exc)) from exc

        # If request included tools, parse model output for tool_call blocks
        tool_calls = _extract_tool_calls(text) if request.tools else []
        if not tool_calls and request.tools:
            # Prose fallback: model described "save as <path>" + code block
            tool_names = {t.function.name for t in request.tools if t.function}
            tool_calls = _extract_prose_write(text, tool_names)
            if tool_calls:
                _log.info("  prose fallback synthesized Write tool_call")
        # Corrective retry: M365 sometimes "creates" a file via its native
        # attachment feature (hosted URL) instead of a tool_call. If it claims a
        # file but emitted none, force one retry demanding a real tool_call.
        if not tool_calls and request.tools and _looks_like_fake_file_claim(text):
            _log.info("  fake file claim detected, forcing corrective retry")
            try:
                retry_text = await client.chat(_RETRY_INSTRUCTION, translated.additional_context, session)
                retry_calls = _extract_tool_calls(retry_text)
                if not retry_calls:
                    tool_names = {t.function.name for t in request.tools if t.function}
                    retry_calls = _extract_prose_write(retry_text, tool_names)
                if retry_calls:
                    _log.info("  retry produced %d tool_call(s)", len(retry_calls))
                    text, tool_calls = retry_text, retry_calls
                    call_record["retried"] = True
            except SubstrateCopilotError:
                pass  # Keep original response if retry fails
        _log.info("[/v1/chat/completions] response len=%d tool_calls=%d", len(text), len(tool_calls))
        if tool_calls:
            _log.info("  parsed tool_calls: %s", [tc["function"]["name"] for tc in tool_calls])
        # Save call record
        call_record["response_len"] = len(text)
        call_record["response_text"] = text[:8000]
        call_record["response_repr"] = repr(text[:2000])
        call_record["tool_calls_result"] = [tc["function"]["name"] for tc in tool_calls] if tool_calls else []
        app.state.call_log.append(call_record)
        if len(app.state.call_log) > 100:
            app.state.call_log = app.state.call_log[-100:]
        if tool_calls:
            remaining = _strip_tool_call_blocks(text)
            msg = {"role": "assistant", "content": remaining or None, "tool_calls": tool_calls}
            return JSONResponse({
                "id": f"chatcmpl_{uuid.uuid4().hex}",
                "object": "chat.completion",
                "created": int(time.time()),
                "model": settings.model_alias,
                "choices": [
                    {
                        "index": 0,
                        "message": msg,
                        "finish_reason": "tool_calls",
                    }
                ],
                "usage": {"prompt_tokens": 0, "completion_tokens": 0, "total_tokens": 0},
            })

        return JSONResponse({
            "id": f"chatcmpl_{uuid.uuid4().hex}",
            "object": "chat.completion",
            "created": int(time.time()),
            "model": settings.model_alias,
            "choices": [
                {
                    "index": 0,
                    "message": {"role": "assistant", "content": text},
                    "finish_reason": "stop",
                }
            ],
            "usage": {"prompt_tokens": 0, "completion_tokens": 0, "total_tokens": 0},
        })

    @app.post("/v1/responses")
    async def openai_responses(
        raw: Request,
        settings: Settings = Depends(get_settings),
        client: SubstrateCopilotClient = Depends(get_copilot_client),
    ):
        body = await raw.json()
        try:
            request = OpenAIResponsesRequest.model_validate(body)
            translated = translate_responses_request(request)
            session = _persistent_session(app, raw, request.model)
        except ValueError as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc

        if request.stream:
            return StreamingResponse(
                _responses_stream(settings.model_alias, client, translated.prompt, translated.additional_context, session),
                media_type="text/event-stream",
            )

        try:
            text = await client.chat(translated.prompt, translated.additional_context, session)
        except SubstrateCopilotError as exc:
            raise HTTPException(status_code=502, detail=str(exc)) from exc

        return JSONResponse({
            "id": f"resp_{uuid.uuid4().hex}",
            "object": "response",
            "created_at": int(time.time()),
            "model": settings.model_alias,
            "output": [{
                "type": "message",
                "id": f"msg_{uuid.uuid4().hex}",
                "role": "assistant",
                "content": [{"type": "output_text", "text": text}],
            }],
            "usage": {"input_tokens": 0, "output_tokens": 0, "total_tokens": 0},
        })

    @app.post("/v1/messages")
    async def anthropic_messages(
        raw_request: Request,
        request: AnthropicMessagesRequest,
        settings: Settings = Depends(get_settings),
        client: SubstrateCopilotClient = Depends(get_copilot_client),
    ):
        try:
            translated = translate_anthropic_request(request)
            session = _persistent_session(app, raw_request, request.model)
        except ValueError as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc

        if request.stream:
            return StreamingResponse(
                _anthropic_stream(settings.model_alias, client, translated.prompt, translated.additional_context, session),
                media_type="text/event-stream",
            )

        try:
            text = await client.chat(translated.prompt, translated.additional_context, session)
        except SubstrateCopilotError as exc:
            raise HTTPException(status_code=502, detail=str(exc)) from exc

        return JSONResponse({
            "id": f"msg_{uuid.uuid4().hex}",
            "type": "message",
            "role": "assistant",
            "model": settings.model_alias,
            "content": [{"type": "text", "text": text}],
            "stop_reason": "end_turn",
            "stop_sequence": None,
            "usage": {"input_tokens": 0, "output_tokens": 0},
        })

    return app


def _persistent_session(
    app: FastAPI,
    raw_request: Request,
    model: str,
    fallback_key: str | None = None,
    request: OpenAIChatRequest | None = None,
) -> PersistentSession | None:
    # Multi-tenant: prefix every session key with the caller's key id (fallback to
    # the bound account id) so two different API keys never share an M365 thread,
    # even when their session ids / opening messages collide.
    key_obj = getattr(raw_request.state, "api_key_obj", None)
    account = getattr(raw_request.state, "account", None)
    tenant = (key_obj.id if key_obj is not None else None) or (account.id if account is not None else "global")
    header_key = (raw_request.headers.get(_SESSION_ID_HEADER) or "").strip()
    if header_key:
        return app.state.session_store.get(f"{tenant}:header:{header_key}")
    if model.endswith(_PERSIST_MODEL_SUFFIX):
        return app.state.session_store.get(f"{tenant}:model:{fallback_key or 'default'}")
    # Auto-detect conversation from the request messages so that all turns of the
    # same Trae conversation reuse one M365 Copilot session (instead of creating a
    # brand-new chat record on every request). A new Trae conversation has a
    # different first user message -> different session key -> new M365 session.
    if request is not None:
        sid, _title = _detect_conversation_session(request)
        # A conversation's opening turn carries no assistant reply yet. If two
        # different conversations happen to share the same first user message
        # (e.g. the same prompt reused to start a new chat), their auto key
        # collides. Reusing the stale M365 thread would feed the model wrong
        # context and make it hallucinate. So on an opening turn, start fresh.
        has_assistant = any(m.role == "assistant" for m in request.messages)
        if not has_assistant:
            return app.state.session_store.reset(f"{tenant}:auto:{sid}")
        return app.state.session_store.get(f"{tenant}:auto:{sid}")
    return None


async def _openai_stream(
    model_alias: str,
    client: SubstrateCopilotClient,
    prompt: str,
    additional_context: list[str],
    session: PersistentSession | None = None,
) -> AsyncIterator[str]:
    completion_id = f"chatcmpl_{uuid.uuid4().hex}"
    created = int(time.time())
    first_chunk = {
        "id": completion_id,
        "object": "chat.completion.chunk",
        "created": created,
        "model": model_alias,
        "choices": [{"index": 0, "delta": {"role": "assistant"}, "finish_reason": None}],
    }
    yield f"data: {json.dumps(first_chunk)}\n\n"
    try:
        async for delta in client.chat_stream(prompt, additional_context, session):
            chunk = {
                "id": completion_id,
                "object": "chat.completion.chunk",
                "created": created,
                "model": model_alias,
                "choices": [{"index": 0, "delta": {"content": delta}, "finish_reason": None}],
            }
            yield f"data: {json.dumps(chunk)}\n\n"
    except SubstrateCopilotError as exc:
        yield f"data: {json.dumps({'error': {'message': str(exc), 'type': 'upstream_error'}})}\n\n"
        yield "data: [DONE]\n\n"
        return
    final_chunk = {
        "id": completion_id,
        "object": "chat.completion.chunk",
        "created": created,
        "model": model_alias,
        "choices": [{"index": 0, "delta": {}, "finish_reason": "stop"}],
    }
    yield f"data: {json.dumps(final_chunk)}\n\n"
    yield "data: [DONE]\n\n"


async def _openai_stream_with_tools(
    model_alias: str,
    client: SubstrateCopilotClient,
    prompt: str,
    additional_context: list[str],
    session: PersistentSession | None = None,
    call_log: list | None = None,
    call_record: dict | None = None,
    tool_names: set | None = None,
) -> AsyncIterator[str]:
    """Buffer full stream, then emit as tool_calls if found, else normal content stream."""
    _log = logging.getLogger("copilot_proxy")
    chunks: list[str] = []
    async for delta in client.chat_stream(prompt, additional_context, session):
        chunks.append(delta)
    full_text = "".join(chunks)

    tool_calls = _extract_tool_calls(full_text)
    if not tool_calls and tool_names:
        # Prose fallback: model described "save as <path>" + code block
        tool_calls = _extract_prose_write(full_text, tool_names)
        if tool_calls:
            _log.info("  prose fallback synthesized Write tool_call")
    # Corrective retry: M365 native file-gen (hosted URL) instead of a tool_call.
    if not tool_calls and tool_names and _looks_like_fake_file_claim(full_text):
        _log.info("  fake file claim detected, forcing corrective retry")
        try:
            retry_chunks: list[str] = []
            async for delta in client.chat_stream(_RETRY_INSTRUCTION, additional_context, session):
                retry_chunks.append(delta)
            retry_text = "".join(retry_chunks)
            retry_calls = _extract_tool_calls(retry_text)
            if not retry_calls:
                retry_calls = _extract_prose_write(retry_text, tool_names)
            if retry_calls:
                _log.info("  retry produced %d tool_call(s)", len(retry_calls))
                full_text, tool_calls = retry_text, retry_calls
                if call_record is not None:
                    call_record["retried"] = True
        except SubstrateCopilotError:
            pass  # Keep original response if retry fails
    _log.info("[stream_with_tools] full_text len=%d tool_calls=%d", len(full_text), len(tool_calls))
    if tool_calls:
        _log.info("  parsed tool_calls: %s", [tc["function"]["name"] for tc in tool_calls])
    # Update call record with results
    if call_record is not None:
        call_record["response_len"] = len(full_text)
        call_record["response_text"] = full_text[:8000]
        call_record["response_repr"] = repr(full_text[:2000])
        call_record["tool_calls_result"] = [tc["function"]["name"] for tc in tool_calls] if tool_calls else []
    completion_id = f"chatcmpl_{uuid.uuid4().hex}"
    created = int(time.time())

    if tool_calls:
        remaining = _strip_tool_call_blocks(full_text)
        # Emit role chunk
        yield f"data: {json.dumps({'id': completion_id, 'object': 'chat.completion.chunk', 'created': created, 'model': model_alias, 'choices': [{'index': 0, 'delta': {'role': 'assistant'}, 'finish_reason': None}]})}\n\n"
        # Emit remaining text content if any
        if remaining:
            yield f"data: {json.dumps({'id': completion_id, 'object': 'chat.completion.chunk', 'created': created, 'model': model_alias, 'choices': [{'index': 0, 'delta': {'content': remaining}, 'finish_reason': None}]})}\n\n"
        # Emit tool_calls chunks — one per tool call
        for i, tc in enumerate(tool_calls):
            delta_tc = [{"index": i, "id": tc["id"], "type": "function", "function": {"name": tc["function"]["name"], "arguments": tc["function"]["arguments"]}}]
            yield f"data: {json.dumps({'id': completion_id, 'object': 'chat.completion.chunk', 'created': created, 'model': model_alias, 'choices': [{'index': 0, 'delta': {'tool_calls': delta_tc}, 'finish_reason': None}]})}\n\n"
        # Final chunk with finish_reason
        yield f"data: {json.dumps({'id': completion_id, 'object': 'chat.completion.chunk', 'created': created, 'model': model_alias, 'choices': [{'index': 0, 'delta': {}, 'finish_reason': 'tool_calls'}]})}\n\n"
        yield "data: [DONE]\n\n"
    else:
        # No tool calls found — re-stream as normal content
        yield f"data: {json.dumps({'id': completion_id, 'object': 'chat.completion.chunk', 'created': created, 'model': model_alias, 'choices': [{'index': 0, 'delta': {'role': 'assistant'}, 'finish_reason': None}]})}\n\n"
        yield f"data: {json.dumps({'id': completion_id, 'object': 'chat.completion.chunk', 'created': created, 'model': model_alias, 'choices': [{'index': 0, 'delta': {'content': full_text}, 'finish_reason': None}]})}\n\n"
        yield f"data: {json.dumps({'id': completion_id, 'object': 'chat.completion.chunk', 'created': created, 'model': model_alias, 'choices': [{'index': 0, 'delta': {}, 'finish_reason': 'stop'}]})}\n\n"
        yield "data: [DONE]\n\n"


async def _responses_stream(
    model_alias: str,
    client: SubstrateCopilotClient,
    prompt: str,
    additional_context: list[str],
    session: PersistentSession | None = None,
) -> AsyncIterator[str]:
    resp_id = f"resp_{uuid.uuid4().hex}"
    item_id = f"msg_{uuid.uuid4().hex}"
    created = int(time.time())

    yield f"data: {json.dumps({'type': 'response.created', 'response': {'id': resp_id, 'object': 'response', 'created_at': created, 'model': model_alias, 'status': 'in_progress', 'output': []}})}\n\n"
    yield f"data: {json.dumps({'type': 'response.output_item.added', 'output_index': 0, 'item': {'id': item_id, 'type': 'message', 'role': 'assistant', 'content': []}})}\n\n"
    yield f"data: {json.dumps({'type': 'response.content_part.added', 'item_id': item_id, 'output_index': 0, 'content_index': 0, 'part': {'type': 'output_text', 'text': ''}})}\n\n"

    full_text = ""
    try:
        async for delta in client.chat_stream(prompt, additional_context, session):
            full_text += delta
            yield f"data: {json.dumps({'type': 'response.output_text.delta', 'item_id': item_id, 'output_index': 0, 'content_index': 0, 'delta': delta})}\n\n"
    except SubstrateCopilotError as exc:
        yield f"data: {json.dumps({'type': 'error', 'error': {'message': str(exc), 'type': 'upstream_error'}})}\n\n"
        return

    yield f"data: {json.dumps({'type': 'response.output_text.done', 'item_id': item_id, 'output_index': 0, 'content_index': 0, 'text': full_text})}\n\n"
    yield f"data: {json.dumps({'type': 'response.completed', 'response': {'id': resp_id, 'object': 'response', 'created_at': created, 'model': model_alias, 'status': 'completed', 'output': [{'id': item_id, 'type': 'message', 'role': 'assistant', 'content': [{'type': 'output_text', 'text': full_text}]}], 'usage': {'input_tokens': 0, 'output_tokens': 0, 'total_tokens': 0}}})}\n\n"


async def _anthropic_stream(
    model_alias: str,
    client: SubstrateCopilotClient,
    prompt: str,
    additional_context: list[str],
    session: PersistentSession | None = None,
) -> AsyncIterator[str]:
    msg_id = f"msg_{uuid.uuid4().hex}"

    def sse(event: str, data: dict) -> str:
        return f"event: {event}\ndata: {json.dumps(data)}\n\n"

    yield sse("message_start", {"type": "message_start", "message": {"id": msg_id, "type": "message", "role": "assistant", "content": [], "model": model_alias, "stop_reason": None, "stop_sequence": None, "usage": {"input_tokens": 0, "output_tokens": 0}}})
    yield sse("content_block_start", {"type": "content_block_start", "index": 0, "content_block": {"type": "text", "text": ""}})
    yield sse("ping", {"type": "ping"})

    try:
        async for delta in client.chat_stream(prompt, additional_context, session):
            yield sse("content_block_delta", {"type": "content_block_delta", "index": 0, "delta": {"type": "text_delta", "text": delta}})
    except SubstrateCopilotError as exc:
        yield sse("error", {"type": "error", "error": {"type": "upstream_error", "message": str(exc)}})
        return

    yield sse("content_block_stop", {"type": "content_block_stop", "index": 0})
    yield sse("message_delta", {"type": "message_delta", "delta": {"stop_reason": "end_turn", "stop_sequence": None}, "usage": {"output_tokens": 0}})
    yield sse("message_stop", {"type": "message_stop"})


_LOGIN_HTML = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Ciallo Ms-365 OpenAI Proxy</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:-apple-system,BlinkMacSystemFont,Segoe UI,Roboto,sans-serif;background:#0f172a;color:#e2e8f0;min-height:100vh;display:flex;align-items:center;justify-content:center}
.login-box{background:#1e293b;border-radius:14px;padding:2.5rem;width:360px;border:1px solid #334155;text-align:center;position:relative}
.login-box h1{font-size:1.3rem;margin-bottom:.5rem;background:linear-gradient(135deg,#06b6d4,#8b5cf6);-webkit-background-clip:text;-webkit-text-fill-color:transparent}
.login-box p{color:#64748b;font-size:.85rem;margin-bottom:1.5rem}
input{width:100%;padding:.75rem 1rem;background:#0f172a;border:1px solid #475569;border-radius:8px;color:#e2e8f0;font-size:.9rem;outline:none;margin-bottom:1rem;transition:border-color .2s}
input:focus{border-color:#06b6d4}
button{width:100%;background:linear-gradient(135deg,#06b6d4,#8b5cf6);color:#fff;border:none;border-radius:8px;padding:.75rem;font-size:.95rem;font-weight:600;cursor:pointer;transition:opacity .2s}
button:hover{opacity:.85}
button:disabled{opacity:.5;cursor:not-allowed}
.msg{padding:.5rem .75rem;border-radius:6px;font-size:.8rem;margin-top:.75rem;display:none}
.msg.err{display:block;background:#450a0a;color:#ef4444;border:1px solid #991b1b}
.lang-btn{position:absolute;top:12px;right:12px;background:linear-gradient(135deg,rgba(6,182,212,0.18),rgba(139,92,246,0.18));border:1px solid rgba(139,92,246,0.5);color:#e2e8f0;font-size:12px;padding:4px 12px;border-radius:16px;cursor:pointer;font-weight:600;width:auto}
</style>
</head>
<body>
<div class="login-box">
<button class="lang-btn" id="lang-toggle" onclick="toggleLang()">&#127760; EN</button>
<h1>Ciallo Ms-365 OpenAI Proxy</h1>
<p id="login-desc" data-i18n="login_desc">输入管理员密码以继续</p>
<input id="pw" type="password" placeholder="API Key / 密码" autofocus onkeydown="if(event.key==='Enter')doLogin()">
<button id="btn" onclick="doLogin()" data-i18n="login_btn">登录</button>
<div id="msg" class="msg"></div>
</div>
<script>
const i18n={
  zh:{login_desc:'输入管理员密码以继续',login_btn:'登录',placeholder:'API Key / 密码',login_failed:'登录失败',network_error:'网络错误',wrong_password:'密码错误'},
  en:{login_desc:'Enter admin password to continue',login_btn:'Login',placeholder:'API Key / Password',login_failed:'Login failed',network_error:'Network error',wrong_password:'Wrong password'}
};
let lang=localStorage.getItem('lang')||'zh';
function t(k){return i18n[lang][k]||k}
function applyLang(){
  const btn=document.getElementById('lang-toggle');
  btn.innerHTML=lang==='zh'?'&#127760; EN':'&#127760; 中文';
  document.querySelectorAll('[data-i18n]').forEach(el=>{const k=el.getAttribute('data-i18n');if(i18n[lang][k])el.textContent=i18n[lang][k]});
  document.getElementById('pw').placeholder=t('placeholder');
}
function toggleLang(){lang=lang==='zh'?'en':'zh';localStorage.setItem('lang',lang);applyLang()}
applyLang();
async function doLogin(){
  const pw=document.getElementById('pw').value;
  const btn=document.getElementById('btn');
  const msg=document.getElementById('msg');
  btn.disabled=true;msg.className='msg';msg.textContent='';
  try{
    const r=await fetch('/admin/login',{method:'POST',credentials:'include',headers:{'Content-Type':'application/json'},body:JSON.stringify({password:pw})});
    const d=await r.json();
    if(r.ok){location.reload()}else{msg.className='msg err';msg.textContent=d.error?.message||t('login_failed')}
  }catch(e){msg.className='msg err';msg.textContent=t('network_error')}
  finally{btn.disabled=false}
}
</script>
</body>
</html>"""

_ADMIN_HTML = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Ciallo Ms-365 OpenAI Proxy</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:-apple-system,BlinkMacSystemFont,Segoe UI,Roboto,sans-serif;background:#0f172a;color:#e2e8f0;min-height:100vh;padding:2rem}
.container{max-width:720px;margin:0 auto}
h1{font-size:1.5rem;margin-bottom:1.5rem;background:linear-gradient(135deg,#06b6d4,#8b5cf6);-webkit-background-clip:text;-webkit-text-fill-color:transparent}
.card{background:#1e293b;border-radius:12px;padding:1.5rem;margin-bottom:1.5rem;border:1px solid #334155}
.card h2{font-size:1.1rem;margin-bottom:1rem;color:#e2e8f0}
.status-row{display:flex;justify-content:space-between;align-items:center;padding:.5rem 0;border-bottom:1px solid #334155}
.status-row:last-child{border:none}
.status-label{color:#94a3b8;font-size:.9rem}
.status-value{font-weight:600;font-size:.9rem}
.valid{color:#22c55e}.invalid{color:#ef4444}.warn{color:#f59e0b}
textarea{width:100%;height:120px;background:#0f172a;border:1px solid #475569;border-radius:8px;color:#e2e8f0;padding:.75rem;font-family:monospace;font-size:.8rem;resize:vertical;margin-bottom:.75rem}
textarea:focus{outline:none;border-color:#06b6d4}
button{background:linear-gradient(135deg,#06b6d4,#8b5cf6);color:#fff;border:none;border-radius:8px;padding:.55rem .9rem;font-size:.8rem;font-weight:600;cursor:pointer;transition:opacity .2s;white-space:nowrap;flex-shrink:0}
button:hover{opacity:.85}
button:disabled{opacity:.5;cursor:not-allowed}
.btn-bar{display:flex;gap:.5rem;margin-bottom:.25rem;flex-wrap:wrap}
.msg{padding:.6rem 1rem;border-radius:8px;font-size:.85rem;margin-top:.5rem;display:none}
.msg.ok{display:block;background:#052e16;color:#22c55e;border:1px solid #166534}
.msg.err{display:block;background:#450a0a;color:#ef4444;border:1px solid #991b1b}
.api-info{margin-top:1rem;padding:.75rem;background:#0f172a;border-radius:8px;font-family:monospace;font-size:.8rem;color:#64748b;line-height:1.6}
a{color:#06b6d4;text-decoration:none}
a:hover{text-decoration:underline}
</style>
</head>
<body>
<div class="container">
<h1>Ciallo Ms-365 OpenAI Proxy <span style="font-size:13px;color:#8b5cf6;font-weight:600" data-i18n="multi_badge">多租户</span> <button id="lang-toggle" onclick="toggleLang()" style="font-size:14px;padding:5px 14px;border:1px solid rgba(139,92,246,0.5);border-radius:20px;background:linear-gradient(135deg,rgba(6,182,212,0.18),rgba(139,92,246,0.18));cursor:pointer;vertical-align:middle;margin-left:12px;transition:all .2s;letter-spacing:1px;font-weight:600;line-height:1">&#127760; EN</button></h1>

<div class="card">
<div style="display:flex;align-items:center;gap:.5rem;margin-bottom:.75rem">
<h2 data-i18n="title_accounts" style="margin:0">账户池</h2>
<button onclick="addAccount()" style="margin-left:auto;font-size:.8rem;padding:5px 12px" data-i18n="btn_add_account">添加账户</button>
</div>
<div style="font-size:.8rem;color:#64748b;margin-bottom:.5rem" data-i18n="accounts_hint">每个账户拥有独立的 M365 Token 与 Chromium 刷新配置。刷新按需串行拉起浏览器，用完即关。</div>
<div id="accounts-content"><span style="color:#64748b" data-i18n="loading">加载中...</span></div>
</div>

<div class="card">
<div style="display:flex;align-items:center;gap:.5rem;margin-bottom:.75rem">
<h2 data-i18n="title_keys" style="margin:0">API Key 管理</h2>
<button onclick="addKey()" style="margin-left:auto;font-size:.8rem;padding:5px 12px" data-i18n="btn_add_key">新建 Key</button>
</div>
<div style="font-size:.8rem;color:#64748b;margin-bottom:.5rem" data-i18n="keys_hint">每个 Key 绑定一个账户，可单独设置对话模式、提示词并随时启用/停用。</div>
<div id="keys-content"><span style="color:#64748b" data-i18n="loading">加载中...</span></div>
</div>

<details class="card">
<summary style="cursor:pointer;font-size:1.1rem;font-weight:600;color:#e2e8f0;list-style:none;display:flex;align-items:center;gap:.5rem">
<span data-i18n="title_legacy">全局 / 兼容 Token（高级）</span>
<span style="font-size:.7rem;color:#475569;margin-left:auto" data-i18n="click_expand">点击展开</span>
</summary>
<div style="margin-top:.75rem">
<h2 data-i18n="title_update_token">更新 Token</h2>
<p style="color:#64748b;font-size:.85rem;margin-bottom:.75rem" data-i18n="desc_paste_token">粘贴 access_token 值或完整的 wss:// URL</p>
<textarea id="token-input" placeholder="eyJ0eXAiOiJKV1QiLCJhbGci...&#10;&#10;or full URL:&#10;wss://substrate.office.com/m365Copilot/Chathub/...?access_token=eyJ..."></textarea>
<div class="btn-bar">
<button id="btn-update" onclick="updateToken()" data-i18n="btn_update">更新 Token</button>
<button id="btn-check" onclick="checkLogin()" style="background:linear-gradient(135deg,#f59e0b,#d97706)" data-i18n="btn_check_login">检查登录</button>
<button id="btn-auto" onclick="autoCapture()" style="background:linear-gradient(135deg,#22c55e,#059669)" data-i18n="btn_auto_capture">自动刷新</button>
<button id="btn-stop-refresh" onclick="toggleAutoRefresh()" style="display:none"></button>
<button id="btn-logout" onclick="logoutUser()" style="display:none;background:linear-gradient(135deg,#ef4444,#dc2626)" data-i18n="btn_logout">登出用户</button>
</div>
<div id="update-msg" class="msg"></div>
</div>
</details>

<div class="card">
<h2 data-i18n="title_status">Token 与 登录状态</h2>
<div id="status-content"><span style="color:#64748b" data-i18n="loading">加载中...</span></div>
</div>

<div class="card">
<div style="display:flex;align-items:center;gap:.5rem">
<h2 data-i18n="title_tone" style="margin:0">对话模式</h2>
<span id="tone-saved" style="font-size:.75rem;color:#22c55e;opacity:0;transition:opacity .3s"></span>
<select id="tone-select" style="margin-left:auto;width:150px;max-width:50%;padding:6px 32px 6px 10px;background:#0f172a;border:1px solid #334155;border-radius:8px;color:#e2e8f0;font-size:.8rem;outline:none;-webkit-appearance:none;-moz-appearance:none;appearance:none;background-image:url(&quot;data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 24 24' fill='none' stroke='%2394a3b8' stroke-width='2'%3E%3Cpath d='M6 9l6 6 6-6'/%3E%3C/svg%3E&quot;);background-repeat:no-repeat;background-position:right 10px center"></select>
</div>
</div>

<div class="card">
<details id="tool-prompt-details" style="cursor:pointer">
<summary style="font-size:1.1rem;font-weight:600;color:#e2e8f0;list-style:none;display:flex;align-items:center;gap:.5rem">
<span data-i18n="title_tool_prompt">提示词微调</span>
<span style="font-size:.7rem;color:#475569;margin-left:auto" data-i18n="click_expand">点击展开</span>
</summary>
<div style="margin-top:.75rem">
<div style="font-size:.8rem;color:#64748b;margin-bottom:.5rem" data-i18n="tool_prompt_hint">追加到工具调用提示词后的自定义指令，用于调教模型的 tool_call 行为。立即生效并持久保存，留空则不追加。</div>
<textarea id="tool-prompt-input" rows="4" style="width:100%;box-sizing:border-box;padding:8px 12px;background:#0f172a;border:1px solid #334155;border-radius:8px;color:#e2e8f0;font-size:.85rem;font-family:monospace;outline:none;resize:vertical" placeholder=""></textarea>
<div style="display:flex;align-items:center;gap:.5rem;margin-top:.5rem">
<button id="tool-prompt-save" onclick="saveToolPrompt()" data-i18n="tool_prompt_save">保存</button>
<button id="tool-prompt-reset" onclick="resetToolPrompt()" style="background:linear-gradient(135deg,#64748b,#475569)" data-i18n="prompt_reset">恢复默认</button>
<span id="tool-prompt-saved" style="font-size:.75rem;color:#22c55e;opacity:0;transition:opacity .3s"></span>
</div>
</div>
</details>
</div>

<div class="card">
<details id="system-prompt-details" style="cursor:pointer">
<summary style="font-size:1.1rem;font-weight:600;color:#e2e8f0;list-style:none;display:flex;align-items:center;gap:.5rem">
<span data-i18n="title_system_prompt">系统级提示词（高级）</span>
<span style="font-size:.7rem;color:#475569;margin-left:auto" data-i18n="click_expand">点击展开</span>
</summary>
<div style="margin-top:.75rem">
<div style="font-size:.8rem;color:#64748b;margin-bottom:.5rem" data-i18n="system_prompt_hint">覆盖工具调用的基础系统提示词（定义 tool_call 格式与规则）。改错会导致工具调用失效，仅供高级用户调试。动态工具列表始终自动追加，不可编辑。留空则使用内置默认。</div>
<div id="system-prompt-locked">
<button id="system-prompt-unlock" onclick="unlockSystemPrompt()" style="background:linear-gradient(135deg,#ef4444,#dc2626)" data-i18n="system_prompt_unlock">解锁编辑（高级）</button>
</div>
<div id="system-prompt-editor" style="display:none">
<textarea id="system-prompt-input" rows="10" style="width:100%;box-sizing:border-box;padding:8px 12px;background:#0f172a;border:1px solid #7f1d1d;border-radius:8px;color:#e2e8f0;font-size:.8rem;font-family:monospace;outline:none;resize:vertical" placeholder=""></textarea>
<div style="display:flex;align-items:center;gap:.5rem;margin-top:.5rem">
<button id="system-prompt-save" onclick="saveSystemPrompt()" data-i18n="system_prompt_save">保存</button>
<button id="system-prompt-reset" onclick="resetSystemPrompt()" style="background:linear-gradient(135deg,#64748b,#475569)" data-i18n="prompt_reset">恢复默认</button>
<span id="system-prompt-saved" style="font-size:.75rem;color:#22c55e;opacity:0;transition:opacity .3s"></span>
</div>
</div>
</div>
</details>
</div>

<div class="card">
<details id="call-log-details" style="cursor:pointer">
<summary style="font-size:1.1rem;font-weight:600;color:#e2e8f0;list-style:none;display:flex;align-items:center;gap:.5rem">
<span data-i18n="title_call_log">API 调用记录</span>
<span id="call-log-count" style="font-size:.75rem;color:#64748b;background:#1e293b;padding:2px 8px;border-radius:8px">0</span>
<span style="font-size:.7rem;color:#475569;margin-left:auto" data-i18n="click_expand">点击展开</span>
</summary>
<div id="call-log-content" style="margin-top:.75rem;max-height:400px;overflow-y:auto;font-family:monospace;font-size:.8rem">
<span style="color:#64748b" data-i18n="no_calls_yet">暂无调用记录</span>
</div>
</details>
</div>

<div class="card">
<details id="capture-details" style="cursor:pointer">
<summary style="font-size:1.1rem;font-weight:600;color:#e2e8f0;list-style:none;display:flex;align-items:center;gap:.5rem">
<span data-i18n="title_capture">模式抓包对比</span>
<span id="capture-count" style="font-size:.75rem;color:#64748b;background:#1e293b;padding:2px 8px;border-radius:8px">0</span>
<span style="font-size:.7rem;color:#475569;margin-left:auto" data-i18n="click_expand">点击展开</span>
</summary>
<div style="margin-top:.5rem;font-size:.75rem;color:#64748b" data-i18n="capture_hint">在 M365 Copilot 切换不同模式（快速答复/深度思考、GPT 5.5/5.2）各发一条消息，用油猴脚本推送抓包，下方对比哪些字段控制模式。</div>
<div id="capture-content" style="margin-top:.75rem;max-height:400px;overflow-y:auto;font-family:monospace;font-size:.78rem">
<span style="color:#64748b" data-i18n="no_capture_yet">暂无抓包数据</span>
</div>
</details>
</div>

<div class="card">
<h2 data-i18n="title_quick_start">快速开始</h2>
<p style="color:#94a3b8;font-size:.85rem;line-height:1.6;margin-bottom:.75rem">
<strong style="color:#22c55e" data-i18n="qs_recommended">推荐：</strong><span data-i18n="qs_install_script">安装油猴脚本（</span><a href="https://gh-proxy.com/https://raw.githubusercontent.com/MurasameCyan/Ciallo-Ms-365-OpenAI-Proxy-Docker/main/get_token.user.js" target="_blank" data-i18n="qs_script_name">一键脚本</a>），<span data-i18n="qs_open_copilot">打开</span> <a href="https://m365.cloud.microsoft/chat" target="_blank">M365 Copilot</a>，<span data-i18n="qs_type_trigger">输入内容触发 WebSocket，然后在脚本面板点击</span> <strong data-i18n="qs_push_token">推送 Token</strong>。<br>
<strong style="color:#f59e0b" data-i18n="qs_alternative">备选：</strong><span data-i18n="qs_manual_copy">在 DevTools（Network → WS → wss://substrate.office.com/...）中手动复制 </span><code>access_token</code>，<span data-i18n="qs_paste_above">然后粘贴到上方。</span>
</p>
<details style="cursor:pointer">
<summary style="font-weight:600;color:#e2e8f0;list-style:none;display:flex;align-items:center;gap:.5rem">
<span data-i18n="title_api_endpoints">API 端点</span>
<span style="font-size:.7rem;color:#475569;margin-left:auto" data-i18n="click_expand">点击展开</span>
</summary>
<div class="api-info" style="margin-top:.5rem">
GET  /healthz<br>
GET  /admin/token/status<br>
POST /admin/token/update<br>
POST /admin/token/auto-capture<br>
POST /admin/cookie/inject<br>
GET  /admin/chromium/login-status<br>
POST /admin/chromium/logout<br>
GET  /admin/call-log<br>
GET  /admin/capture-payload<br>
POST /admin/capture-payload<br>
GET  /admin/tone<br>
POST /admin/tone<br>
GET  /admin/tool-prompt<br>
POST /admin/tool-prompt<br>
GET  /admin/system-prompt<br>
POST /admin/system-prompt<br>
GET  /v1/models<br>
POST /v1/chat/completions<br>
POST /v1/responses<br>
POST /v1/messages
</div>
</details>
</div>

<script>
const i18n={
  zh:{
    multi_badge:'多租户',
    title_accounts:'账户池',btn_add_account:'添加账户',
    accounts_hint:'每个账户拥有独立的 M365 Token 与 Chromium 刷新配置。刷新按需串行拉起浏览器，用完即关。',
    title_keys:'API Key 管理',btn_add_key:'新建 Key',
    keys_hint:'每个 Key 绑定一个账户，可单独设置对话模式、提示词并随时启用/停用。',
    title_legacy:'全局 / 兼容 Token（高级）',
    acct_prompt_name:'账户名称（可选）：',acct_prompt_token:'可选：粘贴该账户的 access_token 或 wss:// URL（留空则稍后用 CDP 刷新）：',
    key_prompt_name:'Key 名称（可选，如用户/用途）：',
    col_name:'名称',col_account:'账户',col_token:'Token',col_status:'状态',col_actions:'操作',col_key:'Key',col_mode:'模式',col_enabled:'启用',
    btn_refresh:'刷新',btn_rebind:'改绑',btn_delete:'删除',btn_copy:'复制',btn_enable:'启用',btn_disable:'停用',btn_push_token:'推送 Token',
    confirm_del_account:'确定删除该账户？绑定它的 Key 将解绑。',confirm_del_key:'确定删除该 Key？',
    valid_short:'有效',invalid_short:'无效',no_accounts:'暂无账户',no_keys:'暂无 Key',unbound:'未绑定',
    rebind_prompt:'输入要绑定的账户 ID（留空则解绑）：',push_token_prompt:'粘贴该账户的 access_token 或 wss:// URL：',
    title_update_token:'更新 Token',btn_update:'更新 Token',btn_check_login:'检查登录',btn_auto_capture:'自动刷新',
    title_status:'Token 与 登录状态',loading:'加载中...',
    title_quick_start:'快速开始',qs_recommended:'推荐：',qs_install_script:'安装油猴脚本（',qs_script_name:'一键脚本',
    qs_open_copilot:'打开',qs_type_trigger:'输入内容触发 WebSocket，然后在脚本面板点击',qs_push_token:'推送 Token',
    qs_alternative:'备选：',qs_manual_copy:'在 DevTools（Network → WS → wss://substrate.office.com/...）中手动复制 ',
    qs_paste_above:'然后粘贴到上方。',title_api_endpoints:'API 端点',
    desc_paste_token:'粘贴 access_token 值或完整的 wss:// URL',
    valid:'有效',invalid:'无效',expires:'过期时间',remaining:'剩余',error:'错误',
    login:'登录',logged_in:'已登录',not_logged_in:'未登录（仅手动推送 Token）',
    btn_logout:'登出用户',logging_out:'登出中...',logout_ok:'已登出',logout_failed:'登出失败',
    page:'页面',title:'标题',chromium_not_running:'Chromium 未运行',
    capturing:'捕获中...',auto_captured:'自动刷新成功！剩余：',auto_capture_failed:'自动刷新失败',
    check_login:'检查登录中...',login_ok:'Chromium 已登录！自动刷新已启用。',
    login_not_ok:'未登录。请先使用油猴脚本推送 Cookie。',check_failed:'检查失败：',
    capturing_btn:'捕获中...',check_btn:'检查中...',
    status_yes:'是',status_no:'否',
    auto_refresh_on:'自动刷新：开',auto_refresh_off:'自动刷新：关',
    btn_stop_refresh:'停止自动刷新',btn_start_refresh:'启动自动刷新',
    auto_refresh_stopped:'自动刷新已停止',auto_refresh_started:'自动刷新已启动',
    auto_refresh_label:'自动刷新',
    username_label:'用户名',
    title_call_log:'API 调用记录',
    click_expand:'点击展开',
    no_calls_yet:'暂无调用记录',
    tool_calls_parsed:'解析出工具调用',
    view_raw:'查看原文',
    copy:'复制',copied:'已复制',copy_record:'复制整条',
    title_capture:'模式抓包对比',
    capture_hint:'在 M365 Copilot 切换不同模式（快速答复/深度思考、GPT 5.5/5.2）各发一条消息，用油猴脚本推送抓包，下方对比哪些字段控制模式。',
    no_capture_yet:'暂无抓包数据',
    title_tone:'对话模式',
    tone_hint:'选择 M365 Copilot 的对话模式（模型），立即生效并持久保存。',
    tone_saved:'已保存',
    title_tool_prompt:'提示词增强',
    tool_prompt_hint:'追加到工具调用提示词后的自定义指令，用于增强模型的 tool_call 行为。立即生效并持久保存，留空则不追加。',
    tool_prompt_save:'保存',
    tool_prompt_saved:'已保存',
    prompt_reset:'恢复默认',
    title_system_prompt:'系统提示词（高级）',
    system_prompt_hint:'覆盖工具调用的基础系统提示词（定义 tool_call 格式与规则）。改错会导致工具调用失效，仅供高级用户调试。动态工具列表始终自动追加，不可编辑。留空则使用内置默认。',
    system_prompt_unlock:'解锁编辑（高级）',
    system_prompt_save:'保存',
    system_prompt_warn:'警告：系统级提示词定义了工具调用（tool_call）的格式与核心规则。修改不当会直接导致工具调用失效、模型无法读写文件。仅在你清楚自己在做什么时继续。\\n\\n确定要解锁编辑吗？',
    system_prompt_reset_confirm:'确定要将系统级提示词恢复为内置默认吗？当前自定义内容将被清空。',
  },
  en:{
    multi_badge:'Multi-tenant',
    title_accounts:'Account Pool',btn_add_account:'Add Account',
    accounts_hint:'Each account owns an isolated M365 token and Chromium refresh profile. Refresh brings one browser up on demand (serial) and tears it down afterwards.',
    title_keys:'API Key Management',btn_add_key:'New Key',
    keys_hint:'Each key is bound to one account, with its own conversation mode and prompts, and can be enabled/disabled anytime.',
    title_legacy:'Global / Legacy Token (Advanced)',
    acct_prompt_name:'Account name (optional):',acct_prompt_token:'Optional: paste this account\\u0027s access_token or wss:// URL (leave empty to refresh via CDP later):',
    key_prompt_name:'Key name (optional, e.g. user/purpose):',
    col_name:'Name',col_account:'Account',col_token:'Token',col_status:'Status',col_actions:'Actions',col_key:'Key',col_mode:'Mode',col_enabled:'Enabled',
    btn_refresh:'Refresh',btn_rebind:'Rebind',btn_delete:'Delete',btn_copy:'Copy',btn_enable:'Enable',btn_disable:'Disable',btn_push_token:'Push Token',
    confirm_del_account:'Delete this account? Keys bound to it will be unbound.',confirm_del_key:'Delete this key?',
    valid_short:'Valid',invalid_short:'Invalid',no_accounts:'No accounts yet',no_keys:'No keys yet',unbound:'Unbound',
    rebind_prompt:'Enter the account ID to bind (leave empty to unbind):',push_token_prompt:'Paste this account\\u0027s access_token or wss:// URL:',
    title_update_token:'Update Token',btn_update:'Update Token',btn_check_login:'Check Login',btn_auto_capture:'Auto Capture',
    title_status:'Token & Login Status',loading:'Loading...',
    title_quick_start:'Quick Start',qs_recommended:'Recommended:',qs_install_script:'Install the Tampermonkey script (',qs_script_name:'one-click script',
    qs_open_copilot:'open',qs_type_trigger:'type something to trigger WebSocket, then click',qs_push_token:'Push Token',
    qs_alternative:'Alternative:',qs_manual_copy:'Manually copy the ',
    qs_paste_above:'from DevTools (Network → WS → wss://substrate.office.com/...), then paste above.',title_api_endpoints:'API Endpoints',
    desc_paste_token:'Paste the access_token value or the full wss:// URL',
    valid:'Valid',invalid:'Invalid',expires:'Expires',remaining:'Remaining',error:'Error',
    login:'Login',logged_in:'Logged In',not_logged_in:'Not Logged In (auto-refresh only)',
    btn_logout:'Logout',logging_out:'Logging out...',logout_ok:'Logged out',logout_failed:'Logout failed',
    page:'Page',title:'Title',chromium_not_running:'Chromium Not Running',
    capturing:'Capturing...',auto_captured:'Auto-captured! Remaining: ',auto_capture_failed:'Auto-capture failed',
    check_login:'Checking...',login_ok:'Chromium is logged in! Auto-refresh is active.',
    login_not_ok:'Not logged in. Use Tampermonkey script to push cookies first.',check_failed:'Check failed: ',
    capturing_btn:'Capturing...',check_btn:'Checking...',
    status_yes:'Yes',status_no:'No',
    auto_refresh_on:'Auto Refresh: On',auto_refresh_off:'Auto Refresh: Off',
    btn_stop_refresh:'Stop Auto Refresh',btn_start_refresh:'Start Auto Refresh',
    auto_refresh_stopped:'Auto refresh stopped',auto_refresh_started:'Auto refresh started',
    auto_refresh_label:'Auto Refresh',
    username_label:'Username',
    title_call_log:'API Call Log',
    click_expand:'Click to expand',
    no_calls_yet:'No calls yet',
    tool_calls_parsed:'Parsed tool calls',
    view_raw:'View raw',
    copy:'Copy',copied:'Copied',copy_record:'Copy record',
    title_capture:'Mode Capture Compare',
    capture_hint:'In M365 Copilot switch between modes (Fast/Think, GPT 5.5/5.2) and send one message each, then push the captures via the Tampermonkey script. Compare which fields control the mode below.',
    no_capture_yet:'No captures yet',
    title_tone:'Conversation Mode',
    tone_hint:'Select the M365 Copilot conversation mode (model). Applies immediately and persists across restarts.',
    tone_saved:'Saved',
    title_tool_prompt:'Prompt Enhancement',
    tool_prompt_hint:'Custom instruction appended after the tool-call prompt to tune the tool_call behavior of the model. Applies immediately and persists across restarts; leave empty to append nothing.',
    tool_prompt_save:'Save',
    tool_prompt_saved:'Saved',
    prompt_reset:'Restore default',
    title_system_prompt:'System Prompt (Advanced)',
    system_prompt_hint:'Overrides the base system prompt for tool calls (defines the tool_call format and rules). A wrong edit will break tool calling. For advanced debugging only. The dynamic tool list is always appended and is not editable. Leave empty to use the built-in default.',
    system_prompt_unlock:'Unlock editing (Advanced)',
    system_prompt_save:'Save',
    system_prompt_warn:'WARNING: the system prompt defines the format and core rules of tool calls (tool_call). An incorrect edit will break tool calling and the model will be unable to read/write files. Continue only if you know what you are doing.\\n\\nUnlock editing?',
    system_prompt_reset_confirm:'Restore the system prompt to the built-in default? Your current custom content will be cleared.',
  }
};
let lang=localStorage.getItem('lang')||'zh';
function t(key){return i18n[lang][key]||key}
function toggleLang(){
  lang=lang==='zh'?'en':'zh';
  localStorage.setItem('lang',lang);
  applyLang();
}
function applyLang(){
  const btn=document.getElementById('lang-toggle');
  btn.innerHTML=lang==='zh'?'&#127760; EN':'&#127760; 中文';
  btn.style.color='transparent';
  btn.style.background='linear-gradient(135deg,rgba(6,182,212,0.18),rgba(139,92,246,0.18))';
  btn.style.webkitBackgroundClip='padding-box';
  // Apply gradient text color matching h1
  const txt=btn.childNodes[btn.childNodes.length-1];
  if(txt&&txt.nodeType===3){
    const span=document.createElement('span');
    span.textContent=txt.textContent;
    span.style.background='linear-gradient(135deg,#06b6d4,#8b5cf6)';
    span.style.webkitBackgroundClip='text';
    span.style.webkitTextFillColor='transparent';
    txt.replaceWith(span);
  }
  document.querySelectorAll('[data-i18n]').forEach(el=>{
    const key=el.getAttribute('data-i18n');
    if(i18n[lang][key])el.textContent=i18n[lang][key];
  });
  loadStatus();loadChromiumStatus();loadTone();
  loadAccounts();loadKeys();
}
applyLang();

function showInlineLogin(){
  const curLang=localStorage.getItem('lang')||'zh';
  const li18n={zh:{desc:'输入管理员密码以继续',btn:'登录',ph:'API Key / 密码'},en:{desc:'Enter admin password to continue',btn:'Login',ph:'API Key / Password'}};
  const lt=k=>li18n[curLang][k]||k;
  document.body.innerHTML='<div style="display:flex;align-items:center;justify-content:center;min-height:100vh;background:#0f172a;color:#e2e8f0;font-family:-apple-system,BlinkMacSystemFont,Segoe UI,Roboto,sans-serif"><div style="background:#1e293b;border-radius:14px;padding:2.5rem 2.5rem 2.5rem 2.5rem;width:360px;border:1px solid #334155;text-align:center;position:relative"><button onclick="toggleInlineLang()" style="position:absolute;top:12px;right:12px;background:linear-gradient(135deg,rgba(6,182,212,0.18),rgba(139,92,246,0.18));border:1px solid rgba(139,92,246,0.5);color:#e2e8f0;font-size:12px;padding:4px 12px;border-radius:16px;cursor:pointer;font-weight:600;width:auto">'+(curLang==='zh'?'&#127760; EN':'&#127760; 中文')+'</button><h1 style="font-size:1.3rem;margin-bottom:.5rem;background:linear-gradient(135deg,#06b6d4,#8b5cf6);-webkit-background-clip:text;-webkit-text-fill-color:transparent">Ciallo Ms-365 OpenAI Proxy</h1><p style="color:#64748b;font-size:.85rem;margin-bottom:1.5rem">'+lt('desc')+'</p><input id="pw" type="password" placeholder="'+lt('ph')+'" autofocus style="width:100%;padding:.75rem 1rem;background:#0f172a;border:1px solid #475569;border-radius:8px;color:#e2e8f0;font-size:.9rem;outline:none;margin-bottom:1rem"><button onclick="doInlineLogin()" style="width:100%;background:linear-gradient(135deg,#06b6d4,#8b5cf6);color:#fff;border:none;border-radius:8px;padding:.75rem;font-size:.95rem;font-weight:600;cursor:pointer">'+lt('btn')+'</button><div id="ilm" style="padding:.5rem .75rem;border-radius:6px;font-size:.8rem;margin-top:.75rem;display:none"></div></div></div>';
  document.getElementById('pw').addEventListener('keydown',function(e){if(e.key==='Enter')doInlineLogin()});
}
function toggleInlineLang(){localStorage.setItem('lang',localStorage.getItem('lang')==='zh'?'en':'zh');showInlineLogin()}

async function doInlineLogin(){
  const pw=document.getElementById('pw').value;
  const btns=document.querySelectorAll('button');
  const btn=btns.length>1?btns[btns.length-1]:btns[0];
  const msg=document.getElementById('ilm');
  const curLang=localStorage.getItem('lang')||'zh';
  const li18n={zh:{fail:'登录失败',neterr:'网络错误'},en:{fail:'Login failed',neterr:'Network error'}};
  const lt=k=>li18n[curLang][k]||k;
  btn.disabled=true;msg.style.display='none';
  try{
    const r=await fetch('/admin/login',{method:'POST',credentials:'include',headers:{'Content-Type':'application/json'},body:JSON.stringify({password:pw})});
    if(r.ok){location.reload();return}
    const d=await r.json();
    msg.style.display='block';msg.style.background='#450a0a';msg.style.color='#ef4444';msg.style.border='1px solid #991b1b';
    msg.textContent=d.error?.message||lt('fail');
  }catch(e){msg.style.display='block';msg.style.background='#450a0a';msg.style.color='#ef4444';msg.style.border='1px solid #991b1b';msg.textContent=lt('neterr')}
  finally{btn.disabled=false}
}

// Merged status: fetch token status + chromium login status, render in fixed order:
// 用户名 > 登录 > 有效 > 过期时间 > 剩余 > 自动刷新 > 标题 > 页面 > 错误
async function loadStatus(){
  try{
    const [tr,cr]=await Promise.all([
      fetch('/admin/token/status',{credentials:'include'}),
      fetch('/admin/chromium/login-status',{credentials:'include'}).catch(()=>null),
    ]);
    if(tr.status===401){showInlineLogin();return}
    const d=await tr.json();
    let c={};
    if(cr&&cr.ok){try{c=await cr.json()}catch(e){c={}}}
    const v=d.valid;
    const cls=v?'valid':'invalid';
    const exp=d.expires_at?new Date(d.expires_at).toLocaleString():'N/A';
    if(d.username)window.__m365_username=d.username;
    const row=(label,val,vcls)=>'<div class="status-row"><span class="status-label">'+label+'</span><span class="status-value '+(vcls||'')+'">'+val+'</span></div>';
    const warnCls=(v&&d.seconds_remaining<600)?'warn':'';
    let html='';
    // 1. 用户名
    if(d.username)html+=row(t('username_label'),d.username,'valid');
    // 2. 登录 (chromium) — 状态显示为 是/否
    if(c.chromium_running===false){
      html+=row(t('login'),t('chromium_not_running'),'invalid');
    }else if(c.chromium_running){
      html+=row(t('login'),c.logged_in?t('status_yes'):t('status_no'),c.logged_in?'valid':'warn');
    }
    const logoutBtn=document.getElementById('btn-logout');
    if(logoutBtn)logoutBtn.style.display=c.logged_in?'inline-block':'none';
    // 3. 自动刷新（紧跟登录下方）
    html+=row(t('auto_refresh_label'),d.auto_refresh?t('status_yes'):t('status_no'),d.auto_refresh?'valid':'warn');
    // 4. 有效
    html+=row(t('valid'),v?t('status_yes'):t('status_no'),cls);
    // 5. 过期时间
    html+=row(t('expires'),exp,warnCls);
    // 6. 剩余
    html+=row(t('remaining'),'<span id="remaining-sec">'+fmtSec(d.seconds_remaining)+'</span>',warnCls);
    // 7. 标题 (chromium)
    if(c.title)html+='<div class="status-row"><span class="status-label">'+t('title')+'</span><span class="status-value" style="font-size:.75rem">'+c.title+'</span></div>';
    // 8. 页面 (chromium)
    if(c.url)html+='<div class="status-row"><span class="status-label">'+t('page')+'</span><span class="status-value" style="font-size:.75rem;word-break:break-all">'+c.url+'</span></div>';
    // 9. 错误
    if(d.error)html+=row(t('error'),d.error,'invalid');
    document.getElementById('status-content').innerHTML=html;
    startCountdown(d.seconds_remaining||0);
    updateRefreshBtn(d.auto_refresh);
  }catch(e){
    document.getElementById('status-content').innerHTML='<span class="invalid">Failed to load</span>';
  }
}

// Kept as a thin alias so existing init/interval calls still work; loadStatus now
// renders both token and chromium status together in the required order.
async function loadChromiumStatus(){return loadStatus()}

function fmtSec(s){
  if(!s&&s!==0)return'N/A';
  const h=Math.floor(s/3600),m=Math.floor(s%3600/60),sec=s%60;
  return(h?h+'h ':'')+(m?m+'m ':'')+sec+'s';
}

function updateRefreshBtn(enabled){
  const btn=document.getElementById('btn-stop-refresh');
  if(enabled){
    btn.style.display='inline-block';
    btn.style.background='linear-gradient(135deg,#ef4444,#dc2626)';
    btn.textContent=t('btn_stop_refresh');
  }else{
    btn.style.display='inline-block';
    btn.style.background='linear-gradient(135deg,#22c55e,#059669)';
    btn.textContent=t('btn_start_refresh');
  }
}

async function toggleAutoRefresh(){
  const msg=document.getElementById('update-msg');
  const btn=document.getElementById('btn-stop-refresh');
  btn.disabled=true;msg.className='msg';msg.textContent='';
  try{
    const r=await fetch('/admin/token/auto-refresh-toggle',{method:'POST',credentials:'include'});
    const d=await r.json();
    if(r.ok){
      msg.className='msg ok';msg.textContent=d.auto_refresh?t('auto_refresh_started'):t('auto_refresh_stopped');
      updateRefreshBtn(d.auto_refresh);
      loadStatus();
    }else{
      msg.className='msg err';msg.textContent=d.error?.message||d.error||'Toggle failed';
    }
  }catch(e){msg.className='msg err';msg.textContent=(lang==='zh'?'网络错误：':'Network error: ')+e}
  finally{btn.disabled=false}
}

async function updateToken(){
  const input=document.getElementById('token-input').value.trim();
  const msg=document.getElementById('update-msg');
  const btn=document.getElementById('btn-update');
  if(!input){msg.className='msg err';msg.textContent=lang==='zh'?'请粘贴 Token':'Please paste a token';return}
  btn.disabled=true;msg.className='msg';msg.textContent='';
  try{
    const r=await fetch('/admin/token/update',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({token:input})});
    const d=await r.json();
    if(r.ok){
      msg.className='msg ok';msg.textContent=(lang==='zh'?'Token 已更新！剩余：':'Token updated! Remaining: ')+fmtSec(d.token_status?.seconds_remaining);
      document.getElementById('token-input').value='';
      loadStatus();
    }else{
      msg.className='msg err';msg.textContent=d.error?.message||d.error||(lang==='zh'?'更新失败':'Update failed');
    }
  }catch(e){msg.className='msg err';msg.textContent=(lang==='zh'?'网络错误：':'Network error: ')+e}
  finally{btn.disabled=false}
}

async function autoCapture(){
  const msg=document.getElementById('update-msg');
  const btn=document.getElementById('btn-auto');
  const upd=document.getElementById('btn-update');
  btn.disabled=true;upd.disabled=true;
  msg.className='msg';msg.textContent='';
  btn.textContent=t('capturing_btn');
  try{
    const r=await fetch('/admin/token/auto-capture',{method:'POST'});
    const d=await r.json();
    if(r.ok){
      msg.className='msg ok';msg.textContent=t('auto_captured')+fmtSec(d.token_status?.seconds_remaining);
      loadStatus();
    }else{
      msg.className='msg err';msg.textContent=d.error?.message||d.error||t('auto_capture_failed');
    }
  }catch(e){msg.className='msg err';msg.textContent=(lang==='zh'?'网络错误：':'Network error: ')+e}
  finally{btn.disabled=false;upd.disabled=false;btn.textContent=t('btn_auto_capture')}
}

async function checkLogin(){
  loadChromiumStatus();
  const msg=document.getElementById('update-msg');
  msg.className='msg';msg.textContent=t('check_login');
  await new Promise(r=>setTimeout(r,1500));
  try{
    const r=await fetch('/admin/chromium/login-status',{credentials:'include'});
    const d=await r.json();
    msg.className=d.logged_in?'msg ok':'msg err';
    msg.textContent=d.logged_in?t('login_ok'):t('login_not_ok');
  }catch(e){msg.className='msg err';msg.textContent=t('check_failed')+e}
}

async function logoutUser(){
  const msg=document.getElementById('update-msg');
  const btn=document.getElementById('btn-logout');
  btn.disabled=true;msg.className='msg';msg.textContent=t('logging_out');
  try{
    const r=await fetch('/admin/chromium/logout',{method:'POST',credentials:'include'});
    const d=await r.json();
    if(r.ok){
      msg.className='msg ok';msg.textContent=t('logout_ok')+(d.message?' — '+d.message:'');
      loadChromiumStatus();loadStatus();
    }else{
      msg.className='msg err';msg.textContent=d.error?.message||d.error||t('logout_failed');
    }
  }catch(e){msg.className='msg err';msg.textContent=(lang==='zh'?'网络错误：':'Network error: ')+e}
  finally{btn.disabled=false}
}

// ============================ Multi-tenant admin JS ============================
function esc(s){return String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;')}
let __accounts=[];
async function loadAccounts(){
  const box=document.getElementById('accounts-content');
  if(!box)return;
  try{
    const r=await fetch('/admin/accounts',{credentials:'include'});
    if(r.status===401){box.innerHTML='<span style="color:#64748b">'+t('loading')+'</span>';return}
    const d=await r.json();
    __accounts=d.accounts||[];
    if(!__accounts.length){box.innerHTML='<span style="color:#64748b">'+t('no_accounts')+'</span>';return}
    let h='<table style="width:100%;border-collapse:collapse;font-size:.82rem"><thead><tr style="color:#94a3b8;text-align:left">'
      +'<th style="padding:.3rem">'+t('col_name')+'</th><th style="padding:.3rem">'+t('col_status')+'</th><th style="padding:.3rem">'+t('col_token')+'</th><th style="padding:.3rem;text-align:right">'+t('col_actions')+'</th></tr></thead><tbody>';
    __accounts.forEach(a=>{
      const st=a.token_status||{};
      const valid=st.valid;
      const rem=valid?(' '+Math.floor((st.seconds_remaining||0)/60)+'m'):'';
      const badge='<span style="padding:.1rem .5rem;border-radius:99px;font-size:.72rem;background:'+(valid?'#065f46':'#7f1d1d')+';color:'+(valid?'#d1fae5':'#fee2e2')+'">'+(valid?t('valid_short'):t('invalid_short'))+rem+'</span>';
      h+='<tr style="border-top:1px solid #334155">'
        +'<td style="padding:.4rem"><div>'+esc(a.name||a.id)+'</div><div style="color:#475569;font-size:.7rem">'+esc(a.id)+' · '+a.key_count+' key</div></td>'
        +'<td style="padding:.4rem">'+badge+'</td>'
        +'<td style="padding:.4rem;color:#64748b">'+esc(a.token_source)+'</td>'
        +'<td style="padding:.4rem;text-align:right;white-space:nowrap">'
        +'<button onclick="refreshAccount(\\''+a.id+'\\')" style="font-size:.72rem;padding:3px 8px">'+t('btn_refresh')+'</button> '
        +'<button onclick="pushAccountToken(\\''+a.id+'\\')" style="font-size:.72rem;padding:3px 8px;background:#334155">'+t('btn_push_token')+'</button> '
        +'<button onclick="delAccount(\\''+a.id+'\\')" style="font-size:.72rem;padding:3px 8px;background:linear-gradient(135deg,#ef4444,#dc2626)">'+t('btn_delete')+'</button>'
        +'</td></tr>';
    });
    h+='</tbody></table>';
    box.innerHTML=h;
  }catch(e){}
}
async function addAccount(){
  const name=prompt(t('acct_prompt_name'));
  if(name===null)return;
  const token=prompt(t('acct_prompt_token'))||'';
  try{
    const r=await fetch('/admin/accounts',{method:'POST',credentials:'include',headers:{'Content-Type':'application/json'},body:JSON.stringify({name:name,token:token})});
    if(!r.ok){const d=await r.json().catch(()=>({}));alert((d.error&&d.error.message)||'error');return}
    loadAccounts();loadKeys();
  }catch(e){}
}
async function refreshAccount(id){
  try{
    const r=await fetch('/admin/accounts/'+id+'/refresh',{method:'POST',credentials:'include'});
    const d=await r.json().catch(()=>({}));
    if(!r.ok)alert((d.error&&d.error.message)||'error');
    loadAccounts();
  }catch(e){}
}
async function pushAccountToken(id){
  const token=prompt(t('push_token_prompt'));
  if(!token)return;
  try{
    const r=await fetch('/admin/accounts/'+id+'/token',{method:'POST',credentials:'include',headers:{'Content-Type':'application/json'},body:JSON.stringify({token:token})});
    if(!r.ok){const d=await r.json().catch(()=>({}));alert((d.error&&d.error.message)||'error');return}
    loadAccounts();
  }catch(e){}
}
async function delAccount(id){
  if(!confirm(t('confirm_del_account')))return;
  try{await fetch('/admin/accounts/'+id,{method:'DELETE',credentials:'include'});loadAccounts();loadKeys()}catch(e){}
}
let __keys=[];
async function loadKeys(){
  const box=document.getElementById('keys-content');
  if(!box)return;
  try{
    const r=await fetch('/admin/keys',{credentials:'include'});
    if(r.status===401){box.innerHTML='<span style="color:#64748b">'+t('loading')+'</span>';return}
    const d=await r.json();
    __keys=d.keys||[];
    if(!__keys.length){box.innerHTML='<span style="color:#64748b">'+t('no_keys')+'</span>';return}
    let h='<table style="width:100%;border-collapse:collapse;font-size:.82rem"><thead><tr style="color:#94a3b8;text-align:left">'
      +'<th style="padding:.3rem">'+t('col_name')+'</th><th style="padding:.3rem">'+t('col_key')+'</th><th style="padding:.3rem">'+t('col_account')+'</th><th style="padding:.3rem">'+t('col_mode')+'</th><th style="padding:.3rem;text-align:right">'+t('col_actions')+'</th></tr></thead><tbody>';
    __keys.forEach(k=>{
      const acc=k.account_id?esc(k.account_name||k.account_id):('<span style="color:#f59e0b">'+t('unbound')+'</span>');
      const en=k.enabled;
      h+='<tr style="border-top:1px solid #334155;'+(en?'':'opacity:.5')+'">'
        +'<td style="padding:.4rem">'+esc(k.name||k.id)+'</td>'
        +'<td style="padding:.4rem"><code style="font-size:.72rem;color:#818cf8">'+esc(k.key.slice(0,10))+'…</code> <button onclick="copyKey(\\''+k.id+'\\')" style="font-size:.68rem;padding:2px 6px;background:#334155">'+t('btn_copy')+'</button></td>'
        +'<td style="padding:.4rem">'+acc+'</td>'
        +'<td style="padding:.4rem;color:#64748b">'+esc(k.tone)+'</td>'
        +'<td style="padding:.4rem;text-align:right;white-space:nowrap">'
        +'<button onclick="rebindKey(\\''+k.id+'\\')" style="font-size:.72rem;padding:3px 8px;background:#334155">'+t('btn_rebind')+'</button> '
        +'<button onclick="toggleKey(\\''+k.id+'\\','+(en?'false':'true')+')" style="font-size:.72rem;padding:3px 8px;background:'+(en?'#b45309':'#059669')+'">'+(en?t('btn_disable'):t('btn_enable'))+'</button> '
        +'<button onclick="delKey(\\''+k.id+'\\')" style="font-size:.72rem;padding:3px 8px;background:linear-gradient(135deg,#ef4444,#dc2626)">'+t('btn_delete')+'</button>'
        +'</td></tr>';
    });
    h+='</tbody></table>';
    box.innerHTML=h;
  }catch(e){}
}
async function addKey(){
  const name=prompt(t('key_prompt_name'));
  if(name===null)return;
  let account_id='';
  if(__accounts.length){account_id=prompt(t('rebind_prompt')+'\\n'+__accounts.map(a=>a.id+' = '+(a.name||'')).join('\\n'))||''}
  try{
    const r=await fetch('/admin/keys',{method:'POST',credentials:'include',headers:{'Content-Type':'application/json'},body:JSON.stringify({name:name,account_id:account_id})});
    if(!r.ok){const d=await r.json().catch(()=>({}));alert((d.error&&d.error.message)||'error');return}
    loadKeys();loadAccounts();
  }catch(e){}
}
function copyKey(id){
  const k=__keys.find(x=>x.id===id);if(!k)return;
  navigator.clipboard.writeText(k.key).then(()=>{},()=>{});
}
async function rebindKey(id){
  const account_id=prompt(t('rebind_prompt')+'\\n'+__accounts.map(a=>a.id+' = '+(a.name||'')).join('\\n'));
  if(account_id===null)return;
  try{
    const r=await fetch('/admin/keys/'+id,{method:'POST',credentials:'include',headers:{'Content-Type':'application/json'},body:JSON.stringify({account_id:account_id})});
    if(!r.ok){const d=await r.json().catch(()=>({}));alert((d.error&&d.error.message)||'error');return}
    loadKeys();loadAccounts();
  }catch(e){}
}
async function toggleKey(id,enabled){
  try{await fetch('/admin/keys/'+id,{method:'POST',credentials:'include',headers:{'Content-Type':'application/json'},body:JSON.stringify({enabled:enabled})});loadKeys()}catch(e){}
}
async function delKey(id){
  if(!confirm(t('confirm_del_key')))return;
  try{await fetch('/admin/keys/'+id,{method:'DELETE',credentials:'include'});loadKeys();loadAccounts()}catch(e){}
}

loadStatus();
loadChromiumStatus();
loadCallLog();
loadCapture();
loadTone();
loadToolPrompt();
loadSystemPrompt();
loadAccounts();
loadKeys();
setInterval(loadStatus,60000);
setInterval(loadChromiumStatus,60000);
setInterval(loadCallLog,5000);
setInterval(loadCapture,5000);

// Client-side countdown timer
let _countdownSec=0;
let _countdownTick=0;
function startCountdown(sec){_countdownSec=sec;_countdownTick=0}
function tickCountdown(){
  if(_countdownSec<=0)return;
  _countdownSec--;_countdownTick++;
  const el=document.getElementById('remaining-sec');
  if(el)el.textContent=fmtSec(_countdownSec);
}
setInterval(tickCountdown,1000);

window.__callTexts={};
function copyCallText(key){
  const txt=window.__callTexts[key];
  if(txt==null)return;
  navigator.clipboard.writeText(txt).then(()=>{
    const b=document.getElementById('copybtn-'+key);
    if(b){const o=b.textContent;b.textContent=t('copied');setTimeout(()=>{b.textContent=o},1200)}
  }).catch(()=>{});
}
async function loadCallLog(){
  try{
    const r=await fetch('/admin/call-log',{credentials:'include'});
    if(r.status===401){showInlineLogin();return}
    const d=await r.json();
    const logs=d.logs||[];
    document.getElementById('call-log-count').textContent=logs.length;
    const el=document.getElementById('call-log-content');
    if(!logs.length){el.innerHTML='<span style="color:#64748b">'+t('no_calls_yet')+'</span>';window.__callLogSig='';return}
    // Skip re-render if nothing changed — prevents open <details> from collapsing every 5s
    const sig=JSON.stringify(logs);
    if(sig===window.__callLogSig)return;
    window.__callLogSig=sig;
    window.__callTexts={};
    let html='';
    for(let i=logs.length-1;i>=0;i--){
      const l=logs[i];
      const esc=s=>String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
      const tc=l.tools&&l.tools.length?l.tools.join(', '):'—';
      const tr=l.tool_calls_result&&l.tool_calls_result.length?
        '<span style="color:#22c55e">'+t('tool_calls_parsed')+': '+l.tool_calls_result.join(', ')+'</span>':'';
      const reprKey='r'+i, textKey='x'+i, fullKey='f'+i;
      if(l.response_repr!=null)window.__callTexts[reprKey]=l.response_repr;
      if(l.response_text!=null)window.__callTexts[textKey]=l.response_text;
      // Full single-record text: call info + repr + text
      const fullParts=[];
      fullParts.push('time: '+l.time);
      fullParts.push('mode: '+(l.stream?'stream':'sync'));
      fullParts.push('tools: '+tc);
      if(l.tool_calls_result&&l.tool_calls_result.length)fullParts.push('tool_calls_result: '+l.tool_calls_result.join(', '));
      if(l.response_len!=null)fullParts.push('resp: '+l.response_len+' chars');
      if(l.response_repr!=null)fullParts.push('repr:\\n'+l.response_repr);
      if(l.response_text!=null)fullParts.push('text:\\n'+l.response_text);
      window.__callTexts[fullKey]=fullParts.join('\\n');
      const copyBtn=(key)=>'<button class="copybtn" id="copybtn-'+key+'" data-key="'+key+'" style="padding:2px 8px;font-size:.65rem;margin-left:6px">'+t('copy')+'</button>';
      const copyFullBtn='<button class="copybtn" id="copybtn-'+fullKey+'" data-key="'+fullKey+'" style="padding:2px 8px;font-size:.65rem">'+t('copy_record')+'</button>';
      const respView=(l.response_repr||l.response_text)?
        '<details style="margin-top:4px"><summary style="cursor:pointer;color:#64748b;font-size:.75rem;list-style:none">'+t('view_raw')+'</summary>'+
        (l.response_repr?'<div style="display:flex;align-items:center;color:#475569;margin-top:4px;font-size:.7rem">repr:'+copyBtn(reprKey)+'</div><pre style="white-space:pre-wrap;word-break:break-all;background:#0f172a;padding:6px;border-radius:6px;color:#94a3b8;margin-top:2px;font-size:.7rem;max-height:200px;overflow:auto">'+esc(l.response_repr)+'</pre>':'')+
        (l.response_text?'<div style="display:flex;align-items:center;color:#475569;margin-top:4px;font-size:.7rem">text:'+copyBtn(textKey)+'</div><pre style="white-space:pre-wrap;word-break:break-all;background:#0f172a;padding:6px;border-radius:6px;color:#e2e8f0;margin-top:2px;font-size:.7rem;max-height:300px;overflow:auto">'+esc(l.response_text)+'</pre>':'')+
        '</details>':'';
      html+='<div style="border-bottom:1px solid #1e293b;padding:6px 0">'+
        '<div style="display:flex;justify-content:space-between;align-items:center;color:#94a3b8">'+
        '<span>'+l.time+'</span><span style="display:flex;align-items:center;gap:6px"><span style="color:#475569">'+(l.stream?'stream':'sync')+'</span>'+copyFullBtn+'</span></div>'+
        '<div style="color:#e2e8f0;margin-top:2px">tools: <span style="color:#38bdf8">'+tc+'</span></div>'+
        (l.incremental!=null?'<div style="color:#475569;margin-top:2px">incremental: <span style="color:'+(l.incremental?'#22c55e':'#f59e0b')+'">'+(l.incremental?'yes':'no')+'</span> &nbsp; turn: '+(l.turn_count==null?'-':l.turn_count)+'</div>':'')+
        (tr?'<div style="margin-top:2px">'+tr+'</div>':'')+
        (l.response_len?'<div style="color:#475569;margin-top:2px">resp: '+l.response_len+' chars</div>':'')+
        respView+
        '</div>';
    }
    el.innerHTML=html;
    el.querySelectorAll('.copybtn').forEach(function(b){
      b.addEventListener('click',function(){copyCallText(b.getAttribute('data-key'))});
    });
  }catch(e){}
}
async function loadCapture(){
  try{
    const r=await fetch('/admin/capture-payload',{credentials:'include'});
    if(r.status===401){return}
    const d=await r.json();
    const ps=d.payloads||[];
    document.getElementById('capture-count').textContent=ps.length;
    const el=document.getElementById('capture-content');
    if(!ps.length){el.innerHTML='<span style="color:#64748b">'+t('no_capture_yet')+'</span>';window.__capSig='';return}
    const sig=JSON.stringify(ps);
    if(sig===window.__capSig)return;
    window.__capSig=sig;
    const esc=s=>String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
    let html='';
    for(let i=0;i<ps.length;i++){
      const p=ps[i];
      const opts=(p.optionsSets||[]).join(', ');
      const gpt=p.gptId&&Object.keys(p.gptId).length?JSON.stringify(p.gptId):'-';
      html+='<div style="border-bottom:1px solid #1e293b;padding:6px 0;line-height:1.5">'+
        '<div style="color:#38bdf8">'+esc(p.time)+' &nbsp; tone: <b>'+esc(p.tone||'-')+'</b> &nbsp; model: <b>'+esc(p.modelId||'-')+'</b></div>'+
        '<div style="color:#94a3b8">gptId: '+esc(gpt)+'</div>'+
        '<div style="color:#64748b;word-break:break-all">optionsSets: '+esc(opts)+'</div>'+
        '<details style="margin-top:4px"><summary style="cursor:pointer;color:#64748b;font-size:.72rem;list-style:none">'+t('view_raw')+'</summary>'+
        '<pre style="white-space:pre-wrap;word-break:break-all;background:#0f172a;padding:6px;border-radius:6px;color:#94a3b8;margin-top:2px;font-size:.7rem;max-height:240px;overflow:auto">'+esc(JSON.stringify(p.raw,null,2))+'</pre></details>'+
        '</div>';
    }
    el.innerHTML=html;
  }catch(e){}
}
async function loadTone(){
  try{
    const r=await fetch('/admin/tone',{credentials:'include'});
    if(r.status===401){return}
    const d=await r.json();
    const sel=document.getElementById('tone-select');
    if(!sel)return;
    const cur=d.tone||'Magic';
    const opts=d.options||[];
    window.__toneOpts=opts;
    // Skip re-render if unchanged (avoids resetting an open dropdown). Signature
    // includes lang so switching language re-renders the localized labels.
    const sig=JSON.stringify(opts)+'|'+cur+'|'+lang;
    if(sig===window.__toneSig)return;
    window.__toneSig=sig;
    const lbl=o=>(lang==='en'?(o.label_en||o.label):(o.label_zh||o.label))||o.label;
    sel.innerHTML=opts.map(o=>'<option value="'+o.value+'"'+(o.value===cur?' selected':'')+'>'+lbl(o)+'</option>').join('');
    sel.onchange=()=>saveTone(sel.value);
  }catch(e){}
}
async function saveTone(tone){
  try{
    const r=await fetch('/admin/tone',{method:'POST',credentials:'include',headers:{'Content-Type':'application/json'},body:JSON.stringify({tone})});
    if(!r.ok)return;
    window.__toneSig='';
    const s=document.getElementById('tone-saved');
    if(s){s.textContent=t('tone_saved');s.style.opacity='1';setTimeout(()=>{s.style.opacity='0'},1500)}
  }catch(e){}
}
async function loadToolPrompt(){
  try{
    const r=await fetch('/admin/tool-prompt',{credentials:'include'});
    if(r.status===401){return}
    const d=await r.json();
    const ta=document.getElementById('tool-prompt-input');
    if(!ta)return;
    if(document.activeElement!==ta)ta.value=d.tool_prompt||'';
  }catch(e){}
}
async function saveToolPrompt(){
  try{
    const ta=document.getElementById('tool-prompt-input');
    if(!ta)return;
    const r=await fetch('/admin/tool-prompt',{method:'POST',credentials:'include',headers:{'Content-Type':'application/json'},body:JSON.stringify({tool_prompt:ta.value})});
    if(!r.ok)return;
    const s=document.getElementById('tool-prompt-saved');
    if(s){s.textContent=t('tool_prompt_saved');s.style.opacity='1';setTimeout(()=>{s.style.opacity='0'},1500)}
  }catch(e){}
}
async function resetToolPrompt(){
  // Extra instruction default is empty.
  const ta=document.getElementById('tool-prompt-input');
  if(ta)ta.value='';
  await saveToolPrompt();
}

let __systemPromptDefault='';
async function loadSystemPrompt(){
  try{
    const r=await fetch('/admin/system-prompt',{credentials:'include'});
    if(r.status===401){return}
    const d=await r.json();
    __systemPromptDefault=d.default||'';
    const ta=document.getElementById('system-prompt-input');
    if(!ta)return;
    // Show the saved override, or fall back to the default text for reference.
    if(document.activeElement!==ta)ta.value=(d.system_prompt&&d.system_prompt.length)?d.system_prompt:__systemPromptDefault;
  }catch(e){}
}
function unlockSystemPrompt(){
  if(!confirm(t('system_prompt_warn')))return;
  const locked=document.getElementById('system-prompt-locked');
  const editor=document.getElementById('system-prompt-editor');
  if(locked)locked.style.display='none';
  if(editor)editor.style.display='block';
}
async function saveSystemPrompt(){
  try{
    const ta=document.getElementById('system-prompt-input');
    if(!ta)return;
    const r=await fetch('/admin/system-prompt',{method:'POST',credentials:'include',headers:{'Content-Type':'application/json'},body:JSON.stringify({system_prompt:ta.value})});
    if(!r.ok)return;
    const s=document.getElementById('system-prompt-saved');
    if(s){s.textContent=t('tool_prompt_saved');s.style.opacity='1';setTimeout(()=>{s.style.opacity='0'},1500)}
  }catch(e){}
}
async function resetSystemPrompt(){
  if(!confirm(t('system_prompt_reset_confirm')))return;
  const ta=document.getElementById('system-prompt-input');
  // Saving an empty override makes the backend fall back to the built-in default.
  if(ta)ta.value=__systemPromptDefault;
  try{
    const r=await fetch('/admin/system-prompt',{method:'POST',credentials:'include',headers:{'Content-Type':'application/json'},body:JSON.stringify({system_prompt:''})});
    if(!r.ok)return;
    const s=document.getElementById('system-prompt-saved');
    if(s){s.textContent=t('tool_prompt_saved');s.style.opacity='1';setTimeout(()=>{s.style.opacity='0'},1500)}
  }catch(e){}
}

</script>
</body>
</html>"""


_USER_HTML = """<!DOCTYPE html>
<html lang="zh">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>M365 Copilot Proxy - User</title>
<style>
*{box-sizing:border-box}
body{margin:0;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;background:#0f172a;color:#e2e8f0;line-height:1.5}
.wrap{max-width:760px;margin:0 auto;padding:1.5rem 1rem 3rem}
h1{font-size:1.3rem;margin:0}
.top{display:flex;align-items:center;justify-content:space-between;margin-bottom:1.2rem}
.card{background:#1e293b;border:1px solid #334155;border-radius:10px;padding:1rem 1.2rem;margin-bottom:1rem}
.card h2{font-size:1rem;margin:0 0 .8rem;color:#e2e8f0}
label{display:block;font-size:.85rem;color:#94a3b8;margin:.6rem 0 .3rem}
input,select,textarea{width:100%;background:#0f172a;border:1px solid #334155;border-radius:6px;color:#e2e8f0;padding:.5rem .6rem;font-size:.9rem;font-family:inherit}
textarea{resize:vertical;min-height:70px}
button{background:linear-gradient(135deg,#6366f1,#4f46e5);color:#fff;border:0;border-radius:6px;padding:.5rem 1rem;font-size:.9rem;cursor:pointer;margin-top:.6rem}
button:disabled{opacity:.5;cursor:not-allowed}
.btn-ghost{background:#334155}
.row{display:flex;gap:.5rem;align-items:center}
.row>*{margin-top:0}
.pill{display:inline-block;font-size:.75rem;padding:.15rem .5rem;border-radius:99px;background:#334155;color:#cbd5e1}
.pill.ok{background:#065f46;color:#d1fae5}
.pill.bad{background:#7f1d1d;color:#fee2e2}
.msg{font-size:.8rem;margin-left:.5rem;opacity:0;transition:opacity .2s;color:#86efac}
.hint{font-size:.8rem;color:#64748b;margin-bottom:.4rem}
.hidden{display:none}
a{color:#818cf8}
</style>
</head>
<body>
<div class="wrap">
  <div class="top">
    <h1 data-i18n="title">M365 Copilot 代理 · 用户</h1>
    <button class="btn-ghost" id="lang-toggle" onclick="toggleLang()">&#127760; EN</button>
  </div>

  <div id="login-card" class="card">
    <h2 data-i18n="login_title">登录</h2>
    <div class="hint" data-i18n="login_hint">输入你的 API Key 以管理自己的对话模式、提示词与账户 Token。</div>
    <input id="apikey" type="password" data-i18n-ph="apikey_ph" placeholder="sk-...">
    <div class="row"><button id="login-btn" onclick="doLogin()" data-i18n="login_btn">登录</button><span id="login-msg" class="msg"></span></div>
  </div>

  <div id="app" class="hidden">
    <div class="card">
      <h2 data-i18n="account_title">账户与 Token</h2>
      <div id="account-info"></div>
      <label data-i18n="push_token_label">推送 / 更新账户 Token</label>
      <div class="hint" data-i18n="push_token_hint">粘贴 access_token 值或完整 wss:// URL。若尚未绑定账户，将自动创建并绑定。</div>
      <textarea id="acct-token" placeholder="access_token / wss://substrate.office.com/..."></textarea>
      <div class="row"><button onclick="pushToken()" data-i18n="push_token_btn">推送 Token</button><span id="token-msg" class="msg"></span></div>
    </div>

    <div class="card">
      <h2 data-i18n="tone_title">对话模式</h2>
      <div class="row"><select id="tone" onchange="saveTone()"></select><span id="tone-msg" class="msg"></span></div>
    </div>

    <div class="card">
      <h2 data-i18n="tool_prompt_title">提示词增强</h2>
      <div class="hint" data-i18n="tool_prompt_hint">追加到工具调用提示词后的自定义指令，仅作用于你自己的 Key。留空则不追加。</div>
      <textarea id="tool-prompt"></textarea>
      <div class="row"><button onclick="saveToolPrompt()" data-i18n="save">保存</button><span id="tool-msg" class="msg"></span></div>
    </div>

    <details class="card">
      <summary style="cursor:pointer;font-weight:600" data-i18n="sys_prompt_title">系统提示词（高级）</summary>
      <div class="hint" style="margin-top:.6rem" data-i18n="sys_prompt_hint">覆盖工具调用的基础系统提示词。改错会导致工具调用失效，仅供高级用户调试。留空则使用内置默认。</div>
      <textarea id="sys-prompt"></textarea>
      <div class="row"><button onclick="saveSysPrompt()" data-i18n="save">保存</button><button class="btn-ghost" onclick="resetSysPrompt()" data-i18n="reset">恢复默认</button><span id="sys-msg" class="msg"></span></div>
    </details>

    <div class="card">
      <h2 data-i18n="endpoints_title">API 端点</h2>
      <div class="hint">Base URL: <code id="base-url"></code></div>
      <div class="hint" data-i18n="endpoints_hint">在你的 OpenAI 兼容客户端里填入上面的 Base URL 和你的 API Key。</div>
    </div>
  </div>
</div>

<script>
const i18n={
  zh:{
    title:'M365 Copilot 代理 · 用户',
    login_title:'登录',login_hint:'输入你的 API Key 以管理自己的对话模式、提示词与账户 Token。',
    apikey_ph:'sk-...',login_btn:'登录',login_failed:'登录失败，请检查 API Key',network_error:'网络错误',
    account_title:'账户与 Token',push_token_label:'推送 / 更新账户 Token',
    push_token_hint:'粘贴 access_token 值或完整 wss:// URL。若尚未绑定账户，将自动创建并绑定。',
    push_token_btn:'推送 Token',saved:'已保存',push_ok:'Token 已更新',
    tone_title:'对话模式',tool_prompt_title:'提示词增强',
    tool_prompt_hint:'追加到工具调用提示词后的自定义指令，仅作用于你自己的 Key。留空则不追加。',
    save:'保存',reset:'恢复默认',
    sys_prompt_title:'系统提示词（高级）',
    sys_prompt_hint:'覆盖工具调用的基础系统提示词。改错会导致工具调用失效，仅供高级用户调试。留空则使用内置默认。',
    sys_prompt_reset_confirm:'确定要将系统提示词恢复为内置默认吗？当前自定义内容将被清空。',
    endpoints_title:'API 端点',endpoints_hint:'在你的 OpenAI 兼容客户端里填入上面的 Base URL 和你的 API Key。',
    logout:'登出',no_account:'尚未绑定账户，推送 Token 后将自动创建。',
    key_name:'名称',bound_account:'绑定账户',token_valid:'有效',token_invalid:'无效/缺失',remaining:'剩余',
  },
  en:{
    title:'M365 Copilot Proxy · User',
    login_title:'Login',login_hint:'Enter your API key to manage your own conversation mode, prompts and account token.',
    apikey_ph:'sk-...',login_btn:'Login',login_failed:'Login failed, check your API key',network_error:'Network error',
    account_title:'Account & Token',push_token_label:'Push / update account token',
    push_token_hint:'Paste the access_token value or the full wss:// URL. If no account is bound yet, one will be created and bound automatically.',
    push_token_btn:'Push Token',saved:'Saved',push_ok:'Token updated',
    tone_title:'Conversation Mode',tool_prompt_title:'Prompt Enhancement',
    tool_prompt_hint:'Custom instruction appended after the tool-call prompt, applies only to your own key. Leave empty to append nothing.',
    save:'Save',reset:'Restore default',
    sys_prompt_title:'System Prompt (Advanced)',
    sys_prompt_hint:'Overrides the base system prompt for tool calls. A wrong edit will break tool calling. For advanced debugging only. Leave empty to use the built-in default.',
    sys_prompt_reset_confirm:'Restore the system prompt to the built-in default? Your current custom content will be cleared.',
    endpoints_title:'API Endpoints',endpoints_hint:'Point your OpenAI-compatible client at the Base URL above with your API key.',
    logout:'Logout',no_account:'No account bound yet. Pushing a token will create one automatically.',
    key_name:'Name',bound_account:'Bound account',token_valid:'Valid',token_invalid:'Invalid/Missing',remaining:'Remaining',
  }
};
let lang=localStorage.getItem('lang')||'zh';
let toneOptions=[];
let sysDefault='';
function t(k){return i18n[lang][k]||k}
function getKey(){return localStorage.getItem('user_api_key')||''}
function authHeaders(){return {'Content-Type':'application/json','Authorization':'Bearer '+getKey()}}
function applyLang(){
  const btn=document.getElementById('lang-toggle');
  btn.innerHTML=lang==='zh'?'&#127760; EN':'&#127760; 中文';
  document.querySelectorAll('[data-i18n]').forEach(el=>{const k=el.getAttribute('data-i18n');if(i18n[lang][k])el.textContent=i18n[lang][k]});
  document.querySelectorAll('[data-i18n-ph]').forEach(el=>{const k=el.getAttribute('data-i18n-ph');if(i18n[lang][k])el.placeholder=i18n[lang][k]});
  renderToneOptions();
}
function toggleLang(){lang=lang==='zh'?'en':'zh';localStorage.setItem('lang',lang);applyLang()}
function renderToneOptions(){
  const sel=document.getElementById('tone');if(!sel||!toneOptions.length)return;
  const cur=sel.value;
  sel.innerHTML='';
  toneOptions.forEach(o=>{
    const opt=document.createElement('option');
    opt.value=o.value;
    opt.textContent=lang==='zh'?(o.label_zh||o.label):(o.label_en||o.label);
    sel.appendChild(opt);
  });
  if(cur)sel.value=cur;
}
function flash(id){const s=document.getElementById(id);if(!s)return;s.textContent=t('saved');s.style.opacity='1';setTimeout(()=>{s.style.opacity='0'},1500)}
async function doLogin(){
  const key=document.getElementById('apikey').value.trim();
  const msg=document.getElementById('login-msg');
  if(!key)return;
  localStorage.setItem('user_api_key',key);
  const ok=await loadMe();
  if(!ok){msg.className='msg';msg.style.color='#fca5a5';msg.style.opacity='1';msg.textContent=t('login_failed');localStorage.removeItem('user_api_key')}
}
function logout(){localStorage.removeItem('user_api_key');document.getElementById('app').classList.add('hidden');document.getElementById('login-card').classList.remove('hidden')}
async function loadMe(){
  if(!getKey())return false;
  try{
    const r=await fetch('/user/me',{headers:authHeaders()});
    if(!r.ok)return false;
    const d=await r.json();
    toneOptions=d.tone_options||[];
    sysDefault=d.default_system_prompt||'';
    document.getElementById('login-card').classList.add('hidden');
    document.getElementById('app').classList.remove('hidden');
    document.getElementById('base-url').textContent=location.origin+'/v1';
    renderToneOptions();
    document.getElementById('tone').value=d.tone||'Magic';
    document.getElementById('tool-prompt').value=d.tool_prompt||'';
    document.getElementById('sys-prompt').value=d.system_prompt||'';
    let acc='';
    if(d.account){
      const st=d.account.token_status||{};
      const valid=st.valid;
      const rem=valid?(' · '+t('remaining')+' '+Math.floor((st.seconds_remaining||0)/60)+'m'):'';
      acc='<div class="row" style="flex-wrap:wrap;gap:.4rem"><span class="pill">'+t('bound_account')+': '+(d.account.name||d.account.id)+'</span>'
        +'<span class="pill '+(valid?'ok':'bad')+'">'+(valid?t('token_valid'):t('token_invalid'))+rem+'</span></div>';
    }else{
      acc='<div class="hint">'+t('no_account')+'</div>';
    }
    acc+='<div style="margin-top:.6rem"><button class="btn-ghost" onclick="logout()">'+t('logout')+'</button></div>';
    document.getElementById('account-info').innerHTML=acc;
    return true;
  }catch(e){return false}
}
async function pushToken(){
  const token=document.getElementById('acct-token').value.trim();
  if(!token)return;
  try{
    const r=await fetch('/user/account/token',{method:'POST',headers:authHeaders(),body:JSON.stringify({token:token})});
    if(r.ok){document.getElementById('acct-token').value='';flash('token-msg');loadMe()}
  }catch(e){}
}
async function saveTone(){
  const tone=document.getElementById('tone').value;
  try{await fetch('/user/tone',{method:'POST',headers:authHeaders(),body:JSON.stringify({tone:tone})});flash('tone-msg')}catch(e){}
}
async function saveToolPrompt(){
  const p=document.getElementById('tool-prompt').value;
  try{await fetch('/user/tool-prompt',{method:'POST',headers:authHeaders(),body:JSON.stringify({tool_prompt:p})});flash('tool-msg')}catch(e){}
}
async function saveSysPrompt(){
  const p=document.getElementById('sys-prompt').value;
  try{await fetch('/user/system-prompt',{method:'POST',headers:authHeaders(),body:JSON.stringify({system_prompt:p})});flash('sys-msg')}catch(e){}
}
async function resetSysPrompt(){
  if(!confirm(t('sys_prompt_reset_confirm')))return;
  document.getElementById('sys-prompt').value='';
  try{await fetch('/user/system-prompt',{method:'POST',headers:authHeaders(),body:JSON.stringify({system_prompt:''})});flash('sys-msg')}catch(e){}
}
applyLang();
loadMe();
</script>
</body>
</html>"""
