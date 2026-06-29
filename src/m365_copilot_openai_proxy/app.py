from __future__ import annotations

import json
import re
import time
import uuid
from collections.abc import AsyncIterator, Callable
from pathlib import Path

from fastapi import Depends, FastAPI, HTTPException, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse, JSONResponse, StreamingResponse

from .config import Settings
from .session_store import PersistentSession, PersistentSessionStore
from .substrate_client import SubstrateCopilotClient, SubstrateCopilotError
from .token_store import AccessTokenStore
from .models import AnthropicMessagesRequest, OpenAIChatRequest, OpenAIResponsesRequest
from .translator import translate_anthropic_request, translate_openai_request, translate_responses_request

_PERSIST_MODEL_SUFFIX = ":persist"
_SESSION_ID_HEADER = "x-m365-session-id"


def create_app(
    settings: Settings | None = None,
    copilot_client_factory: Callable[[], SubstrateCopilotClient] | None = None,
) -> FastAPI:
    app = FastAPI(title="Microsoft 365 Copilot OpenAI Proxy")
    resolved_settings = settings or Settings()
    app.state.settings = resolved_settings
    app.state.token_store = AccessTokenStore(resolved_settings.access_token)
    app.state.session_store = PersistentSessionStore()
    app.state.copilot_client_factory = copilot_client_factory or (
        lambda: SubstrateCopilotClient(app.state.token_store.get(), resolved_settings.time_zone)
    )

    # CORS middleware
    app.add_middleware(
        CORSMiddleware,
        allow_origins=["*"],
        allow_methods=["GET", "POST", "OPTIONS"],
        allow_headers=["Content-Type", "Authorization", "x-m365-session-id"],
        max_age=86400,
    )

    # API Key authentication middleware
    @app.middleware("http")
    async def api_key_auth(request: Request, call_next):
        if not resolved_settings.api_key:
            return await call_next(request)
        # Skip auth for admin page and health endpoints
        path = request.url.path
        if path in ("/", "/healthz", "/v1/token/status", "/v1/token/update", "/v1/token/auto-capture", "/v1/cookie/inject"):
            return await call_next(request)
        auth = request.headers.get("Authorization", "")
        match = re.match(r"^Bearer\s+(.+)$", auth, re.IGNORECASE)
        if match and match.group(1) == resolved_settings.api_key:
            return await call_next(request)
        return JSONResponse(
            status_code=401,
            content={"error": {"message": "Invalid API key", "type": "auth_error"}},
        )

    def get_settings() -> Settings:
        return app.state.settings

    def get_copilot_client() -> SubstrateCopilotClient:
        return app.state.copilot_client_factory()

    @app.get("/healthz")
    async def healthz() -> dict:
        return {"status": "ok", "token": app.state.token_store.status()}

    @app.get("/v1/token/status")
    async def token_status() -> dict:
        return app.state.token_store.status()

    @app.post("/v1/token/update")
    async def update_token(request: Request) -> dict:
        body = await request.json()
        token = body.get("token", "").strip()
        if not token:
            return JSONResponse(status_code=400, content={"error": "Token is empty"})
        # Extract token from full WebSocket URL if needed
        match = re.search(r"access_token=([^&\s]+)", token)
        if match:
            token = match.group(1)
        if not token.startswith("eyJ"):
            return JSONResponse(status_code=400, content={"error": "Not a valid JWT token"})
        # Write to .env
        env_path = Path(".env")
        token_line_pattern = r"(?m)^M365_ACCESS_TOKEN=.*$"
        if env_path.exists():
            text = env_path.read_text(encoding="utf-8")
            if re.search(token_line_pattern, text):
                text = re.sub(token_line_pattern, f"M365_ACCESS_TOKEN={token}", text)
            else:
                text += f"\nM365_ACCESS_TOKEN={token}\n"
        else:
            text = f"M365_ACCESS_TOKEN={token}\n"
        env_path.write_text(text, encoding="utf-8")
        # Update in-memory store
        app.state.token_store._token = token
        app.state.token_store._mtime_ns = None
        return {"status": "ok", "message": "Token updated", "token_status": app.state.token_store.status()}

    @app.post("/v1/token/auto-capture")
    async def auto_capture_token() -> dict:
        """Auto-capture token from Chromium CDP running inside the container."""
        import asyncio
        from .cli import _cdp_extract_token, _write_token
        cdp_port = int(resolved_settings.api_key or "9222")  # not api_key, get from env
        cdp_port = 9222  # always use default CDP port inside container
        try:
            token = await _cdp_extract_token(cdp_port, allow_nudge=True)
        except Exception as exc:
            return JSONResponse(status_code=502, content={"error": f"CDP capture failed: {exc}"})
        if not token:
            return JSONResponse(status_code=404, content={"error": "No substrate token found. Make sure M365 Copilot is open and logged in in Chromium."})
        # Write to .env and update in-memory
        _write_token(token)
        app.state.token_store._token = token
        app.state.token_store._mtime_ns = None
        return {"status": "ok", "message": "Token auto-captured", "token_status": app.state.token_store.status()}

    @app.post("/v1/cookie/inject")
    async def inject_cookie(request: Request) -> dict:
        """Inject cookies into Chromium via CDP to log in to M365."""
        body = await request.json()
        cookies = body.get("cookies", [])
        if not cookies:
            return JSONResponse(status_code=400, content={"error": "No cookies provided"})
        import asyncio as _async
        import httpx as _httpx
        import websockets as _ws

        cdp_port = 9222
        try:
            async with _httpx.AsyncClient(timeout=3) as client:
                tabs = (await client.get(f"http://localhost:{cdp_port}/json")).json()
        except Exception as exc:
            return JSONResponse(status_code=502, content={"error": f"Cannot connect to Chromium CDP: {exc}"})

        tab = next((t for t in tabs if t.get("type") == "page" and t.get("url", "").startswith("https://m365.cloud.microsoft/")), None)
        if not tab:
            tab = next((t for t in tabs if t.get("type") == "page"), None)
        if not tab:
            return JSONResponse(status_code=404, content={"error": "No browser tab found in Chromium"})

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
                    if cookie.get("sameSite"):
                        ss = cookie["sameSite"].capitalize()
                        if ss in ("Strict", "Lax", "None"):
                            cookie_params["sameSite"] = ss
                    if cookie.get("expirationDate"):
                        cookie_params["expires"] = cookie["expirationDate"]
                    await ws.send(json.dumps({"id": 100 + i, "method": "Network.setCookie", "params": cookie_params}))
                    try:
                        resp = await _async.wait_for(ws.recv(), timeout=5)
                        result = json.loads(resp)
                        if result.get("result", {}).get("success"):
                            injected += 1
                    except (_async.TimeoutError, Exception):
                        pass

                await ws.send(json.dumps({"id": 999, "method": "Page.reload"}))
        except Exception as exc:
            return JSONResponse(status_code=502, content={"error": f"CDP cookie injection failed: {exc}"})

        return {"status": "ok", "message": f"Injected {injected}/{len(cookies)} cookies. Page reloading.", "injected": injected, "total": len(cookies)}

    @app.get("/", response_class=HTMLResponse)
    async def admin_page() -> str:
        return _ADMIN_HTML

    @app.get("/v1/models")
    async def list_models(settings: Settings = Depends(get_settings)) -> dict:
        return {
            "object": "list",
            "data": [
                {
                    "id": settings.model_alias,
                    "object": "model",
                    "owned_by": "microsoft-365-copilot",
                }
            ],
        }

    @app.post("/v1/chat/completions")
    async def chat_completions(
        raw_request: Request,
        request: OpenAIChatRequest,
        settings: Settings = Depends(get_settings),
        client: SubstrateCopilotClient = Depends(get_copilot_client),
    ):
        try:
            translated = translate_openai_request(request)
            session = _persistent_session(app, raw_request, request.model, request.user)
            if request.stream:
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
) -> PersistentSession | None:
    header_key = (raw_request.headers.get(_SESSION_ID_HEADER) or "").strip()
    if header_key:
        return app.state.session_store.get(f"header:{header_key}")
    if model.endswith(_PERSIST_MODEL_SUFFIX):
        return app.state.session_store.get(f"model:{fallback_key or 'default'}")
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


_ADMIN_HTML = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>M365 Copilot Proxy Admin</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:-apple-system,BlinkMacSystemFont,Segoe UI,Roboto,sans-serif;background:#0f172a;color:#e2e8f0;min-height:100vh;padding:2rem}
.container{max-width:720px;margin:0 auto}
h1{font-size:1.5rem;margin-bottom:1.5rem;background:linear-gradient(135deg,#06b6d4,#8b5cf6);-webkit-background-clip:text;-webkit-text-fill-color:transparent}
.card{background:#1e293b;border-radius:12px;padding:1.5rem;margin-bottom:1.5rem;border:1px solid #334155}
.card h2{font-size:1.1rem;margin-bottom:1rem;color:#94a3b8}
.status-row{display:flex;justify-content:space-between;align-items:center;padding:.5rem 0;border-bottom:1px solid #334155}
.status-row:last-child{border:none}
.status-label{color:#94a3b8;font-size:.9rem}
.status-value{font-weight:600;font-size:.9rem}
.valid{color:#22c55e}.invalid{color:#ef4444}.warn{color:#f59e0b}
textarea{width:100%;height:120px;background:#0f172a;border:1px solid #475569;border-radius:8px;color:#e2e8f0;padding:.75rem;font-family:monospace;font-size:.8rem;resize:vertical;margin-bottom:.75rem}
textarea:focus{outline:none;border-color:#06b6d4}
button{background:linear-gradient(135deg,#06b6d4,#8b5cf6);color:#fff;border:none;border-radius:8px;padding:.7rem 1.5rem;font-size:.9rem;font-weight:600;cursor:pointer;transition:opacity .2s}
button:hover{opacity:.85}
button:disabled{opacity:.5;cursor:not-allowed}
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
<h1>M365 Copilot Proxy</h1>

<div class="card">
<h2>Token Status</h2>
<div id="status-content"><span style="color:#64748b">Loading...</span></div>
</div>

<div class="card">
<h2>Update Token</h2>
<p style="color:#64748b;font-size:.85rem;margin-bottom:.75rem">Paste the access_token value or the full wss:// URL from <a href="https://m365.cloud.microsoft/chat" target="_blank">M365 Copilot</a></p>
<textarea id="token-input" placeholder="eyJ0eXAiOiJKV1QiLCJhbGci...&#10;&#10;or full URL:&#10;wss://substrate.office.com/m365Copilot/Chathub/...?access_token=eyJ..."></textarea>
<div style="display:flex;gap:.75rem;margin-bottom:.25rem">
<button id="btn-update" onclick="updateToken()">Update Token</button>
<button id="btn-auto" onclick="autoCapture()" style="background:linear-gradient(135deg,#22c55e,#059669)">Auto Capture</button>
</div>
<div id="update-msg" class="msg"></div>
</div>

<div class="card">
<h2>API Endpoints</h2>
<div class="api-info">
GET  /healthz<br>
GET  /v1/token/status<br>
POST /v1/token/update<br>
POST /v1/token/auto-capture<br>
POST /v1/cookie/inject<br>
GET  /v1/models<br>
POST /v1/chat/completions<br>
POST /v1/responses<br>
POST /v1/messages
</div>
</div>

<div class="card">
<h2>Cookie Login</h2>
<p style="color:#64748b;font-size:.85rem;margin-bottom:.75rem">Inject M365 cookies into Chromium to log in. <a href="https://m365.cloud.microsoft/chat" target="_blank">Open M365 Copilot</a> in your browser first.</p>
<div style="display:flex;gap:.75rem;margin-bottom:.75rem">
<button id="btn-export" onclick="copyCookieScript()" style="background:linear-gradient(135deg,#f59e0b,#d97706);font-size:.8rem;padding:.5rem 1rem">1. Copy Cookie Script</button>
<button id="btn-inject" onclick="injectCookies()" style="background:linear-gradient(135deg,#8b5cf6,#6d28d9);font-size:.8rem;padding:.5rem 1rem">2. Inject Cookies</button>
</div>
<textarea id="cookie-input" placeholder="Paste cookies JSON here, e.g.:&#10;{&#10;  "cookies": [&#10;    {"name": "...", "value": "...", "domain": ".microsoft.com"},&#10;    ...&#10;  ]&#10;}" style="height:100px"></textarea>
<div id="cookie-msg" class="msg"></div>
</div>
</div>

<script>
async function loadStatus(){
  try{
    const r=await fetch('/v1/token/status');
    const d=await r.json();
    const v=d.valid;
    const cls=v?'valid':'invalid';
    const exp=d.expires_at?new Date(d.expires_at).toLocaleString():'N/A';
    document.getElementById('status-content').innerHTML=
      '<div class="status-row"><span class="status-label">Valid</span><span class="status-value '+cls+'">'+(v?'Yes':'No')+'</span></div>'+
      '<div class="status-row"><span class="status-label">Expires</span><span class="status-value '+(v&&d.seconds_remaining<600?'warn':'')+'">'+exp+'</span></div>'+
      '<div class="status-row"><span class="status-label">Remaining</span><span class="status-value '+(v&&d.seconds_remaining<600?'warn':'')+'">'+fmtSec(d.seconds_remaining)+'</span></div>'+
      (d.error?'<div class="status-row"><span class="status-label">Error</span><span class="status-value invalid">'+d.error+'</span></div>':'');
  }catch(e){
    document.getElementById('status-content').innerHTML='<span class="invalid">Failed to load</span>';
  }
}

function fmtSec(s){
  if(!s&&s!==0)return'N/A';
  const h=Math.floor(s/3600),m=Math.floor(s%3600/60),sec=s%60;
  return(h?h+'h ':'')+(m?m+'m ':'')+sec+'s';
}

async function updateToken(){
  const input=document.getElementById('token-input').value.trim();
  const msg=document.getElementById('update-msg');
  const btn=document.getElementById('btn-update');
  if(!input){msg.className='msg err';msg.textContent='Please paste a token';return}
  btn.disabled=true;msg.className='msg';msg.textContent='';
  try{
    const r=await fetch('/v1/token/update',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({token:input})});
    const d=await r.json();
    if(r.ok){
      msg.className='msg ok';msg.textContent='Token updated! Remaining: '+fmtSec(d.token_status?.seconds_remaining);
      document.getElementById('token-input').value='';
      loadStatus();
    }else{
      msg.className='msg err';msg.textContent=d.error||'Update failed';
    }
  }catch(e){msg.className='msg err';msg.textContent='Network error: '+e}
  finally{btn.disabled=false}
}

async function autoCapture(){
  const msg=document.getElementById('update-msg');
  const btn=document.getElementById('btn-auto');
  const upd=document.getElementById('btn-update');
  btn.disabled=true;upd.disabled=true;
  msg.className='msg';msg.textContent='';
  btn.textContent='Capturing...';
  try{
    const r=await fetch('/v1/token/auto-capture',{method:'POST'});
    const d=await r.json();
    if(r.ok){
      msg.className='msg ok';msg.textContent='Auto-captured! Remaining: '+fmtSec(d.token_status?.seconds_remaining);
      loadStatus();
    }else{
      msg.className='msg err';msg.textContent=d.error||'Auto-capture failed';
    }
  }catch(e){msg.className='msg err';msg.textContent='Network error: '+e}
  finally{btn.disabled=false;upd.disabled=false;btn.textContent='Auto Capture'}
}

function copyCookieScript(){
  const script=`(function(){const c=document.cookie.split(';').map(x=>{const[n,...v]=x.trim().split('=');return{name:n,value:v.join('='),domain:location.hostname,path:'/',secure:location.protocol==='https:',httpOnly:false,sameSite:'Lax'}});const d=JSON.stringify({cookies:c});navigator.clipboard.writeText(d).then(()=>alert('Cookies copied! Go back to Admin page and paste into Cookie input.')).catch(()=>{const t=document.createElement('textarea');t.value=d;document.body.appendChild(t);t.select();document.execCommand('copy');document.body.removeChild(t);alert('Cookies copied!');})})()`;
  navigator.clipboard.writeText(script).then(()=>{document.getElementById('cookie-msg').className='msg ok';document.getElementById('cookie-msg').textContent='Script copied! Now: 1) Open M365 Copilot, 2) Press F12 -> Console, 3) Paste & run the script, 4) Come back and paste result into Cookie input.'}).catch(()=>{document.getElementById('cookie-msg').className='msg err';document.getElementById('cookie-msg').textContent='Copy failed. Please copy manually.'});
}

async function injectCookies(){
  const msg=document.getElementById('cookie-msg');
  const btn=document.getElementById('btn-inject');
  const input=document.getElementById('cookie-input').value.trim();
  if(!input){msg.className='msg err';msg.textContent='Please paste cookies JSON first';return}
  let data;
  try{data=JSON.parse(input)}catch(e){msg.className='msg err';msg.textContent='Invalid JSON format';return}
  if(!data.cookies||!Array.isArray(data.cookies)){msg.className='msg err';msg.textContent='JSON must have a "cookies" array';return}
  btn.disabled=true;btn.textContent='Injecting...';msg.className='msg';msg.textContent='';
  try{
    const r=await fetch('/v1/cookie/inject',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(data)});
    const d=await r.json();
    if(r.ok){msg.className='msg ok';msg.textContent=d.message}
    else{msg.className='msg err';msg.textContent=d.error||'Injection failed'}
  }catch(e){msg.className='msg err';msg.textContent='Network error: '+e}
  finally{btn.disabled=false;btn.textContent='2. Inject Cookies'}
}

loadStatus();
setInterval(loadStatus,60000);
</script>
</body>
</html>"""
