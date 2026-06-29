// ==UserScript==
// @name         M365 Copilot Token & Cookie Extractor
// @namespace    https://m365.cloud.microsoft
// @version      2.0
// @description  拦截 M365 Copilot Substrate WebSocket 连接，提取 access_token 和 Cookie
// @match        https://m365.cloud.microsoft/*
// @grant        none
// ==/UserScript==

(function() {
    'use strict';

    const SUBSTRATE_WS_RE = /wss:\/\/substrate\.office\.com\/.*[?&]access_token=([^&]+)/;
    const PROXY_BASE = ''; // 留空则从面板输入框读取，或填入你的代理地址如 http://192.168.1.100:8000

    // 拦截 WebSocket 构造
    const OrigWebSocket = window.WebSocket;
    window.WebSocket = function(url, protocols) {
        const match = url.match(SUBSTRATE_WS_RE);
        if (match) {
            const token = match[1];
            showTokenPanel(token);
        }
        return new OrigWebSocket(url, protocols);
    };
    window.WebSocket.prototype = OrigWebSocket.prototype;
    window.WebSocket.CONNECTING = OrigWebSocket.CONNECTING;
    window.WebSocket.OPEN = OrigWebSocket.OPEN;
    window.WebSocket.CLOSING = OrigWebSocket.CLOSING;
    window.WebSocket.CLOSED = OrigWebSocket.CLOSED;

    function getCookies() {
        return document.cookie.split(';').map(c => {
            const [name, ...rest] = c.trim().split('=');
            return {
                name,
                value: rest.join('='),
                domain: location.hostname,
                path: '/',
                secure: location.protocol === 'https:',
                httpOnly: false,
                sameSite: 'Lax'
            };
        });
    }

    function getProxyBase() {
        const input = document.getElementById('m365-proxy-url');
        return input ? input.value.trim().replace(/\/+$/, '') : PROXY_BASE;
    }

    // 推送 Token 到代理
    async function pushToken(token) {
        const base = getProxyBase();
        if (!base) { alert('Please enter proxy URL first'); return; }
        try {
            const r = await fetch(base + '/v1/token/update', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ token })
            });
            const d = await r.json();
            alert(r.ok ? `Token pushed! Remaining: ${d.token_status?.seconds_remaining}s` : `Failed: ${d.error}`);
        } catch (e) { alert('Network error: ' + e); }
    }

    // 推送 Cookie 到代理
    async function pushCookies() {
        const base = getProxyBase();
        if (!base) { alert('Please enter proxy URL first'); return; }
        const cookies = getCookies();
        try {
            const r = await fetch(base + '/v1/cookie/inject', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ cookies })
            });
            const d = await r.json();
            alert(r.ok ? `Cookies pushed! ${d.message}` : `Failed: ${d.error}`);
        } catch (e) { alert('Network error: ' + e); }
    }

    // 复制 Cookie JSON
    function copyCookies() {
        const cookies = getCookies();
        const data = JSON.stringify({ cookies }, null, 2);
        navigator.clipboard.writeText(data).then(() => alert('Cookies JSON copied!')).catch(() => alert('Copy failed'));
    }

    function showTokenPanel(token) {
        if (document.getElementById('m365-token-panel')) {
            document.getElementById('m365-token-panel').remove();
        }

        const panel = document.createElement('div');
        panel.id = 'm365-token-panel';
        panel.innerHTML = `
            <div style="position:fixed; top:10px; right:10px; z-index:99999;
                        background:#1a1a2e; color:#e0e0e0; padding:16px 20px;
                        border-radius:10px; font-family:monospace; font-size:13px;
                        box-shadow:0 4px 20px rgba(0,0,0,0.5); max-width:520px;
                        border:1px solid #16213e; max-height:90vh; overflow-y:auto;">
                <div style="font-weight:bold; font-size:15px; margin-bottom:8px; color:#00d2ff;">
                    M365 Copilot Proxy Tool
                </div>

                <div style="margin-bottom:10px;">
                    <div style="font-size:11px; color:#8892b0; margin-bottom:4px;">Proxy URL</div>
                    <input id="m365-proxy-url" type="text" placeholder="http://your-server:8000"
                        value="${PROXY_BASE}"
                        style="width:100%; padding:6px 10px; background:#0f0f23; border:1px solid #475569;
                               border-radius:6px; color:#e0e0e0; font-size:12px; font-family:monospace;">
                </div>

                <div style="font-size:11px; color:#8892b0; margin-bottom:4px;">Token</div>
                <div style="word-break:break-all; max-height:80px; overflow-y:auto;
                            background:#0f0f23; padding:8px; border-radius:6px;
                            font-size:10px; color:#a8b2d1; line-height:1.4;">
                    ${token.slice(0, 80)}...
                </div>

                <div style="margin-top:10px; display:flex; flex-wrap:wrap; gap:6px;">
                    <button id="m365-copy-token" style="padding:5px 12px; border:none;
                            border-radius:6px; background:#00d2ff; color:#1a1a2e;
                            cursor:pointer; font-weight:bold; font-size:12px;">
                        Copy Token
                    </button>
                    <button id="m365-push-token" style="padding:5px 12px; border:none;
                            border-radius:6px; background:#22c55e; color:#fff;
                            cursor:pointer; font-weight:bold; font-size:12px;">
                        Push Token to Proxy
                    </button>
                </div>

                <div style="border-top:1px solid #334155; margin:12px 0 10px; padding-top:10px;">
                    <div style="font-size:11px; color:#8892b0; margin-bottom:6px;">Cookie Tools</div>
                    <div style="display:flex; flex-wrap:wrap; gap:6px;">
                        <button id="m365-copy-cookies" style="padding:5px 12px; border:none;
                                border-radius:6px; background:#f59e0b; color:#1a1a2e;
                                cursor:pointer; font-weight:bold; font-size:12px;">
                            Copy Cookies
                        </button>
                        <button id="m365-push-cookies" style="padding:5px 12px; border:none;
                                border-radius:6px; background:#8b5cf6; color:#fff;
                                cursor:pointer; font-weight:bold; font-size:12px;">
                            Push Cookies to Proxy
                        </button>
                        <button id="m365-close-panel" style="padding:5px 12px; border:none;
                                border-radius:6px; background:#e94560; color:#fff;
                                cursor:pointer; font-weight:bold; font-size:12px;">
                            Close
                        </button>
                    </div>
                </div>
            </div>
        `;
        document.body.appendChild(panel);

        document.getElementById('m365-copy-token').onclick = () => {
            navigator.clipboard.writeText(token).then(() => {
                const btn = document.getElementById('m365-copy-token');
                btn.textContent = 'Copied!';
                setTimeout(() => { btn.textContent = 'Copy Token'; }, 1500);
            });
        };

        document.getElementById('m365-push-token').onclick = () => pushToken(token);
        document.getElementById('m365-copy-cookies').onclick = () => copyCookies();
        document.getElementById('m365-push-cookies').onclick = () => pushCookies();
        document.getElementById('m365-close-panel').onclick = () => panel.remove();
    }
})();
