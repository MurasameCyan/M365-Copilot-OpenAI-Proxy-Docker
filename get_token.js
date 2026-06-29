// ==UserScript==
// @name         M365 Copilot Token & Cookie Extractor
// @namespace    https://m365.cloud.microsoft
// @version      3.1
// @description  拦截 M365 Copilot Substrate WebSocket 连接，提取 access_token 并推送到代理服务
// @match        https://m365.cloud.microsoft/*
// @grant        none
// ==/UserScript==

(function() {
    'use strict';

    const SUBSTRATE_WS_RE = /wss:\/\/substrate\.office\.com\/.*[?&]access_token=([^&]+)/;
    const PROXY_BASE = ''; // 留空则从面板输入框读取，或填入你的代理地址如 http://192.168.1.100:8000

    // Store the latest token
    let latestToken = '';

    // Intercept WebSocket construction
    const OrigWebSocket = window.WebSocket;
    window.WebSocket = function(url, protocols) {
        const match = url.match(SUBSTRATE_WS_RE);
        if (match) {
            latestToken = match[1];
            showPanel();
        }
        return new OrigWebSocket(url, protocols);
    };
    window.WebSocket.prototype = OrigWebSocket.prototype;
    window.WebSocket.CONNECTING = OrigWebSocket.CONNECTING;
    window.WebSocket.OPEN = OrigWebSocket.OPEN;
    window.WebSocket.CLOSING = OrigWebSocket.CLOSING;
    window.WebSocket.CLOSED = OrigWebSocket.CLOSED;

    // Get cookies that document.cookie can see (non-httpOnly only)
    function getVisibleCookies() {
        return document.cookie.split(';').map(c => {
            const [name, ...rest] = c.trim().split('=');
            return {
                name,
                value: rest.join('='),
                domain: location.hostname,
                path: '/',
                secure: true,
                httpOnly: false,
                sameSite: 'None'
            };
        });
    }

    function getProxyBase() {
        const input = document.getElementById('m365-proxy-url');
        return input ? input.value.trim().replace(/\/+$/, '') : PROXY_BASE;
    }

    // Push Token to proxy
    async function pushToken() {
        const base = getProxyBase();
        if (!base) { alert('Please enter proxy URL first'); return; }
        if (!latestToken) { alert('No token captured yet. Type something in Copilot to trigger WebSocket.'); return; }
        try {
            const r = await fetch(base + '/v1/token/update', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ token: latestToken })
            });
            const d = await r.json();
            alert(r.ok ? `Token pushed! Remaining: ${d.token_status?.seconds_remaining}s` : `Failed: ${d.error?.message || d.error}`);
        } catch (e) { alert('Network error: ' + e + '\n\nMake sure the proxy server has CORS enabled and is reachable.'); }
    }

    // Push visible cookies to proxy (press button in panel)
    async function pushCookies() {
        const base = getProxyBase();
        if (!base) { alert('Please enter proxy URL first'); return; }
        const cookies = getVisibleCookies();
        if (!cookies.length) { alert('No cookies found on this page.'); return; }
        try {
            const r = await fetch(base + '/v1/cookie/inject', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ cookies })
            });
            const d = await r.json();
            alert(r.ok ? `Cookies pushed! ${d.message}` : `Failed: ${d.error?.message || d.error}`);
        } catch (e) { alert('Network error: ' + e); }
    }

    // Copy visible cookies JSON
    function copyCookies() {
        const cookies = getVisibleCookies();
        const data = JSON.stringify({ cookies }, null, 2);
        navigator.clipboard.writeText(data).then(() => alert('Cookies JSON copied! (httpOnly cookies not included - use browser extension for those)')).catch(() => alert('Copy failed'));
    }

    // Copy token to clipboard
    function copyToken() {
        if (!latestToken) { alert('No token captured yet'); return; }
        navigator.clipboard.writeText(latestToken).then(() => alert('Token copied!')).catch(() => alert('Copy failed'));
    }

    // One-click: push token then auto-capture (no cookie push - httpOnly cookies not accessible via document.cookie)
    async function oneClickSetup() {
        const base = getProxyBase();
        if (!base) { alert('Please enter proxy URL first'); return; }
        if (!latestToken) { alert('No token captured yet. Type something in Copilot to trigger WebSocket first.'); return; }

        const btn = document.getElementById('m365-one-click');
        btn.textContent = 'Working...';
        btn.disabled = true;

        try {
            // Step 1: Push token
            btn.textContent = 'Pushing token...';
            const tr = await fetch(base + '/v1/token/update', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ token: latestToken })
            });
            const td = await tr.json();
            if (!tr.ok) { alert('Token push failed: ' + (td.error?.message || td.error)); return; }

            // Step 2: Auto-capture (optional, to sync Chromium state)
            btn.textContent = 'Auto-capturing...';
            await new Promise(r => setTimeout(r, 2000));
            const cr = await fetch(base + '/v1/token/auto-capture', { method: 'POST' });
            const cd = await cr.json();

            if (cr.ok) {
                alert(`Setup complete! Token remaining: ${cd.token_status?.seconds_remaining}s`);
            } else {
                alert(`Token pushed OK (${td.token_status?.seconds_remaining}s remaining). Auto-capture skipped: ${cd.error?.message || cd.error}`);
            }
        } catch (e) {
            alert('Error: ' + e);
        } finally {
            btn.textContent = 'One-Click Setup';
            btn.disabled = false;
        }
    }

    function showPanel() {
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

                <div style="font-size:11px; color:#8892b0; margin-bottom:4px;">Token<span style="color:#64748b"> (truncated)</span></div>
                <div style="word-break:break-all; max-height:60px; overflow-y:auto;
                            background:#0f0f23; padding:8px; border-radius:6px;
                            font-size:10px; color:#a8b2d1; line-height:1.4;">
                    ${latestToken ? latestToken.slice(0, 80) + '...' : 'No token captured yet'}
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
                        Push Token
                    </button>
                </div>

                <div style="border-top:1px solid #334155; margin:12px 0 10px; padding-top:10px;">
                    <div style="font-size:11px; color:#8892b0; margin-bottom:6px;">Cookie Tools <span style="color:#f59e0b">NOTE: httpOnly cookies not accessible via document.cookie</span></div>
                    <div style="display:flex; flex-wrap:wrap; gap:6px;">
                        <button id="m365-copy-cookies" style="padding:5px 12px; border:none;
                                border-radius:6px; background:#f59e0b; color:#1a1a2e;
                                cursor:pointer; font-weight:bold; font-size:12px;">
                            Copy Cookies
                        </button>
                        <button id="m365-push-cookies" style="padding:5px 12px; border:none;
                                border-radius:6px; background:#8b5cf6; color:#fff;
                                cursor:pointer; font-weight:bold; font-size:12px;">
                            Push Cookies
                        </button>
                    </div>
                </div>

                <div style="border-top:1px solid #334155; margin:12px 0 10px; padding-top:10px;">
                    <div style="font-size:11px; color:#22c55e; margin-bottom:6px; font-weight:bold;">Quick Setup</div>
                    <div style="font-size:10px; color:#8892b0; margin-bottom:6px;">Push captured token to proxy then auto-capture</div>
                    <button id="m365-one-click" style="padding:5px 12px; border:none;
                            border-radius:6px; background:linear-gradient(135deg,#06b6d4,#22c55e); color:#fff;
                            cursor:pointer; font-weight:bold; font-size:12px; width:100%;">
                        One-Click Setup
                    </button>
                </div>

                <div style="border-top:1px solid #334155; margin:12px 0 0; padding-top:10px;">
                    <button id="m365-close-panel" style="padding:5px 12px; border:none;
                            border-radius:6px; background:#e94560; color:#fff;
                            cursor:pointer; font-weight:bold; font-size:12px;">
                        Close
                    </button>
                </div>
            </div>
        `;
        document.body.appendChild(panel);

        document.getElementById('m365-copy-token').onclick = () => copyToken();
        document.getElementById('m365-push-token').onclick = () => pushToken();
        document.getElementById('m365-copy-cookies').onclick = () => copyCookies();
        document.getElementById('m365-push-cookies').onclick = () => pushCookies();
        document.getElementById('m365-one-click').onclick = () => oneClickSetup();
        document.getElementById('m365-close-panel').onclick = () => panel.remove();
    }

    // Show panel on demand via keyboard shortcut (Ctrl+Shift+M)
    document.addEventListener('keydown', (e) => {
        if (e.ctrlKey && e.shiftKey && e.key === 'M') {
            showPanel();
        }
    });
})();
