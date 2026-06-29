// ==UserScript==
// @name         M365 Copilot Token & Cookie Extractor
// @namespace    https://m365.cloud.microsoft
// @version      4.0
// @description  提取 M365 Copilot 完整 Cookie（含 httpOnly）推送到代理服务实现登录
// @match        https://m365.cloud.microsoft/*
// @match        https://login.microsoftonline.com/*
// @match        https://microsoftonline.com/*
// @match        https://www.office.com/*
// @match        https://office.com/*
// @match        https://microsoft.com/*
// @grant        GM_cookie
// @grant        GM_xmlhttpRequest
// @grant        unsafeWindow
// @connect      *
// ==/UserScript==

(function() {
    'use strict';

    const SUBSTRATE_WS_RE = /wss:\/\/substrate\.office\.com\/.*[?&]access_token=([^&]+)/;
    const PROXY_BASE = ''; // 留空则从面板输入框读取，或填入你的代理地址如 http://192.168.1.100:8000

    // Domains whose cookies are needed for M365 login
    const COOKIE_DOMAINS = [
        'https://m365.cloud.microsoft',
        'https://login.microsoftonline.com',
        'https://microsoftonline.com',
        'https://microsoft.com',
        'https://office.com',
        'https://www.office.com',
    ];

    // Store the latest token
    let latestToken = '';

    // Intercept WebSocket construction on the real page (not in sandbox)
    const pageWindow = typeof unsafeWindow !== 'undefined' ? unsafeWindow : window;

    const OrigWebSocket = pageWindow.WebSocket;
    pageWindow.WebSocket = function(url, protocols) {
        const match = url.match(SUBSTRATE_WS_RE);
        if (match) {
            latestToken = match[1];
            showPanel();
        }
        return new OrigWebSocket(url, protocols);
    };
    pageWindow.WebSocket.prototype = OrigWebSocket.prototype;
    pageWindow.WebSocket.CONNECTING = OrigWebSocket.CONNECTING;
    pageWindow.WebSocket.OPEN = OrigWebSocket.OPEN;
    pageWindow.WebSocket.CLOSING = OrigWebSocket.CLOSING;
    pageWindow.WebSocket.CLOSED = OrigWebSocket.CLOSED;

    function getProxyBase() {
        const input = document.getElementById('m365-proxy-url');
        return input ? input.value.trim().replace(/\/+$/, '') : PROXY_BASE;
    }

    // Cross-origin fetch via GM_xmlhttpRequest
    function gmFetch(url, options) {
        return new Promise((resolve, reject) => {
            GM_xmlhttpRequest({
                method: options.method || 'GET',
                url: url,
                headers: options.headers || {},
                data: options.body || null,
                responseType: 'json',
                onload: (resp) => {
                    resolve({
                        ok: resp.status >= 200 && resp.status < 300,
                        status: resp.status,
                        json: () => Promise.resolve(resp.response || {}),
                    });
                },
                onerror: (err) => reject(new Error('GM_xmlhttpRequest error: ' + err)),
                ontimeout: () => reject(new Error('GM_xmlhttpRequest timeout')),
            });
        });
    }

    // Get ALL cookies (including httpOnly) via GM_cookie
    async function getAllCookies() {
        const allCookies = [];
        const seen = new Set();

        function addCookie(c) {
            const key = c.name + '@' + c.domain;
            if (seen.has(key)) return;
            seen.add(key);
            allCookies.push({
                name: c.name || '',
                value: c.value || '',
                domain: c.domain || '',
                path: c.path || '/',
                secure: c.secure !== false,
                httpOnly: !!c.httpOnly,
                sameSite: (c.sameSite || '').charAt(0).toUpperCase() + (c.sameSite || '').slice(1).toLowerCase() || 'None',
                expires: c.expirationDate || c.expires || undefined,
            });
        }

        function gmCookieList(details) {
            return new Promise((resolve) => {
                const timer = setTimeout(() => resolve([]), 1500);
                try {
                    GM_cookie.list(details, (c, err) => {
                        clearTimeout(timer);
                        if (err) { resolve([]); }
                        else { resolve(c || []); }
                    });
                } catch(e) { clearTimeout(timer); resolve([]); }
            });
        }

        // All queries to run in parallel
        const queries = [
            {},  // current document URL
            { url: 'https://m365.cloud.microsoft/' },
            { url: 'https://login.microsoftonline.com/' },
            { url: 'https://microsoftonline.com/' },
            { url: 'https://microsoft.com/' },
            { url: 'https://office.com/' },
            { url: 'https://www.office.com/' },
            { domain: '.login.microsoftonline.com' },
            { domain: '.microsoft.com' },
            { domain: '.microsoftonline.com' },
        ];

        // Run all queries in parallel
        const results = await Promise.all(queries.map(q => gmCookieList(q)));

        for (const cookies of results) {
            for (const c of (cookies || [])) {
                addCookie(c);
            }
        }

        console.log(`[M365 Proxy] Total cookies:`, allCookies.length, '(httpOnly:', allCookies.filter(c=>c.httpOnly).length, ')');
        return allCookies;
    }

    // Check if GM_cookie is available
    function hasGMCookie() {
        return (typeof GM_cookie !== 'undefined' && typeof GM_cookie.list === 'function') ||
            (typeof GM !== 'undefined' && GM.cookie && typeof GM.cookie.list === 'function');
    }

    // Push Token to proxy
    async function pushToken() {
        const base = getProxyBase();
        if (!base) { alert('Please enter proxy URL first'); return; }
        if (!latestToken) { alert('No token captured yet. Type something in Copilot to trigger WebSocket.'); return; }
        try {
            const r = await gmFetch(base + '/v1/token/update', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ token: latestToken })
            });
            const d = await r.json();
            alert(r.ok ? `Token pushed! Remaining: ${d.token_status?.seconds_remaining}s` : `Failed: ${d.error?.message || d.error}`);
        } catch (e) { alert('Network error: ' + e); }
    }

    // Push ALL cookies (including httpOnly) to proxy for Chromium login
    async function pushCookies() {
        const base = getProxyBase();
        if (!base) { alert('Please enter proxy URL first'); return; }

        if (!hasGMCookie()) {
            alert('GM_cookie API not available.\n\nPlease use Tampermonkey Beta or enable "Allow scripts to access HttpOnly cookies" in Tampermonkey settings:\nSettings > Security > "Allow scripts to access cookies"');
            return;
        }

        const btn = document.getElementById('m365-push-cookies');
        if (btn) { btn.disabled = true; btn.textContent = 'Fetching...'; }

        try {
            const cookies = await getAllCookies();
            if (!cookies.length) { alert('No cookies found.'); return; }

            if (btn) btn.textContent = 'Pushing...';
            const r = await gmFetch(base + '/v1/cookie/inject', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ cookies })
            });
            const d = await r.json();
            alert(r.ok ? `Cookies pushed! ${d.message}\n(httpOnly included: ${cookies.filter(c => c.httpOnly).length})` : `Failed: ${d.error?.message || d.error}`);
        } catch (e) {
            alert('Error: ' + e);
        } finally {
            if (btn) { btn.disabled = false; btn.textContent = 'Push Cookies'; }
        }
    }

    // Copy token to clipboard
    function copyToken() {
        if (!latestToken) { alert('No token captured yet'); return; }
        navigator.clipboard.writeText(latestToken).then(() => alert('Token copied!')).catch(() => alert('Copy failed'));
    }

    // One-click: push cookies first (to login Chromium), then push token
    async function oneClickSetup() {
        const base = getProxyBase();
        if (!base) { alert('Please enter proxy URL first'); return; }
        if (!latestToken) { alert('No token captured yet. Type something in Copilot to trigger WebSocket first.'); return; }

        const btn = document.getElementById('m365-one-click');
        btn.textContent = 'Working...';
        btn.disabled = true;

        try {
            // Step 1: Push cookies (if GM_cookie available) to login Chromium
            if (hasGMCookie()) {
                btn.textContent = '1/2 Pushing cookies...';
                const cookies = await getAllCookies();
                if (cookies.length) {
                    await gmFetch(base + '/v1/cookie/inject', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({ cookies })
                    });
                    // Wait for Chromium to process cookies and reload
                    await new Promise(r => setTimeout(r, 3000));
                }
            }

            // Step 2: Push token
            btn.textContent = '2/2 Pushing token...';
            const r = await gmFetch(base + '/v1/token/update', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ token: latestToken })
            });
            const d = await r.json();
            if (r.ok) {
                alert(`Setup complete! Token remaining: ${d.token_status?.seconds_remaining}s\nProxy is ready to use.`);
            } else {
                alert('Token push failed: ' + (d.error?.message || d.error));
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

        const gmCookieNote = hasGMCookie()
        ? '<span style="color:#22c55e">&#10003; GM_cookie available</span>'
        : '<span style="color:#f59e0b">&#9888; GM_cookie unavailable. Use Tampermonkey Beta.</span>';

        const panel = document.createElement('div');
        panel.id = 'm365-token-panel';
        panel.innerHTML = `
            <div style="position:fixed; top:10px; right:10px; z-index:99999;
                        background:linear-gradient(180deg,#0f172a 0%,#1e293b 100%);
                        color:#e2e8f0; padding:20px 24px;
                        border-radius:14px; font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',monospace; font-size:13px;
                        box-shadow:0 8px 32px rgba(0,0,0,0.6),0 0 0 1px rgba(148,163,184,0.1);
                        max-width:480px; width:calc(100vw - 20px); max-height:90vh; overflow-y:auto;
                        backdrop-filter:blur(12px);">
                <div style="font-weight:700; font-size:16px; margin-bottom:12px; color:#38bdf8;
                            letter-spacing:0.5px; display:flex; align-items:center; gap:8px;">
                    <span style="font-size:18px">&#9889;</span> Ciallo Ms-365 Proxy
                </div>

                <div style="margin-bottom:12px;">
                    <div style="font-size:11px; color:#94a3b8; margin-bottom:5px; font-weight:500;">Proxy URL</div>
                    <input id="m365-proxy-url" type="text" placeholder="http://your-server:8000"
                        value="${PROXY_BASE}"
                        style="width:100%; box-sizing:border-box; padding:8px 12px; background:#0f172a; border:1px solid #334155;
                               border-radius:8px; color:#e2e8f0; font-size:12px; font-family:monospace;
                               outline:none; transition:border-color 0.2s;"
                        onfocus="this.style.borderColor='#38bdf8'" onblur="this.style.borderColor='#334155'">
                </div>

                <div style="font-size:11px; color:#94a3b8; margin-bottom:5px; font-weight:500;">Token <span style="color:#475569">truncated</span></div>
                <div style="word-break:break-all; max-height:56px; overflow-y:auto;
                            background:#0f172a; padding:8px 12px; border-radius:8px;
                            font-size:11px; color:#a8b2d1; line-height:1.5;
                            border:1px solid #334155;">
                    ${latestToken ? latestToken.slice(0, 80) + '...' : '<span style="color:#475569">No token captured yet</span>'}
                </div>

                <div style="margin-top:12px; display:flex; gap:8px;">
                    <button id="m365-copy-token" style="flex:1; padding:8px 0; border:none;
                            border-radius:8px; background:#0ea5e9; color:#fff;
                            cursor:pointer; font-weight:600; font-size:12px;
                            transition:opacity 0.2s;" onmouseover="this.style.opacity=0.85" onmouseout="this.style.opacity=1">
                        &#128203; Copy Token
                    </button>
                    <button id="m365-push-token" style="flex:1; padding:8px 0; border:none;
                            border-radius:8px; background:#22c55e; color:#fff;
                            cursor:pointer; font-weight:600; font-size:12px;
                            transition:opacity 0.2s;" onmouseover="this.style.opacity=0.85" onmouseout="this.style.opacity=1">
                        &#128228; Push Token
                    </button>
                </div>

                <div style="border-top:1px solid #1e293b; margin:14px 0 12px; padding-top:12px;">
                    <div style="font-size:11px; color:#94a3b8; margin-bottom:8px; font-weight:500;">Cookie Login ${gmCookieNote}</div>
                    <button id="m365-push-cookies" style="width:100%; padding:8px 0; border:none;
                            border-radius:8px; background:linear-gradient(135deg,#8b5cf6,#7c3aed); color:#fff;
                            cursor:pointer; font-weight:600; font-size:12px;
                            transition:opacity 0.2s;" onmouseover="this.style.opacity=0.85" onmouseout="this.style.opacity=1">
                        &#127850; Push All Cookies
                    </button>
                </div>

                <div style="border-top:1px solid #1e293b; margin:0 0 12px; padding-top:12px;">
                    <div style="font-size:11px; color:#22c55e; margin-bottom:8px; font-weight:700;">&#9889; Quick Setup</div>
                    <div style="font-size:10px; color:#64748b; margin-bottom:8px;">Push cookies + token to proxy for Chromium login and auto-refresh</div>
                    <button id="m365-one-click" style="width:100%; padding:10px 0; border:none;
                            border-radius:8px; background:linear-gradient(135deg,#8b5cf6,#06b6d4,#22c55e); color:#fff;
                            cursor:pointer; font-weight:700; font-size:13px; letter-spacing:0.3px;
                            transition:opacity 0.2s;" onmouseover="this.style.opacity=0.85" onmouseout="this.style.opacity=1">
                        &#128640; One-Click Setup
                    </button>
                </div>

                <div style="border-top:1px solid #1e293b; margin:0; padding-top:12px; display:flex; justify-content:space-between; align-items:center;">
                    <span style="font-size:10px; color:#475569">Ctrl+Shift+M to toggle</span>
                    <button id="m365-close-panel" style="padding:6px 16px; border:1px solid #334155;
                            border-radius:8px; background:transparent; color:#94a3b8;
                            cursor:pointer; font-weight:500; font-size:12px;
                            transition:all 0.2s;" onmouseover="this.style.borderColor=#ef4444;this.style.color=#ef4444" onmouseout="this.style.borderColor=#334155;this.style.color=#94a3b8">
                        Close
                    </button>
                </div>
            </div>
        `;
        document.body.appendChild(panel);

        document.getElementById('m365-copy-token').onclick = () => copyToken();
        document.getElementById('m365-push-token').onclick = () => pushToken();
        document.getElementById('m365-push-cookies').onclick = () => pushCookies();
        document.getElementById('m365-one-click').onclick = () => oneClickSetup();
        document.getElementById('m365-close-panel').onclick = () => panel.remove();
    }

    // Show panel on demand via keyboard shortcut (Ctrl+Shift+M)
    pageWindow.addEventListener('keydown', (e) => {
        if (e.ctrlKey && e.shiftKey && e.key === 'M') {
            showPanel();
        }
    });
})();
