#!/bin/bash
set -e

CDP_PORT="${CHROME_CDP_PORT:-9222}"
CHROME_PROFILE="/chrome-profile"
AUTO_REFRESH="${AUTO_REFRESH:-true}"

# --- Root-only section: fix permissions and clean stale locks ---
if [ "$(id -u)" = "0" ]; then
    # Fix volume ownership (Docker volumes default to root:root on first mount)
    chown -R app:app "$CHROME_PROFILE" 2>/dev/null || true
    mkdir -p "$CHROME_PROFILE" 2>/dev/null || true
    rm -f "$CHROME_PROFILE/SingletonLock" "$CHROME_PROFILE/SingletonCookie" "$CHROME_PROFILE/SingletonSocket" 2>/dev/null || true

    # Prepare token storage on tmpfs-backed volume
    mkdir -p /home/app/token 2>/dev/null || true
    chown -R app:app /home/app/token 2>/dev/null || true

    # Re-exec as app user via gosu (does NOT preserve environment wholesale — safer)
    export HOME=/home/app
    exec gosu app "$0" "$@"
fi

# --- Below runs as app user ---

# Detect Chromium binary (name varies by distro)
if command -v chromium &> /dev/null; then
    CHROME_BIN="chromium"
elif command -v chromium-browser &> /dev/null; then
    CHROME_BIN="chromium-browser"
elif command -v google-chrome-stable &> /dev/null; then
    CHROME_BIN="google-chrome-stable"
elif command -v google-chrome &> /dev/null; then
    CHROME_BIN="google-chrome"
else
    echo "WARNING: No Chrome/Chromium binary found. Starting server without auto-refresh."
    CHROME_BIN=""
fi

# Start Chrome headless + CDP (only if binary found and AUTO_REFRESH is true)
if [ -n "$CHROME_BIN" ] && [ "$AUTO_REFRESH" = "true" ]; then
    echo "Starting $CHROME_BIN headless on CDP port $CDP_PORT ..."
    "$CHROME_BIN" \
        --headless=new \
        --remote-debugging-address=127.0.0.1 \
        --remote-debugging-port="$CDP_PORT" \
        --user-data-dir="$CHROME_PROFILE" \
        --no-first-run \
        --disable-gpu \
        --disable-dev-shm-usage \
        --disable-software-rasterizer \
        --disable-background-networking \
        --disable-sync \
        --no-default-browser-check \
        --disable-features=InfiniteRestore,MediaRouter,DialMediaRouteProvider,TranslateUI \
        --disable-breakpad \
        --no-experiments \
        --crash-dumps-dir=/tmp \
        "https://m365.cloud.microsoft/chat" 2>&1 | grep -v -E '(dbus|system_bus_socket|DEPRECATED_ENDPOINT|NameHasOwner|Properties\.GetAll|crashpad)' &

    CHROME_PID=$!
    echo "Chromium started with PID $CHROME_PID"

    # Wait for Chrome CDP to be ready
    echo "Waiting for Chromium CDP on port $CDP_PORT ..."
    for i in $(seq 1 30); do
        if python3 -c "import urllib.request; urllib.request.urlopen('http://localhost:$CDP_PORT/json', timeout=2)" 2>/dev/null; then
            echo "Chromium CDP ready."
            break
        fi
        if [ $i -eq 30 ]; then
            echo "WARNING: Chromium CDP did not become ready in 30s. Continuing without CDP."
        fi
        sleep 1
    done
else
    echo "Chromium headless not started. Token auto-refresh disabled."
fi

# Build serve command arguments
# Use --no-launch-edge to prevent Python from launching another Chromium instance
SERVE_ARGS="--host 0.0.0.0 --port 8000 --no-launch-edge"

if [ -n "$CHROME_BIN" ] && [ "$AUTO_REFRESH" = "true" ]; then
    SERVE_ARGS="$SERVE_ARGS --cdp-port $CDP_PORT --refresh-before-seconds ${REFRESH_BEFORE_SECONDS:-300}"
else
    SERVE_ARGS="$SERVE_ARGS --no-auto-refresh --no-capture-on-start"
fi

# If no token is set, log a hint
if [ -z "$M365_ACCESS_TOKEN" ]; then
    echo "WARNING: M365_ACCESS_TOKEN is not set."
    echo "Please get a token via the Tampermonkey script (get_token.js) and set it in .env"
fi

echo "Starting copilot-openai-proxy serve $SERVE_ARGS"
exec uv run copilot-openai-proxy serve $SERVE_ARGS
