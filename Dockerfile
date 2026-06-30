FROM python:3.11-slim-bookworm

# Install Chromium (available on both amd64 and arm64 via Debian repos)
# NOTE: Chromium version is pinned by Debian repo snapshot; for reproducible builds
# consider pinning or using Chrome for Testing with explicit version.
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
        chromium \
        chromium-common \
        fonts-wqy-zenhei \
        gosu \
        curl \
    && rm -rf /var/lib/apt/lists/* \
    && echo "Chromium version: $(chromium --version || echo 'unknown')"

# Install uv package manager
RUN pip install --no-cache-dir uv

# Create non-root user and directories (merged into single RUN to reduce layers)
RUN groupadd -r app && useradd -r -g app -d /home/app -s /sbin/nologin app && \
    mkdir -p /chrome-profile /home/app/token /home/app /app && \
    chown -R app:app /chrome-profile /home/app /app

WORKDIR /app

# Copy dependency files first for Docker layer caching
COPY --chown=app:app pyproject.toml .
COPY --chown=app:app uv.lock .

# Install Python dependencies as app user (avoids chown -R producing extra layer)
USER app
RUN uv sync --frozen --no-dev
USER root

# Copy project source and entrypoint
COPY --chown=app:app src/ src/
COPY --chown=app:app entrypoint.sh /entrypoint.sh
RUN chmod +x /entrypoint.sh

# Persist Chrome user data (login state)
VOLUME /chrome-profile

# NOTE: /home/app/token is managed as tmpfs via docker-compose (not VOLUME here).
# Declaring VOLUME would conflict with tmpfs, causing Docker to mount a named volume
# owned by root:root that shadows the tmpfs mount and breaks app-user write access.

# Environment variables (do NOT set M365_ACCESS_TOKEN or ADMIN_PASSWORD here — they may leak into image layers)
ENV M365_TIME_ZONE="Asia/Shanghai"
ENV M365_MODEL_ALIAS="m365-copilot"
ENV CHROME_CDP_PORT=9222
ENV AUTO_REFRESH="true"
ENV REFRESH_BEFORE_SECONDS=300
ENV IDLE_TIMEOUT_MINUTES=30
ENV TOKEN_DIR="/home/app/token"

EXPOSE 8000

HEALTHCHECK --interval=30s --timeout=5s --start-period=15s --retries=3 \
    CMD curl -sf http://localhost:8000/healthz || exit 1

# Start as root to fix volume permissions, then drop to app user via gosu in entrypoint
ENTRYPOINT ["/entrypoint.sh"]
