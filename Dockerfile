FROM python:3.11-slim

# Install Chromium (available on both amd64 and arm64 via Debian repos)
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
        chromium \
        chromium-common \
        fonts-noto-cjk \
    && rm -rf /var/lib/apt/lists/*

# Install uv package manager
RUN pip install --no-cache-dir uv

# Create non-root user and directories
RUN groupadd -r app && useradd -r -g app -d /home/app -s /sbin/nologin app && \
    mkdir -p /chrome-profile /home/app && chown -R app:app /chrome-profile /home/app

WORKDIR /app

# Copy dependency files first for Docker layer caching
COPY --chown=app:app pyproject.toml .
COPY --chown=app:app uv.lock .

# Install Python dependencies
RUN uv sync --frozen --no-dev && chown -R app:app /app

# Copy project source and entrypoint
COPY --chown=app:app src/ src/
COPY --chown=app:app entrypoint.sh /entrypoint.sh
RUN chmod +x /entrypoint.sh

# Persist Chrome user data (login state)
VOLUME /chrome-profile

# Environment variables
ENV M365_ACCESS_TOKEN=""
ENV M365_TIME_ZONE="Asia/Shanghai"
ENV M365_MODEL_ALIAS="m365-copilot"
ENV CHROME_CDP_PORT=9222
ENV AUTO_REFRESH="true"
ENV REFRESH_BEFORE_SECONDS=300
ENV IDLE_TIMEOUT_MINUTES=30

EXPOSE 8000

# Start as root to fix volume permissions, then drop to app user in entrypoint
ENTRYPOINT ["/entrypoint.sh"]
