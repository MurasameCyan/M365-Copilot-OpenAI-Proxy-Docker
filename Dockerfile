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

WORKDIR /app

# Copy dependency files first for Docker layer caching
COPY pyproject.toml .
COPY uv.lock .

# Install Python dependencies
RUN uv sync --frozen --no-dev

# Copy project source
COPY src/ src/

# Copy entrypoint script
COPY entrypoint.sh /entrypoint.sh
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

EXPOSE 8000

ENTRYPOINT ["/entrypoint.sh"]
