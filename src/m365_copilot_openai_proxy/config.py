from __future__ import annotations

from pydantic import Field
from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
        extra="ignore",
    )

    access_token: str = Field(default="", alias="M365_ACCESS_TOKEN")
    time_zone: str = Field(default="Asia/Shanghai", alias="M365_TIME_ZONE")
    model_alias: str = Field(default="m365-copilot", alias="M365_MODEL_ALIAS")
    api_key: str = Field(default="", alias="API_KEY")
    admin_password: str = Field(default="", alias="ADMIN_PASSWORD")
    idle_timeout_minutes: int = Field(default=30, alias="IDLE_TIMEOUT_MINUTES")
    token_dir: str = Field(default="/home/app/token", alias="TOKEN_DIR")
