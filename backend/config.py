from __future__ import annotations

from pathlib import Path

from pydantic import Field
from pydantic_settings import BaseSettings, SettingsConfigDict

ROOT = Path(__file__).resolve().parent.parent


class Settings(BaseSettings):
    model_config = SettingsConfigDict(
        env_file=ROOT / ".env",
        env_file_encoding="utf-8",
        extra="ignore",
        populate_by_name=True,
        env_ignore_empty=True,
    )

    header_name: str = Field(
        default="X-AssignLetters-Id",
        validation_alias="ASSIGNLETTERS_HEADER_NAME",
    )
    missing_header_message: str = Field(
        default=(
            "This message does not include the required assignment header, "
            "so AssignLetters cannot open the form."
        ),
        validation_alias="ASSIGNLETTERS_MISSING_HEADER_MESSAGE",
    )
    public_base_url: str = Field(
        default="https://localhost:8000",
        validation_alias="ASSIGNLETTERS_PUBLIC_BASE_URL",
    )
    api_url: str = Field(
        default="",
        validation_alias="ASSIGNLETTERS_API_URL",
    )
    log_path: Path = Field(
        default=ROOT / "data" / "logs" / "assignletters.log",
        validation_alias="ASSIGNLETTERS_LOG_PATH",
    )
    staff_path: Path = Field(
        default=ROOT / "data" / "staff.json",
        validation_alias="ASSIGNLETTERS_STAFF_PATH",
    )
    cors_allow_origins: str = Field(
        default="*",
        validation_alias="ASSIGNLETTERS_CORS_ALLOW_ORIGINS",
    )
    host: str = Field(default="0.0.0.0", validation_alias="HOST")
    port: int = Field(default=8000, validation_alias="PORT")

    @property
    def resolved_api_url(self) -> str:
        return (self.api_url or self.public_base_url).rstrip("/")

    @property
    def cors_origins(self) -> list[str]:
        raw = self.cors_allow_origins.strip()
        if raw == "*":
            return ["*"]
        return [part.strip() for part in raw.split(",") if part.strip()]


def get_settings() -> Settings:
    return Settings()
