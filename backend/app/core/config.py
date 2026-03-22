from pathlib import Path
from typing import List

from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    openai_api_key: str = ""
    openai_model: str = "gpt-4o-mini"
    openai_base_url: str = "https://api.openai.com/v1"
    openai_api_key: str | None = None
    openai_image_model: str = "gpt-image-1"
    openai_image_size: str = "1024x1024"
    app_cors_origins: str = "http://127.0.0.1:5173,http://localhost:5173"
    chroma_dir: str = "./data/chroma"
    upload_dir: str = "./data/uploads"
    generated_dir: str = "./data/generated"

    model_config = SettingsConfigDict(env_file=".env", env_file_encoding="utf-8", extra="ignore")

    @property
    def cors_origins(self) -> List[str]:
        return [item.strip() for item in self.app_cors_origins.split(",") if item.strip()]

    def ensure_dirs(self) -> None:
        for path in [self.chroma_dir, self.upload_dir, self.generated_dir]:
            Path(path).mkdir(parents=True, exist_ok=True)


settings = Settings()
settings.ensure_dirs()
