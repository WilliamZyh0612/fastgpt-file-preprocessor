from __future__ import annotations

import json
import os
from dataclasses import dataclass, field
from pathlib import Path


@dataclass(frozen=True)
class Settings:
    unknown_value: str = "待人工确认"
    api_base_url: str = ""
    api_key_env: str = "FASTGPT_PREPROCESSOR_API_KEY"
    text_model: str = ""
    vision_model: str = ""
    embedding_model: str = ""
    ai_enabled: bool = False
    timeout_seconds: int = 45
    retries: int = 2
    concurrency: int = 2
    similarity_threshold: float = 0.72
    semantic_similarity_enabled: bool = False
    max_file_size_mb: int = 100
    min_pdf_text_chars: int = 80
    legacy_office_conversion_enabled: bool = False
    legacy_office_timeout_seconds: int = 60
    chunk_size_chars: int = 12000
    chunk_overlap_chars: int = 400
    embedding_batch_size: int = 32
    ai_schema_mode: str = "strict"
    input_token_price_per_million: float = 0.0
    output_token_price_per_million: float = 0.0
    model_aliases: dict[str, list[str]] = field(default_factory=dict)
    cnc_aliases: dict[str, list[str]] = field(default_factory=dict)

    @property
    def api_key(self) -> str:
        """Secrets intentionally stay in process memory and are never serialized."""
        return os.environ.get(self.api_key_env, "")


def load_settings(path: Path | None = None) -> Settings:
    config_path = path or Path.cwd() / "config.json"
    if not config_path.exists():
        raise FileNotFoundError(f"未找到配置文件：{config_path}")
    raw = json.loads(config_path.read_text(encoding="utf-8"))
    if not isinstance(raw, dict):
        raise ValueError("config.json 顶层必须是对象")
    allowed = {field.name for field in Settings.__dataclass_fields__.values()}
    data = {key: value for key, value in raw.items() if key in allowed}
    settings = Settings(**data)
    if settings.timeout_seconds < 1 or settings.concurrency < 1 or settings.chunk_size_chars < 500 or not 0 <= settings.chunk_overlap_chars < settings.chunk_size_chars or settings.input_token_price_per_million < 0 or settings.output_token_price_per_million < 0 or not 0 < settings.similarity_threshold <= 1 or settings.ai_schema_mode not in {"strict", "compat"}:
        raise ValueError("超时、并发或相似度阈值配置不合法")
    return settings
