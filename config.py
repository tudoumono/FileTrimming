"""
config.py - .env から設定を読み込み、型付きで提供する
"""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path

try:
    from dotenv import load_dotenv
except ImportError:
    def load_dotenv(*_args, **_kwargs) -> bool:
        return False


_PROJECT_ROOT = Path(__file__).parent


def _load_env_file(env_file: str | Path | None = None) -> None:
    env_path = Path(env_file) if env_file else _PROJECT_ROOT / ".env"
    if env_path.exists():
        load_dotenv(env_path, override=True)


_load_env_file()


def _env(key: str, default: str = "") -> str:
    return os.getenv(key, default)


def _env_int(key: str, default: int = 0) -> int:
    return int(os.getenv(key, str(default)))


@dataclass(frozen=True)
class PathConfig:
    input_dir: Path = field(default_factory=lambda: Path(_env("INPUT_DIR", "./input")))
    work_dir: Path = field(default_factory=lambda: Path(_env("WORK_DIR", "./work")))
    output_dir: Path = field(default_factory=lambda: Path(_env("OUTPUT_DIR", "./output")))

    def step_dir(self, step_number: int) -> Path:
        """各ステップの中間出力フォルダ: work/step{N}/"""
        d = self.work_dir / f"step{step_number}"
        d.mkdir(parents=True, exist_ok=True)
        return d

    def ensure_dirs(self):
        for d in (self.work_dir, self.output_dir):
            d.mkdir(parents=True, exist_ok=True)


@dataclass(frozen=True)
class LLMConfig:
    provider: str = field(default_factory=lambda: _env("LLM_PROVIDER", "openai"))
    # OpenAI
    openai_api_key: str = field(default_factory=lambda: _env("OPENAI_API_KEY"))
    openai_model: str = field(default_factory=lambda: _env("OPENAI_MODEL", "gpt-4o-mini"))
    openai_base_url: str = field(default_factory=lambda: _env("OPENAI_BASE_URL", "https://api.openai.com/v1"))
    # Ollama
    ollama_base_url: str = field(default_factory=lambda: _env("OLLAMA_BASE_URL", "http://localhost:11434"))
    ollama_model: str = field(default_factory=lambda: _env("OLLAMA_MODEL", "llama3"))
    # 共通
    temperature: float = field(default_factory=lambda: float(_env("LLM_TEMPERATURE", "0.1")))


@dataclass(frozen=True)
class ProcessingConfig:
    token_limit: int = field(default_factory=lambda: _env_int("TOKEN_LIMIT", 80000))
    quality_threshold_high: int = field(default_factory=lambda: _env_int("QUALITY_THRESHOLD_HIGH", 70))
    quality_threshold_medium: int = field(default_factory=lambda: _env_int("QUALITY_THRESHOLD_MEDIUM", 40))


@dataclass(frozen=True)
class ExecutionConfig:
    start_step: int = field(default_factory=lambda: _env_int("START_STEP", 1))
    file_conflict_mode: str = field(default_factory=lambda: _env("FILE_CONFLICT_MODE", "skip"))
    log_level: str = field(default_factory=lambda: _env("LOG_LEVEL", "INFO"))
    log_file: Path = field(default_factory=lambda: Path(_env("LOG_FILE", "./work/logs/pipeline.log")))


@dataclass(frozen=True)
class EncodingConfig:
    default: str = field(default_factory=lambda: _env("DEFAULT_ENCODING", "shift_jis"))
    fallbacks: list[str] = field(default_factory=lambda: _env("FALLBACK_ENCODINGS", "utf-8,shift_jis,cp932,euc-jp").split(","))


@dataclass(frozen=True)
class AppConfig:
    paths: PathConfig = field(default_factory=PathConfig)
    llm: LLMConfig = field(default_factory=LLMConfig)
    processing: ProcessingConfig = field(default_factory=ProcessingConfig)
    execution: ExecutionConfig = field(default_factory=ExecutionConfig)
    encoding: EncodingConfig = field(default_factory=EncodingConfig)


def load_config(env_file: str | Path | None = None) -> AppConfig:
    """設定を読み込んで AppConfig を返す"""
    if env_file is not None:
        _load_env_file(env_file)
    return AppConfig()
