"""パイプライン設定

パス設定、LLM バックエンド選択、各種閾値を管理する。
"""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path


def _parse_bool_env(value: str) -> bool:
    return value.strip().lower() in {"1", "true", "yes", "on"}


def _parse_int_env(value: str, minimum: int = 1) -> int:
    parsed = int(value.strip())
    if parsed < minimum:
        raise ValueError(f"value must be >= {minimum}: {parsed}")
    return parsed


@dataclass
class PipelineConfig:
    """パイプライン全体の設定"""

    # --- パス設定 ---
    input_dir: Path = field(default_factory=lambda: Path("input"))
    intermediate_base: Path = field(default_factory=lambda: Path("intermediate"))
    output_base: Path = field(default_factory=lambda: Path("output"))
    run_id: str = ""  # タイムスタンプ (例: "20260313_175250")。空なら自動生成
    normalize_workers: int = 1
    extract_workers: int = 1
    transform_workers: int = 1

    def _ensure_run_id(self) -> str:
        if not self.run_id:
            self.run_id = datetime.now().strftime("%Y%m%d_%H%M%S")
        return self.run_id

    @property
    def intermediate_dir(self) -> Path:
        return self.intermediate_base / self._ensure_run_id()

    @property
    def output_dir(self) -> Path:
        return self.output_base / self._ensure_run_id()

    @property
    def normalized_dir(self) -> Path:
        return self.intermediate_dir / "01_normalized"

    @property
    def extracted_dir(self) -> Path:
        return self.intermediate_dir / "02_extracted"

    @property
    def transformed_dir(self) -> Path:
        return self.intermediate_dir / "03_transformed"

    @property
    def review_dir(self) -> Path:
        return self.intermediate_dir / "04_review"

    # --- LLM 設定 ---
    llm_backend: str = "noop"  # "noop", "openai", "local"
    openai_api_key: str = ""
    openai_base_url: str = ""
    openai_model: str = "gpt-4o-mini"
    ollama_base_url: str = "http://localhost:11434"
    ollama_model: str = "llama-3-elyza-8b"
    llm_proxy_url: str = ""
    llm_skip_ssl_verify: bool = False
    llm_observation_only: bool = False

    # --- Step1: 正規化 ---
    # COM 変換対象の拡張子 (.doc/.rtf → Word COM → .docx, .xls → Excel COM → .xlsx)
    com_normalize_exts: set[str] = field(
        default_factory=lambda: {".doc", ".rtf", ".xls"}
    )
    # そのまま通す拡張子（正規化不要でコピーのみ）
    passthrough_exts: set[str] = field(
        default_factory=lambda: {".docx", ".xlsx", ".xlsm"}
    )

    # --- Step2: 構造抽出 ---
    # 疑似見出し検出: 本文推定フォントサイズ (pt)
    body_font_size_pt: float = 10.5
    # フォントサイズがこれ以上大きければ見出し候補
    heading_font_size_min_pt: float = 11.0
    # 疑似見出しの最大文字数
    heading_max_chars: int = 80
    # 変更履歴テーブル検出キーワード
    change_history_keywords: set[str] = field(
        default_factory=lambda: {"ページ", "種別", "年月", "記事"}
    )
    change_history_min_match: int = 3

    # --- Step2: Excel 構造抽出 ---
    # 大きなシートの警告閾値
    excel_large_sheet_rows: int = 500
    excel_large_sheet_cols: int = 30

    # --- Step3: 変換 ---
    # Dify のファイルサイズ上限 (bytes)
    max_file_size_bytes: int = 15 * 1024 * 1024  # 15MB

    # --- Word 系対象拡張子（全ステップ共通） ---
    word_exts: set[str] = field(
        default_factory=lambda: {".doc", ".docx", ".rtf"}
    )

    # --- Excel 系対象拡張子（全ステップ共通） ---
    excel_exts: set[str] = field(
        default_factory=lambda: {".xls", ".xlsx", ".xlsm"}
    )

    def load_env(self, env_path: Path | None = None) -> None:
        """`.env` ファイルから環境変数を読み込み、設定に反映する。

        `.env` の形式: KEY=VALUE (1行1エントリ、# コメント、空行無視)
        python-dotenv に依存せず、自前でパースする。

        対応する環境変数:
          OPENAI_API_KEY    → openai_api_key
          OPENAI_BASE_URL   → openai_base_url
          OPENAI_MODEL      → openai_model
          OLLAMA_BASE_URL   → ollama_base_url
          OLLAMA_MODEL      → ollama_model
          LLM_PROXY_URL     → llm_proxy_url
          LLM_SKIP_SSL_VERIFY → llm_skip_ssl_verify
          LLM_OBSERVATION_ONLY → llm_observation_only
          LLM_BACKEND       → llm_backend
          NORMALIZE_WORKERS → normalize_workers
          EXTRACT_WORKERS   → extract_workers
          TRANSFORM_WORKERS → transform_workers
        """
        if env_path is None:
            env_path = Path(".env")

        # .env ファイルがあれば読み込んで環境変数にセット
        if env_path.exists():
            for line in env_path.read_text(encoding="utf-8").splitlines():
                line = line.strip()
                if not line or line.startswith("#"):
                    continue
                if "=" not in line:
                    continue
                key, _, value = line.partition("=")
                key = key.strip()
                value = value.strip().strip("'\"")  # クォート除去
                os.environ.setdefault(key, value)

        # 環境変数から設定に反映（環境変数が優先）
        env_map = {
            "OPENAI_API_KEY": "openai_api_key",
            "OPENAI_BASE_URL": "openai_base_url",
            "OPENAI_MODEL": "openai_model",
            "OLLAMA_BASE_URL": "ollama_base_url",
            "OLLAMA_MODEL": "ollama_model",
            "LLM_PROXY_URL": "llm_proxy_url",
            "LLM_BACKEND": "llm_backend",
        }
        for env_key, attr in env_map.items():
            val = os.environ.get(env_key)
            if val:
                setattr(self, attr, val)

        bool_env_map = {
            "LLM_SKIP_SSL_VERIFY": "llm_skip_ssl_verify",
            "LLM_OBSERVATION_ONLY": "llm_observation_only",
        }
        for env_key, attr in bool_env_map.items():
            val = os.environ.get(env_key)
            if val:
                setattr(self, attr, _parse_bool_env(val))

        int_env_map = {
            "NORMALIZE_WORKERS": "normalize_workers",
            "EXTRACT_WORKERS": "extract_workers",
            "TRANSFORM_WORKERS": "transform_workers",
        }
        for env_key, attr in int_env_map.items():
            val = os.environ.get(env_key)
            if val:
                setattr(self, attr, _parse_int_env(val))

    def validate(self) -> None:
        """設定値の整合性を検証する。"""
        for attr in (
            "normalize_workers",
            "extract_workers",
            "transform_workers",
        ):
            value = getattr(self, attr)
            if value < 1:
                raise ValueError(f"{attr} must be >= 1: {value}")
