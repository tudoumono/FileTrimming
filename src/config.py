"""パイプライン設定

パス設定、LLM バックエンド選択、各種閾値を管理する。
"""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path


@dataclass
class PipelineConfig:
    """パイプライン全体の設定"""

    # --- パス設定 ---
    input_dir: Path = field(default_factory=lambda: Path("input"))
    intermediate_dir: Path = field(default_factory=lambda: Path("intermediate"))
    output_dir: Path = field(default_factory=lambda: Path("output"))

    @property
    def normalized_dir(self) -> Path:
        return self.intermediate_dir / "01_normalized"

    @property
    def extracted_dir(self) -> Path:
        return self.intermediate_dir / "02_extracted"

    @property
    def transformed_dir(self) -> Path:
        return self.intermediate_dir / "03_transformed"

    # --- LLM 設定 ---
    llm_backend: str = "noop"  # "noop", "openai", "local"
    openai_api_key: str = ""
    openai_model: str = "gpt-4o-mini"
    ollama_base_url: str = "http://localhost:11434"
    ollama_model: str = "llama-3-elyza-8b"

    # --- Step1: 正規化 ---
    # COM 変換対象の拡張子
    com_normalize_exts: set[str] = field(
        default_factory=lambda: {".doc", ".rtf"}
    )
    # そのまま通す拡張子
    passthrough_exts: set[str] = field(
        default_factory=lambda: {".docx"}
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

    # --- Step3: 変換 ---
    # Dify のファイルサイズ上限 (bytes)
    max_file_size_bytes: int = 15 * 1024 * 1024  # 15MB

    # --- Word 系対象拡張子（全ステップ共通） ---
    word_exts: set[str] = field(
        default_factory=lambda: {".doc", ".docx", ".rtf"}
    )

    def load_env(self, env_path: Path | None = None) -> None:
        """`.env` ファイルから環境変数を読み込み、設定に反映する。

        `.env` の形式: KEY=VALUE (1行1エントリ、# コメント、空行無視)
        python-dotenv に依存せず、自前でパースする。

        対応する環境変数:
          OPENAI_API_KEY    → openai_api_key
          OPENAI_MODEL      → openai_model
          OLLAMA_BASE_URL   → ollama_base_url
          OLLAMA_MODEL      → ollama_model
          LLM_BACKEND       → llm_backend
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
            "OPENAI_MODEL": "openai_model",
            "OLLAMA_BASE_URL": "ollama_base_url",
            "OLLAMA_MODEL": "ollama_model",
            "LLM_BACKEND": "llm_backend",
        }
        for env_key, attr in env_map.items():
            val = os.environ.get(env_key)
            if val:
                setattr(self, attr, val)
