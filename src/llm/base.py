"""LLM バックエンド共通インターフェース

パイプライン内の LLM 呼び出し部分は全てこのインターフェースを通す。
バックエンド（OpenAI / ローカル LLM / なし）を設定で切り替え可能にする。
"""

from __future__ import annotations

from dataclasses import asdict, dataclass, field
from abc import ABC, abstractmethod
from typing import Any


@dataclass(frozen=True)
class ReconstructionUnit:
    """テーブル単位の再構成入力。"""

    schema_version: str
    unit_id: str
    source_path: str
    source_ext: str
    sheet_name: str
    table_caption: str
    rows: list[list[dict[str, Any]]]
    context: dict[str, Any] = field(default_factory=dict)
    hints: dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(frozen=True)
class TableInterpretationResult:
    """テーブル解釈結果の共通契約。"""

    schema_version: str
    unit_id: str
    table_type: str
    render_strategy: str
    header_rows: list[int] = field(default_factory=list)
    data_start_row: int = 0
    column_labels: list[str] = field(default_factory=list)
    active_columns: list[int] = field(default_factory=list)
    render_plan: dict[str, Any] = field(default_factory=dict)
    notes: list[str] = field(default_factory=list)
    self_assessment: dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


class LLMBackend(ABC):
    """LLM バックエンドの抽象基底クラス"""

    @abstractmethod
    def generate(self, prompt: str, system: str = "") -> str:
        """テキスト生成

        Args:
            prompt: ユーザープロンプト
            system: システムプロンプト（任意）

        Returns:
            生成されたテキスト
        """
        ...

    def supports_table_interpretation(self) -> bool:
        """構造化テーブル解釈をサポートしているか。"""
        return False

    def backend_name(self) -> str:
        """レビュー成果物向けのバックエンド識別子を返す。"""
        name = self.__class__.__name__
        if name.endswith("Backend"):
            name = name[:-7]
        return name.lower()

    def model_name(self) -> str:
        """使用モデル名を返す。未使用時は空文字列。"""
        return ""

    def prompt_version(self) -> str:
        """使用プロンプト版を返す。未使用時は空文字列。"""
        return ""

    def close(self) -> None:
        """必要なら保持中のリソースを解放する。"""
        return None

    def interpret_table(
        self, unit: ReconstructionUnit, system: str = "",
    ) -> TableInterpretationResult:
        """構造化されたテーブル解釈結果を返す。

        初期段階ではバックエンドごとの実装待ちとし、
        `supports_table_interpretation()` が False の場合は
        呼び出し側が `LLMなし` の既存解釈へフォールバックする。
        """
        raise NotImplementedError(
            f"{self.__class__.__name__} does not support structured table interpretation yet."
        )
