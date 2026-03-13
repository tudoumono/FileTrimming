"""LLM バックエンド共通インターフェース

パイプライン内の LLM 呼び出し部分は全てこのインターフェースを通す。
バックエンド（OpenAI / ローカル LLM / なし）を設定で切り替え可能にする。
"""

from __future__ import annotations

from abc import ABC, abstractmethod


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
