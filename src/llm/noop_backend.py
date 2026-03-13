"""LLM なしフォールバック

LLM を使わずにパイプラインを完走させるためのダミーバックエンド。
generate() は常に空文字列を返す。
"""

from __future__ import annotations

from src.llm.base import LLMBackend


class NoopBackend(LLMBackend):
    """LLM を使わないダミーバックエンド"""

    def generate(self, prompt: str, system: str = "") -> str:
        return ""
