"""OpenAI API バックエンド

前処理はオンライン PC で実施するため OpenAI API を利用可能。
API キーは .env ファイルまたは環境変数 OPENAI_API_KEY で設定する。
"""

from __future__ import annotations

from logging import getLogger

from src.llm.base import LLMBackend

logger = getLogger(__name__)


class OpenAIBackend(LLMBackend):
    """OpenAI API を使った LLM バックエンド"""

    def __init__(self, api_key: str, model: str = "gpt-4o-mini") -> None:
        if not api_key:
            raise ValueError(
                "OpenAI API キーが設定されていません。\n"
                ".env ファイルに OPENAI_API_KEY=sk-... を記載するか、\n"
                "環境変数 OPENAI_API_KEY を設定してください。"
            )

        try:
            from openai import OpenAI
        except ImportError as e:
            raise ImportError(
                "openai パッケージがインストールされていません。\n"
                "pip install openai でインストールしてください。"
            ) from e

        self._client = OpenAI(api_key=api_key)
        self._model = model
        logger.info("OpenAI バックエンド初期化: model=%s", model)

    def generate(self, prompt: str, system: str = "") -> str:
        messages = []
        if system:
            messages.append({"role": "system", "content": system})
        messages.append({"role": "user", "content": prompt})

        response = self._client.chat.completions.create(
            model=self._model,
            messages=messages,
        )
        return response.choices[0].message.content or ""
