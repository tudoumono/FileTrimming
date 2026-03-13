"""ローカル LLM バックエンド (Ollama)

Ollama の OpenAI 互換 API を使用する。
デフォルトで http://localhost:11434 に接続。API キー不要。
"""

from __future__ import annotations

from logging import getLogger

from src.llm.base import LLMBackend

logger = getLogger(__name__)


class LocalBackend(LLMBackend):
    """Ollama (ローカル LLM) バックエンド

    Ollama は OpenAI 互換 API を提供しているため、
    openai パッケージの base_url を差し替えて使う。
    """

    def __init__(
        self,
        base_url: str = "http://localhost:11434",
        model: str = "llama-3-elyza-8b",
    ) -> None:
        try:
            from openai import OpenAI
        except ImportError as e:
            raise ImportError(
                "openai パッケージがインストールされていません。\n"
                "pip install openai でインストールしてください。"
            ) from e

        # Ollama の OpenAI 互換エンドポイント
        self._client = OpenAI(
            base_url=f"{base_url}/v1",
            api_key="ollama",  # Ollama は API キー不要だがダミー値が必要
        )
        self._model = model
        logger.info("Ollama バックエンド初期化: base_url=%s, model=%s", base_url, model)

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
