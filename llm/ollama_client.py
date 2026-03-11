"""
llm/ollama_client.py - Ollama（ローカルLLM）実装

Ollama は OpenAI 互換の /v1/chat/completions エンドポイントを提供する。
openai ライブラリの base_url を差し替えて動作させる。
"""

import logging
from typing import Optional
import openai
from .base import LLMClient, LLMResponse

logger = logging.getLogger(__name__)


class OllamaClient(LLMClient):
    def __init__(self, base_url: str, model: str, default_temperature: float = 0.1):
        api_base = base_url.rstrip("/")
        if not api_base.endswith("/v1"):
            api_base += "/v1"
        self._client = openai.OpenAI(api_key="ollama", base_url=api_base)
        self._model = model
        self._default_temperature = default_temperature

    def chat(self, system_prompt, user_message, max_tokens=16000,
             temperature=None, response_format=None) -> LLMResponse:
        temp = temperature if temperature is not None else self._default_temperature
        kwargs = dict(
            model=self._model,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_message},
            ],
            max_tokens=max_tokens,
            temperature=temp,
        )
        if response_format:
            kwargs["response_format"] = response_format

        logger.debug("Ollama request: model=%s, max_tokens=%d", self._model, max_tokens)
        try:
            response = self._client.chat.completions.create(**kwargs)
        except openai.APIConnectionError:
            raise ConnectionError(
                f"Ollama に接続できません。起動を確認してください: {self._client.base_url}"
            )

        choice = response.choices[0]
        usage = response.usage
        return LLMResponse(
            content=choice.message.content or "",
            model=response.model or self._model,
            prompt_tokens=usage.prompt_tokens if usage else 0,
            completion_tokens=usage.completion_tokens if usage else 0,
            total_tokens=usage.total_tokens if usage else 0,
        )

    def provider_name(self) -> str:
        return f"ollama ({self._model})"
