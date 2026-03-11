"""llm/openai_client.py - OpenAI API 実装"""

import logging
from typing import Optional
import openai
from .base import LLMClient, LLMResponse

logger = logging.getLogger(__name__)


class OpenAIClient(LLMClient):
    def __init__(self, api_key: str, model: str, base_url: str, default_temperature: float = 0.1):
        self._client = openai.OpenAI(api_key=api_key, base_url=base_url)
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

        logger.debug("OpenAI request: model=%s, max_tokens=%d", self._model, max_tokens)
        response = self._client.chat.completions.create(**kwargs)
        choice = response.choices[0]
        usage = response.usage
        return LLMResponse(
            content=choice.message.content or "",
            model=response.model,
            prompt_tokens=usage.prompt_tokens if usage else 0,
            completion_tokens=usage.completion_tokens if usage else 0,
            total_tokens=usage.total_tokens if usage else 0,
        )

    def provider_name(self) -> str:
        return f"openai ({self._model})"
