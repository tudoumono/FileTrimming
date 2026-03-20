"""OpenAI 互換 API を使う LLM バックエンド共通実装。"""

from __future__ import annotations

from logging import getLogger

import httpx

from src.llm.base import LLMBackend, ReconstructionUnit, TableInterpretationResult
from src.llm.table_interpretation import (
    TABLE_INTERPRETATION_PROMPT_VERSION,
    TABLE_INTERPRETATION_SYSTEM_PROMPT,
    build_table_interpretation_prompt,
    parse_table_interpretation_response,
)

logger = getLogger(__name__)


class OpenAICompatibleBackend(LLMBackend):
    """OpenAI 互換 Chat Completions API 向け共通実装。"""

    def __init__(
        self,
        *,
        client_kwargs: dict[str, object],
        model: str,
        log_message: str,
        log_args: tuple[object, ...],
    ) -> None:
        try:
            from openai import OpenAI
        except ImportError as e:
            raise ImportError(
                "openai パッケージがインストールされていません。\n"
                "pip install openai でインストールしてください。"
            ) from e

        raw_http_client = client_kwargs.get("http_client")
        self._http_client = (
            raw_http_client if isinstance(raw_http_client, httpx.Client) else None
        )
        self._client = OpenAI(**client_kwargs)
        self._model = model
        logger.info(log_message, *log_args)

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

    def supports_table_interpretation(self) -> bool:
        return True

    def interpret_table(
        self, unit: ReconstructionUnit, system: str = "",
    ) -> TableInterpretationResult:
        prompt = build_table_interpretation_prompt(unit)
        response_text = self.generate(
            prompt,
            system=system or TABLE_INTERPRETATION_SYSTEM_PROMPT,
        )
        return parse_table_interpretation_response(response_text, unit.unit_id)

    def model_name(self) -> str:
        return self._model

    def prompt_version(self) -> str:
        return TABLE_INTERPRETATION_PROMPT_VERSION

    def close(self) -> None:
        if self._http_client is not None:
            self._http_client.close()
            self._http_client = None
