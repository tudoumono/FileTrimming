"""OpenAI API バックエンド

前処理はオンライン PC で実施するため OpenAI API を利用可能。
API キーは .env ファイルまたは環境変数 OPENAI_API_KEY で設定する。
"""

from __future__ import annotations

from logging import getLogger

from src.llm.base import LLMBackend, ReconstructionUnit, TableInterpretationResult
from src.llm.http_client import build_http_client
from src.llm.table_interpretation import (
    TABLE_INTERPRETATION_SYSTEM_PROMPT,
    build_table_interpretation_prompt,
    parse_table_interpretation_response,
)

logger = getLogger(__name__)


class OpenAIBackend(LLMBackend):
    """OpenAI API を使った LLM バックエンド"""

    def __init__(
        self,
        api_key: str,
        model: str = "gpt-4o-mini",
        base_url: str = "",
        proxy_url: str = "",
        skip_ssl_verify: bool = False,
    ) -> None:
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

        client_kwargs: dict[str, object] = {"api_key": api_key}
        if base_url:
            client_kwargs["base_url"] = base_url
        http_client = build_http_client(
            proxy_url=proxy_url,
            skip_ssl_verify=skip_ssl_verify,
        )
        if http_client is not None:
            client_kwargs["http_client"] = http_client

        self._client = OpenAI(**client_kwargs)
        self._model = model
        logger.info(
            "OpenAI バックエンド初期化: model=%s, base_url=%s, proxy=%s, skip_ssl_verify=%s",
            model,
            base_url or "(default)",
            proxy_url or "(none)",
            skip_ssl_verify,
        )

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
