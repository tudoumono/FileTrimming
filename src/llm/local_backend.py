"""ローカル LLM バックエンド (Ollama)

Ollama の OpenAI 互換 API を使用する。
デフォルトで http://localhost:11434 に接続。API キー不要。
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


class LocalBackend(LLMBackend):
    """Ollama (ローカル LLM) バックエンド

    Ollama は OpenAI 互換 API を提供しているため、
    openai パッケージの base_url を差し替えて使う。
    """

    def __init__(
        self,
        base_url: str = "http://localhost:11434",
        model: str = "llama-3-elyza-8b",
        proxy_url: str = "",
        skip_ssl_verify: bool = False,
    ) -> None:
        try:
            from openai import OpenAI
        except ImportError as e:
            raise ImportError(
                "openai パッケージがインストールされていません。\n"
                "pip install openai でインストールしてください。"
            ) from e

        # Ollama の OpenAI 互換エンドポイント
        client_kwargs: dict[str, object] = {
            "base_url": f"{base_url}/v1",
            "api_key": "ollama",
        }
        http_client = build_http_client(
            proxy_url=proxy_url,
            skip_ssl_verify=skip_ssl_verify,
        )
        if http_client is not None:
            client_kwargs["http_client"] = http_client

        self._client = OpenAI(**client_kwargs)
        self._model = model
        logger.info(
            "Ollama バックエンド初期化: base_url=%s, model=%s, proxy=%s, skip_ssl_verify=%s",
            base_url,
            model,
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
