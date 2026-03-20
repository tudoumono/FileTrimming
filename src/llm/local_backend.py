"""ローカル LLM バックエンド (Ollama)

Ollama の OpenAI 互換 API を使用する。
デフォルトで http://localhost:11434 に接続。API キー不要。
"""

from __future__ import annotations

from src.llm.http_client import build_http_client
from src.llm.openai_compatible_backend import OpenAICompatibleBackend


class LocalBackend(OpenAICompatibleBackend):
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

        super().__init__(
            client_kwargs=client_kwargs,
            model=model,
            log_message=(
                "Ollama バックエンド初期化: base_url=%s, model=%s, proxy=%s, "
                "skip_ssl_verify=%s"
            ),
            log_args=(
                base_url,
                model,
                proxy_url or "(none)",
                skip_ssl_verify,
            ),
        )
