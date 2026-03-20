"""OpenAI API バックエンド

前処理はオンライン PC で実施するため OpenAI API を利用可能。
API キーは .env ファイルまたは環境変数 OPENAI_API_KEY で設定する。
"""

from __future__ import annotations

from src.llm.http_client import build_http_client
from src.llm.openai_compatible_backend import OpenAICompatibleBackend


class OpenAIBackend(OpenAICompatibleBackend):
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

        client_kwargs: dict[str, object] = {"api_key": api_key}
        if base_url:
            client_kwargs["base_url"] = base_url
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
                "OpenAI バックエンド初期化: model=%s, base_url=%s, proxy=%s, "
                "skip_ssl_verify=%s"
            ),
            log_args=(
                model,
                base_url or "(default)",
                proxy_url or "(none)",
                skip_ssl_verify,
            ),
        )
