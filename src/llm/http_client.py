"""LLM バックエンド用の HTTP クライアント生成。"""

from __future__ import annotations

import httpx


def build_http_client(
    proxy_url: str = "",
    skip_ssl_verify: bool = False,
) -> httpx.Client | None:
    """必要なオプションがあるときだけ httpx.Client を生成する。"""
    if not proxy_url and not skip_ssl_verify:
        return None
    return httpx.Client(
        proxy=proxy_url or None,
        verify=not skip_ssl_verify,
    )
