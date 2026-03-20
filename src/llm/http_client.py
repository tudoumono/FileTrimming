"""LLM バックエンド用の HTTP クライアント生成。"""

from __future__ import annotations

import atexit
from threading import Lock

import httpx

_http_clients: list[httpx.Client] = []
_http_clients_lock = Lock()


def _register_http_client(client: httpx.Client) -> httpx.Client:
    with _http_clients_lock:
        _http_clients.append(client)
    return client


def _close_registered_http_clients() -> None:
    with _http_clients_lock:
        clients = list(_http_clients)
        _http_clients.clear()

    for client in clients:
        try:
            client.close()
        except Exception:
            continue


atexit.register(_close_registered_http_clients)


def build_http_client(
    proxy_url: str = "",
    skip_ssl_verify: bool = False,
) -> httpx.Client | None:
    """必要なオプションがあるときだけ httpx.Client を生成する。"""
    if not proxy_url and not skip_ssl_verify:
        return None
    return _register_http_client(httpx.Client(
        proxy=proxy_url or None,
        verify=not skip_ssl_verify,
    ))
