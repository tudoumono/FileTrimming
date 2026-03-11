"""utils/token_utils.py - トークン推定ユーティリティ"""

from __future__ import annotations

try:
    import tiktoken
except ImportError:
    tiktoken = None

_encoder: tiktoken.Encoding | None = None


def _get_encoder() -> tiktoken.Encoding:
    global _encoder
    if tiktoken is None:
        raise RuntimeError(
            "tiktoken が見つかりません。`pip install -r requirements.txt` を実行してください。"
        )
    if _encoder is None:
        _encoder = tiktoken.encoding_for_model("gpt-4o")
    return _encoder


def estimate_tokens(text: str) -> int:
    return len(_get_encoder().encode(text))


def exceeds_limit(text: str, limit: int) -> bool:
    return estimate_tokens(text) > limit
