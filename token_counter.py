"""utils/token_counter.py - トークン数推定"""
import tiktoken
from pathlib import Path

_encoder = None

def get_encoder():
    global _encoder
    if _encoder is None:
        _encoder = tiktoken.encoding_for_model("gpt-4o")
    return _encoder

def count_tokens(text: str) -> int:
    return len(get_encoder().encode(text))

def estimate_file_tokens(filepath, encoding="utf-8") -> int:
    return count_tokens(Path(filepath).read_text(encoding=encoding, errors="replace"))
