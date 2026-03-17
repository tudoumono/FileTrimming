"""LLM バックエンド生成ユーティリティ。"""

from __future__ import annotations

from src.config import PipelineConfig
from src.llm.base import LLMBackend
from src.llm.local_backend import LocalBackend
from src.llm.noop_backend import NoopBackend
from src.llm.openai_backend import OpenAIBackend


def create_backend(config: PipelineConfig) -> LLMBackend:
    """設定値から LLM バックエンドを生成する。"""
    if config.llm_backend == "noop":
        return NoopBackend()
    if config.llm_backend == "openai":
        return OpenAIBackend(
            api_key=config.openai_api_key,
            base_url=config.openai_base_url,
            model=config.openai_model,
            proxy_url=config.llm_proxy_url,
            skip_ssl_verify=config.llm_skip_ssl_verify,
        )
    if config.llm_backend == "local":
        return LocalBackend(
            base_url=config.ollama_base_url,
            model=config.ollama_model,
            proxy_url=config.llm_proxy_url,
            skip_ssl_verify=config.llm_skip_ssl_verify,
        )
    raise ValueError(f"unsupported llm_backend: {config.llm_backend}")
