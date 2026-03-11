"""llm/factory.py - LLM クライアントのファクトリ"""

from config import LLMConfig
from .base import LLMClient


def create_llm_client(config: LLMConfig) -> LLMClient:
    provider = config.provider.lower().strip()

    if provider == "openai":
        from .openai_client import OpenAIClient
        if not config.openai_api_key:
            raise ValueError("OPENAI_API_KEY が未設定です。.env を確認してください。")
        return OpenAIClient(
            api_key=config.openai_api_key, model=config.openai_model,
            base_url=config.openai_base_url, default_temperature=config.temperature,
        )
    elif provider == "ollama":
        from .ollama_client import OllamaClient
        return OllamaClient(
            base_url=config.ollama_base_url, model=config.ollama_model,
            default_temperature=config.temperature,
        )
    else:
        raise ValueError(f"未対応の LLM_PROVIDER: '{provider}'")
