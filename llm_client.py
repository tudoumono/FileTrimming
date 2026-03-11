"""
llm_client.py - LLM プロバイダー抽象化レイヤー

設計方針:
    共通インターフェース (BaseLLMClient) を定義し、
    OpenAI / Ollama / Azure OpenAI をプロバイダーとして差し替え可能にする。
    Ollama は OpenAI 互換エンドポイント (/v1/chat/completions) を持つため、
    OpenAI SDK の base_url を差し替えるだけで接続できる。

追加方法:
    1. BaseLLMClient を継承
    2. chat() を実装
    3. _PROVIDERS に登録
"""
from __future__ import annotations
import json, logging
from abc import ABC, abstractmethod
from dataclasses import dataclass
from typing import Optional
from config import PipelineConfig

logger = logging.getLogger(__name__)

@dataclass
class LLMResponse:
    content: str
    model: str
    provider: str
    usage: Optional[dict] = None

class BaseLLMClient(ABC):
    def __init__(self, config: PipelineConfig):
        self.config = config

    @abstractmethod
    def chat(self, system_prompt: str, user_message: str, *,
             temperature: Optional[float] = None, max_tokens: int = 16_000,
             response_format: Optional[dict] = None) -> LLMResponse: ...

    def chat_json(self, system_prompt: str, user_message: str, **kw) -> dict:
        resp = self.chat(system_prompt, user_message,
                         response_format={"type": "json_object"}, **kw)
        return json.loads(resp.content)

class OpenAIClient(BaseLLMClient):
    def __init__(self, config: PipelineConfig):
        super().__init__(config)
        import openai
        self._client = openai.OpenAI(api_key=config.openai_api_key,
                                     base_url=config.openai_base_url)
        self._model = config.openai_model

    def chat(self, system_prompt, user_message, *, temperature=None,
             max_tokens=16_000, response_format=None) -> LLMResponse:
        temp = temperature if temperature is not None else self.config.llm_temperature
        kw = dict(model=self._model, messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_message}],
            max_tokens=max_tokens, temperature=temp)
        if response_format:
            kw["response_format"] = response_format
        resp = self._client.chat.completions.create(**kw)
        usage = ({"prompt_tokens": resp.usage.prompt_tokens,
                  "completion_tokens": resp.usage.completion_tokens}
                 if resp.usage else None)
        return LLMResponse(content=resp.choices[0].message.content or "",
                           model=resp.model, provider="openai", usage=usage)

class OllamaClient(BaseLLMClient):
    """Ollama: OpenAI互換エンドポイントを利用"""
    def __init__(self, config: PipelineConfig):
        super().__init__(config)
        import openai
        self._client = openai.OpenAI(api_key="ollama",
                                     base_url=f"{config.ollama_base_url}/v1")
        self._model = config.ollama_model

    def chat(self, system_prompt, user_message, *, temperature=None,
             max_tokens=16_000, response_format=None) -> LLMResponse:
        temp = temperature if temperature is not None else self.config.llm_temperature
        kw = dict(model=self._model, messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_message}],
            max_tokens=max_tokens, temperature=temp)
        if response_format:
            kw["response_format"] = response_format
        resp = self._client.chat.completions.create(**kw)
        usage = ({"prompt_tokens": resp.usage.prompt_tokens,
                  "completion_tokens": resp.usage.completion_tokens}
                 if resp.usage else None)
        return LLMResponse(content=resp.choices[0].message.content or "",
                           model=resp.model or self._model,
                           provider="ollama", usage=usage)

class AzureOpenAIClient(BaseLLMClient):
    """Azure OpenAI（将来用スタブ）"""
    def __init__(self, config: PipelineConfig):
        super().__init__(config)
        import openai
        self._client = openai.AzureOpenAI(
            api_key=config.azure_openai_api_key,
            azure_endpoint=config.azure_openai_endpoint,
            api_version="2024-02-01")
        self._deployment = config.azure_openai_deployment

    def chat(self, system_prompt, user_message, *, temperature=None,
             max_tokens=16_000, response_format=None) -> LLMResponse:
        temp = temperature if temperature is not None else self.config.llm_temperature
        kw = dict(model=self._deployment, messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_message}],
            max_tokens=max_tokens, temperature=temp)
        if response_format:
            kw["response_format"] = response_format
        resp = self._client.chat.completions.create(**kw)
        return LLMResponse(content=resp.choices[0].message.content or "",
                           model=resp.model or self._deployment,
                           provider="azure_openai")

_PROVIDERS: dict[str, type[BaseLLMClient]] = {
    "openai": OpenAIClient,
    "ollama": OllamaClient,
    "azure_openai": AzureOpenAIClient,
}

def create_llm_client(config: PipelineConfig) -> BaseLLMClient:
    provider = config.llm_provider
    if provider not in _PROVIDERS:
        raise ValueError(f"未対応 LLM: {provider!r} (選択肢: {list(_PROVIDERS)})")
    cls = _PROVIDERS[provider]
    logger.info("LLM: %s (%s)", provider, cls.__name__)
    return cls(config)
