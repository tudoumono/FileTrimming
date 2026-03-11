"""
llm/base.py - LLM クライアントの抽象基底クラス

【設計方針: OpenAI -> Ollama 対応】

1. LLMClient (抽象基底クラス)
   - chat() メソッドだけを公開インタフェースとする
   - 各ステップは LLMClient のみに依存し、プロバイダを知らない

2. .env の LLM_PROVIDER 切り替えだけで OpenAI <-> Ollama を変更可能

3. Ollama は OpenAI 互換 API を提供するため、
   openai ライブラリの base_url 差し替えで動作する。
   プロバイダ固有の調整は各実装クラス内に閉じ込める
"""

from abc import ABC, abstractmethod
from dataclasses import dataclass
from typing import Optional


@dataclass
class LLMResponse:
    """LLM からの応答"""
    content: str
    model: str
    prompt_tokens: int = 0
    completion_tokens: int = 0
    total_tokens: int = 0


class LLMClient(ABC):
    """LLM クライアントの抽象インタフェース"""

    @abstractmethod
    def chat(
        self,
        system_prompt: str,
        user_message: str,
        max_tokens: int = 16000,
        temperature: Optional[float] = None,
        response_format: Optional[dict] = None,
    ) -> LLMResponse:
        ...

    @abstractmethod
    def provider_name(self) -> str:
        ...
