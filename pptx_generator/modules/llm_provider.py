"""
LLM Provider Module

Provides a unified interface for multiple LLM providers:
- Anthropic (Claude)
- OpenAI (GPT)
- Google (Gemini)
- Ollama (Local models)
- LiteLLM (Unified interface for 100+ providers)

Supports easy switching between providers for testing and comparison.
API keys are loaded from .env file in the project root.

Phase 1 Enhancement (2025-12-29):
- Added LiteLLM integration for unified multi-provider access
- Added tone and verbosity controls for content generation
- Updated model IDs to current versions
- Added Ollama support for local models
"""

import os
import json
import logging
from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from typing import Optional, List, Dict, Any
from enum import Enum
from pathlib import Path

logger = logging.getLogger(__name__)


# =============================================================================
# Environment Configuration
# =============================================================================

def load_env_file():
    """Load environment variables from .env file."""
    env_path = Path(__file__).parent.parent.parent / ".env"
    if env_path.exists():
        logger.info(f"Loading API keys from: {env_path}")
        with open(env_path, 'r') as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith('#') and '=' in line:
                    key, value = line.split('=', 1)
                    key = key.strip()
                    value = value.strip()
                    # Only set if not empty and not already set
                    if value and not os.environ.get(key):
                        os.environ[key] = value
    else:
        logger.warning(f".env file not found at: {env_path}")

# Load .env on module import
load_env_file()


# =============================================================================
# Tone and Verbosity Controls (Phase 1 Enhancement)
# =============================================================================

class Tone(Enum):
    """Content tone options for generation."""
    DEFAULT = "default"
    PROFESSIONAL = "professional"
    CASUAL = "casual"
    SALES_PITCH = "sales_pitch"
    EDUCATIONAL = "educational"
    EXECUTIVE = "executive"

TONE_PROMPTS = {
    Tone.DEFAULT: "",
    Tone.PROFESSIONAL: "Use formal business language with precise terminology. Be direct and data-driven. Avoid colloquialisms.",
    Tone.CASUAL: "Use conversational, approachable language. Include relatable examples. Keep it friendly but informative.",
    Tone.SALES_PITCH: "Emphasize benefits and value propositions. Use persuasive language with clear calls to action. Create urgency.",
    Tone.EDUCATIONAL: "Explain concepts clearly with examples. Build understanding progressively. Define technical terms.",
    Tone.EXECUTIVE: "Lead with key insights and recommendations. Be concise and strategic. Focus on business impact and decisions.",
}


class Verbosity(Enum):
    """Content verbosity options."""
    CONCISE = "concise"
    STANDARD = "standard"
    DETAILED = "detailed"

VERBOSITY_PROMPTS = {
    Verbosity.CONCISE: "Be extremely brief. Use bullet points. Maximum 3-4 words per bullet. Eliminate filler words.",
    Verbosity.STANDARD: "Provide balanced detail. Use complete sentences for key points. Include supporting context.",
    Verbosity.DETAILED: "Provide comprehensive coverage. Include supporting data, examples, and context. Elaborate on implications.",
}


@dataclass
class GenerationConfig:
    """Configuration for content generation."""
    tone: Tone = Tone.PROFESSIONAL
    verbosity: Verbosity = Verbosity.STANDARD
    max_bullets_per_slide: int = 6
    max_words_per_bullet: int = 15
    include_sources: bool = True

    def get_style_prompt(self) -> str:
        """Generate style instructions for LLM prompts."""
        parts = []

        tone_prompt = TONE_PROMPTS.get(self.tone, "")
        if tone_prompt:
            parts.append(f"Tone: {tone_prompt}")

        verbosity_prompt = VERBOSITY_PROMPTS.get(self.verbosity, "")
        if verbosity_prompt:
            parts.append(f"Verbosity: {verbosity_prompt}")

        parts.append(f"Constraints: Maximum {self.max_bullets_per_slide} bullets per slide, {self.max_words_per_bullet} words per bullet.")

        return "\n".join(parts)


# =============================================================================
# Model Configuration
# =============================================================================

class ModelProvider(Enum):
    """Supported LLM providers."""
    ANTHROPIC = "anthropic"
    OPENAI = "openai"
    GOOGLE = "google"
    OLLAMA = "ollama"
    LITELLM = "litellm"  # Unified interface


@dataclass
class ModelConfig:
    """Configuration for a specific model."""
    provider: ModelProvider
    model_id: str
    display_name: str
    max_tokens: int = 4096
    temperature: float = 0.7
    litellm_model: str = None  # LiteLLM model identifier

    def __post_init__(self):
        """Set LiteLLM model identifier if not provided."""
        if self.litellm_model is None:
            # Map to LiteLLM format
            if self.provider == ModelProvider.ANTHROPIC:
                self.litellm_model = self.model_id
            elif self.provider == ModelProvider.OPENAI:
                self.litellm_model = self.model_id
            elif self.provider == ModelProvider.GOOGLE:
                self.litellm_model = f"gemini/{self.model_id}"
            elif self.provider == ModelProvider.OLLAMA:
                self.litellm_model = f"ollama/{self.model_id}"


# Available models - Current models as of December 2025
AVAILABLE_MODELS = {
    # Anthropic models (Claude 3.5 / Claude 4 series)
    "claude-sonnet-4": ModelConfig(
        provider=ModelProvider.ANTHROPIC,
        model_id="claude-sonnet-4-20250514",
        display_name="Claude Sonnet 4",
        max_tokens=8192,
        litellm_model="claude-sonnet-4-20250514"
    ),
    "claude-3.5-sonnet": ModelConfig(
        provider=ModelProvider.ANTHROPIC,
        model_id="claude-3-5-sonnet-20241022",
        display_name="Claude 3.5 Sonnet",
        max_tokens=8192,
        litellm_model="claude-3-5-sonnet-20241022"
    ),
    "claude-3.5-haiku": ModelConfig(
        provider=ModelProvider.ANTHROPIC,
        model_id="claude-3-5-haiku-20241022",
        display_name="Claude 3.5 Haiku",
        max_tokens=8192,
        litellm_model="claude-3-5-haiku-20241022"
    ),
    "claude-3-opus": ModelConfig(
        provider=ModelProvider.ANTHROPIC,
        model_id="claude-3-opus-20240229",
        display_name="Claude 3 Opus",
        max_tokens=4096,
        litellm_model="claude-3-opus-20240229"
    ),
    # OpenAI models (GPT-4 series)
    "gpt-4o": ModelConfig(
        provider=ModelProvider.OPENAI,
        model_id="gpt-4o",
        display_name="GPT-4o",
        max_tokens=4096,
        litellm_model="gpt-4o"
    ),
    "gpt-4o-mini": ModelConfig(
        provider=ModelProvider.OPENAI,
        model_id="gpt-4o-mini",
        display_name="GPT-4o Mini",
        max_tokens=4096,
        litellm_model="gpt-4o-mini"
    ),
    "gpt-4-turbo": ModelConfig(
        provider=ModelProvider.OPENAI,
        model_id="gpt-4-turbo",
        display_name="GPT-4 Turbo",
        max_tokens=4096,
        litellm_model="gpt-4-turbo"
    ),
    "o1": ModelConfig(
        provider=ModelProvider.OPENAI,
        model_id="o1",
        display_name="OpenAI o1",
        max_tokens=8192,
        litellm_model="o1"
    ),
    "o1-mini": ModelConfig(
        provider=ModelProvider.OPENAI,
        model_id="o1-mini",
        display_name="OpenAI o1-mini",
        max_tokens=8192,
        litellm_model="o1-mini"
    ),
    # Google models (Gemini series)
    "gemini-2.0-flash": ModelConfig(
        provider=ModelProvider.GOOGLE,
        model_id="gemini-2.0-flash-exp",
        display_name="Gemini 2.0 Flash",
        max_tokens=8192,
        litellm_model="gemini/gemini-2.0-flash-exp"
    ),
    "gemini-1.5-pro": ModelConfig(
        provider=ModelProvider.GOOGLE,
        model_id="gemini-1.5-pro",
        display_name="Gemini 1.5 Pro",
        max_tokens=8192,
        litellm_model="gemini/gemini-1.5-pro"
    ),
    "gemini-1.5-flash": ModelConfig(
        provider=ModelProvider.GOOGLE,
        model_id="gemini-1.5-flash",
        display_name="Gemini 1.5 Flash",
        max_tokens=8192,
        litellm_model="gemini/gemini-1.5-flash"
    ),
    # Ollama (Local models)
    "ollama-llama3.2": ModelConfig(
        provider=ModelProvider.OLLAMA,
        model_id="llama3.2",
        display_name="Llama 3.2 (Local)",
        max_tokens=4096,
        litellm_model="ollama/llama3.2"
    ),
    "ollama-mistral": ModelConfig(
        provider=ModelProvider.OLLAMA,
        model_id="mistral",
        display_name="Mistral (Local)",
        max_tokens=4096,
        litellm_model="ollama/mistral"
    ),
    "ollama-qwen2.5": ModelConfig(
        provider=ModelProvider.OLLAMA,
        model_id="qwen2.5",
        display_name="Qwen 2.5 (Local)",
        max_tokens=4096,
        litellm_model="ollama/qwen2.5"
    ),
}


@dataclass
class LLMResponse:
    """Standardized response from any LLM provider."""
    content: str
    model: str
    provider: str
    usage: Dict[str, int]
    raw_response: Any = None


class BaseLLMProvider(ABC):
    """Abstract base class for LLM providers."""

    @abstractmethod
    def generate(
        self,
        prompt: str,
        system_prompt: Optional[str] = None,
        max_tokens: int = 4096,
        temperature: float = 0.7
    ) -> LLMResponse:
        """Generate a response from the LLM."""
        pass

    @abstractmethod
    def is_available(self) -> bool:
        """Check if the provider is properly configured."""
        pass


class AnthropicProvider(BaseLLMProvider):
    """Anthropic Claude provider."""

    def __init__(self, model_id: str = "claude-sonnet-4-5"):
        self.model_id = model_id
        self.api_key = os.environ.get("ANTHROPIC_API_KEY")
        self._client = None

    def _get_client(self):
        """Lazy initialization of Anthropic client."""
        if self._client is None:
            try:
                import anthropic
                self._client = anthropic.Anthropic(api_key=self.api_key)
            except ImportError:
                raise ImportError("anthropic package not installed. Run: pip install anthropic")
        return self._client

    def is_available(self) -> bool:
        """Check if Anthropic API key is configured."""
        return bool(self.api_key)

    def generate(
        self,
        prompt: str,
        system_prompt: Optional[str] = None,
        max_tokens: int = 4096,
        temperature: float = 0.7
    ) -> LLMResponse:
        """Generate a response using Claude."""
        if not self.is_available():
            raise ValueError("ANTHROPIC_API_KEY not set in .env file")

        client = self._get_client()

        kwargs = {
            "model": self.model_id,
            "max_tokens": max_tokens,
            "messages": [{"role": "user", "content": prompt}]
        }

        if system_prompt:
            kwargs["system"] = system_prompt

        if temperature is not None:
            kwargs["temperature"] = temperature

        response = client.messages.create(**kwargs)

        return LLMResponse(
            content=response.content[0].text,
            model=self.model_id,
            provider="anthropic",
            usage={
                "input_tokens": response.usage.input_tokens,
                "output_tokens": response.usage.output_tokens
            },
            raw_response=response
        )


class OpenAIProvider(BaseLLMProvider):
    """OpenAI GPT provider."""

    def __init__(self, model_id: str = "gpt-5"):
        self.model_id = model_id
        self.api_key = os.environ.get("OPENAI_API_KEY")
        self._client = None

    def _get_client(self):
        """Lazy initialization of OpenAI client."""
        if self._client is None:
            try:
                import openai
                self._client = openai.OpenAI(api_key=self.api_key)
            except ImportError:
                raise ImportError("openai package not installed. Run: pip install openai")
        return self._client

    def is_available(self) -> bool:
        """Check if OpenAI API key is configured."""
        return bool(self.api_key)

    def generate(
        self,
        prompt: str,
        system_prompt: Optional[str] = None,
        max_tokens: int = 4096,
        temperature: float = 0.7
    ) -> LLMResponse:
        """Generate a response using GPT."""
        if not self.is_available():
            raise ValueError("OPENAI_API_KEY not set in .env file")

        client = self._get_client()

        messages = []
        if system_prompt:
            messages.append({"role": "system", "content": system_prompt})
        messages.append({"role": "user", "content": prompt})

        response = client.chat.completions.create(
            model=self.model_id,
            messages=messages,
            max_tokens=max_tokens,
            temperature=temperature
        )

        return LLMResponse(
            content=response.choices[0].message.content,
            model=self.model_id,
            provider="openai",
            usage={
                "input_tokens": response.usage.prompt_tokens,
                "output_tokens": response.usage.completion_tokens
            },
            raw_response=response
        )


class GoogleProvider(BaseLLMProvider):
    """Google Gemini provider."""

    def __init__(self, model_id: str = "gemini-1.5-flash"):
        self.model_id = model_id
        self.api_key = os.environ.get("GOOGLE_API_KEY")
        self._client = None

    def _get_client(self):
        """Lazy initialization of Google Generative AI client."""
        if self._client is None:
            try:
                import google.generativeai as genai
                genai.configure(api_key=self.api_key)
                self._client = genai
            except ImportError:
                raise ImportError("google-generativeai package not installed. Run: pip install google-generativeai")
        return self._client

    def is_available(self) -> bool:
        """Check if Google API key is configured."""
        return bool(self.api_key)

    def generate(
        self,
        prompt: str,
        system_prompt: Optional[str] = None,
        max_tokens: int = 4096,
        temperature: float = 0.7
    ) -> LLMResponse:
        """Generate a response using Gemini."""
        if not self.is_available():
            raise ValueError("GOOGLE_API_KEY not set in .env file")

        genai = self._get_client()

        # Configure model
        generation_config = {
            "temperature": temperature,
            "max_output_tokens": max_tokens,
        }

        model = genai.GenerativeModel(
            model_name=self.model_id,
            generation_config=generation_config,
            system_instruction=system_prompt if system_prompt else None
        )

        response = model.generate_content(prompt)

        # Extract usage stats if available
        usage = {"input_tokens": 0, "output_tokens": 0}
        if hasattr(response, 'usage_metadata'):
            usage = {
                "input_tokens": getattr(response.usage_metadata, 'prompt_token_count', 0),
                "output_tokens": getattr(response.usage_metadata, 'candidates_token_count', 0)
            }

        return LLMResponse(
            content=response.text,
            model=self.model_id,
            provider="google",
            usage=usage,
            raw_response=response
        )


class OllamaProvider(BaseLLMProvider):
    """Ollama local model provider."""

    def __init__(self, model_id: str = "llama3.2"):
        self.model_id = model_id
        self.base_url = os.environ.get("OLLAMA_URL", "http://localhost:11434")

    def is_available(self) -> bool:
        """Check if Ollama is running."""
        try:
            import requests
            response = requests.get(f"{self.base_url}/api/tags", timeout=2)
            return response.status_code == 200
        except Exception:
            return False

    def generate(
        self,
        prompt: str,
        system_prompt: Optional[str] = None,
        max_tokens: int = 4096,
        temperature: float = 0.7
    ) -> LLMResponse:
        """Generate a response using Ollama."""
        try:
            import requests
        except ImportError:
            raise ImportError("requests package not installed. Run: pip install requests")

        messages = []
        if system_prompt:
            messages.append({"role": "system", "content": system_prompt})
        messages.append({"role": "user", "content": prompt})

        response = requests.post(
            f"{self.base_url}/api/chat",
            json={
                "model": self.model_id,
                "messages": messages,
                "options": {
                    "num_predict": max_tokens,
                    "temperature": temperature
                },
                "stream": False
            }
        )
        response.raise_for_status()
        data = response.json()

        return LLMResponse(
            content=data.get("message", {}).get("content", ""),
            model=self.model_id,
            provider="ollama",
            usage={
                "input_tokens": data.get("prompt_eval_count", 0),
                "output_tokens": data.get("eval_count", 0)
            },
            raw_response=data
        )


class LiteLLMProvider(BaseLLMProvider):
    """
    LiteLLM unified provider - supports 100+ models through a single interface.

    This is the recommended provider for production use as it:
    - Provides consistent API across all providers
    - Handles retries and fallbacks
    - Supports streaming, async, and caching
    - Has built-in cost tracking
    """

    def __init__(self, model: str = "claude-3-5-sonnet-20241022"):
        self.model = model
        self._initialized = False

    def _ensure_initialized(self):
        """Lazy initialization of LiteLLM."""
        if not self._initialized:
            try:
                import litellm
                # Suppress verbose logging
                litellm.set_verbose = False
                self._initialized = True
            except ImportError:
                raise ImportError(
                    "litellm package not installed. Run: pip install litellm"
                )

    def is_available(self) -> bool:
        """Check if the model's provider is configured."""
        try:
            self._ensure_initialized()
            # Check based on model prefix
            model_lower = self.model.lower()
            if "claude" in model_lower or "anthropic" in model_lower:
                return bool(os.environ.get("ANTHROPIC_API_KEY"))
            elif "gpt" in model_lower or "o1" in model_lower:
                return bool(os.environ.get("OPENAI_API_KEY"))
            elif "gemini" in model_lower:
                return bool(os.environ.get("GOOGLE_API_KEY") or os.environ.get("GEMINI_API_KEY"))
            elif "ollama" in model_lower:
                return True  # Ollama doesn't need API key
            return True  # Assume available for other models
        except ImportError:
            return False

    def generate(
        self,
        prompt: str,
        system_prompt: Optional[str] = None,
        max_tokens: int = 4096,
        temperature: float = 0.7
    ) -> LLMResponse:
        """Generate a response using LiteLLM."""
        self._ensure_initialized()
        import litellm

        messages = []
        if system_prompt:
            messages.append({"role": "system", "content": system_prompt})
        messages.append({"role": "user", "content": prompt})

        response = litellm.completion(
            model=self.model,
            messages=messages,
            max_tokens=max_tokens,
            temperature=temperature
        )

        return LLMResponse(
            content=response.choices[0].message.content,
            model=self.model,
            provider="litellm",
            usage={
                "input_tokens": response.usage.prompt_tokens,
                "output_tokens": response.usage.completion_tokens
            },
            raw_response=response
        )

    async def agenerate(
        self,
        prompt: str,
        system_prompt: Optional[str] = None,
        max_tokens: int = 4096,
        temperature: float = 0.7
    ) -> LLMResponse:
        """Generate a response asynchronously using LiteLLM."""
        self._ensure_initialized()
        import litellm

        messages = []
        if system_prompt:
            messages.append({"role": "system", "content": system_prompt})
        messages.append({"role": "user", "content": prompt})

        response = await litellm.acompletion(
            model=self.model,
            messages=messages,
            max_tokens=max_tokens,
            temperature=temperature
        )

        return LLMResponse(
            content=response.choices[0].message.content,
            model=self.model,
            provider="litellm",
            usage={
                "input_tokens": response.usage.prompt_tokens,
                "output_tokens": response.usage.completion_tokens
            },
            raw_response=response
        )


class LLMManager:
    """
    Manages LLM providers and provides a unified interface.

    Usage:
        # Using LiteLLM (recommended - unified interface)
        manager = LLMManager(use_litellm=True)
        manager.set_model("claude-3.5-sonnet")
        response = manager.generate("Write a summary about industrial real estate")

        # Using native providers
        manager = LLMManager(use_litellm=False)
        manager.set_model("gpt-4o")
        response = manager.generate("Write a summary about industrial real estate")

        # With tone and verbosity controls
        from .llm_provider import GenerationConfig, Tone, Verbosity
        config = GenerationConfig(tone=Tone.EXECUTIVE, verbosity=Verbosity.CONCISE)
        response = manager.generate(
            "Write a summary about industrial real estate",
            generation_config=config
        )
    """

    def __init__(
        self,
        default_model: str = "claude-3.5-sonnet",
        use_litellm: bool = True,
        generation_config: GenerationConfig = None
    ):
        self.providers: Dict[ModelProvider, BaseLLMProvider] = {}
        self.current_model: Optional[ModelConfig] = None
        self.use_litellm = use_litellm
        self.generation_config = generation_config or GenerationConfig()
        self._litellm_provider: Optional[LiteLLMProvider] = None
        self.set_model(default_model)

    def set_model(self, model_name: str) -> None:
        """
        Set the current model to use.

        Args:
            model_name: One of the keys in AVAILABLE_MODELS, or a direct LiteLLM model string
        """
        if model_name in AVAILABLE_MODELS:
            self.current_model = AVAILABLE_MODELS[model_name]
        else:
            # Allow direct LiteLLM model strings (e.g., "anthropic/claude-3-sonnet")
            if self.use_litellm:
                self.current_model = ModelConfig(
                    provider=ModelProvider.LITELLM,
                    model_id=model_name,
                    display_name=model_name,
                    litellm_model=model_name
                )
            else:
                available = ", ".join(AVAILABLE_MODELS.keys())
                raise ValueError(f"Unknown model '{model_name}'. Available: {available}")

        if self.use_litellm:
            # Use LiteLLM unified provider
            litellm_model = self.current_model.litellm_model or self.current_model.model_id
            self._litellm_provider = LiteLLMProvider(model=litellm_model)
        else:
            # Initialize native provider if needed
            if self.current_model.provider not in self.providers:
                if self.current_model.provider == ModelProvider.ANTHROPIC:
                    self.providers[ModelProvider.ANTHROPIC] = AnthropicProvider(
                        self.current_model.model_id
                    )
                elif self.current_model.provider == ModelProvider.OPENAI:
                    self.providers[ModelProvider.OPENAI] = OpenAIProvider(
                        self.current_model.model_id
                    )
                elif self.current_model.provider == ModelProvider.GOOGLE:
                    self.providers[ModelProvider.GOOGLE] = GoogleProvider(
                        self.current_model.model_id
                    )
                elif self.current_model.provider == ModelProvider.OLLAMA:
                    self.providers[ModelProvider.OLLAMA] = OllamaProvider(
                        self.current_model.model_id
                    )
            else:
                # Update model ID on existing provider
                provider = self.providers[self.current_model.provider]
                provider.model_id = self.current_model.model_id

        logger.info(f"Set model to: {self.current_model.display_name} (LiteLLM: {self.use_litellm})")

    def get_provider(self) -> BaseLLMProvider:
        """Get the current provider instance."""
        if not self.current_model:
            raise ValueError("No model selected")
        if self.use_litellm and self._litellm_provider:
            return self._litellm_provider
        return self.providers[self.current_model.provider]

    def is_available(self) -> bool:
        """Check if current provider is properly configured."""
        try:
            return self.get_provider().is_available()
        except (ValueError, KeyError):
            return False

    def set_generation_config(self, config: GenerationConfig) -> None:
        """Set the generation configuration for tone and verbosity."""
        self.generation_config = config

    def generate(
        self,
        prompt: str,
        system_prompt: Optional[str] = None,
        max_tokens: Optional[int] = None,
        temperature: Optional[float] = None,
        generation_config: Optional[GenerationConfig] = None
    ) -> LLMResponse:
        """
        Generate a response using the current model.

        Args:
            prompt: The user prompt
            system_prompt: Optional system/context prompt
            max_tokens: Override default max tokens
            temperature: Override default temperature
            generation_config: Override generation config for tone/verbosity

        Returns:
            LLMResponse with the generated content
        """
        provider = self.get_provider()
        config = generation_config or self.generation_config

        # Build enhanced system prompt with style instructions
        enhanced_system = system_prompt or ""
        style_prompt = config.get_style_prompt()
        if style_prompt:
            if enhanced_system:
                enhanced_system = f"{enhanced_system}\n\n{style_prompt}"
            else:
                enhanced_system = style_prompt

        return provider.generate(
            prompt=prompt,
            system_prompt=enhanced_system if enhanced_system else None,
            max_tokens=max_tokens or self.current_model.max_tokens,
            temperature=temperature if temperature is not None else self.current_model.temperature
        )

    async def agenerate(
        self,
        prompt: str,
        system_prompt: Optional[str] = None,
        max_tokens: Optional[int] = None,
        temperature: Optional[float] = None,
        generation_config: Optional[GenerationConfig] = None
    ) -> LLMResponse:
        """
        Generate a response asynchronously using the current model.

        Only available when using LiteLLM provider.

        Args:
            prompt: The user prompt
            system_prompt: Optional system/context prompt
            max_tokens: Override default max tokens
            temperature: Override default temperature
            generation_config: Override generation config for tone/verbosity

        Returns:
            LLMResponse with the generated content
        """
        if not self.use_litellm or not self._litellm_provider:
            raise NotImplementedError("Async generation only available with LiteLLM provider")

        config = generation_config or self.generation_config

        # Build enhanced system prompt with style instructions
        enhanced_system = system_prompt or ""
        style_prompt = config.get_style_prompt()
        if style_prompt:
            if enhanced_system:
                enhanced_system = f"{enhanced_system}\n\n{style_prompt}"
            else:
                enhanced_system = style_prompt

        return await self._litellm_provider.agenerate(
            prompt=prompt,
            system_prompt=enhanced_system if enhanced_system else None,
            max_tokens=max_tokens or self.current_model.max_tokens,
            temperature=temperature if temperature is not None else self.current_model.temperature
        )

    def list_models(self) -> List[Dict[str, Any]]:
        """List all available models with their status."""
        models = []
        for name, config in AVAILABLE_MODELS.items():
            # Check if provider is configured
            if config.provider == ModelProvider.ANTHROPIC:
                available = bool(os.environ.get("ANTHROPIC_API_KEY"))
            elif config.provider == ModelProvider.OPENAI:
                available = bool(os.environ.get("OPENAI_API_KEY"))
            elif config.provider == ModelProvider.GOOGLE:
                available = bool(os.environ.get("GOOGLE_API_KEY") or os.environ.get("GEMINI_API_KEY"))
            elif config.provider == ModelProvider.OLLAMA:
                # Check if Ollama is running
                try:
                    import requests
                    ollama_url = os.environ.get("OLLAMA_URL", "http://localhost:11434")
                    response = requests.get(f"{ollama_url}/api/tags", timeout=1)
                    available = response.status_code == 200
                except Exception:
                    available = False
            else:
                available = False

            models.append({
                "name": name,
                "display_name": config.display_name,
                "provider": config.provider.value,
                "model_id": config.model_id,
                "litellm_model": config.litellm_model,
                "available": available,
                "current": self.current_model and name == self._get_current_model_name()
            })

        return models

    def _get_current_model_name(self) -> Optional[str]:
        """Get the name of the current model."""
        if not self.current_model:
            return None
        for name, config in AVAILABLE_MODELS.items():
            if config.model_id == self.current_model.model_id:
                return name
        return None

    def get_status(self) -> Dict[str, Any]:
        """Get status of all providers."""
        # Check Ollama availability
        ollama_available = False
        try:
            import requests
            ollama_url = os.environ.get("OLLAMA_URL", "http://localhost:11434")
            response = requests.get(f"{ollama_url}/api/tags", timeout=1)
            ollama_available = response.status_code == 200
        except Exception:
            pass

        return {
            "current_model": self._get_current_model_name(),
            "current_display_name": self.current_model.display_name if self.current_model else None,
            "use_litellm": self.use_litellm,
            "generation_config": {
                "tone": self.generation_config.tone.value,
                "verbosity": self.generation_config.verbosity.value,
            },
            "anthropic_configured": bool(os.environ.get("ANTHROPIC_API_KEY")),
            "openai_configured": bool(os.environ.get("OPENAI_API_KEY")),
            "google_configured": bool(os.environ.get("GOOGLE_API_KEY") or os.environ.get("GEMINI_API_KEY")),
            "ollama_available": ollama_available,
            "models": self.list_models()
        }


# =============================================================================
# Convenience Functions
# =============================================================================

def get_llm(
    model: str = "claude-3.5-sonnet",
    use_litellm: bool = True,
    tone: str = "professional",
    verbosity: str = "standard"
) -> LLMManager:
    """
    Get an LLM manager with the specified model and configuration.

    Args:
        model: Model name from AVAILABLE_MODELS or direct LiteLLM model string
        use_litellm: Whether to use LiteLLM unified interface (recommended)
        tone: Content tone (default, professional, casual, sales_pitch, educational, executive)
        verbosity: Content verbosity (concise, standard, detailed)

    Returns:
        Configured LLMManager instance
    """
    config = GenerationConfig(
        tone=Tone(tone),
        verbosity=Verbosity(verbosity)
    )
    return LLMManager(default_model=model, use_litellm=use_litellm, generation_config=config)


def quick_generate(
    prompt: str,
    model: str = "claude-3.5-sonnet",
    tone: str = "professional",
    verbosity: str = "standard"
) -> str:
    """
    Quick helper for one-off generation without managing LLMManager.

    Args:
        prompt: The prompt to send to the LLM
        model: Model name to use
        tone: Content tone
        verbosity: Content verbosity

    Returns:
        Generated text content
    """
    manager = get_llm(model=model, tone=tone, verbosity=verbosity)
    response = manager.generate(prompt)
    return response.content


def main():
    """Test the LLM providers."""
    import argparse

    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

    parser = argparse.ArgumentParser(description="LLM Provider Test")
    parser.add_argument("--model", "-m", default="claude-3.5-sonnet", help="Model to use")
    parser.add_argument("--list", "-l", action="store_true", help="List available models")
    parser.add_argument("--status", "-s", action="store_true", help="Show provider status")
    parser.add_argument("--test", "-t", action="store_true", help="Run a test generation")
    parser.add_argument("--prompt", "-p", default="Say hello in one sentence.", help="Test prompt")
    parser.add_argument("--tone", default="professional",
                       choices=["default", "professional", "casual", "sales_pitch", "educational", "executive"],
                       help="Content tone")
    parser.add_argument("--verbosity", default="standard",
                       choices=["concise", "standard", "detailed"],
                       help="Content verbosity")
    parser.add_argument("--no-litellm", action="store_true", help="Use native providers instead of LiteLLM")

    args = parser.parse_args()

    config = GenerationConfig(
        tone=Tone(args.tone),
        verbosity=Verbosity(args.verbosity)
    )
    manager = LLMManager(
        default_model=args.model,
        use_litellm=not args.no_litellm,
        generation_config=config
    )

    if args.list:
        print("\nAvailable Models:")
        print("-" * 80)
        for model in manager.list_models():
            status = "OK" if model["available"] else "X"
            current = " (current)" if model["current"] else ""
            print(f"  [{status}] {model['name']}: {model['display_name']}{current}")
            print(f"       Provider: {model['provider']}, LiteLLM: {model['litellm_model']}")
        print()

    if args.status:
        status = manager.get_status()
        print("\nProvider Status:")
        print("-" * 60)
        print(f"  Current model: {status['current_display_name']}")
        print(f"  Using LiteLLM: {status['use_litellm']}")
        print(f"  Tone: {status['generation_config']['tone']}")
        print(f"  Verbosity: {status['generation_config']['verbosity']}")
        print()
        print("  API Keys:")
        print(f"    Anthropic: {'Configured' if status['anthropic_configured'] else 'Not configured'}")
        print(f"    OpenAI: {'Configured' if status['openai_configured'] else 'Not configured'}")
        print(f"    Google: {'Configured' if status['google_configured'] else 'Not configured'}")
        print(f"    Ollama: {'Running' if status['ollama_available'] else 'Not running'}")
        print()

    if args.test:
        if not manager.is_available():
            print(f"\nError: {args.model} provider not configured (missing API key in .env)")
            return

        print(f"\nTesting {manager.current_model.display_name}...")
        print(f"Prompt: {args.prompt}")
        print(f"Tone: {args.tone}, Verbosity: {args.verbosity}")
        print("-" * 60)

        try:
            response = manager.generate(args.prompt)
            print(f"Response: {response.content}")
            print(f"\nTokens: {response.usage}")
            print(f"Provider: {response.provider}")
        except Exception as e:
            print(f"Error: {e}")
            import traceback
            traceback.print_exc()


if __name__ == "__main__":
    main()
