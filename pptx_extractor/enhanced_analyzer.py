"""
Enhanced Slide Analyzer with Optimized LLM Vision Analysis

This module provides advanced slide analysis capabilities:
- Extended thinking for complex analysis
- Multi-pass analysis strategy
- Structured output validation
- Caching and batch processing
- Image preprocessing integration
"""

import json
import hashlib
import logging
import os
import time
from pathlib import Path
from typing import Optional, Dict, Any, List, Tuple
from dataclasses import dataclass, field, asdict
from enum import Enum
from concurrent.futures import ThreadPoolExecutor, as_completed

logger = logging.getLogger(__name__)


class AnalysisMode(Enum):
    """Analysis modes for different accuracy/speed tradeoffs."""
    FAST = "fast"           # Single pass, basic model
    STANDARD = "standard"   # Single pass, best model
    THOROUGH = "thorough"   # Multi-pass with validation
    ULTRA = "ultra"         # Extended thinking + multi-pass


class ModelTier(Enum):
    """Model tiers for different tasks."""
    FAST = "fast"       # Quick, cheap (Haiku 4.5/GPT-5-nano)
    BALANCED = "balanced"  # Good balance (Sonnet 4.5/GPT-5)
    BEST = "best"       # Highest quality (Opus 4.5/GPT-5.2)


@dataclass
class AnalysisConfig:
    """Configuration for enhanced analysis."""
    mode: AnalysisMode = AnalysisMode.STANDARD

    # Model selection
    preferred_provider: str = "google"  # google, anthropic, openai
    model_tier: ModelTier = ModelTier.BALANCED

    # Extended thinking
    use_extended_thinking: bool = False
    thinking_budget_tokens: int = 10000

    # Multi-pass
    num_passes: int = 1
    refine_on_low_confidence: bool = True
    confidence_threshold: float = 0.85

    # Preprocessing
    preprocess_image: bool = True
    add_measurement_grid: bool = True
    add_rulers: bool = True

    # Validation
    validate_output: bool = True
    auto_correct_errors: bool = True

    # Caching
    use_cache: bool = True
    cache_dir: Optional[Path] = None

    # Batch processing
    max_parallel: int = 3


@dataclass
class AnalysisResult:
    """Result of slide analysis."""
    description: Dict[str, Any]
    confidence: float
    model_used: str
    tokens_used: int
    processing_time_seconds: float
    passes_completed: int
    validation_status: str
    corrections_applied: List[str] = field(default_factory=list)
    cache_hit: bool = False
    raw_response: Optional[str] = None


# Model mappings by tier and provider
MODEL_MAPPINGS = {
    "anthropic": {
        ModelTier.FAST: "claude-haiku-4.5",
        ModelTier.BALANCED: "claude-sonnet-4.5",
        ModelTier.BEST: "claude-opus-4.5"
    },
    "openai": {
        ModelTier.FAST: "gpt-5-nano",
        ModelTier.BALANCED: "gpt-5",
        ModelTier.BEST: "gpt-5.2"
    },
    "google": {
        ModelTier.FAST: "gemini-3-flash",
        ModelTier.BALANCED: "gemini-3-pro",
        ModelTier.BEST: "gemini-3-pro"
    }
}


class AnalysisCache:
    """
    Cache for analysis results to avoid redundant API calls.
    """

    def __init__(self, cache_dir: Optional[Path] = None):
        """Initialize cache."""
        self.cache_dir = cache_dir or Path.home() / ".pptx_extractor_cache"
        self.cache_dir.mkdir(parents=True, exist_ok=True)
        self.memory_cache: Dict[str, AnalysisResult] = {}

    def _get_image_hash(self, image_path: Path) -> str:
        """Generate hash for an image file."""
        with open(image_path, 'rb') as f:
            return hashlib.md5(f.read()).hexdigest()

    def _get_cache_key(self, image_path: Path, config: AnalysisConfig) -> str:
        """Generate cache key from image and config."""
        image_hash = self._get_image_hash(image_path)
        config_hash = hashlib.md5(
            f"{config.mode.value}_{config.model_tier.value}_{config.num_passes}".encode()
        ).hexdigest()[:8]
        return f"{image_hash}_{config_hash}"

    def get(self, image_path: Path, config: AnalysisConfig) -> Optional[AnalysisResult]:
        """Get cached result if available."""
        key = self._get_cache_key(image_path, config)

        # Check memory cache first
        if key in self.memory_cache:
            result = self.memory_cache[key]
            result.cache_hit = True
            return result

        # Check disk cache
        cache_file = self.cache_dir / f"{key}.json"
        if cache_file.exists():
            try:
                with open(cache_file, 'r') as f:
                    data = json.load(f)
                result = AnalysisResult(**data)
                result.cache_hit = True
                self.memory_cache[key] = result
                return result
            except Exception as e:
                logger.warning(f"Failed to load cache: {e}")

        return None

    def set(self, image_path: Path, config: AnalysisConfig, result: AnalysisResult):
        """Store result in cache."""
        key = self._get_cache_key(image_path, config)

        # Memory cache
        self.memory_cache[key] = result

        # Disk cache
        cache_file = self.cache_dir / f"{key}.json"
        try:
            with open(cache_file, 'w') as f:
                json.dump(asdict(result), f, indent=2)
        except Exception as e:
            logger.warning(f"Failed to save cache: {e}")

    def clear(self):
        """Clear all caches."""
        self.memory_cache.clear()
        for f in self.cache_dir.glob("*.json"):
            f.unlink()


class OutputValidator:
    """
    Validates and corrects LLM output for slide descriptions.
    """

    # Required fields in output
    REQUIRED_FIELDS = ["slide_dimensions", "background", "elements"]

    # Valid ranges
    VALID_RANGES = {
        "left_inches": (0, 14),
        "top_inches": (0, 8),
        "width_inches": (0, 14),
        "height_inches": (0, 8),
        "font_size_pt": (6, 200),
        "border_width_pt": (0, 20),
        "rotation_degrees": (-360, 360),
        "transparency": (0, 1),
        "z_order": (0, 1000)
    }

    def validate(self, description: Dict[str, Any]) -> Tuple[bool, List[str]]:
        """
        Validate a slide description.

        Returns:
            Tuple of (is_valid, list_of_errors)
        """
        errors = []

        # Check required fields
        for field in self.REQUIRED_FIELDS:
            if field not in description:
                errors.append(f"Missing required field: {field}")

        # Validate slide dimensions
        if "slide_dimensions" in description:
            dims = description["slide_dimensions"]
            if not 10 < dims.get("width_inches", 0) < 15:
                errors.append(f"Invalid slide width: {dims.get('width_inches')}")
            if not 5 < dims.get("height_inches", 0) < 10:
                errors.append(f"Invalid slide height: {dims.get('height_inches')}")

        # Validate elements
        if "elements" in description:
            for i, elem in enumerate(description["elements"]):
                elem_errors = self._validate_element(elem, i)
                errors.extend(elem_errors)

        # Validate colors
        color_errors = self._validate_colors(description)
        errors.extend(color_errors)

        return len(errors) == 0, errors

    def _validate_element(self, element: Dict[str, Any], index: int) -> List[str]:
        """Validate a single element."""
        errors = []
        prefix = f"Element {index}"

        # Check required element fields
        if "type" not in element:
            errors.append(f"{prefix}: Missing type")

        # Validate position
        if "position" in element:
            pos = element["position"]
            for field, (min_val, max_val) in self.VALID_RANGES.items():
                if field in pos:
                    val = pos[field]
                    if not min_val <= val <= max_val:
                        errors.append(f"{prefix}: {field}={val} out of range [{min_val}, {max_val}]")

        return errors

    def _validate_colors(self, description: Dict[str, Any]) -> List[str]:
        """Validate all colors are valid hex format."""
        errors = []

        def check_color(path: str, value: Any):
            if isinstance(value, str) and value.startswith('#'):
                # Check hex format
                hex_part = value[1:]
                if len(hex_part) not in [6, 8]:
                    errors.append(f"{path}: Invalid color format {value}")
                try:
                    int(hex_part, 16)
                except ValueError:
                    errors.append(f"{path}: Invalid hex color {value}")

        def walk(obj, path=""):
            if isinstance(obj, dict):
                for k, v in obj.items():
                    new_path = f"{path}.{k}" if path else k
                    if "color" in k.lower():
                        check_color(new_path, v)
                    walk(v, new_path)
            elif isinstance(obj, list):
                for i, item in enumerate(obj):
                    walk(item, f"{path}[{i}]")

        walk(description)
        return errors

    def auto_correct(self, description: Dict[str, Any]) -> Tuple[Dict[str, Any], List[str]]:
        """
        Automatically correct common errors in the description.

        Returns:
            Tuple of (corrected_description, list_of_corrections)
        """
        corrections = []
        desc = json.loads(json.dumps(description))  # Deep copy

        # Fix slide dimensions if missing or invalid
        if "slide_dimensions" not in desc:
            desc["slide_dimensions"] = {
                "width_inches": 13.333,
                "height_inches": 7.5,
                "aspect_ratio": "16:9"
            }
            corrections.append("Added default slide dimensions")

        # Fix elements
        if "elements" in desc:
            for i, elem in enumerate(desc["elements"]):
                # Clamp positions to valid ranges
                if "position" in elem:
                    pos = elem["position"]
                    for field, (min_val, max_val) in self.VALID_RANGES.items():
                        if field in pos:
                            original = pos[field]
                            pos[field] = max(min_val, min(max_val, pos[field]))
                            if pos[field] != original:
                                corrections.append(
                                    f"Element {i}: Clamped {field} from {original} to {pos[field]}"
                                )

                # Add missing type
                if "type" not in elem:
                    elem["type"] = "shape"
                    corrections.append(f"Element {i}: Added default type 'shape'")

                # Add missing z_order
                if "z_order" not in elem:
                    elem["z_order"] = i + 1
                    corrections.append(f"Element {i}: Added z_order={i + 1}")

        # Fix color format (ensure # prefix)
        def fix_colors(obj):
            if isinstance(obj, dict):
                for k, v in obj.items():
                    if "color" in k.lower() and isinstance(v, str):
                        if len(v) == 6 and not v.startswith('#'):
                            try:
                                int(v, 16)
                                obj[k] = f"#{v}"
                                corrections.append(f"Fixed color format: {v} -> #{v}")
                            except ValueError:
                                pass
                    else:
                        fix_colors(v)
            elif isinstance(obj, list):
                for item in obj:
                    fix_colors(item)

        fix_colors(desc)

        return desc, corrections


class EnhancedAnalyzer:
    """
    Enhanced slide analyzer with optimized LLM vision analysis.
    """

    def __init__(self, config: Optional[AnalysisConfig] = None):
        """Initialize the enhanced analyzer."""
        self.config = config or AnalysisConfig()
        self.cache = AnalysisCache(self.config.cache_dir)
        self.validator = OutputValidator()
        self._preprocessor = None

    @property
    def preprocessor(self):
        """Lazy load preprocessor."""
        if self._preprocessor is None:
            from pptx_extractor.image_preprocessor import ImagePreprocessor, PreprocessingConfig
            preprocess_config = PreprocessingConfig(
                add_grid=self.config.add_measurement_grid,
                add_rulers=self.config.add_rulers,
                normalize_contrast=True,
                normalize_brightness=True,
                scale_to_optimal=True
            )
            self._preprocessor = ImagePreprocessor(preprocess_config)
        return self._preprocessor

    def analyze(
        self,
        image_path: Path,
        config: Optional[AnalysisConfig] = None
    ) -> AnalysisResult:
        """
        Analyze a slide image with optimized LLM vision.

        Args:
            image_path: Path to the slide image
            config: Optional override configuration

        Returns:
            AnalysisResult with description and metadata
        """
        cfg = config or self.config
        image_path = Path(image_path)
        start_time = time.time()

        # Check cache
        if cfg.use_cache:
            cached = self.cache.get(image_path, cfg)
            if cached:
                logger.info(f"Cache hit for {image_path.name}")
                return cached

        # Determine analysis approach based on mode
        if cfg.mode == AnalysisMode.FAST:
            result = self._analyze_fast(image_path, cfg)
        elif cfg.mode == AnalysisMode.STANDARD:
            result = self._analyze_standard(image_path, cfg)
        elif cfg.mode == AnalysisMode.THOROUGH:
            result = self._analyze_thorough(image_path, cfg)
        elif cfg.mode == AnalysisMode.ULTRA:
            result = self._analyze_ultra(image_path, cfg)
        else:
            result = self._analyze_standard(image_path, cfg)

        result.processing_time_seconds = time.time() - start_time

        # Cache result
        if cfg.use_cache:
            self.cache.set(image_path, cfg, result)

        return result

    def _analyze_fast(self, image_path: Path, config: AnalysisConfig) -> AnalysisResult:
        """Fast single-pass analysis with basic model."""
        model = self._get_model(config.preferred_provider, ModelTier.FAST)
        return self._single_pass_analysis(image_path, model, config, use_v2_prompt=False)

    def _analyze_standard(self, image_path: Path, config: AnalysisConfig) -> AnalysisResult:
        """Standard single-pass analysis with balanced model."""
        model = self._get_model(config.preferred_provider, config.model_tier)
        return self._single_pass_analysis(image_path, model, config, use_v2_prompt=True)

    def _analyze_thorough(self, image_path: Path, config: AnalysisConfig) -> AnalysisResult:
        """Multi-pass analysis with validation and refinement."""
        model = self._get_model(config.preferred_provider, ModelTier.BEST)

        # First pass with preprocessing
        result = self._single_pass_analysis(image_path, model, config, use_v2_prompt=True)

        # Refine if confidence is low
        if result.confidence < config.confidence_threshold and config.refine_on_low_confidence:
            result = self._refine_analysis(image_path, result, model, config)

        return result

    def _analyze_ultra(self, image_path: Path, config: AnalysisConfig) -> AnalysisResult:
        """Ultra analysis with extended thinking and multi-pass."""
        model = self._get_model(config.preferred_provider, ModelTier.BEST)

        # Use extended thinking for Claude
        if config.preferred_provider == "anthropic":
            result = self._analyze_with_extended_thinking(image_path, model, config)
        else:
            # Fall back to thorough for non-Anthropic
            result = self._analyze_thorough(image_path, config)

        return result

    def _single_pass_analysis(
        self,
        image_path: Path,
        model: str,
        config: AnalysisConfig,
        use_v2_prompt: bool = True
    ) -> AnalysisResult:
        """Perform single-pass analysis."""
        # Preprocess image if enabled
        if config.preprocess_image:
            preprocessed_path, _ = self.preprocessor.preprocess(image_path)
            analysis_image = preprocessed_path
        else:
            analysis_image = image_path

        # Load prompt
        prompt_name = "description_prompt_v2" if use_v2_prompt else "description_prompt"
        prompt = self._load_prompt(prompt_name)

        # Call LLM
        response, tokens = self._call_llm_vision(analysis_image, prompt, model, config)

        # Parse response
        description = self._parse_json_response(response)

        # Extract confidence
        confidence = description.get("metadata", {}).get("analysis_confidence", 0.7)

        # Validate and correct
        if config.validate_output:
            is_valid, errors = self.validator.validate(description)
            if not is_valid and config.auto_correct_errors:
                description, corrections = self.validator.auto_correct(description)
            else:
                corrections = []
        else:
            corrections = []

        return AnalysisResult(
            description=description,
            confidence=confidence,
            model_used=model,
            tokens_used=tokens,
            processing_time_seconds=0,  # Set by caller
            passes_completed=1,
            validation_status="valid" if config.validate_output else "unchecked",
            corrections_applied=corrections,
            raw_response=response
        )

    def _refine_analysis(
        self,
        image_path: Path,
        initial_result: AnalysisResult,
        model: str,
        config: AnalysisConfig
    ) -> AnalysisResult:
        """Refine analysis through additional passes."""
        current_description = initial_result.description
        total_tokens = initial_result.tokens_used
        passes = 1

        for _ in range(config.num_passes - 1):
            # Create refinement prompt
            refinement_prompt = self._create_refinement_prompt(current_description)

            # Preprocess if needed
            if config.preprocess_image:
                preprocessed_path, _ = self.preprocessor.preprocess(image_path)
                analysis_image = preprocessed_path
            else:
                analysis_image = image_path

            # Call LLM for refinement
            response, tokens = self._call_llm_vision(
                analysis_image, refinement_prompt, model, config
            )
            total_tokens += tokens
            passes += 1

            # Parse refined description
            refined = self._parse_json_response(response)

            # Check if confidence improved
            new_confidence = refined.get("metadata", {}).get("analysis_confidence", 0)
            if new_confidence > current_description.get("metadata", {}).get("analysis_confidence", 0):
                current_description = refined

        # Final validation
        corrections = []
        if config.validate_output:
            is_valid, errors = self.validator.validate(current_description)
            if not is_valid and config.auto_correct_errors:
                current_description, corrections = self.validator.auto_correct(current_description)

        return AnalysisResult(
            description=current_description,
            confidence=current_description.get("metadata", {}).get("analysis_confidence", 0.7),
            model_used=model,
            tokens_used=total_tokens,
            processing_time_seconds=0,
            passes_completed=passes,
            validation_status="valid" if config.validate_output else "unchecked",
            corrections_applied=corrections
        )

    def _analyze_with_extended_thinking(
        self,
        image_path: Path,
        model: str,
        config: AnalysisConfig
    ) -> AnalysisResult:
        """
        Analyze using Claude's extended thinking for complex reasoning.
        """
        from pptx_generator.modules.llm_provider import load_env_file
        load_env_file()

        import anthropic
        from pptx_extractor.comparator import image_to_base64

        # Preprocess
        if config.preprocess_image:
            preprocessed_path, _ = self.preprocessor.preprocess(image_path)
            image_b64 = image_to_base64(preprocessed_path)
        else:
            image_b64 = image_to_base64(image_path)

        # Load prompt with thinking instructions
        prompt = self._load_prompt("description_prompt_v2")
        thinking_prompt = f"""Think through this slide analysis step by step:

1. First, identify the overall layout and structure
2. Then, catalog each visual element systematically
3. Measure positions relative to the slide dimensions
4. Extract exact colors and typography
5. Finally, compile everything into the JSON format

{prompt}"""

        api_key = os.environ.get("ANTHROPIC_API_KEY")
        client = anthropic.Anthropic(api_key=api_key)

        # Determine model ID
        from pptx_generator.modules.llm_provider import AVAILABLE_MODELS
        model_config = AVAILABLE_MODELS.get(model)
        model_id = model_config.model_id if model_config else "claude-sonnet-4-5-20250514"

        try:
            # Use extended thinking with streaming
            response = client.messages.create(
                model=model_id,
                max_tokens=16000,
                thinking={
                    "type": "enabled",
                    "budget_tokens": config.thinking_budget_tokens
                },
                messages=[
                    {
                        "role": "user",
                        "content": [
                            {
                                "type": "image",
                                "source": {
                                    "type": "base64",
                                    "media_type": "image/png",
                                    "data": image_b64
                                }
                            },
                            {
                                "type": "text",
                                "text": thinking_prompt
                            }
                        ]
                    }
                ]
            )

            # Extract text response (skip thinking blocks)
            response_text = ""
            for block in response.content:
                if block.type == "text":
                    response_text = block.text
                    break

            tokens = response.usage.input_tokens + response.usage.output_tokens

        except Exception as e:
            logger.warning(f"Extended thinking failed, falling back to standard: {e}")
            return self._analyze_thorough(image_path, config)

        # Parse response
        description = self._parse_json_response(response_text)
        confidence = description.get("metadata", {}).get("analysis_confidence", 0.85)

        # Validate
        corrections = []
        if config.validate_output:
            is_valid, errors = self.validator.validate(description)
            if not is_valid and config.auto_correct_errors:
                description, corrections = self.validator.auto_correct(description)

        return AnalysisResult(
            description=description,
            confidence=confidence,
            model_used=model,
            tokens_used=tokens,
            processing_time_seconds=0,
            passes_completed=1,
            validation_status="valid" if config.validate_output else "unchecked",
            corrections_applied=corrections,
            raw_response=response_text
        )

    def analyze_batch(
        self,
        image_paths: List[Path],
        config: Optional[AnalysisConfig] = None
    ) -> List[AnalysisResult]:
        """
        Analyze multiple images in parallel.

        Args:
            image_paths: List of image paths
            config: Optional configuration

        Returns:
            List of AnalysisResults in same order as inputs
        """
        cfg = config or self.config
        results = [None] * len(image_paths)

        with ThreadPoolExecutor(max_workers=cfg.max_parallel) as executor:
            future_to_index = {
                executor.submit(self.analyze, path, cfg): i
                for i, path in enumerate(image_paths)
            }

            for future in as_completed(future_to_index):
                index = future_to_index[future]
                try:
                    results[index] = future.result()
                except Exception as e:
                    logger.error(f"Failed to analyze image {index}: {e}")
                    results[index] = AnalysisResult(
                        description={"error": str(e)},
                        confidence=0,
                        model_used="none",
                        tokens_used=0,
                        processing_time_seconds=0,
                        passes_completed=0,
                        validation_status="error"
                    )

        return results

    def _get_model(self, provider: str, tier: ModelTier) -> str:
        """Get model name for provider and tier."""
        return MODEL_MAPPINGS.get(provider, MODEL_MAPPINGS["google"]).get(
            tier, "gemini-3-flash"
        )

    def _load_prompt(self, prompt_name: str) -> str:
        """Load a prompt template."""
        prompt_dir = Path(__file__).parent / "prompts"
        prompt_path = prompt_dir / f"{prompt_name}.txt"

        if not prompt_path.exists():
            # Fall back to original prompt
            prompt_path = prompt_dir / "description_prompt.txt"

        with open(prompt_path, 'r', encoding='utf-8') as f:
            return f.read()

    def _create_refinement_prompt(self, current_description: Dict[str, Any]) -> str:
        """Create a prompt for refining an existing description."""
        return f"""Review and improve this slide description. Focus on:
1. Verifying all positions are accurate to 0.1 inch
2. Confirming all colors are exact hex values
3. Ensuring no elements are missing
4. Checking typography details

Current description:
```json
{json.dumps(current_description, indent=2)}
```

Provide an improved JSON description with higher confidence. Output ONLY JSON."""

    def _call_llm_vision(
        self,
        image_path: Path,
        prompt: str,
        model: str,
        config: AnalysisConfig
    ) -> Tuple[str, int]:
        """Call LLM with vision capability."""
        from pptx_extractor.descriptor import (
            _call_llm_vision,
            _call_anthropic_vision
        )
        from pptx_extractor.comparator import image_to_base64

        image_b64 = image_to_base64(image_path)

        # Use the descriptor's LLM vision call
        try:
            result = _call_llm_vision(prompt, image_b64, model)
            # Estimate tokens (rough approximation)
            tokens = len(prompt) // 4 + len(str(result)) // 4
            return json.dumps(result), tokens
        except Exception as e:
            logger.error(f"LLM vision call failed: {e}")
            raise

    def _parse_json_response(self, response_text: str) -> Dict[str, Any]:
        """Parse JSON from response text."""
        import re

        # Try to extract JSON from markdown code blocks
        json_match = re.search(r'```(?:json)?\s*([\s\S]*?)\s*```', response_text)
        if json_match:
            json_str = json_match.group(1)
        else:
            json_str = response_text.strip()

        try:
            return json.loads(json_str)
        except json.JSONDecodeError as e:
            logger.error(f"JSON parse error: {e}")
            return {"parse_error": str(e), "raw": response_text[:1000]}


# Convenience functions
def analyze_slide(
    image_path: Path,
    mode: str = "standard",
    provider: str = "anthropic"
) -> AnalysisResult:
    """
    Quick function to analyze a slide image.

    Args:
        image_path: Path to slide image
        mode: Analysis mode (fast, standard, thorough, ultra)
        provider: LLM provider (anthropic, openai, google)

    Returns:
        AnalysisResult
    """
    config = AnalysisConfig(
        mode=AnalysisMode(mode),
        preferred_provider=provider
    )
    analyzer = EnhancedAnalyzer(config)
    return analyzer.analyze(image_path)


def analyze_slides_batch(
    image_paths: List[Path],
    mode: str = "standard",
    max_parallel: int = 3
) -> List[AnalysisResult]:
    """
    Analyze multiple slide images in parallel.

    Args:
        image_paths: List of image paths
        mode: Analysis mode
        max_parallel: Maximum parallel requests

    Returns:
        List of AnalysisResults
    """
    config = AnalysisConfig(
        mode=AnalysisMode(mode),
        max_parallel=max_parallel
    )
    analyzer = EnhancedAnalyzer(config)
    return analyzer.analyze_batch(image_paths)


if __name__ == "__main__":
    import sys

    logging.basicConfig(level=logging.INFO)

    if len(sys.argv) > 1:
        image_path = Path(sys.argv[1])
        mode = sys.argv[2] if len(sys.argv) > 2 else "standard"

        print(f"Analyzing {image_path} with mode={mode}...")
        result = analyze_slide(image_path, mode=mode)

        print(f"\nConfidence: {result.confidence:.2f}")
        print(f"Model: {result.model_used}")
        print(f"Tokens: {result.tokens_used}")
        print(f"Time: {result.processing_time_seconds:.2f}s")
        print(f"Passes: {result.passes_completed}")
        print(f"Validation: {result.validation_status}")

        if result.corrections_applied:
            print(f"Corrections: {result.corrections_applied}")

        print(f"\nDescription:\n{json.dumps(result.description, indent=2)}")
    else:
        print("Usage: python enhanced_analyzer.py <image_path> [mode]")
        print("Modes: fast, standard, thorough, ultra")
