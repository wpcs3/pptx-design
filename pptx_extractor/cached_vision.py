"""
Cached Vision Module

Provides vision API calls with prompt caching for significant cost reduction.
Prompt caching reduces the cost of repeated system prompts from $3/MTok to $0.30/MTok
(90% savings on prompt tokens).

Usage:
    client = CachedVisionClient()
    result = client.analyze_image(image_path, mode='detailed')
"""

import base64
import hashlib
import json
import logging
import os
from pathlib import Path
from typing import Optional, Dict, Any, List, Literal
from dataclasses import dataclass

logger = logging.getLogger(__name__)


# System prompts for caching - these remain constant across calls
SYSTEM_PROMPTS = {
    "categorize": """You are a PowerPoint slide classifier. Your task is to quickly categorize slides by type and layout.

Always output valid JSON with these fields:
- slide_type: one of [title, section_divider, content, data_chart, comparison, timeline, process, quote, image_focus, blank]
- layout_category: one of [single_column, two_column, grid, centered, asymmetric]
- element_count: integer count of visual elements
- has_chart: boolean
- has_table: boolean
- has_image: boolean (photos/graphics, not icons)
- primary_colors: array of 2-3 hex colors
- complexity_score: 1-5 (1=simple, 5=complex)

Be concise. Output only the JSON object.""",

    "standard": """You are a PowerPoint slide analyzer. Extract complete specifications for recreating slides programmatically.

REFERENCE DIMENSIONS:
- Slide width: 13.333 inches (960 points)
- Slide height: 7.5 inches (540 points)
- 1 inch = 72 points

MEASUREMENT GUIDELINES:
- All positions in inches from top-left origin
- Font sizes in points (pt)
- Line weights in points
- Colors as hex (#RRGGBB)

OUTPUT STRUCTURE:
{
    "metadata": {"slide_type": "", "complexity_score": 1-5, "analysis_confidence": 0.0-1.0},
    "slide_dimensions": {"width_inches": 13.333, "height_inches": 7.5, "aspect_ratio": "16:9"},
    "background": {"type": "solid|gradient", "color": "#FFFFFF"},
    "elements": [
        {
            "id": "unique_id",
            "type": "textbox|shape|line|image|chart|table",
            "z_order": 1,
            "position": {"left_inches": 0, "top_inches": 0, "width_inches": 0, "height_inches": 0},
            "text_content": {
                "text": "full text",
                "paragraphs": [{"text": "", "alignment": "left|center|right", "runs": [...]}]
            },
            "shape_properties": {"fill": {"type": "solid|none", "color": "#"}, "border": {...}}
        }
    ],
    "color_palette": {"primary": "#", "secondary": "#", "background": "#", "text_primary": "#"},
    "typography_system": {"title_font": "", "body_font": "", "title_size_pt": 0, "body_size_pt": 0},
    "layout_grid": {"columns": 1, "margins": {"left_inches": 0, "right_inches": 0, "top_inches": 0, "bottom_inches": 0}},
    "design_notes": "Brief description of the slide design and purpose"
}

Be precise with measurements. Output only valid JSON.""",

    "detailed": """You are a PowerPoint slide analyzer creating specifications for automated slide generation systems.

REFERENCE DIMENSIONS:
- Slide width: 13.333 inches (960 points)
- Slide height: 7.5 inches (540 points)
- 1 inch = 72 points

Your output has TWO sections:

SECTION 1 - SLIDE SPECIFICATION:
Complete technical specification for recreating the slide pixel-perfectly.

SECTION 2 - GENERATOR HINTS:
Metadata to help automated systems reuse this template:
- template_category: What type of slide template this represents
- reusability_score: 1-5 how reusable is this as a template
- content_placeholders: Which elements contain variable content vs fixed design
- chrome_elements: Decorative/structural elements that stay constant
- style_tokens: Key style characteristics for matching
- layout_zones: Named regions for content placement

OUTPUT STRUCTURE:
{
    "metadata": {"slide_type": "", "complexity_score": 1-5, "analysis_confidence": 0.0-1.0},
    "slide_dimensions": {"width_inches": 13.333, "height_inches": 7.5, "aspect_ratio": "16:9"},
    "background": {"type": "solid|gradient", "color": "#FFFFFF"},
    "elements": [...],
    "color_palette": {...},
    "typography_system": {...},
    "layout_grid": {...},
    "design_notes": "",
    "generator_hints": {
        "template_category": "section_divider|title_slide|content_slide|data_visualization|comparison|process_flow",
        "reusability_score": 1-5,
        "content_placeholders": [
            {"id": "element_id", "purpose": "main_title|subtitle|section_number|body_text|data_label|footer", "editable": true, "sample_content": ""}
        ],
        "chrome_elements": ["element_ids"],
        "style_tokens": {
            "primary_font": "",
            "heading_weight": "bold|normal",
            "color_scheme": "monochrome|corporate|colorful|dark|light",
            "visual_style": "minimal|classic|modern|bold"
        },
        "layout_zones": [
            {"name": "header|content|sidebar|footer", "bounds": {"left": 0, "top": 0, "width": 0, "height": 0}, "purpose": ""}
        ]
    }
}

Be precise with measurements. Identify all content placeholders accurately. Output only valid JSON."""
}


@dataclass
class CacheStats:
    """Statistics about cache usage."""
    cache_writes: int = 0
    cache_reads: int = 0
    total_input_tokens: int = 0
    total_output_tokens: int = 0
    estimated_savings: float = 0.0  # In dollars


class CachedVisionClient:
    """
    Vision client with prompt caching support.

    Prompt caching stores the system prompt server-side, reducing costs for
    repeated calls with the same prompt from $3/MTok to $0.30/MTok.
    """

    def __init__(
        self,
        model: str = "claude-haiku-4-5",
        cache_ttl: Literal["5m", "1h"] = "5m"
    ):
        """
        Initialize cached vision client.

        Args:
            model: Model to use (Anthropic models only - uses prompt caching)
            cache_ttl: Cache time-to-live ("5m" for 5 minutes, "1h" for 1 hour)

        Note: This client is Anthropic-specific for prompt caching benefits.
              Default is Haiku for cost efficiency.
        """
        self.model = model
        self.cache_ttl = cache_ttl
        self._client = None
        self.stats = CacheStats()

        # Model ID mapping
        self._model_ids = {
            "claude-opus-4.5": "claude-opus-4-5",
            "claude-sonnet-4.5": "claude-sonnet-4-5",
            "claude-haiku-4.5": "claude-haiku-4-5"
        }

    def _get_client(self):
        """Lazy initialization of Anthropic client."""
        if self._client is None:
            import anthropic
            from pptx_generator.modules.llm_provider import load_env_file
            load_env_file()
            self._client = anthropic.Anthropic()
        return self._client

    def _get_model_id(self) -> str:
        """Get the API model ID."""
        return self._model_ids.get(self.model, self.model)

    def _image_to_base64(self, image_path: Path) -> str:
        """Convert image to base64 string."""
        with open(image_path, 'rb') as f:
            return base64.standard_b64encode(f.read()).decode('utf-8')

    def _get_media_type(self, image_path: Path) -> str:
        """Get media type from file extension."""
        ext = image_path.suffix.lower()
        types = {
            '.png': 'image/png',
            '.jpg': 'image/jpeg',
            '.jpeg': 'image/jpeg',
            '.gif': 'image/gif',
            '.webp': 'image/webp'
        }
        return types.get(ext, 'image/png')

    def analyze_image(
        self,
        image_path: Path,
        mode: Literal["categorize", "standard", "detailed"] = "standard",
        additional_context: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Analyze a slide image with prompt caching.

        Args:
            image_path: Path to the image file
            mode: Analysis mode (affects prompt and output detail)
            additional_context: Optional additional instructions

        Returns:
            Parsed JSON result from the analysis
        """
        client = self._get_client()
        image_path = Path(image_path)

        # Get system prompt for caching
        system_prompt = SYSTEM_PROMPTS[mode]

        # Build user message with image
        image_b64 = self._image_to_base64(image_path)
        media_type = self._get_media_type(image_path)

        user_content = [
            {
                "type": "image",
                "source": {
                    "type": "base64",
                    "media_type": media_type,
                    "data": image_b64
                }
            },
            {
                "type": "text",
                "text": "Analyze this PowerPoint slide." +
                       (f"\n\nAdditional context: {additional_context}" if additional_context else "")
            }
        ]

        try:
            # Make API call with prompt caching
            response = client.messages.create(
                model=self._get_model_id(),
                max_tokens=4096,
                system=[
                    {
                        "type": "text",
                        "text": system_prompt,
                        "cache_control": {"type": "ephemeral"}
                    }
                ],
                messages=[
                    {"role": "user", "content": user_content}
                ]
            )

            # Track cache statistics
            usage = response.usage
            self.stats.total_input_tokens += usage.input_tokens
            self.stats.total_output_tokens += usage.output_tokens

            # Check for cache usage (if available in response)
            if hasattr(usage, 'cache_creation_input_tokens'):
                self.stats.cache_writes += usage.cache_creation_input_tokens
            if hasattr(usage, 'cache_read_input_tokens'):
                self.stats.cache_reads += usage.cache_read_input_tokens
                # Calculate savings: cache reads cost 0.1x vs 1x for regular input
                savings_per_token = 0.9 * (3.0 / 1_000_000)  # $3/MTok * 0.9 savings
                self.stats.estimated_savings += usage.cache_read_input_tokens * savings_per_token

            # Parse response
            content = response.content[0].text

            # Try to extract JSON from response
            result = self._parse_json_response(content)
            result['_usage'] = {
                'input_tokens': usage.input_tokens,
                'output_tokens': usage.output_tokens,
                'cache_read_tokens': getattr(usage, 'cache_read_input_tokens', 0),
                'cache_write_tokens': getattr(usage, 'cache_creation_input_tokens', 0)
            }

            return result

        except Exception as e:
            logger.error(f"Vision analysis failed: {e}")
            raise

    def _parse_json_response(self, content: str) -> Dict[str, Any]:
        """Parse JSON from LLM response, handling common issues."""
        # Try direct parse first
        try:
            return json.loads(content)
        except json.JSONDecodeError:
            pass

        # Try to extract JSON from markdown code blocks
        import re
        json_match = re.search(r'```(?:json)?\s*\n?(.*?)\n?```', content, re.DOTALL)
        if json_match:
            try:
                return json.loads(json_match.group(1))
            except json.JSONDecodeError:
                pass

        # Try to find JSON object in text
        json_match = re.search(r'\{.*\}', content, re.DOTALL)
        if json_match:
            try:
                return json.loads(json_match.group(0))
            except json.JSONDecodeError:
                pass

        # Return raw content wrapped
        logger.warning("Could not parse JSON from response")
        return {"raw_response": content, "parse_error": True}

    def analyze_batch(
        self,
        image_paths: List[Path],
        mode: Literal["categorize", "standard", "detailed"] = "standard",
        progress_callback=None
    ) -> List[Dict[str, Any]]:
        """
        Analyze multiple images with prompt caching.

        The first call caches the system prompt, subsequent calls read from cache.

        Args:
            image_paths: List of image paths
            mode: Analysis mode
            progress_callback: Optional callback(completed, total)

        Returns:
            List of analysis results
        """
        results = []

        for i, path in enumerate(image_paths):
            try:
                result = self.analyze_image(path, mode)
                result['_source_image'] = str(path)
                results.append(result)
            except Exception as e:
                logger.error(f"Failed to analyze {path}: {e}")
                results.append({
                    '_source_image': str(path),
                    '_error': str(e)
                })

            if progress_callback:
                progress_callback(i + 1, len(image_paths))

        return results

    def get_stats(self) -> Dict[str, Any]:
        """Get cache usage statistics."""
        return {
            'cache_writes': self.stats.cache_writes,
            'cache_reads': self.stats.cache_reads,
            'total_input_tokens': self.stats.total_input_tokens,
            'total_output_tokens': self.stats.total_output_tokens,
            'estimated_savings_usd': round(self.stats.estimated_savings, 4),
            'cache_hit_rate': (
                self.stats.cache_reads /
                (self.stats.cache_reads + self.stats.cache_writes)
                if (self.stats.cache_reads + self.stats.cache_writes) > 0
                else 0
            )
        }

    def reset_stats(self):
        """Reset cache statistics."""
        self.stats = CacheStats()


def estimate_batch_cost(
    num_slides: int,
    mode: Literal["categorize", "standard", "detailed"] = "standard",
    model: str = "claude-haiku-4.5",
    use_batch_api: bool = True,
    use_prompt_cache: bool = True,
    avg_image_tokens: int = 2765  # 1920x1080 / 750
) -> Dict[str, Any]:
    """
    Estimate the cost of processing a batch of slides.

    Args:
        num_slides: Number of slides to process
        mode: Analysis mode
        model: Model to use
        use_batch_api: Whether using Batch API (50% discount)
        use_prompt_cache: Whether using prompt caching
        avg_image_tokens: Average tokens per image

    Returns:
        Cost breakdown dictionary
    """
    # Model pricing (per million tokens)
    pricing = {
        "claude-opus-4.5": {"input": 5.0, "output": 25.0},
        "claude-sonnet-4.5": {"input": 3.0, "output": 15.0},
        "claude-haiku-4.5": {"input": 1.0, "output": 5.0}
    }

    # Output tokens by mode
    output_tokens_by_mode = {
        "categorize": 150,
        "standard": 2500,
        "detailed": 3500
    }

    # System prompt tokens (approximate)
    system_prompt_tokens = {
        "categorize": 200,
        "standard": 800,
        "detailed": 1200
    }

    model_pricing = pricing.get(model, pricing["claude-haiku-4.5"])
    output_tokens = output_tokens_by_mode[mode]
    prompt_tokens = system_prompt_tokens[mode]

    # Calculate base costs
    total_input_tokens = num_slides * (avg_image_tokens + prompt_tokens)
    total_output_tokens = num_slides * output_tokens

    base_input_cost = (total_input_tokens / 1_000_000) * model_pricing["input"]
    base_output_cost = (total_output_tokens / 1_000_000) * model_pricing["output"]

    # Apply discounts
    input_cost = base_input_cost
    output_cost = base_output_cost

    savings = {}

    # Batch API discount (50% off everything)
    if use_batch_api:
        input_cost *= 0.5
        output_cost *= 0.5
        savings["batch_api"] = (base_input_cost + base_output_cost) * 0.5

    # Prompt caching discount (90% off prompt tokens after first call)
    if use_prompt_cache and num_slides > 1:
        # First call pays full price, rest get 90% discount on prompt
        cached_prompt_tokens = (num_slides - 1) * prompt_tokens
        prompt_savings = (cached_prompt_tokens / 1_000_000) * model_pricing["input"] * 0.9
        if use_batch_api:
            prompt_savings *= 0.5  # Batch discount applies first
        input_cost -= prompt_savings
        savings["prompt_cache"] = prompt_savings

    total_cost = input_cost + output_cost

    return {
        "num_slides": num_slides,
        "mode": mode,
        "model": model,
        "total_input_tokens": total_input_tokens,
        "total_output_tokens": total_output_tokens,
        "base_cost_usd": round(base_input_cost + base_output_cost, 2),
        "final_cost_usd": round(total_cost, 2),
        "savings_breakdown": {k: round(v, 2) for k, v in savings.items()},
        "total_savings_usd": round(sum(savings.values()), 2),
        "savings_percentage": round(
            (1 - total_cost / (base_input_cost + base_output_cost)) * 100, 1
        ) if (base_input_cost + base_output_cost) > 0 else 0
    }
