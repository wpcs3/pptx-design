"""
LLM-Powered Content Generator

Generates presentation content using LLMs with awareness of:
- ComponentLibrary patterns and styles
- Slide type templates
- Real estate domain knowledge
- User context and preferences

Phase 1 Enhancement (2025-12-29):
- Integrated tone and verbosity controls
- Updated to use LiteLLM unified interface
- Added async generation support
"""

import json
import logging
from pathlib import Path
from typing import Optional, List, Dict, Any
from dataclasses import dataclass

from .llm_provider import (
    LLMManager, LLMResponse, AVAILABLE_MODELS,
    GenerationConfig, Tone, Verbosity
)
from .component_library import ComponentLibrary
from .library_enhancer import LibraryEnhancer

logger = logging.getLogger(__name__)


# System prompts for different content types
SYSTEM_PROMPTS = {
    "outline": """You are an expert presentation strategist specializing in real estate investment presentations.
Your task is to create detailed presentation outlines that are professional, data-driven, and compelling.

When creating outlines:
- Structure content logically with clear sections
- Include specific data points and metrics placeholders
- Consider the target audience (institutional investors, LPs, etc.)
- Follow industry best practices for pitch decks

Output your response as valid JSON matching the requested schema.""",

    "slide_content": """You are an expert presentation content writer specializing in real estate investment materials.
Your task is to write compelling, professional content for presentation slides.

Guidelines:
- Be concise - bullet points should be 5-10 words max
- Use specific numbers and metrics when possible
- Maintain professional tone appropriate for institutional investors
- Focus on value propositions and differentiators

Output your response as valid JSON matching the requested schema.""",

    "executive_summary": """You are an expert at writing executive summaries for real estate investment presentations.
Create compelling summaries that:
- Lead with the key value proposition
- Include critical metrics (fund size, target returns, strategy)
- Highlight competitive advantages
- Create urgency and interest

Keep summaries to 3-5 bullet points, each 10-15 words maximum.""",

    "market_analysis": """You are a real estate market analyst creating content for investor presentations.
Generate market analysis content that:
- Uses current market data and trends
- Compares to relevant benchmarks
- Identifies opportunities and risks
- Supports the investment thesis

Be specific with data points and cite market sources where appropriate.""",

    "risk_factors": """You are a risk analyst for real estate investments.
Create risk factor content that:
- Identifies key risks clearly
- Provides mitigation strategies
- Is honest but not alarmist
- Follows SEC/regulatory guidelines for disclosure

Present risks professionally with corresponding mitigations."""
}


@dataclass
class GeneratedContent:
    """Container for generated content with metadata."""
    content: Dict[str, Any]
    model_used: str
    prompt_tokens: int
    completion_tokens: int
    raw_response: Optional[str] = None


class ContentGenerator:
    """
    LLM-powered content generator for presentations.

    Uses the ComponentLibrary to provide context about available
    patterns, styles, and examples to the LLM.

    Supports tone and verbosity controls for customizing output style.
    """

    def __init__(
        self,
        model: str = "claude-3.5-sonnet",
        library: Optional[ComponentLibrary] = None,
        enhancer: Optional[LibraryEnhancer] = None,
        tone: str = "professional",
        verbosity: str = "standard",
        use_litellm: bool = True
    ):
        """
        Initialize the content generator.

        Args:
            model: Model name from AVAILABLE_MODELS or direct LiteLLM model string
            library: Optional ComponentLibrary instance
            enhancer: Optional LibraryEnhancer instance
            tone: Content tone (default, professional, casual, sales_pitch, educational, executive)
            verbosity: Content verbosity (concise, standard, detailed)
            use_litellm: Whether to use LiteLLM unified interface (recommended)
        """
        # Create generation config
        self.generation_config = GenerationConfig(
            tone=Tone(tone),
            verbosity=Verbosity(verbosity)
        )

        self.llm = LLMManager(
            default_model=model,
            use_litellm=use_litellm,
            generation_config=self.generation_config
        )
        self.library = library or ComponentLibrary()
        self.enhancer = enhancer or LibraryEnhancer(self.library)

        # Load content patterns
        config_dir = Path(__file__).parent.parent / "config"
        patterns_path = config_dir / "content_patterns.json"

        if patterns_path.exists():
            with open(patterns_path, 'r') as f:
                self.content_patterns = json.load(f)
        else:
            self.content_patterns = {}

        logger.info(f"ContentGenerator initialized with model: {model}, tone: {tone}, verbosity: {verbosity}")

    def set_model(self, model: str) -> None:
        """Switch to a different model."""
        self.llm.set_model(model)
        logger.info(f"Switched to model: {model}")

    def set_tone(self, tone: str) -> None:
        """Set the content tone."""
        self.generation_config.tone = Tone(tone)
        self.llm.set_generation_config(self.generation_config)
        logger.info(f"Set tone to: {tone}")

    def set_verbosity(self, verbosity: str) -> None:
        """Set the content verbosity."""
        self.generation_config.verbosity = Verbosity(verbosity)
        self.llm.set_generation_config(self.generation_config)
        logger.info(f"Set verbosity to: {verbosity}")

    def get_library_context(self) -> str:
        """
        Build context string about available library components.

        This helps the LLM understand what visual elements are available.
        """
        context_parts = []

        # Domain tags available
        domain_stats = self.enhancer.get_domain_stats()
        if domain_stats:
            context_parts.append("Available domain categories for content:")
            for domain, count in sorted(domain_stats.items(), key=lambda x: -x[1])[:10]:
                context_parts.append(f"  - {domain}: {count} components")

        # Chart types available
        chart_info = self._get_chart_summary()
        if chart_info:
            context_parts.append("\nAvailable chart types:")
            for chart_type, count in chart_info.items():
                context_parts.append(f"  - {chart_type}: {count} examples")

        # Table structures
        table_info = self._get_table_summary()
        if table_info:
            context_parts.append("\nAvailable table structures:")
            for structure, count in list(table_info.items())[:5]:
                context_parts.append(f"  - {structure}: {count} examples")

        return "\n".join(context_parts)

    def _get_chart_summary(self) -> Dict[str, int]:
        """Get summary of available chart types."""
        charts = self.library.get_charts()
        type_counts = {}
        for chart in charts:
            chart_type = chart.get('chart_type', 'unknown')
            type_counts[chart_type] = type_counts.get(chart_type, 0) + 1
        return type_counts

    def _get_table_summary(self) -> Dict[str, int]:
        """Get summary of available table structures."""
        tables = self.library.get_tables()
        structure_counts = {}
        for table in tables:
            cols = table.get('columns', 0)
            rows = table.get('rows', 0)
            structure = f"{rows}x{cols}"
            structure_counts[structure] = structure_counts.get(structure, 0) + 1
        return structure_counts

    def generate_outline(
        self,
        request: str,
        presentation_type: Optional[str] = None,
        context: Optional[Dict[str, Any]] = None
    ) -> GeneratedContent:
        """
        Generate a presentation outline using LLM.

        Args:
            request: Natural language description of the presentation needed
            presentation_type: Optional type (investment_pitch, market_analysis, etc.)
            context: Optional additional context (fund details, market data, etc.)

        Returns:
            GeneratedContent with the outline structure
        """
        # Build the prompt
        prompt_parts = [
            f"Create a detailed presentation outline for the following request:",
            f"\n{request}\n"
        ]

        # Add presentation type context if available
        if presentation_type and presentation_type in self.content_patterns.get("presentation_types", {}):
            type_config = self.content_patterns["presentation_types"][presentation_type]
            prompt_parts.append(f"\nPresentation type: {presentation_type}")
            prompt_parts.append(f"Description: {type_config.get('description', '')}")
            prompt_parts.append(f"Typical sections: {', '.join(type_config.get('typical_sections', []))}")

        # Add user context
        if context:
            prompt_parts.append(f"\nAdditional context:")
            prompt_parts.append(json.dumps(context, indent=2))

        # Add library context
        library_context = self.get_library_context()
        if library_context:
            prompt_parts.append(f"\nAvailable visual components in the library:")
            prompt_parts.append(library_context)

        # Add output schema
        prompt_parts.append("""
\nGenerate the outline as JSON with this structure:
{
  "presentation_type": "type name",
  "title": "presentation title",
  "description": "brief description",
  "target_audience": "who this is for",
  "sections": [
    {
      "name": "Section Name",
      "purpose": "what this section accomplishes",
      "slides": [
        {
          "slide_type": "title_content|two_column|data_chart|table_slide|key_metrics|section_divider",
          "slide_number": 1,
          "title": "Slide Title",
          "content_description": "what content goes here",
          "suggested_visuals": ["chart type", "table structure", etc.],
          "key_points": ["point 1", "point 2"]
        }
      ]
    }
  ],
  "estimated_slide_count": 20
}""")

        prompt = "\n".join(prompt_parts)

        # Generate
        response = self.llm.generate(
            prompt=prompt,
            system_prompt=SYSTEM_PROMPTS["outline"],
            temperature=0.7
        )

        # Parse the response
        try:
            # Extract JSON from response (handle markdown code blocks)
            content_text = response.content
            if "```json" in content_text:
                content_text = content_text.split("```json")[1].split("```")[0]
            elif "```" in content_text:
                content_text = content_text.split("```")[1].split("```")[0]

            outline = json.loads(content_text.strip())
        except json.JSONDecodeError as e:
            logger.warning(f"Failed to parse outline JSON: {e}")
            outline = {"raw_response": response.content, "parse_error": str(e)}

        return GeneratedContent(
            content=outline,
            model_used=response.model,
            prompt_tokens=response.usage.get("input_tokens", 0),
            completion_tokens=response.usage.get("output_tokens", 0),
            raw_response=response.content
        )

    def generate_slide_content(
        self,
        slide_spec: Dict[str, Any],
        presentation_context: Optional[Dict[str, Any]] = None
    ) -> GeneratedContent:
        """
        Generate content for a specific slide.

        Args:
            slide_spec: Slide specification with type, title, description
            presentation_context: Context about the overall presentation

        Returns:
            GeneratedContent with slide content
        """
        slide_type = slide_spec.get("slide_type", "title_content")
        title = slide_spec.get("title", "")
        description = slide_spec.get("content_description", "")

        # Build prompt
        prompt_parts = [
            f"Generate content for a '{slide_type}' slide.",
            f"\nTitle: {title}",
            f"Purpose: {description}"
        ]

        if presentation_context:
            prompt_parts.append(f"\nPresentation context:")
            prompt_parts.append(json.dumps(presentation_context, indent=2))

        # Add schema based on slide type
        if slide_type == "title_content":
            prompt_parts.append("""
\nGenerate JSON:
{
  "title": "Slide Title",
  "bullets": ["bullet 1", "bullet 2", "bullet 3", "bullet 4"],
  "speaker_notes": "notes for presenter"
}""")
        elif slide_type == "two_column":
            prompt_parts.append("""
\nGenerate JSON:
{
  "title": "Slide Title",
  "left_column": {
    "header": "Left Header",
    "bullets": ["point 1", "point 2"]
  },
  "right_column": {
    "header": "Right Header",
    "bullets": ["point 1", "point 2"]
  },
  "speaker_notes": "notes for presenter"
}""")
        elif slide_type == "key_metrics":
            prompt_parts.append("""
\nGenerate JSON:
{
  "title": "Slide Title",
  "metrics": [
    {"label": "Metric Name", "value": "$X", "description": "context"},
    {"label": "Metric Name", "value": "X%", "description": "context"}
  ],
  "speaker_notes": "notes for presenter"
}""")
        elif slide_type == "data_chart":
            prompt_parts.append("""
\nGenerate JSON:
{
  "title": "Slide Title",
  "chart_type": "bar|line|pie|column",
  "chart_title": "Chart Title",
  "data_description": "what the chart shows",
  "categories": ["Cat 1", "Cat 2", "Cat 3"],
  "series": [
    {"name": "Series 1", "values": [10, 20, 30]}
  ],
  "speaker_notes": "notes for presenter"
}""")
        elif slide_type == "table_slide":
            prompt_parts.append("""
\nGenerate JSON:
{
  "title": "Slide Title",
  "headers": ["Col 1", "Col 2", "Col 3"],
  "data": [
    ["Row 1 Val 1", "Row 1 Val 2", "Row 1 Val 3"],
    ["Row 2 Val 1", "Row 2 Val 2", "Row 2 Val 3"]
  ],
  "speaker_notes": "notes for presenter"
}""")
        else:
            prompt_parts.append("""
\nGenerate JSON:
{
  "title": "Slide Title",
  "content": "main content",
  "speaker_notes": "notes for presenter"
}""")

        prompt = "\n".join(prompt_parts)

        response = self.llm.generate(
            prompt=prompt,
            system_prompt=SYSTEM_PROMPTS["slide_content"],
            temperature=0.7
        )

        # Parse response
        try:
            content_text = response.content
            if "```json" in content_text:
                content_text = content_text.split("```json")[1].split("```")[0]
            elif "```" in content_text:
                content_text = content_text.split("```")[1].split("```")[0]

            content = json.loads(content_text.strip())
        except json.JSONDecodeError as e:
            logger.warning(f"Failed to parse slide content JSON: {e}")
            content = {"raw_response": response.content, "parse_error": str(e)}

        return GeneratedContent(
            content=content,
            model_used=response.model,
            prompt_tokens=response.usage.get("input_tokens", 0),
            completion_tokens=response.usage.get("output_tokens", 0),
            raw_response=response.content
        )

    def generate_section_content(
        self,
        section_type: str,
        context: Dict[str, Any]
    ) -> GeneratedContent:
        """
        Generate content for a specific section type.

        Args:
            section_type: Type of section (executive_summary, market_analysis, etc.)
            context: Context about the presentation and section

        Returns:
            GeneratedContent with section content
        """
        system_prompt = SYSTEM_PROMPTS.get(section_type, SYSTEM_PROMPTS["slide_content"])

        prompt = f"""Generate content for a '{section_type}' section.

Context:
{json.dumps(context, indent=2)}

Generate professional, data-driven content appropriate for institutional investors.
Output as JSON with relevant fields for the section type."""

        response = self.llm.generate(
            prompt=prompt,
            system_prompt=system_prompt,
            temperature=0.7
        )

        try:
            content_text = response.content
            if "```json" in content_text:
                content_text = content_text.split("```json")[1].split("```")[0]
            elif "```" in content_text:
                content_text = content_text.split("```")[1].split("```")[0]

            content = json.loads(content_text.strip())
        except json.JSONDecodeError:
            content = {"text": response.content}

        return GeneratedContent(
            content=content,
            model_used=response.model,
            prompt_tokens=response.usage.get("input_tokens", 0),
            completion_tokens=response.usage.get("output_tokens", 0),
            raw_response=response.content
        )

    def enrich_outline(self, outline: Dict[str, Any]) -> Dict[str, Any]:
        """
        Enrich an outline by generating content for each slide.

        Args:
            outline: Presentation outline structure

        Returns:
            Enriched outline with generated content for each slide
        """
        enriched = outline.copy()
        total_tokens = {"input": 0, "output": 0}

        presentation_context = {
            "title": outline.get("title", ""),
            "type": outline.get("presentation_type", ""),
            "description": outline.get("description", "")
        }

        for section in enriched.get("sections", []):
            for slide in section.get("slides", []):
                # Generate content for this slide
                result = self.generate_slide_content(slide, presentation_context)

                # Merge generated content into slide
                slide["content"] = result.content
                slide["generated_by"] = result.model_used

                total_tokens["input"] += result.prompt_tokens
                total_tokens["output"] += result.completion_tokens

        enriched["generation_stats"] = {
            "model": self.llm.current_model.display_name if self.llm.current_model else "unknown",
            "total_input_tokens": total_tokens["input"],
            "total_output_tokens": total_tokens["output"]
        }

        return enriched

    def get_model_status(self) -> Dict[str, Any]:
        """Get current model status and available models."""
        return self.llm.get_status()


def main():
    """Test the content generator."""
    import argparse

    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

    parser = argparse.ArgumentParser(description="Content Generator Test")
    parser.add_argument("--model", "-m", default="claude-3.5-sonnet", help="Model to use")
    parser.add_argument("--outline", "-o", action="store_true", help="Generate a test outline")
    parser.add_argument("--slide", "-s", action="store_true", help="Generate test slide content")
    parser.add_argument("--request", "-r", default="Create a pitch deck for a $100M industrial fund", help="Request")
    parser.add_argument("--tone", default="professional",
                       choices=["default", "professional", "casual", "sales_pitch", "educational", "executive"],
                       help="Content tone")
    parser.add_argument("--verbosity", default="standard",
                       choices=["concise", "standard", "detailed"],
                       help="Content verbosity")
    parser.add_argument("--no-litellm", action="store_true", help="Use native providers instead of LiteLLM")

    args = parser.parse_args()

    generator = ContentGenerator(
        model=args.model,
        tone=args.tone,
        verbosity=args.verbosity,
        use_litellm=not args.no_litellm
    )

    print(f"\nUsing model: {generator.llm.current_model.display_name}")
    print(f"Tone: {args.tone}, Verbosity: {args.verbosity}")
    print("-" * 60)

    if args.outline:
        print(f"\nGenerating outline for: {args.request}")
        result = generator.generate_outline(args.request)
        print(f"\nGenerated outline:")
        print(json.dumps(result.content, indent=2))
        print(f"\nTokens used: {result.prompt_tokens} input, {result.completion_tokens} output")

    if args.slide:
        print("\nGenerating slide content...")
        slide_spec = {
            "slide_type": "key_metrics",
            "title": "Fund Highlights",
            "content_description": "Key metrics for a $100M industrial real estate fund"
        }
        result = generator.generate_slide_content(slide_spec)
        print(f"\nGenerated content:")
        print(json.dumps(result.content, indent=2))
        print(f"\nTokens used: {result.prompt_tokens} input, {result.completion_tokens} output")


if __name__ == "__main__":
    main()
