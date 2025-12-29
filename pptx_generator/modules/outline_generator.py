"""
Outline Generator Module

Creates presentation outlines from user descriptions using content patterns.
"""

import json
import logging
import re
from pathlib import Path
from typing import Optional

logger = logging.getLogger(__name__)


class OutlineGenerator:
    """Generates presentation outlines from user descriptions."""

    def __init__(self, content_patterns: dict, slide_catalog: dict):
        """
        Initialize the outline generator.

        Args:
            content_patterns: Content patterns dictionary
            slide_catalog: Slide catalog dictionary
        """
        self.content_patterns = content_patterns
        self.slide_catalog = slide_catalog

        self.presentation_types = content_patterns.get("presentation_types", {})
        self.reusable_sections = content_patterns.get("reusable_sections", {})
        self.research_categories = content_patterns.get("research_categories", {})

    def generate_outline(self, user_request: str) -> dict:
        """
        Generate a presentation outline from a user description.

        Args:
            user_request: Natural language description of the presentation needed

        Returns:
            Structured outline dictionary
        """
        # Classify the presentation type
        pres_type = self._classify_presentation_type(user_request)
        logger.info(f"Classified presentation type: {pres_type}")

        # Get the template for this presentation type
        type_config = self.presentation_types.get(pres_type, {})
        typical_sections = type_config.get("typical_sections", [])

        # Extract key information from the request
        context = self._extract_context(user_request)

        # Build the outline
        outline = {
            "presentation_type": pres_type,
            "title": self._generate_title(user_request, context),
            "description": type_config.get("description", ""),
            "context": context,
            "sections": [],
            "estimated_slide_count": 0
        }

        total_slides = 0

        for section_config in typical_sections:
            section = self._build_section(section_config, context)
            outline["sections"].append(section)
            total_slides += len(section.get("slides", []))

        outline["estimated_slide_count"] = total_slides

        return outline

    def _classify_presentation_type(self, request: str) -> str:
        """Classify the presentation type from the user request."""
        request_lower = request.lower()

        # Keywords for each presentation type
        type_keywords = {
            "investment_pitch": [
                "pitch", "fund", "investor", "raise", "capital",
                "fundraising", "lp", "limited partner", "commitment"
            ],
            "market_analysis": [
                "market analysis", "market overview", "market update",
                "economic", "macro", "outlook", "quarterly update"
            ],
            "due_diligence": [
                "due diligence", "property", "acquisition", "asset",
                "dd", "underwriting", "deal"
            ],
            "business_case": [
                "business case", "proposal", "initiative", "project",
                "budget", "approval", "recommendation"
            ],
            "consulting_framework": [
                "framework", "analysis", "strategy", "consulting",
                "assessment", "strategic"
            ]
        }

        # Score each type
        scores = {}
        for pres_type, keywords in type_keywords.items():
            score = sum(1 for kw in keywords if kw in request_lower)
            scores[pres_type] = score

        # Return the highest scoring type
        best_type = max(scores.items(), key=lambda x: x[1])
        if best_type[1] > 0:
            return best_type[0]

        # Default to investment_pitch
        return "investment_pitch"

    def _extract_context(self, request: str) -> dict:
        """Extract key context from the user request."""
        context = {
            "raw_request": request,
            "fund_size": None,
            "strategy": None,
            "geography": None,
            "sector": None,
            "timeframe": None,
            "key_topics": []
        }

        # Extract fund size
        fund_match = re.search(r'\$(\d+(?:\.\d+)?)\s*(M|B|million|billion)?', request, re.I)
        if fund_match:
            amount = float(fund_match.group(1))
            unit = fund_match.group(2)
            if unit and unit.lower().startswith('b'):
                amount *= 1000
            context["fund_size"] = f"${amount:.0f}M"

        # Extract strategy keywords
        strategy_keywords = [
            "logistics", "industrial", "multifamily", "office", "retail",
            "value-add", "core", "core-plus", "opportunistic", "development",
            "last-mile", "data center", "life sciences", "student housing",
            "senior housing", "self-storage", "hospitality"
        ]
        for kw in strategy_keywords:
            if kw.lower() in request.lower():
                context["strategy"] = kw
                context["key_topics"].append(kw)
                break

        # Extract geography
        geo_patterns = [
            r'(US|United States|nationwide)',
            r'(Europe|European)',
            r'(Asia|Asian|APAC)',
            r'(secondary markets?)',
            r'(primary markets?)',
            r'(gateway markets?|gateway cities)',
            r'(sun ?belt)',
            r'(midwest|west coast|east coast|northeast|southeast|southwest)'
        ]
        for pattern in geo_patterns:
            match = re.search(pattern, request, re.I)
            if match:
                context["geography"] = match.group(1)
                break

        # Extract sector
        sector_keywords = {
            "industrial": ["industrial", "warehouse", "logistics", "distribution"],
            "multifamily": ["multifamily", "residential", "apartment"],
            "office": ["office"],
            "retail": ["retail", "shopping"],
            "hospitality": ["hotel", "hospitality"]
        }
        for sector, keywords in sector_keywords.items():
            if any(kw in request.lower() for kw in keywords):
                context["sector"] = sector
                break

        return context

    def _generate_title(self, request: str, context: dict) -> str:
        """Generate a presentation title from the request."""
        # Try to extract an explicit title
        title_match = re.search(r'(?:titled?|called?|named?)\s*["\']([^"\']+)["\']', request, re.I)
        if title_match:
            return title_match.group(1)

        # Generate based on context
        parts = []

        if context.get("fund_size"):
            parts.append(context["fund_size"])

        if context.get("strategy"):
            parts.append(context["strategy"].title())

        if context.get("sector"):
            parts.append(context["sector"].title())

        if "fund" in request.lower() or "pitch" in request.lower():
            parts.append("Fund")
        elif "analysis" in request.lower():
            parts.append("Market Analysis")

        if parts:
            return " ".join(parts)

        return "Investor Presentation"

    def _build_section(self, section_config: dict, context: dict) -> dict:
        """Build a section from its configuration."""
        section = {
            "name": section_config["name"],
            "slides": [],
            "content_source": self._determine_content_source(section_config),
            "is_reusable": "reusable" in section_config.get("content_sources", [])
        }

        # Add research topics if applicable
        if section_config.get("research_topics"):
            section["research_topics"] = self._customize_research_topics(
                section_config["research_topics"],
                context
            )

        # Add reusable section reference
        if section_config.get("reusable_section"):
            section["reusable_section_id"] = section_config["reusable_section"]
            section["is_reusable"] = True

        # Build individual slides
        slide_count = section_config.get("slide_count", {})
        target_count = (slide_count.get("min", 1) + slide_count.get("max", 3)) // 2

        slide_types = section_config.get("slide_types", ["title_content"])

        for i in range(target_count):
            slide_type = slide_types[i % len(slide_types)]
            slide = {
                "slide_type": slide_type,
                "slide_number": i + 1,
                "content_source": section["content_source"],
                "content": {}  # To be filled during content assembly
            }

            # Add placeholder content hints
            if slide_type == "title_slide":
                slide["content"]["title"] = context.get("raw_request", "")[:50]
            elif slide_type == "section_divider":
                slide["content"]["title"] = section_config["name"]

            section["slides"].append(slide)

        return section

    def _determine_content_source(self, section_config: dict) -> str:
        """Determine the primary content source for a section."""
        sources = section_config.get("content_sources", ["user_input"])

        if "reusable" in sources:
            return "reusable"
        elif "research" in sources:
            return "research"
        elif "internal_data" in sources:
            return "internal_data"
        else:
            return "user_input"

    def _customize_research_topics(self, topics: list, context: dict) -> list:
        """Customize research topics based on context."""
        customized = []

        sector = context.get("sector", "")
        geography = context.get("geography", "")

        for topic in topics:
            customized_topic = topic

            # Add sector context
            if sector and "market" in topic.lower():
                customized_topic = f"{sector} {topic}"

            # Add geographic context
            if geography:
                customized_topic = f"{customized_topic} in {geography}"

            customized.append(customized_topic)

        return customized

    def refine_outline(self, outline: dict, user_feedback: str) -> dict:
        """
        Refine an outline based on user feedback.

        Args:
            outline: Existing outline dictionary
            user_feedback: User's modification request

        Returns:
            Updated outline dictionary
        """
        feedback_lower = user_feedback.lower()

        # Handle adding sections
        add_match = re.search(r'add\s+(?:a\s+)?(?:section\s+(?:on|about|for)\s+)?(.+?)(?:\s+section)?$', feedback_lower)
        if add_match:
            section_name = add_match.group(1).strip().title()
            new_section = self._create_custom_section(section_name)
            outline["sections"].append(new_section)
            outline["estimated_slide_count"] += len(new_section["slides"])
            logger.info(f"Added section: {section_name}")

        # Handle removing sections
        remove_match = re.search(r'remove\s+(?:the\s+)?(?:section\s+(?:on|about|for)\s+)?(.+?)(?:\s+section)?$', feedback_lower)
        if remove_match:
            section_name = remove_match.group(1).strip().lower()
            original_count = len(outline["sections"])
            outline["sections"] = [
                s for s in outline["sections"]
                if section_name not in s["name"].lower()
            ]
            removed_count = original_count - len(outline["sections"])
            if removed_count > 0:
                logger.info(f"Removed {removed_count} section(s) matching: {section_name}")
                # Recalculate slide count
                outline["estimated_slide_count"] = sum(
                    len(s.get("slides", [])) for s in outline["sections"]
                )

        # Handle reordering
        if "before" in feedback_lower or "after" in feedback_lower:
            self._reorder_sections(outline, feedback_lower)

        # Handle slide count adjustments
        if "more slides" in feedback_lower or "fewer slides" in feedback_lower:
            self._adjust_slide_counts(outline, feedback_lower)

        return outline

    def _create_custom_section(self, section_name: str) -> dict:
        """Create a custom section based on the name."""
        # Determine likely slide types based on section name
        name_lower = section_name.lower()

        if "risk" in name_lower:
            slide_types = ["section_divider", "title_content", "two_column"]
        elif "data" in name_lower or "analysis" in name_lower:
            slide_types = ["section_divider", "data_chart", "title_content"]
        elif "summary" in name_lower or "overview" in name_lower:
            slide_types = ["title_content", "key_metrics"]
        else:
            slide_types = ["section_divider", "title_content", "title_content"]

        section = {
            "name": section_name,
            "slides": [],
            "content_source": "user_input",
            "is_reusable": False
        }

        for i, slide_type in enumerate(slide_types):
            slide = {
                "slide_type": slide_type,
                "slide_number": i + 1,
                "content_source": "user_input",
                "content": {}
            }
            if slide_type == "section_divider":
                slide["content"]["title"] = section_name
            section["slides"].append(slide)

        return section

    def _reorder_sections(self, outline: dict, feedback: str) -> None:
        """Reorder sections based on feedback."""
        # This is a simplified implementation
        # Full implementation would parse "move X before Y" patterns
        pass

    def _adjust_slide_counts(self, outline: dict, feedback: str) -> None:
        """Adjust slide counts based on feedback."""
        if "more slides" in feedback:
            for section in outline["sections"]:
                if len(section["slides"]) < 6:
                    new_slide = {
                        "slide_type": "title_content",
                        "slide_number": len(section["slides"]) + 1,
                        "content_source": section["content_source"],
                        "content": {}
                    }
                    section["slides"].append(new_slide)
        elif "fewer slides" in feedback:
            for section in outline["sections"]:
                if len(section["slides"]) > 1:
                    section["slides"].pop()

        # Recalculate total
        outline["estimated_slide_count"] = sum(
            len(s.get("slides", [])) for s in outline["sections"]
        )

    def outline_to_text(self, outline: dict) -> str:
        """Convert an outline to a human-readable text format."""
        lines = []

        lines.append(f"# {outline.get('title', 'Presentation Outline')}")
        lines.append(f"Type: {outline.get('presentation_type', 'unknown')}")
        lines.append(f"Estimated Slides: {outline.get('estimated_slide_count', 0)}")
        lines.append("")

        for i, section in enumerate(outline.get("sections", []), 1):
            source_indicator = ""
            if section.get("is_reusable"):
                source_indicator = " [REUSABLE]"
            elif section.get("content_source") == "research":
                source_indicator = " [RESEARCH]"

            lines.append(f"## {i}. {section['name']}{source_indicator}")

            for slide in section.get("slides", []):
                lines.append(f"   - {slide['slide_type']}")

            if section.get("research_topics"):
                lines.append(f"   Research: {', '.join(section['research_topics'][:3])}")

            lines.append("")

        return "\n".join(lines)

    def save_outline(self, outline: dict, output_path: str) -> Path:
        """Save an outline to a JSON file."""
        path = Path(output_path)
        path.parent.mkdir(parents=True, exist_ok=True)

        with open(path, "w", encoding="utf-8") as f:
            json.dump(outline, f, indent=2)

        logger.info(f"Saved outline to: {path}")
        return path

    def load_outline(self, path: str) -> dict:
        """Load an outline from a JSON file."""
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)


def main():
    """Test the outline generator."""
    import argparse

    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

    parser = argparse.ArgumentParser(description="Outline Generator")
    parser.add_argument(
        "--patterns",
        default="C:/Users/wpcol/claudecode/pptx-design/pptx_generator/config/content_patterns.json",
        help="Path to content patterns"
    )
    parser.add_argument(
        "--catalog",
        default="C:/Users/wpcol/claudecode/pptx-design/pptx_generator/config/slide_catalog.json",
        help="Path to slide catalog"
    )
    parser.add_argument(
        "--request",
        default="Create a pitch deck for our new $200M industrial logistics fund targeting last-mile distribution centers in secondary markets",
        help="User request"
    )
    parser.add_argument(
        "--output",
        default="C:/Users/wpcol/claudecode/pptx-design/pptx_generator/output/test_outline.json",
        help="Output file path"
    )

    args = parser.parse_args()

    # Load configs
    with open(args.patterns, "r") as f:
        content_patterns = json.load(f)
    with open(args.catalog, "r") as f:
        slide_catalog = json.load(f)

    # Generate outline
    generator = OutlineGenerator(content_patterns, slide_catalog)
    outline = generator.generate_outline(args.request)

    # Print outline
    print(generator.outline_to_text(outline))

    # Save
    generator.save_outline(outline, args.output)
    print(f"\nSaved outline to: {args.output}")


if __name__ == "__main__":
    main()
