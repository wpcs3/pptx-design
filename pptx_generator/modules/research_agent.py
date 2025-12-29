"""
Research Agent Module

Performs deep research to generate content for presentation sections.
Uses LLM (Claude/GPT) for intelligent content generation.
"""

import json
import logging
import os
from datetime import datetime
from pathlib import Path
from typing import Any, Optional

logger = logging.getLogger(__name__)


class ResearchAgent:
    """Performs research and generates content for presentations using LLM."""

    def __init__(
        self,
        content_patterns: dict,
        cache_dir: Optional[str] = None,
        use_web_search: bool = True,
        use_llm: bool = True,
        llm_model: str = "claude-sonnet"
    ):
        """
        Initialize the research agent.

        Args:
            content_patterns: Content patterns dictionary with research categories
            cache_dir: Directory to cache research results
            use_web_search: Whether to use web search for research
            use_llm: Whether to use LLM for content generation
            llm_model: Which LLM model to use (claude-sonnet, gpt-4o, etc.)
        """
        self.content_patterns = content_patterns
        self.research_categories = content_patterns.get("research_categories", {})
        self.cache_dir = Path(cache_dir) if cache_dir else None
        self.use_web_search = use_web_search
        self.use_llm = use_llm
        self.llm_model = llm_model

        # Initialize LLM if enabled
        self._llm = None
        if self.use_llm:
            try:
                from .llm_provider import LLMManager
                self._llm = LLMManager(llm_model)
                if not self._llm.is_available():
                    logger.warning(f"LLM {llm_model} not available (API key missing)")
                    self._llm = None
            except ImportError:
                logger.warning("LLM provider not available")

        if self.cache_dir:
            self.cache_dir.mkdir(parents=True, exist_ok=True)

    async def research_section(
        self,
        section_name: str,
        research_topics: list[str],
        context: dict
    ) -> dict:
        """
        Perform deep research for a presentation section.

        Args:
            section_name: Name of the section (e.g., "Macro Economic Overview")
            research_topics: List of topics to research
            context: Additional context (sector, geography, etc.)

        Returns:
            Structured content ready for slide generation:
            {
                "slides": [
                    {
                        "slide_type": "data_chart",
                        "title": "GDP Growth Outlook",
                        "content": {...},
                        "sources": [...]
                    },
                    ...
                ]
            }
        """
        logger.info(f"Researching section: {section_name}")
        logger.info(f"Topics: {research_topics}")

        # Check cache first
        cache_key = self._get_cache_key(section_name, research_topics, context)
        cached = self._load_from_cache(cache_key)
        if cached:
            logger.info("Using cached research results")
            return cached

        # Perform research for each topic
        research_results = []
        for topic in research_topics:
            result = await self._research_topic(topic, context)
            research_results.append(result)

        # Format results into slides
        slides = self._format_for_slides(section_name, research_results, context)

        result = {
            "section_name": section_name,
            "slides": slides,
            "research_date": datetime.now().isoformat(),
            "topics_researched": research_topics
        }

        # Cache the results
        self._save_to_cache(cache_key, result)

        return result

    async def _research_topic(self, topic: str, context: dict) -> dict:
        """Research a specific topic."""
        logger.info(f"Researching: {topic}")

        # Determine the research category
        category = self._categorize_topic(topic)

        # Build search query
        query = self._build_search_query(topic, context)

        # Perform web search if enabled
        search_results = []
        if self.use_web_search:
            search_results = await self._web_search(query)

        # Extract and structure the data
        structured_data = self._extract_data(topic, search_results, category)

        return {
            "topic": topic,
            "category": category,
            "query": query,
            "data": structured_data,
            "sources": [r.get("url", "") for r in search_results[:3]]
        }

    def _categorize_topic(self, topic: str) -> str:
        """Categorize a research topic."""
        topic_lower = topic.lower()

        if any(kw in topic_lower for kw in ["gdp", "inflation", "interest", "employment"]):
            return "macroeconomic"
        elif any(kw in topic_lower for kw in ["cap rate", "rent", "vacancy", "transaction"]):
            return "real_estate_market"
        elif any(kw in topic_lower for kw in ["industrial", "logistics", "warehouse"]):
            return "sector_specific.industrial"
        elif any(kw in topic_lower for kw in ["multifamily", "apartment", "residential"]):
            return "sector_specific.multifamily"
        elif any(kw in topic_lower for kw in ["office"]):
            return "sector_specific.office"
        elif any(kw in topic_lower for kw in ["retail", "shopping"]):
            return "sector_specific.retail"
        elif any(kw in topic_lower for kw in ["competitor", "market share"]):
            return "competitive_landscape"

        return "general"

    def _build_search_query(self, topic: str, context: dict) -> str:
        """Build a search query from topic and context."""
        parts = [topic]

        if context.get("sector"):
            parts.append(context["sector"])
        if context.get("geography"):
            parts.append(context["geography"])

        # Add current year for recent data
        parts.append(str(datetime.now().year))

        return " ".join(parts)

    async def _web_search(self, query: str) -> list[dict]:
        """
        Perform a web search.

        This is a placeholder that returns mock data.
        In production, integrate with a search API.
        """
        logger.info(f"Web search: {query}")

        # Mock search results for demonstration
        # In production, use a real search API
        return [
            {
                "title": f"Research on {query}",
                "url": f"https://example.com/research/{query.replace(' ', '-')}",
                "snippet": f"Latest data and analysis on {query}..."
            }
        ]

    def _extract_data(
        self,
        topic: str,
        search_results: list[dict],
        category: str
    ) -> dict:
        """Extract structured data from search results using LLM."""
        # If LLM is available, use it for intelligent content generation
        if self._llm:
            return self._extract_data_with_llm(topic, search_results, category)

        # Fallback to mock data generation
        topic_lower = topic.lower()

        # Generate appropriate mock data based on topic
        if "gdp" in topic_lower:
            return {
                "type": "time_series",
                "title": "GDP Growth Rate",
                "data": {
                    "categories": ["Q1 2024", "Q2 2024", "Q3 2024", "Q4 2024"],
                    "values": [2.1, 2.3, 2.5, 2.4]
                },
                "unit": "%",
                "source": "Bureau of Economic Analysis"
            }
        elif "cap rate" in topic_lower:
            return {
                "type": "comparison",
                "title": "Cap Rates by Market",
                "data": {
                    "categories": ["Gateway", "Secondary", "Tertiary"],
                    "values": [4.5, 5.2, 6.1]
                },
                "unit": "%",
                "source": "CBRE Research"
            }
        elif "rent" in topic_lower:
            return {
                "type": "time_series",
                "title": "Rent Growth",
                "data": {
                    "categories": ["2021", "2022", "2023", "2024"],
                    "values": [8.5, 12.3, 6.2, 4.8]
                },
                "unit": "% YoY",
                "source": "CoStar"
            }
        elif "vacancy" in topic_lower:
            return {
                "type": "metric",
                "title": "Vacancy Rate",
                "value": "4.2%",
                "trend": "down",
                "source": "JLL Research"
            }
        elif "transaction" in topic_lower:
            return {
                "type": "time_series",
                "title": "Transaction Volume",
                "data": {
                    "categories": ["Q1", "Q2", "Q3", "Q4"],
                    "values": [45.2, 52.1, 48.7, 55.3]
                },
                "unit": "$B",
                "source": "Real Capital Analytics"
            }
        else:
            return {
                "type": "text",
                "title": topic,
                "content": f"Key insights on {topic}...",
                "bullets": [
                    f"Point 1 about {topic}",
                    f"Point 2 about {topic}",
                    f"Point 3 about {topic}"
                ],
                "source": "Industry Research"
            }

    def _extract_data_with_llm(
        self,
        topic: str,
        search_results: list[dict],
        category: str
    ) -> dict:
        """Use LLM to generate intelligent content for a topic."""
        system_prompt = """You are an expert real estate analyst generating content for investor presentations.
Generate accurate, professional content with realistic data points. Use current market knowledge.
Output must be valid JSON matching the requested schema."""

        prompt = f"""Generate slide content for the topic: "{topic}"
Category: {category}

Based on your knowledge, create appropriate content. Choose the best format:

For time series data (trends over time), use:
{{
  "type": "time_series",
  "title": "Title",
  "data": {{
    "categories": ["Period 1", "Period 2", "Period 3", "Period 4"],
    "values": [1.0, 2.0, 3.0, 4.0]
  }},
  "unit": "% or $B etc",
  "source": "Data Source"
}}

For comparison data (comparing categories), use:
{{
  "type": "comparison",
  "title": "Title",
  "data": {{
    "categories": ["Cat A", "Cat B", "Cat C"],
    "values": [10, 20, 30]
  }},
  "unit": "unit",
  "source": "Data Source"
}}

For single metric, use:
{{
  "type": "metric",
  "title": "Metric Name",
  "value": "X%",
  "trend": "up|down|stable",
  "context": "Brief context",
  "source": "Data Source"
}}

For text content, use:
{{
  "type": "text",
  "title": "Title",
  "content": "Main narrative",
  "bullets": ["Key point 1", "Key point 2", "Key point 3"],
  "source": "Data Source"
}}

Generate realistic data based on current market conditions (2024-2025).
Output only the JSON, no explanation."""

        try:
            response = self._llm.generate(
                prompt=prompt,
                system_prompt=system_prompt,
                temperature=0.7,
                max_tokens=1000
            )

            # Parse JSON from response
            content_text = response.content
            if "```json" in content_text:
                content_text = content_text.split("```json")[1].split("```")[0]
            elif "```" in content_text:
                content_text = content_text.split("```")[1].split("```")[0]

            data = json.loads(content_text.strip())
            data["generated_by"] = response.model
            return data

        except Exception as e:
            logger.warning(f"LLM extraction failed: {e}, falling back to mock data")
            # Fall back to mock data
            return {
                "type": "text",
                "title": topic,
                "content": f"Key insights on {topic}...",
                "bullets": [
                    f"Point 1 about {topic}",
                    f"Point 2 about {topic}",
                    f"Point 3 about {topic}"
                ],
                "source": "Industry Research"
            }

    def _format_for_slides(
        self,
        section_name: str,
        research_results: list[dict],
        context: dict
    ) -> list[dict]:
        """Format research results into slide specifications."""
        slides = []

        # Add section divider
        slides.append({
            "slide_type": "section_divider",
            "title": section_name,
            "content": {}
        })

        # Create slides from research results
        for result in research_results:
            data = result.get("data", {})
            data_type = data.get("type", "text")

            if data_type in ["time_series", "comparison"]:
                # Create a chart slide
                slide = {
                    "slide_type": "data_chart",
                    "title": data.get("title", result["topic"]),
                    "content": {
                        "chart_data": {
                            "type": "column" if data_type == "comparison" else "line",
                            "categories": data.get("data", {}).get("categories", []),
                            "series": [{
                                "name": data.get("title", "Data"),
                                "values": data.get("data", {}).get("values", [])
                            }]
                        },
                        "narrative": f"Source: {data.get('source', 'Research')}"
                    },
                    "sources": result.get("sources", [])
                }
            elif data_type == "metric":
                # Create a key metrics slide
                slide = {
                    "slide_type": "key_metrics",
                    "title": data.get("title", result["topic"]),
                    "content": {
                        "metrics": [
                            {
                                "label": data.get("title", ""),
                                "value": data.get("value", "")
                            }
                        ]
                    },
                    "sources": result.get("sources", [])
                }
            else:
                # Create a content slide
                slide = {
                    "slide_type": "title_content",
                    "title": data.get("title", result["topic"]),
                    "content": {
                        "body": data.get("content", ""),
                        "bullets": data.get("bullets", [])
                    },
                    "sources": result.get("sources", [])
                }

            slides.append(slide)

        return slides

    def _get_cache_key(
        self,
        section_name: str,
        topics: list[str],
        context: dict
    ) -> str:
        """Generate a cache key for research results."""
        import hashlib

        key_data = json.dumps({
            "section": section_name,
            "topics": sorted(topics),
            "sector": context.get("sector"),
            "geography": context.get("geography")
        }, sort_keys=True)

        return hashlib.md5(key_data.encode()).hexdigest()

    def _load_from_cache(self, cache_key: str) -> Optional[dict]:
        """Load research results from cache."""
        if not self.cache_dir:
            return None

        cache_file = self.cache_dir / f"{cache_key}.json"
        if not cache_file.exists():
            return None

        try:
            with open(cache_file, "r", encoding="utf-8") as f:
                cached = json.load(f)

            # Check if cache is still valid (e.g., less than 24 hours old)
            cache_date = datetime.fromisoformat(cached.get("research_date", "2000-01-01"))
            if (datetime.now() - cache_date).days < 1:
                return cached
        except Exception as e:
            logger.warning(f"Error loading cache: {e}")

        return None

    def _save_to_cache(self, cache_key: str, data: dict) -> None:
        """Save research results to cache."""
        if not self.cache_dir:
            return

        cache_file = self.cache_dir / f"{cache_key}.json"
        try:
            with open(cache_file, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2)
        except Exception as e:
            logger.warning(f"Error saving to cache: {e}")

    def format_for_slides(
        self,
        research_results: dict,
        slide_types: list[str]
    ) -> list[dict]:
        """
        Transform research results into slide-ready content.

        Args:
            research_results: Results from research_section
            slide_types: Target slide types to format for

        Returns:
            List of slide content dictionaries
        """
        return research_results.get("slides", [])


def main():
    """Test the research agent."""
    import asyncio
    import argparse

    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

    parser = argparse.ArgumentParser(description="Research Agent")
    parser.add_argument(
        "--patterns",
        default="C:/Users/wpcol/claudecode/pptx-design/pptx_generator/config/content_patterns.json",
        help="Path to content patterns"
    )
    parser.add_argument(
        "--cache-dir",
        default="C:/Users/wpcol/claudecode/pptx-design/pptx_generator/cache/research",
        help="Cache directory"
    )
    parser.add_argument(
        "--section",
        default="Macro Economic Overview",
        help="Section to research"
    )
    parser.add_argument(
        "--output",
        default="C:/Users/wpcol/claudecode/pptx-design/pptx_generator/output/test_research.json",
        help="Output file path"
    )

    args = parser.parse_args()

    # Load content patterns
    with open(args.patterns, "r") as f:
        content_patterns = json.load(f)

    # Create agent
    agent = ResearchAgent(content_patterns, cache_dir=args.cache_dir)

    # Run research
    async def run_research():
        result = await agent.research_section(
            section_name=args.section,
            research_topics=["GDP growth", "inflation trends", "interest rates"],
            context={"sector": "industrial", "geography": "US"}
        )
        return result

    result = asyncio.run(run_research())

    # Save results
    Path(args.output).parent.mkdir(parents=True, exist_ok=True)
    with open(args.output, "w") as f:
        json.dump(result, f, indent=2)

    print(f"Research completed: {len(result['slides'])} slides generated")
    print(f"Saved to: {args.output}")


if __name__ == "__main__":
    main()
