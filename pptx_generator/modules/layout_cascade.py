"""
Cascading Layout Handler

Automatically selects the optimal slide layout based on content analysis.
Uses a cascade of handlers that check content characteristics in priority order.

Inspired by SlideDeck AI's cascading layout selection pattern.

Phase 2 Enhancement (2025-12-29):
- Content-based layout detection
- Handler cascade with priority ordering
- Smart matching for data-heavy vs text-heavy content
- Icon and visual detection
"""

import logging
import re
from dataclasses import dataclass
from enum import Enum
from typing import Any, Callable, Dict, List, Optional, Tuple

logger = logging.getLogger(__name__)


# =============================================================================
# Layout Types
# =============================================================================

class LayoutType(Enum):
    """Available layout types in order of specificity."""
    TITLE_SLIDE = "title_slide"
    SECTION_DIVIDER = "section_divider"
    AGENDA = "agenda"
    KEY_METRICS = "key_metrics"
    DATA_CHART = "data_chart"
    TABLE_SLIDE = "table_slide"
    TWO_COLUMN = "two_column"
    ICON_GRID = "icon_grid"
    TIMELINE = "timeline"
    STEP_BY_STEP = "step_by_step"
    COMPARISON = "comparison"
    TITLE_CONTENT = "title_content"
    BLANK = "blank"
    DEFAULT = "default"


# =============================================================================
# Content Analysis
# =============================================================================

@dataclass
class ContentAnalysis:
    """Results of content analysis for layout selection."""
    has_chart_data: bool = False
    has_table_data: bool = False
    has_metrics: bool = False
    has_icons: bool = False
    has_comparison: bool = False
    has_timeline: bool = False
    has_steps: bool = False
    has_agenda_markers: bool = False
    bullet_count: int = 0
    word_count: int = 0
    column_count: int = 0
    is_title_only: bool = False
    is_section_header: bool = False
    data_density: float = 0.0  # 0-1, how much numerical data

    def __str__(self) -> str:
        features = []
        if self.has_chart_data:
            features.append("chart")
        if self.has_table_data:
            features.append("table")
        if self.has_metrics:
            features.append("metrics")
        if self.has_icons:
            features.append("icons")
        if self.has_comparison:
            features.append("comparison")
        if self.has_timeline:
            features.append("timeline")
        if self.has_steps:
            features.append("steps")
        return f"ContentAnalysis({', '.join(features) or 'text-only'})"


def analyze_content(content: Dict[str, Any]) -> ContentAnalysis:
    """
    Analyze slide content to determine characteristics.

    Args:
        content: Slide content dictionary

    Returns:
        ContentAnalysis with detected features
    """
    analysis = ContentAnalysis()

    # Check for chart data
    if "chart_data" in content or "chart" in content:
        chart = content.get("chart_data", content.get("chart", {}))
        if chart and (chart.get("categories") or chart.get("series")):
            analysis.has_chart_data = True

    # Check for table data
    if "table_data" in content or ("headers" in content and "data" in content):
        analysis.has_table_data = True
    if "rows" in content and "columns" in content:
        analysis.has_table_data = True

    # Check for metrics
    if "metrics" in content:
        metrics = content.get("metrics", [])
        if metrics and len(metrics) >= 2:
            analysis.has_metrics = True

    # Check for icons
    if "icons" in content:
        analysis.has_icons = True

    # Check for comparison structure
    if "left_column" in content and "right_column" in content:
        analysis.has_comparison = True
        analysis.column_count = 2
    if "columns" in content:
        cols = content.get("columns", [])
        if len(cols) == 2:
            analysis.has_comparison = True
        analysis.column_count = len(cols)

    # Count bullets
    bullets = content.get("bullets", content.get("body", []))
    if isinstance(bullets, list):
        analysis.bullet_count = len(bullets)

        # Check for step indicators (numbered, ">>" prefix)
        step_pattern = re.compile(r'^(\d+[\.\):]|>>|Step\s+\d+|Phase\s+\d+)', re.IGNORECASE)
        step_count = sum(1 for b in bullets if step_pattern.match(str(b).strip()))
        if step_count >= 2:
            analysis.has_steps = True

        # Check for timeline indicators
        timeline_pattern = re.compile(r'(Q[1-4]|20\d{2}|Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)', re.IGNORECASE)
        timeline_count = sum(1 for b in bullets if timeline_pattern.search(str(b)))
        if timeline_count >= 2:
            analysis.has_timeline = True

        # Check for agenda markers (numbered items, topic-like)
        agenda_keywords = ["overview", "introduction", "agenda", "outline", "topics"]
        title = str(content.get("title", "")).lower()
        if any(kw in title for kw in agenda_keywords):
            analysis.has_agenda_markers = True

    # Count words
    all_text = []
    for key in ["title", "subtitle", "body"]:
        val = content.get(key, "")
        if isinstance(val, str):
            all_text.append(val)
    for bullet in content.get("bullets", []):
        all_text.append(str(bullet))

    analysis.word_count = sum(len(t.split()) for t in all_text)

    # Check for title-only slide
    # Only mark as title-only or section header if there's no other content
    title = content.get("title", "")
    subtitle = content.get("subtitle", "")
    has_other_content = (
        analysis.has_chart_data or
        analysis.has_table_data or
        analysis.has_metrics or
        analysis.has_icons or
        analysis.has_comparison or
        "metrics" in content or
        "chart_data" in content or
        "table_data" in content
    )
    if title and not content.get("bullets") and not content.get("body") and not has_other_content:
        if subtitle:
            analysis.is_title_only = True
        elif analysis.word_count < 10:
            analysis.is_section_header = True

    # Calculate data density (percentage of numbers/metrics in content)
    number_pattern = re.compile(r'[\$%]?\d[\d,\.]*[%KMB]?')
    text_combined = " ".join(all_text)
    numbers_found = number_pattern.findall(text_combined)
    if text_combined:
        analysis.data_density = min(1.0, len(numbers_found) * 5 / max(1, analysis.word_count))

    return analysis


# =============================================================================
# Layout Handlers
# =============================================================================

# Type alias for handler function
LayoutHandler = Callable[[Dict[str, Any], ContentAnalysis], bool]


def _is_title_slide(content: Dict[str, Any], analysis: ContentAnalysis) -> bool:
    """Check if content is for a title slide."""
    if content.get("slide_type") == "title_slide":
        return True
    return analysis.is_title_only and "subtitle" in content


def _is_section_divider(content: Dict[str, Any], analysis: ContentAnalysis) -> bool:
    """Check if content is for a section divider."""
    if content.get("slide_type") in ["section_divider", "section_breaker"]:
        return True
    return analysis.is_section_header


def _is_agenda(content: Dict[str, Any], analysis: ContentAnalysis) -> bool:
    """Check if content is for an agenda slide."""
    if content.get("slide_type") == "agenda":
        return True

    # Check title keywords
    title = str(content.get("title", "")).lower()
    agenda_keywords = ["agenda", "outline", "overview", "today", "topics", "contents"]
    if any(kw in title for kw in agenda_keywords):
        if analysis.bullet_count >= 3:
            return True

    return analysis.has_agenda_markers and analysis.bullet_count >= 3


def _is_key_metrics(content: Dict[str, Any], analysis: ContentAnalysis) -> bool:
    """Check if content is for a key metrics slide."""
    if content.get("slide_type") == "key_metrics":
        return True
    if "metrics" in content and len(content.get("metrics", [])) >= 2:
        return True

    # Check for metrics pattern in bullets
    title = str(content.get("title", "")).lower()
    metric_keywords = ["metrics", "kpi", "highlights", "numbers", "key figures", "performance"]
    if any(kw in title for kw in metric_keywords):
        if analysis.data_density > 0.3:
            return True

    return analysis.has_metrics


def _is_data_chart(content: Dict[str, Any], analysis: ContentAnalysis) -> bool:
    """Check if content is for a chart slide."""
    if content.get("slide_type") == "data_chart":
        return True
    return analysis.has_chart_data


def _is_table_slide(content: Dict[str, Any], analysis: ContentAnalysis) -> bool:
    """Check if content is for a table slide."""
    if content.get("slide_type") == "table_slide":
        return True
    return analysis.has_table_data


def _is_two_column(content: Dict[str, Any], analysis: ContentAnalysis) -> bool:
    """Check if content is for a two-column slide."""
    if content.get("slide_type") in ["two_column", "comparison"]:
        return True
    return analysis.has_comparison


def _is_icon_grid(content: Dict[str, Any], analysis: ContentAnalysis) -> bool:
    """Check if content should use an icon grid layout."""
    if content.get("slide_type") == "icon_grid":
        return True
    return analysis.has_icons


def _is_timeline(content: Dict[str, Any], analysis: ContentAnalysis) -> bool:
    """Check if content is for a timeline slide."""
    if content.get("slide_type") == "timeline":
        return True

    title = str(content.get("title", "")).lower()
    timeline_keywords = ["timeline", "roadmap", "milestones", "schedule", "history"]
    if any(kw in title for kw in timeline_keywords):
        return True

    return analysis.has_timeline


def _is_step_by_step(content: Dict[str, Any], analysis: ContentAnalysis) -> bool:
    """Check if content is for a step-by-step process slide."""
    if content.get("slide_type") == "step_by_step":
        return True

    title = str(content.get("title", "")).lower()
    step_keywords = ["process", "steps", "how to", "workflow", "phases"]
    if any(kw in title for kw in step_keywords):
        return True

    return analysis.has_steps


# =============================================================================
# Cascade Definition
# =============================================================================

# Handler cascade in priority order
# Each tuple is (layout_type, handler_function)
LAYOUT_CASCADE: List[Tuple[LayoutType, LayoutHandler]] = [
    (LayoutType.TITLE_SLIDE, _is_title_slide),
    (LayoutType.SECTION_DIVIDER, _is_section_divider),
    (LayoutType.KEY_METRICS, _is_key_metrics),
    (LayoutType.DATA_CHART, _is_data_chart),
    (LayoutType.TABLE_SLIDE, _is_table_slide),
    (LayoutType.ICON_GRID, _is_icon_grid),
    (LayoutType.TIMELINE, _is_timeline),
    (LayoutType.STEP_BY_STEP, _is_step_by_step),
    (LayoutType.TWO_COLUMN, _is_two_column),
    (LayoutType.AGENDA, _is_agenda),
    (LayoutType.TITLE_CONTENT, lambda c, a: True),  # Fallback
]


# =============================================================================
# Main Layout Selector
# =============================================================================

class LayoutCascade:
    """
    Cascading layout selector.

    Usage:
        cascade = LayoutCascade()
        layout = cascade.select_layout(content)
        print(f"Selected: {layout}")

        # With custom handlers
        cascade.add_handler(LayoutType.CUSTOM, my_handler, priority=5)
    """

    def __init__(self, cascade: List[Tuple[LayoutType, LayoutHandler]] = None):
        """
        Initialize with optional custom cascade.

        Args:
            cascade: Custom handler cascade (uses default if not provided)
        """
        self.cascade = list(cascade or LAYOUT_CASCADE)

    def add_handler(
        self,
        layout_type: LayoutType,
        handler: LayoutHandler,
        priority: int = None
    ) -> None:
        """
        Add a custom handler to the cascade.

        Args:
            layout_type: Layout type this handler detects
            handler: Handler function
            priority: Position in cascade (lower = higher priority)
        """
        entry = (layout_type, handler)
        if priority is not None:
            self.cascade.insert(priority, entry)
        else:
            # Insert before DEFAULT handler
            self.cascade.insert(-1, entry)

    def select_layout(
        self,
        content: Dict[str, Any],
        analysis: ContentAnalysis = None
    ) -> LayoutType:
        """
        Select the best layout for the given content.

        Args:
            content: Slide content dictionary
            analysis: Pre-computed analysis (will compute if not provided)

        Returns:
            Selected LayoutType
        """
        # Use explicit slide_type if provided and valid
        explicit_type = content.get("slide_type", "")
        if explicit_type:
            try:
                return LayoutType(explicit_type)
            except ValueError:
                pass  # Not a valid LayoutType, continue with cascade

        # Analyze content if not provided
        if analysis is None:
            analysis = analyze_content(content)

        # Run through cascade
        for layout_type, handler in self.cascade:
            try:
                if handler(content, analysis):
                    logger.debug(f"Layout cascade selected: {layout_type.value} for {analysis}")
                    return layout_type
            except Exception as e:
                logger.warning(f"Handler for {layout_type.value} failed: {e}")
                continue

        # Fallback to default
        return LayoutType.DEFAULT

    def explain_selection(
        self,
        content: Dict[str, Any]
    ) -> Dict[str, Any]:
        """
        Explain why a particular layout was selected.

        Args:
            content: Slide content dictionary

        Returns:
            Dictionary with selection explanation
        """
        analysis = analyze_content(content)
        selected = self.select_layout(content, analysis)

        # Find which handlers matched
        matches = []
        for layout_type, handler in self.cascade:
            try:
                if handler(content, analysis):
                    matches.append(layout_type.value)
            except Exception:
                pass

        return {
            "selected": selected.value,
            "analysis": str(analysis),
            "matching_handlers": matches,
            "content_features": {
                "has_chart_data": analysis.has_chart_data,
                "has_table_data": analysis.has_table_data,
                "has_metrics": analysis.has_metrics,
                "has_comparison": analysis.has_comparison,
                "has_timeline": analysis.has_timeline,
                "has_steps": analysis.has_steps,
                "bullet_count": analysis.bullet_count,
                "word_count": analysis.word_count,
                "data_density": analysis.data_density,
            }
        }


# =============================================================================
# Convenience Functions
# =============================================================================

_default_cascade = LayoutCascade()


def select_layout(content: Dict[str, Any]) -> str:
    """
    Quick layout selection using default cascade.

    Args:
        content: Slide content dictionary

    Returns:
        Layout type string
    """
    return _default_cascade.select_layout(content).value


def explain_layout(content: Dict[str, Any]) -> Dict[str, Any]:
    """
    Explain layout selection for content.

    Args:
        content: Slide content dictionary

    Returns:
        Explanation dictionary
    """
    return _default_cascade.explain_selection(content)


# =============================================================================
# CLI
# =============================================================================

def main():
    """CLI for testing layout cascade."""
    import argparse
    import json

    logging.basicConfig(level=logging.DEBUG, format="%(levelname)s: %(message)s")

    parser = argparse.ArgumentParser(description="Layout Cascade Test")
    parser.add_argument("--content", "-c", help="JSON content string or file path")
    parser.add_argument("--explain", "-e", action="store_true", help="Show explanation")

    args = parser.parse_args()

    # Test cases
    test_cases = [
        {"title": "Q4 2024 Results", "subtitle": "Investment Review"},
        {"title": "Agenda", "bullets": ["Overview", "Market Analysis", "Financial Projections", "Q&A"]},
        {"title": "Key Metrics", "metrics": [
            {"label": "Revenue", "value": "$1.2M"},
            {"label": "Growth", "value": "25%"},
            {"label": "Users", "value": "10K"}
        ]},
        {"title": "Revenue Trend", "chart_data": {
            "type": "line",
            "categories": ["Q1", "Q2", "Q3", "Q4"],
            "series": [{"name": "Revenue", "values": [100, 120, 140, 180]}]
        }},
        {"title": "Comparison", "left_column": {"header": "Before"}, "right_column": {"header": "After"}},
        {"title": "Process", "bullets": [
            "Step 1: Research the market",
            "Step 2: Identify opportunities",
            "Step 3: Execute strategy"
        ]},
        {"title": "Market Overview", "bullets": ["Point 1", "Point 2", "Point 3"]},
    ]

    if args.content:
        # Load custom content
        if args.content.endswith(".json"):
            with open(args.content, "r") as f:
                test_cases = [json.load(f)]
        else:
            test_cases = [json.loads(args.content)]

    cascade = LayoutCascade()

    print("Layout Cascade Test Results")
    print("=" * 60)

    for content in test_cases:
        title = content.get("title", "Untitled")[:40]
        layout = cascade.select_layout(content)
        print(f"\nTitle: {title}")
        print(f"  Layout: {layout.value}")

        if args.explain:
            explanation = cascade.explain_selection(content)
            print(f"  Analysis: {explanation['analysis']}")
            print(f"  Matching: {', '.join(explanation['matching_handlers'])}")


if __name__ == "__main__":
    main()
