"""
Content Generation Pipeline

Separates content generation from layout selection:
1. Content Generation: LLM generates structured content only
2. Layout Selection: Rule-based system matches content to layouts
3. Rendering: Template-based rendering fills placeholders

This reduces API costs and improves consistency.
"""

import json
import logging
import re
from dataclasses import dataclass, field
from enum import Enum
from typing import Any, Dict, List, Optional

logger = logging.getLogger(__name__)


class SlideType(Enum):
    """Types of slides that can be generated."""
    TITLE = "title"
    SECTION = "section"
    AGENDA = "agenda"
    CONTENT = "content"
    TWO_COLUMN = "two_column"
    METRICS = "metrics"
    TABLE = "table"
    CHART = "chart"
    TIMELINE = "timeline"
    END = "end"


@dataclass
class SlideContent:
    """Structured content for a single slide."""
    slide_type: SlideType
    title: str = ""
    subtitle: str = ""
    body: str = ""
    bullets: List[str] = field(default_factory=list)
    data: Dict[str, Any] = field(default_factory=dict)
    notes: str = ""

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for rendering."""
        result = {
            "type": self.slide_type.value,
            "content": {}
        }
        if self.title:
            result["content"]["title"] = self.title
        if self.subtitle:
            result["content"]["subtitle"] = self.subtitle
        if self.body:
            result["content"]["body"] = self.body
        if self.bullets:
            result["content"]["bullets"] = self.bullets
        if self.data:
            result["content"].update(self.data)
        return result


@dataclass
class PresentationOutline:
    """Complete presentation outline with all slides."""
    title: str
    purpose: str
    audience: str
    slides: List[SlideContent] = field(default_factory=list)
    metadata: Dict[str, Any] = field(default_factory=dict)

    def add_slide(self, slide: SlideContent) -> None:
        """Add a slide to the outline."""
        self.slides.append(slide)

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary."""
        return {
            "title": self.title,
            "purpose": self.purpose,
            "audience": self.audience,
            "slides": [s.to_dict() for s in self.slides],
            "metadata": self.metadata,
        }

    def to_json(self, indent: int = 2) -> str:
        """Convert to JSON string."""
        return json.dumps(self.to_dict(), indent=indent)


class LayoutMatcher:
    """
    Rule-based layout matcher.

    Maps slide content to appropriate template layouts based on:
    - Slide type
    - Content structure (bullets, data, etc.)
    - Template availability
    """

    # Default layout mapping
    DEFAULT_LAYOUTS = {
        SlideType.TITLE: ["Frontpage", "Title Slide", "Cover"],
        SlideType.SECTION: ["Section breaker", "Section Divider", "Divider"],
        SlideType.AGENDA: ["Agenda", "Table of Contents", "Default"],
        SlideType.CONTENT: ["Default", "Content", "Title and Content"],
        SlideType.TWO_COLUMN: ["1/2 grey", "Two Column", "Comparison"],
        SlideType.METRICS: ["Default", "Key Metrics", "Dashboard"],
        SlideType.TABLE: ["Default", "Table Slide", "Data"],
        SlideType.CHART: ["Default", "Chart Slide", "Data"],
        SlideType.TIMELINE: ["Default", "Timeline", "Process"],
        SlideType.END: ["End", "Thank You", "Closing", "Frontpage"],
    }

    def __init__(self, available_layouts: List[str] = None):
        """
        Initialize matcher with available layouts.

        Args:
            available_layouts: List of layout names from template
        """
        self.available_layouts = available_layouts or []

    def match(self, slide: SlideContent) -> str:
        """
        Find the best layout for a slide.

        Args:
            slide: SlideContent to match

        Returns:
            Layout name
        """
        preferred_layouts = self.DEFAULT_LAYOUTS.get(slide.slide_type, ["Default"])

        # Try each preferred layout in order
        for layout_name in preferred_layouts:
            # Exact match
            if layout_name in self.available_layouts:
                return layout_name

            # Case-insensitive match
            for available in self.available_layouts:
                if layout_name.lower() == available.lower():
                    return available

            # Partial match
            for available in self.available_layouts:
                if layout_name.lower() in available.lower():
                    return available

        # Fallback to Default or first available
        if "Default" in self.available_layouts:
            return "Default"
        if self.available_layouts:
            return self.available_layouts[0]

        return "Default"

    def match_all(self, outline: PresentationOutline) -> List[Dict[str, Any]]:
        """
        Match layouts for all slides in an outline.

        Returns:
            List of slide specs with layout assignments
        """
        results = []
        for slide in outline.slides:
            slide_dict = slide.to_dict()
            slide_dict["layout"] = self.match(slide)
            results.append(slide_dict)
        return results


class ContentParser:
    """
    Parse structured content from various input formats.

    Supports:
    - Natural language requests
    - JSON outlines
    - Markdown documents
    """

    # Patterns for extracting structure
    SECTION_PATTERN = re.compile(r'^#{1,2}\s+(.+)$', re.MULTILINE)
    BULLET_PATTERN = re.compile(r'^[-*]\s+(.+)$', re.MULTILINE)
    NUMBERED_PATTERN = re.compile(r'^\d+\.\s+(.+)$', re.MULTILINE)

    def parse_markdown(self, markdown: str) -> PresentationOutline:
        """
        Parse markdown into a presentation outline.

        Args:
            markdown: Markdown text

        Returns:
            PresentationOutline
        """
        lines = markdown.strip().split('\n')
        outline = PresentationOutline(title="", purpose="", audience="")

        current_slide = None
        current_bullets = []

        for line in lines:
            line = line.strip()

            # H1 = Presentation title
            if line.startswith('# ') and not line.startswith('## '):
                outline.title = line[2:].strip()
                continue

            # H2 = New slide
            if line.startswith('## '):
                # Save previous slide
                if current_slide:
                    if current_bullets:
                        current_slide.bullets = current_bullets
                    outline.add_slide(current_slide)

                # Start new slide
                title = line[3:].strip()
                slide_type = self._infer_slide_type(title, [])
                current_slide = SlideContent(slide_type=slide_type, title=title)
                current_bullets = []
                continue

            # Bullets
            if line.startswith('- ') or line.startswith('* '):
                current_bullets.append(line[2:].strip())
                continue

            # Numbered items
            match = self.NUMBERED_PATTERN.match(line)
            if match:
                current_bullets.append(match.group(1))
                continue

            # Regular text (body)
            if line and current_slide and not current_slide.body:
                current_slide.body = line

        # Save last slide
        if current_slide:
            if current_bullets:
                current_slide.bullets = current_bullets
            outline.add_slide(current_slide)

        return outline

    def parse_json(self, json_str: str) -> PresentationOutline:
        """
        Parse JSON into a presentation outline.

        Args:
            json_str: JSON string

        Returns:
            PresentationOutline
        """
        data = json.loads(json_str)

        outline = PresentationOutline(
            title=data.get("title", ""),
            purpose=data.get("purpose", ""),
            audience=data.get("audience", ""),
            metadata=data.get("metadata", {})
        )

        for slide_data in data.get("slides", []):
            slide_type = SlideType(slide_data.get("type", "content"))
            content = slide_data.get("content", {})

            slide = SlideContent(
                slide_type=slide_type,
                title=content.get("title", ""),
                subtitle=content.get("subtitle", ""),
                body=content.get("body", ""),
                bullets=content.get("bullets", []),
                data={k: v for k, v in content.items()
                      if k not in ["title", "subtitle", "body", "bullets"]},
            )
            outline.add_slide(slide)

        return outline

    def _infer_slide_type(self, title: str, bullets: List[str]) -> SlideType:
        """Infer slide type from title and content."""
        title_lower = title.lower()

        if any(word in title_lower for word in ["agenda", "contents", "overview"]):
            return SlideType.AGENDA
        if any(word in title_lower for word in ["thank", "questions", "end", "closing"]):
            return SlideType.END
        if any(word in title_lower for word in ["vs", "comparison", "compare"]):
            return SlideType.TWO_COLUMN
        if any(word in title_lower for word in ["metrics", "kpi", "performance"]):
            return SlideType.METRICS
        if any(word in title_lower for word in ["timeline", "roadmap", "schedule"]):
            return SlideType.TIMELINE

        return SlideType.CONTENT


class ContentPipeline:
    """
    Main pipeline for content generation and layout matching.

    Usage:
        pipeline = ContentPipeline(template="consulting_toolkit")
        outline = pipeline.parse_request("Create a pitch deck for a SaaS startup")
        slides = pipeline.prepare_for_rendering(outline)
        # slides can now be passed to PresentationBuilder
    """

    def __init__(self, template: str = None):
        """
        Initialize pipeline.

        Args:
            template: Template name for layout matching
        """
        self.parser = ContentParser()
        self.matcher = None

        if template:
            self._load_template_layouts(template)

    def _load_template_layouts(self, template: str) -> None:
        """Load available layouts from template."""
        from .registry import TemplateRegistry

        registry = TemplateRegistry()
        info = registry.get_template(template)
        if info:
            layouts = info.get("layout_names", [])
            self.matcher = LayoutMatcher(layouts)
            logger.info(f"Loaded {len(layouts)} layouts from {template}")
        else:
            self.matcher = LayoutMatcher([])
            logger.warning(f"Template not found: {template}")

    def parse_markdown(self, markdown: str) -> PresentationOutline:
        """Parse markdown into outline."""
        return self.parser.parse_markdown(markdown)

    def parse_json(self, json_str: str) -> PresentationOutline:
        """Parse JSON into outline."""
        return self.parser.parse_json(json_str)

    def prepare_for_rendering(self, outline: PresentationOutline) -> List[Dict[str, Any]]:
        """
        Prepare outline for rendering by matching layouts.

        Args:
            outline: PresentationOutline

        Returns:
            List of slide specs ready for PresentationBuilder
        """
        if not self.matcher:
            self.matcher = LayoutMatcher([])

        return self.matcher.match_all(outline)

    def create_standard_outline(
        self,
        title: str,
        sections: List[str],
        purpose: str = "",
        audience: str = ""
    ) -> PresentationOutline:
        """
        Create a standard presentation outline.

        Args:
            title: Presentation title
            sections: List of section names
            purpose: Purpose of the presentation
            audience: Target audience

        Returns:
            PresentationOutline with standard structure
        """
        outline = PresentationOutline(
            title=title,
            purpose=purpose,
            audience=audience
        )

        # Title slide
        outline.add_slide(SlideContent(
            slide_type=SlideType.TITLE,
            title=title,
            subtitle=purpose or ""
        ))

        # Agenda slide
        outline.add_slide(SlideContent(
            slide_type=SlideType.AGENDA,
            title="Agenda",
            bullets=sections
        ))

        # Section slides
        for section in sections:
            # Section divider
            outline.add_slide(SlideContent(
                slide_type=SlideType.SECTION,
                title=section
            ))
            # Content placeholder
            outline.add_slide(SlideContent(
                slide_type=SlideType.CONTENT,
                title=section,
                bullets=["Key point 1", "Key point 2", "Key point 3"]
            ))

        # End slide
        outline.add_slide(SlideContent(
            slide_type=SlideType.END,
            title="Thank You",
            subtitle="Questions?"
        ))

        return outline
