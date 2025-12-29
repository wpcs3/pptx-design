"""
Markdown Parser for Presentation Outlines

Converts Markdown content to presentation outline JSON format.
Supports standard Markdown and Marp-style extensions.

Phase 4 Enhancement (2025-12-29):
- Standard Markdown parsing (# titles, ## slides, - bullets)
- Marp frontmatter support (---) for metadata
- Slide separators (---) for explicit slide breaks
- Code blocks for charts/tables
- Two-column layout via special syntax

Usage:
    from pptx_generator.modules.markdown_parser import MarkdownParser, markdown_to_outline

    # Quick conversion
    outline = markdown_to_outline(markdown_text)

    # Full parser with options
    parser = MarkdownParser()
    outline = parser.parse(markdown_text)

Supported Markdown Syntax:
    ---
    title: Presentation Title
    template: consulting_toolkit
    author: Author Name
    ---

    # Title Slide Title

    ## Content Slide Title
    - Bullet point 1
    - Bullet point 2
      - Nested bullet

    ---

    ## Another Slide
    Content paragraph here.

    ```chart:column
    categories: Q1, Q2, Q3, Q4
    series:
      Revenue: 100, 150, 200, 250
      Profit: 20, 35, 50, 70
    ```

    ## Two Column Slide
    ::: columns
    :::: left
    ### Left Header
    - Left bullet 1
    - Left bullet 2
    ::::
    :::: right
    ### Right Header
    - Right bullet 1
    - Right bullet 2
    ::::
    :::

    ## Metrics Slide
    ```metrics
    Revenue: $1.2M
    Growth: 25%
    Users: 10K
    ```

    ## Table Slide
    | Header 1 | Header 2 | Header 3 |
    |----------|----------|----------|
    | Cell 1   | Cell 2   | Cell 3   |
    | Cell 4   | Cell 5   | Cell 6   |
"""

import logging
import re
import yaml
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple

logger = logging.getLogger(__name__)


@dataclass
class ParsedSlide:
    """Represents a parsed slide from Markdown."""
    slide_type: str
    content: Dict[str, Any]
    raw_markdown: str = ""


@dataclass
class ParsedPresentation:
    """Represents a fully parsed presentation."""
    title: str = ""
    template: str = "consulting_toolkit"
    author: str = ""
    metadata: Dict[str, Any] = field(default_factory=dict)
    slides: List[ParsedSlide] = field(default_factory=list)

    def to_outline(self) -> Dict[str, Any]:
        """Convert to presentation outline JSON format."""
        sections = []
        current_section = {"name": "Main", "slides": []}

        for slide in self.slides:
            if slide.slide_type == "section_divider":
                if current_section["slides"]:
                    sections.append(current_section)
                current_section = {
                    "name": slide.content.get("title", "Section"),
                    "slides": []
                }
            else:
                current_section["slides"].append({
                    "slide_type": slide.slide_type,
                    "content": slide.content
                })

        if current_section["slides"]:
            sections.append(current_section)

        return {
            "title": self.title,
            "template": self.template,
            "author": self.author,
            "metadata": self.metadata,
            "sections": sections
        }


class MarkdownParser:
    """
    Parse Markdown content into presentation outlines.

    Supports:
    - YAML frontmatter for metadata
    - Headings for slide structure
    - Lists for bullet points
    - Code blocks for charts, tables, metrics
    - Two-column layouts
    - Markdown tables
    """

    # Regex patterns
    FRONTMATTER_PATTERN = re.compile(r'^---\s*\n(.*?)\n---\s*\n', re.DOTALL)
    SLIDE_SEPARATOR = re.compile(r'\n---\n')
    HEADING_PATTERN = re.compile(r'^(#{1,6})\s+(.+)$', re.MULTILINE)
    BULLET_PATTERN = re.compile(r'^(\s*)-\s+(.+)$', re.MULTILINE)
    CODE_BLOCK_PATTERN = re.compile(r'```(\w+(?::\w+)?)\n(.*?)```', re.DOTALL)
    TABLE_PATTERN = re.compile(r'^\|(.+)\|$\n^\|[-| ]+\|$\n((?:^\|.+\|$\n?)+)', re.MULTILINE)
    COLUMNS_PATTERN = re.compile(r'::: columns\n(.*?)\n:::(?!:)', re.DOTALL)
    COLUMN_PATTERN = re.compile(r':::: (left|right)\n(.*?)\n::::(?!:)', re.DOTALL)

    def __init__(self, default_template: str = "consulting_toolkit"):
        """
        Initialize the Markdown parser.

        Args:
            default_template: Default template to use if not specified in frontmatter.
        """
        self.default_template = default_template

    def parse(self, markdown: str) -> ParsedPresentation:
        """
        Parse Markdown content into a presentation structure.

        Args:
            markdown: Markdown text to parse.

        Returns:
            ParsedPresentation object with slides.
        """
        presentation = ParsedPresentation(template=self.default_template)

        # Extract frontmatter
        markdown = self._parse_frontmatter(markdown, presentation)

        # Split into raw slides by separator or headings
        raw_slides = self._split_into_slides(markdown)

        # Parse each slide
        for raw_slide in raw_slides:
            parsed_slide = self._parse_slide(raw_slide)
            if parsed_slide:
                presentation.slides.append(parsed_slide)

        # Set title from first slide if not in frontmatter
        if not presentation.title and presentation.slides:
            first_slide = presentation.slides[0]
            if first_slide.slide_type == "title_slide":
                presentation.title = first_slide.content.get("title", "")

        logger.info(f"Parsed {len(presentation.slides)} slides from Markdown")
        return presentation

    def _parse_frontmatter(self, markdown: str, presentation: ParsedPresentation) -> str:
        """Extract YAML frontmatter and return remaining content."""
        match = self.FRONTMATTER_PATTERN.match(markdown)
        if match:
            try:
                frontmatter = yaml.safe_load(match.group(1))
                if isinstance(frontmatter, dict):
                    presentation.title = frontmatter.get("title", "")
                    presentation.template = frontmatter.get("template", self.default_template)
                    presentation.author = frontmatter.get("author", "")
                    presentation.metadata = {
                        k: v for k, v in frontmatter.items()
                        if k not in ("title", "template", "author")
                    }
                return markdown[match.end():]
            except yaml.YAMLError as e:
                logger.warning(f"Failed to parse frontmatter: {e}")
        return markdown

    def _split_into_slides(self, markdown: str) -> List[str]:
        """Split Markdown into individual slide content blocks."""
        # First, split by explicit slide separators
        parts = self.SLIDE_SEPARATOR.split(markdown)

        slides = []
        for part in parts:
            part = part.strip()
            if not part:
                continue

            # Check if this part has multiple top-level headings
            headings = list(self.HEADING_PATTERN.finditer(part))
            h1_h2_headings = [h for h in headings if len(h.group(1)) <= 2]

            if len(h1_h2_headings) <= 1:
                slides.append(part)
            else:
                # Split by headings
                for i, heading in enumerate(h1_h2_headings):
                    start = heading.start()
                    end = h1_h2_headings[i + 1].start() if i + 1 < len(h1_h2_headings) else len(part)
                    slide_content = part[start:end].strip()
                    if slide_content:
                        slides.append(slide_content)

        return slides

    def _parse_slide(self, raw_slide: str) -> Optional[ParsedSlide]:
        """Parse a single slide's content."""
        if not raw_slide.strip():
            return None

        # Check for special content types first
        code_blocks = list(self.CODE_BLOCK_PATTERN.finditer(raw_slide))
        tables = list(self.TABLE_PATTERN.finditer(raw_slide))
        columns = self.COLUMNS_PATTERN.search(raw_slide)

        # Extract heading
        heading_match = self.HEADING_PATTERN.search(raw_slide)
        title = ""
        heading_level = 0
        if heading_match:
            heading_level = len(heading_match.group(1))
            title = heading_match.group(2).strip()

        # Determine slide type and parse content
        if code_blocks:
            return self._parse_code_block_slide(raw_slide, code_blocks, title)
        elif tables:
            return self._parse_table_slide(raw_slide, tables[0], title)
        elif columns:
            return self._parse_two_column_slide(raw_slide, columns, title)
        elif heading_level == 1:
            return self._parse_title_slide(raw_slide, title)
        elif heading_level == 3:
            return self._parse_section_slide(raw_slide, title)
        else:
            return self._parse_content_slide(raw_slide, title)

    def _parse_title_slide(self, raw_slide: str, title: str) -> ParsedSlide:
        """Parse a title slide (# heading)."""
        # Look for subtitle (next line or ## subheading)
        lines = raw_slide.split('\n')
        subtitle = ""

        for i, line in enumerate(lines):
            line = line.strip()
            if line.startswith('# '):
                # Check next line for subtitle
                if i + 1 < len(lines):
                    next_line = lines[i + 1].strip()
                    if next_line and not next_line.startswith('#') and not next_line.startswith('-'):
                        subtitle = next_line
                break

        return ParsedSlide(
            slide_type="title_slide",
            content={"title": title, "subtitle": subtitle},
            raw_markdown=raw_slide
        )

    def _parse_section_slide(self, raw_slide: str, title: str) -> ParsedSlide:
        """Parse a section divider slide (### heading)."""
        # Look for subtitle
        subtitle = ""
        lines = raw_slide.split('\n')
        for i, line in enumerate(lines):
            if line.strip().startswith('### '):
                if i + 1 < len(lines):
                    next_line = lines[i + 1].strip()
                    if next_line and not next_line.startswith('#') and not next_line.startswith('-'):
                        subtitle = next_line
                break

        return ParsedSlide(
            slide_type="section_divider",
            content={"title": title, "subtitle": subtitle},
            raw_markdown=raw_slide
        )

    def _parse_content_slide(self, raw_slide: str, title: str) -> ParsedSlide:
        """Parse a standard content slide with bullets."""
        bullets = []
        body_text = []

        # Extract bullets
        for match in self.BULLET_PATTERN.finditer(raw_slide):
            indent = len(match.group(1))
            text = match.group(2).strip()

            if indent == 0:
                bullets.append(text)
            elif indent <= 2 and bullets:
                # Sub-bullet - append to last bullet
                bullets[-1] = bullets[-1] + f" ({text})"
            else:
                bullets.append(text)

        # Extract non-bullet text
        lines = raw_slide.split('\n')
        for line in lines:
            line = line.strip()
            if (line and
                not line.startswith('#') and
                not line.startswith('-') and
                not line.startswith('|') and
                not line.startswith('```')):
                body_text.append(line)

        content = {"title": title}
        if bullets:
            content["bullets"] = bullets
        if body_text and not bullets:
            content["body"] = "\n".join(body_text)

        # Use agenda slide type for agenda-related titles
        slide_type = "title_content"
        if title.lower() in ("agenda", "outline", "topics", "contents"):
            slide_type = "agenda"
            content["items"] = bullets
            content.pop("bullets", None)

        return ParsedSlide(
            slide_type=slide_type,
            content=content,
            raw_markdown=raw_slide
        )

    def _parse_code_block_slide(
        self,
        raw_slide: str,
        code_blocks: List[re.Match],
        title: str
    ) -> ParsedSlide:
        """Parse a slide with code blocks (chart, metrics, etc.)."""
        for block in code_blocks:
            block_type = block.group(1).lower()
            block_content = block.group(2).strip()

            if block_type.startswith("chart:"):
                return self._parse_chart_block(block_type, block_content, title, raw_slide)
            elif block_type == "metrics":
                return self._parse_metrics_block(block_content, title, raw_slide)
            elif block_type == "table":
                return self._parse_yaml_table_block(block_content, title, raw_slide)

        # Default to content slide
        return self._parse_content_slide(raw_slide, title)

    def _parse_chart_block(
        self,
        block_type: str,
        content: str,
        title: str,
        raw_slide: str
    ) -> ParsedSlide:
        """Parse a chart code block."""
        chart_type = block_type.split(":")[1] if ":" in block_type else "column"

        try:
            chart_data = yaml.safe_load(content)
            if not isinstance(chart_data, dict):
                chart_data = {}
        except yaml.YAMLError:
            # Try simple format: categories and series lines
            chart_data = self._parse_simple_chart_format(content)

        # Normalize series format: convert dict to list of dicts
        # YAML format: series: {Revenue: [1,2,3], Target: [4,5,6]}
        # Or string format: series: {Revenue: "1, 2, 3", Target: "4, 5, 6"}
        # Expected format: series: [{"name": "Revenue", "values": [1,2,3]}, ...]
        if "series" in chart_data and isinstance(chart_data["series"], dict):
            series_dict = chart_data["series"]
            normalized_series = []
            for name, values in series_dict.items():
                # Handle comma-separated string format
                if isinstance(values, str):
                    values = [float(v.strip()) for v in values.split(',') if v.strip()]
                elif isinstance(values, (int, float)):
                    values = [float(values)]
                normalized_series.append({"name": name, "values": values})
            chart_data["series"] = normalized_series

        chart_data["type"] = chart_type

        return ParsedSlide(
            slide_type="data_chart",
            content={
                "title": title,
                "chart_data": chart_data
            },
            raw_markdown=raw_slide
        )

    def _parse_simple_chart_format(self, content: str) -> Dict[str, Any]:
        """Parse simple chart format (categories: a, b, c)."""
        result = {"categories": [], "series": []}

        for line in content.split('\n'):
            line = line.strip()
            if not line:
                continue

            if ':' in line:
                key, value = line.split(':', 1)
                key = key.strip().lower()
                values = [v.strip() for v in value.split(',')]

                if key == "categories":
                    result["categories"] = values
                else:
                    # It's a series
                    try:
                        numeric_values = [float(v) for v in values]
                        result["series"].append({
                            "name": key.title(),
                            "values": numeric_values
                        })
                    except ValueError:
                        pass

        return result

    def _parse_metrics_block(
        self,
        content: str,
        title: str,
        raw_slide: str
    ) -> ParsedSlide:
        """Parse a metrics code block."""
        metrics = []

        for line in content.split('\n'):
            line = line.strip()
            if ':' in line:
                label, value = line.split(':', 1)
                metrics.append({
                    "label": label.strip(),
                    "value": value.strip()
                })

        return ParsedSlide(
            slide_type="key_metrics",
            content={
                "title": title or "Key Metrics",
                "metrics": metrics
            },
            raw_markdown=raw_slide
        )

    def _parse_yaml_table_block(
        self,
        content: str,
        title: str,
        raw_slide: str
    ) -> ParsedSlide:
        """Parse a YAML-formatted table block."""
        try:
            table_data = yaml.safe_load(content)
            # Normalize format: extract headers and data/rows
            headers = table_data.get("headers", [])
            data = table_data.get("data", table_data.get("rows", []))
            return ParsedSlide(
                slide_type="table_slide",
                content={
                    "title": title,
                    "headers": headers,
                    "data": data,
                },
                raw_markdown=raw_slide
            )
        except yaml.YAMLError:
            return self._parse_content_slide(raw_slide, title)

    def _parse_table_slide(
        self,
        raw_slide: str,
        table_match: re.Match,
        title: str
    ) -> ParsedSlide:
        """Parse a Markdown table into a table slide."""
        header_row = table_match.group(1)
        data_rows = table_match.group(2)

        headers = [h.strip() for h in header_row.split('|') if h.strip()]
        data = []

        for row_line in data_rows.strip().split('\n'):
            cells = [c.strip() for c in row_line.split('|') if c.strip()]
            if cells:
                data.append(cells)

        return ParsedSlide(
            slide_type="table_slide",
            content={
                "title": title,
                "headers": headers,
                "data": data,
            },
            raw_markdown=raw_slide
        )

    def _parse_two_column_slide(
        self,
        raw_slide: str,
        columns_match: re.Match,
        title: str
    ) -> ParsedSlide:
        """Parse a two-column layout slide."""
        columns_content = columns_match.group(1)

        left_content = {"header": "", "bullets": []}
        right_content = {"header": "", "bullets": []}

        for col_match in self.COLUMN_PATTERN.finditer(columns_content):
            side = col_match.group(1)
            content = col_match.group(2)

            # Extract header
            header_match = re.search(r'^###\s+(.+)$', content, re.MULTILINE)
            if header_match:
                header = header_match.group(1).strip()
            else:
                header = ""

            # Extract bullets
            bullets = []
            for bullet_match in self.BULLET_PATTERN.finditer(content):
                bullets.append(bullet_match.group(2).strip())

            if side == "left":
                left_content = {"header": header, "bullets": bullets}
            else:
                right_content = {"header": header, "bullets": bullets}

        return ParsedSlide(
            slide_type="two_column",
            content={
                "title": title,
                "left": left_content,
                "right": right_content
            },
            raw_markdown=raw_slide
        )


def markdown_to_outline(markdown: str, template: str = "consulting_toolkit") -> Dict[str, Any]:
    """
    Convert Markdown text to presentation outline JSON.

    This is the main convenience function for Markdown → Outline conversion.

    Args:
        markdown: Markdown text to convert.
        template: Template name to use (default: consulting_toolkit).

    Returns:
        Presentation outline dictionary ready for PresentationBuilder.

    Example:
        >>> md = '''
        ... ---
        ... title: Q4 Review
        ... ---
        ...
        ... # Q4 Review
        ... Strategic Analysis
        ...
        ... ## Agenda
        ... - Overview
        ... - Analysis
        ... - Recommendations
        ...
        ... ## Key Metrics
        ... ```metrics
        ... Revenue: $1.2M
        ... Growth: 25%
        ... ```
        ... '''
        >>> outline = markdown_to_outline(md)
        >>> print(outline['title'])
        'Q4 Review'
    """
    parser = MarkdownParser(default_template=template)
    presentation = parser.parse(markdown)
    return presentation.to_outline()


def parse_marp_file(filepath: str) -> Dict[str, Any]:
    """
    Parse a Marp-formatted Markdown file.

    Marp is a popular Markdown presentation framework.
    This function provides compatibility with Marp files.

    Args:
        filepath: Path to the Marp Markdown file.

    Returns:
        Presentation outline dictionary.
    """
    from pathlib import Path

    content = Path(filepath).read_text(encoding='utf-8')
    return markdown_to_outline(content)


# CLI interface for testing
if __name__ == "__main__":
    import json
    import sys

    if len(sys.argv) > 1:
        filepath = sys.argv[1]
        outline = parse_marp_file(filepath)
        print(json.dumps(outline, indent=2))
    else:
        # Demo with sample Markdown
        sample_md = """---
title: Q4 2025 Review
template: consulting_toolkit
author: Strategy Team
---

# Q4 2025 Review
Strategic Business Analysis

## Agenda
- Financial Overview
- Market Analysis
- Strategic Recommendations
- Next Steps

## Financial Highlights

```metrics
Revenue: $1.2M
Growth: 25%
EBITDA: $400K
Customers: 150+
```

## Revenue Trends

```chart:column
categories: Q1, Q2, Q3, Q4
series:
  Revenue: 800000, 950000, 1100000, 1200000
  Profit: 150000, 200000, 280000, 350000
```

## Market Comparison

| Metric | Us | Competitor A | Competitor B |
|--------|-----|--------------|--------------|
| Revenue | $1.2M | $2.1M | $0.8M |
| Growth | 25% | 15% | 30% |
| Market Share | 18% | 32% | 12% |

## Strategic Options

::: columns
:::: left
### Option A: Expand
- Enter new markets
- Increase sales team
- Higher risk, higher reward
::::
:::: right
### Option B: Optimize
- Improve efficiency
- Focus on retention
- Lower risk, steady growth
::::
:::

### Next Steps

## Recommended Actions
- Finalize budget allocation
- Begin market research
- Prepare launch timeline
- Schedule follow-up review

# Thank You
Questions and Discussion
"""

        parser = MarkdownParser()
        presentation = parser.parse(sample_md)
        outline = presentation.to_outline()

        print("Parsed Presentation:")
        print(f"  Title: {presentation.title}")
        print(f"  Template: {presentation.template}")
        print(f"  Slides: {len(presentation.slides)}")
        print()

        for i, slide in enumerate(presentation.slides):
            print(f"  Slide {i + 1}: {slide.slide_type}")
            print(f"    Title: {slide.content.get('title', 'N/A')}")

        print()
        print("JSON Outline:")
        print(json.dumps(outline, indent=2))
