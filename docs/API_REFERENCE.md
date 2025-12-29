# PPTX Design System - API Reference

## Overview

The PPTX Design System provides a unified API for creating professional PowerPoint presentations using template-based generation.

## Installation

```bash
# Clone the repository
git clone <repo-url>
cd pptx-design

# Install dependencies
pip install -r requirements.txt
```

## Quick Start

```python
from pptx_design import PresentationBuilder

# Create a presentation
builder = PresentationBuilder("consulting_toolkit")
builder.add_title_slide("Quarterly Review", "Q4 2025")
builder.add_agenda(["Overview", "Analysis", "Recommendations"])
builder.add_content_slide("Key Findings", bullets=["Finding 1", "Finding 2"])
builder.save("presentation.pptx")
```

---

## PresentationBuilder

The main class for building presentations.

### Constructor

```python
PresentationBuilder(template: str, registry: TemplateRegistry = None)
```

**Parameters:**
- `template`: Template name (e.g., "consulting_toolkit") or path to .pptx file
- `registry`: Optional TemplateRegistry instance

**Example:**
```python
builder = PresentationBuilder("consulting_toolkit")
# or
builder = PresentationBuilder("path/to/template.pptx")
```

### Methods

#### Metadata

```python
builder.set_company(name: str) -> PresentationBuilder
builder.set_author(author: str) -> PresentationBuilder
builder.set_date(date: str) -> PresentationBuilder
```

#### Slide Addition

| Method | Description |
|--------|-------------|
| `add_title_slide(title, subtitle, layout)` | Add cover slide |
| `add_section_divider(title, layout)` | Add section break |
| `add_agenda(items, title, layout)` | Add agenda slide |
| `add_content_slide(title, body, bullets, layout)` | Add content slide |
| `add_two_column(title, left_*, right_*, layout)` | Add comparison slide |
| `add_metrics(title, metrics, layout)` | Add KPI boxes |
| `add_table(title, headers, data, layout)` | Add data table |
| `add_chart(title, chart_type, categories, series, layout)` | Add chart |
| `add_blank(layout)` | Add blank slide |
| `add_end_slide(title, subtitle, layout)` | Add closing slide |

#### Building

```python
builder.build() -> Presentation  # Build and return Presentation object
builder.save(output_path) -> Path  # Build and save to file
builder.preview() -> str  # Get text preview of structure
```

### Examples

**Basic Presentation:**
```python
from pptx_design import PresentationBuilder

builder = PresentationBuilder("consulting_toolkit")
builder.set_company("Acme Corp")

builder.add_title_slide("Strategic Review", "2025 Planning")
builder.add_agenda(["Current State", "Opportunities", "Recommendations"])
builder.add_section_divider("Current State")
builder.add_content_slide("Market Position", bullets=[
    "Leading market share in core segment",
    "Strong brand recognition",
    "Expanding customer base"
])
builder.add_end_slide("Thank You", "Questions?")

builder.save("strategic_review.pptx")
```

**With Charts and Tables:**
```python
builder.add_chart("Revenue Growth",
    chart_type="column",
    categories=["Q1", "Q2", "Q3", "Q4"],
    series=[
        {"name": "2024", "values": [100, 120, 140, 160]},
        {"name": "2025", "values": [150, 180, 200, 220]}
    ]
)

builder.add_table("Feature Comparison",
    headers=["Feature", "Plan A", "Plan B", "Plan C"],
    data=[
        ["Users", "10", "50", "Unlimited"],
        ["Storage", "5GB", "50GB", "500GB"],
        ["Support", "Email", "Phone", "24/7"]
    ]
)
```

---

## TemplateRegistry

Catalog of available templates with metadata.

### Constructor

```python
TemplateRegistry(registry_path: Path = None)
```

### Methods

```python
registry.list_templates() -> List[str]  # List all template names
registry.get_template(name) -> Dict  # Get template info
registry.get_layouts(name) -> List[str]  # Get layout names
registry.find_template(use_case, min_layouts) -> str  # Find matching template
registry.refresh() -> None  # Rebuild registry from templates
registry.save_registry(path) -> Path  # Save to JSON
```

### Example

```python
from pptx_design import TemplateRegistry

registry = TemplateRegistry()
print(f"Available templates: {registry.list_templates()}")

info = registry.get_template("consulting_toolkit")
print(f"Layouts: {info['layout_names']}")
print(f"Use cases: {info.get('use_cases', [])}")
```

---

## ContentPipeline

Separates content generation from layout selection.

### Constructor

```python
ContentPipeline(template: str = None)
```

### Methods

```python
pipeline.parse_markdown(markdown: str) -> PresentationOutline
pipeline.parse_json(json_str: str) -> PresentationOutline
pipeline.prepare_for_rendering(outline) -> List[Dict]
pipeline.create_standard_outline(title, sections, purpose, audience) -> PresentationOutline
```

### Example

```python
from pptx_design import ContentPipeline, PresentationBuilder

# From markdown
pipeline = ContentPipeline("consulting_toolkit")
outline = pipeline.parse_markdown("""
# Product Strategy

## Executive Summary
- Key insight 1
- Key insight 2

## Market Analysis
- TAM: $50B
- Growth: 15% CAGR
""")

# Get layout assignments
slides = pipeline.prepare_for_rendering(outline)

# Use with builder
builder = PresentationBuilder("consulting_toolkit")
for slide in slides:
    if slide["type"] == "content":
        content = slide["content"]
        builder.add_content_slide(
            content.get("title", ""),
            bullets=content.get("bullets", [])
        )
builder.save("output.pptx")
```

---

## VisualTester

Automated visual testing using SSIM comparison.

### Constructor

```python
VisualTester(
    output_dir: Path = None,
    reference_dir: Path = None,
    threshold: float = 0.90
)
```

### Methods

```python
tester.compare_images(img1, img2) -> Tuple[float, Path]
tester.pptx_to_png(pptx_path, output_dir, slide_index) -> Path
tester.test_slide(description_path, reference_path, template, threshold) -> TestResult
tester.run_suite(test_cases, suite_name) -> TestSuite
tester.create_reference(pptx_path, output_path, slide_index) -> Path
```

### Example

```python
from pptx_design import VisualTester, quick_compare

# Quick comparison
score = quick_compare("generated.png", "reference.png")
print(f"SSIM: {score:.4f}")

# Full testing
tester = VisualTester(threshold=0.85)
result = tester.test_slide(
    description_path=Path("slide.json"),
    reference_path=Path("reference.png"),
    template="consulting_toolkit"
)
print(result)  # [PASS] slide: SSIM=0.9234 (threshold=0.85)
```

---

## Data Classes

### SlideContent

```python
@dataclass
class SlideContent:
    slide_type: SlideType
    title: str = ""
    subtitle: str = ""
    body: str = ""
    bullets: List[str] = field(default_factory=list)
    data: Dict[str, Any] = field(default_factory=dict)
    notes: str = ""
```

### SlideType (Enum)

```python
class SlideType(Enum):
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
```

### PresentationOutline

```python
@dataclass
class PresentationOutline:
    title: str
    purpose: str
    audience: str
    slides: List[SlideContent]
    metadata: Dict[str, Any]
```

---

## Available Templates

| Template | Layouts | Best For |
|----------|---------|----------|
| `consulting_toolkit` | 31 | Strategy, analysis, business review |
| `business_case` | 30 | Proposals, investment decisions |
| `due_diligence` | 29 | Audits, reviews, analysis |
| `market_analysis` | 27 | Research, competitive analysis |

---

## Layout Names

Common layouts available across templates:

| Layout | Description |
|--------|-------------|
| `Frontpage` | Title/cover slide |
| `Default` | Standard content slide |
| `Agenda` | Table of contents |
| `Section breaker` | Section divider |
| `1/2 grey` | Two-column comparison |
| `1/3 grey` | Sidebar left |
| `2/3 grey` | Sidebar right |
| `Blank` | Empty slide |
| `End` | Closing slide |

---

## Error Handling

```python
try:
    builder = PresentationBuilder("nonexistent_template")
except FileNotFoundError as e:
    print(f"Template not found: {e}")

try:
    builder.save("/invalid/path/presentation.pptx")
except PermissionError as e:
    print(f"Cannot save: {e}")
```

---

## Best Practices

1. **Use templates consistently** - Pick one template per deck for visual consistency
2. **Fill all placeholders** - Empty placeholders may show default text
3. **Test visually** - Use VisualTester to verify output matches expectations
4. **Keep content concise** - Bullet points work better than paragraphs
5. **Use appropriate layouts** - Match content type to layout type

---

## Version History

- **1.0.0** (2025-12-27)
  - Initial release
  - PresentationBuilder with fluent API
  - TemplateRegistry for template management
  - ContentPipeline for content/layout separation
  - VisualTester for automated testing
