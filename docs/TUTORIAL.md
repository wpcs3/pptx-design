# PPTX Design System - Tutorial

This tutorial walks you through creating your first presentation using the PPTX Design System.

## Prerequisites

- Python 3.8+
- LibreOffice (for PDF/PNG conversion)
- Required packages: `python-pptx`, `pillow`, `scikit-image`

## Step 1: Import the Library

```python
from pptx_design import PresentationBuilder
```

## Step 2: Choose a Template

List available templates:

```python
from pptx_design import TemplateRegistry

registry = TemplateRegistry()
print(registry.list_templates())
# ['business_case', 'business_consulting_toolkit', 'due_diligence', 'market_analysis']
```

## Step 3: Create a Builder

```python
builder = PresentationBuilder("consulting_toolkit")
builder.set_company("Your Company Name")
```

## Step 4: Add Slides

### Title Slide

```python
builder.add_title_slide(
    title="Q4 Strategic Review",
    subtitle="December 2025"
)
```

### Agenda

```python
builder.add_agenda([
    "Executive Summary",
    "Market Analysis",
    "Financial Performance",
    "Strategic Recommendations"
])
```

### Content Slides

```python
builder.add_content_slide(
    title="Executive Summary",
    bullets=[
        "Revenue grew 25% year-over-year",
        "Market share increased to 18%",
        "Customer satisfaction at all-time high (NPS: 72)"
    ]
)
```

### Section Dividers

```python
builder.add_section_divider("Market Analysis")
```

### Two-Column Comparison

```python
builder.add_two_column(
    title="Competitive Analysis",
    left_header="Our Strengths",
    left_bullets=["Strong brand", "Technical excellence", "Customer focus"],
    right_header="Market Challenges",
    right_bullets=["Pricing pressure", "New entrants", "Regulatory changes"]
)
```

### Key Metrics

```python
builder.add_metrics(
    title="Key Performance Indicators",
    metrics=[
        {"label": "Revenue", "value": "$12.5M"},
        {"label": "Growth", "value": "+25%"},
        {"label": "NPS", "value": "72"},
        {"label": "Retention", "value": "95%"}
    ]
)
```

### Tables

```python
builder.add_table(
    title="Quarterly Results",
    headers=["Metric", "Q1", "Q2", "Q3", "Q4"],
    data=[
        ["Revenue", "$2.8M", "$3.0M", "$3.2M", "$3.5M"],
        ["Customers", "450", "520", "580", "650"],
        ["ARPU", "$6.2K", "$5.8K", "$5.5K", "$5.4K"]
    ]
)
```

### Charts

```python
builder.add_chart(
    title="Revenue Trend",
    chart_type="column",
    categories=["Q1", "Q2", "Q3", "Q4"],
    series=[
        {"name": "2024", "values": [2.5, 2.7, 2.9, 3.1]},
        {"name": "2025", "values": [2.8, 3.0, 3.2, 3.5]}
    ]
)
```

### Closing Slide

```python
builder.add_end_slide("Thank You", "Questions?")
```

## Step 5: Preview and Save

```python
# Preview structure
print(builder.preview())

# Save to file
builder.save("q4_review.pptx")
```

## Complete Example

```python
from pptx_design import PresentationBuilder

# Create builder
builder = PresentationBuilder("consulting_toolkit")
builder.set_company("Acme Corporation")

# Build presentation
builder.add_title_slide("Strategic Business Review", "Q4 2025")
builder.add_agenda([
    "Executive Summary",
    "Market Position",
    "Financial Performance",
    "2026 Strategy"
])

builder.add_section_divider("Executive Summary")
builder.add_content_slide("Key Highlights", bullets=[
    "Record revenue quarter: $3.5M (+15% QoQ)",
    "Customer base grew to 650 accounts",
    "Launched 3 new product features",
    "Expanded into 2 new markets"
])

builder.add_section_divider("Market Position")
builder.add_two_column("Competitive Landscape",
    left_header="Strengths",
    left_bullets=["Market leader in core segment", "Strong technical team", "High customer satisfaction"],
    right_header="Opportunities",
    right_bullets=["Enterprise segment expansion", "International markets", "Product diversification"]
)

builder.add_section_divider("Financial Performance")
builder.add_metrics("Key Metrics", [
    {"label": "Revenue", "value": "$12.5M"},
    {"label": "Growth", "value": "+25%"},
    {"label": "Margin", "value": "42%"},
    {"label": "ARR", "value": "$15M"}
])

builder.add_chart("Revenue by Quarter",
    chart_type="column",
    categories=["Q1", "Q2", "Q3", "Q4"],
    series=[{"name": "2025", "values": [2.8, 3.0, 3.2, 3.5]}]
)

builder.add_section_divider("2026 Strategy")
builder.add_content_slide("Strategic Priorities", bullets=[
    "Launch enterprise tier (Q1)",
    "Expand APAC presence (Q2)",
    "Release mobile platform (Q3)",
    "Achieve $20M ARR (Q4)"
])

builder.add_end_slide("Thank You", "Questions & Discussion")

# Save
output = builder.save("strategic_review_2025.pptx")
print(f"Saved: {output}")
```

## Using the Content Pipeline

For more complex workflows, use the ContentPipeline:

```python
from pptx_design import ContentPipeline, PresentationBuilder

# Create pipeline
pipeline = ContentPipeline("consulting_toolkit")

# Generate standard outline
outline = pipeline.create_standard_outline(
    title="Product Launch Plan",
    sections=["Market Research", "Go-to-Market", "Timeline", "Budget"],
    purpose="Board approval for Q1 launch"
)

# Get layout assignments
slides = pipeline.prepare_for_rendering(outline)

# Build with assignments
builder = PresentationBuilder("consulting_toolkit")
for slide in slides:
    slide_type = slide["type"]
    content = slide["content"]

    if slide_type == "title":
        builder.add_title_slide(content.get("title"), content.get("subtitle", ""))
    elif slide_type == "agenda":
        builder.add_agenda(content.get("body", []), content.get("title", "Agenda"))
    elif slide_type == "section":
        builder.add_section_divider(content.get("title"))
    elif slide_type == "end":
        builder.add_end_slide(content.get("title"), content.get("subtitle", ""))
    else:
        builder.add_content_slide(content.get("title"), bullets=content.get("bullets", []))

builder.save("product_launch.pptx")
```

## Tips

1. **Check available layouts** before using custom names
2. **Use section dividers** to organize content
3. **Keep bullet points concise** (5-7 words each)
4. **Limit metrics to 4-5** per slide
5. **Preview before saving** to verify structure

## Next Steps

- See [API Reference](API_REFERENCE.md) for complete method documentation
- Check [IMPROVEMENTS.md](IMPROVEMENTS.md) for roadmap
- Explore the `pptx_templates/` directory for template examples
