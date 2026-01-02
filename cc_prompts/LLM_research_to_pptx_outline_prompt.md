# Prompt for Reformatting Deep Research into PPTX Generator Format

Use this prompt with Claude or ChatGPT to reorganize deep research into a structured JSON format for the pptx-generator.

---

```
I have deep research on [TOPIC] that I need to reorganize into a structured JSON format for my PowerPoint presentation generator.

## About the PPTX Generator

The system generates professional investor presentations from structured JSON outlines. It uses template-based rendering with these slide types:

### Available Slide Types:

1. **title_slide** - Opening slide with title and subtitle
2. **key_metrics** - 3-4 key statistics displayed prominently
3. **title_content** - Title with bullet points (most common)
4. **section_divider** - Section break slide with background image
5. **two_column** - Side-by-side comparison (pros/cons, before/after)
6. **data_chart** - Charts with data (column, bar, line, pie)
7. **table_slide** - Data tables with headers and rows

### JSON Schema Structure:

```json
{
  "presentation_type": "investment_pitch",
  "title": "Presentation Title",
  "subtitle": "Subtitle | Additional Context",
  "template": "consulting_toolkit",
  "sections": [
    {
      "name": "Section Name",
      "slides": [
        {
          "slide_type": "title_content",
          "content": {
            "title": "Slide Title (max 60 chars)",
            "takeaway": "One-sentence key insight for this slide",
            "bullets": [
              "Bullet point with specific data or insight",
              "Include numbers, percentages, dollar amounts where possible",
              "Maximum 6-7 bullets per slide"
            ]
          }
        }
      ]
    }
  ]
}
```

### Slide Content Guidelines:

**title_content:**
```json
{
  "slide_type": "title_content",
  "content": {
    "title": "Clear, Concise Title",
    "takeaway": "The key insight readers should remember",
    "bullets": ["Point 1", "Point 2", "Point 3"]
  }
}
```

**key_metrics:**
```json
{
  "slide_type": "key_metrics",
  "content": {
    "title": "Investment Highlights",
    "takeaway": "Summary of key metrics",
    "metrics": [
      {"label": "Target IRR", "value": "12-15%"},
      {"label": "Occupancy", "value": "95%+"},
      {"label": "Cap Rate", "value": "5.5-6.0%"},
      {"label": "Rent Growth", "value": "4-5%"}
    ]
  }
}
```

**two_column:**
```json
{
  "slide_type": "two_column",
  "content": {
    "title": "SFR vs. Traditional Multifamily",
    "takeaway": "SFR offers distinct advantages for certain demographics",
    "left_column": {
      "header": "Single Family Rental",
      "bullets": ["Point 1", "Point 2", "Point 3"]
    },
    "right_column": {
      "header": "Traditional Apartments",
      "bullets": ["Point 1", "Point 2", "Point 3"]
    }
  }
}
```

**table_slide:**
```json
{
  "slide_type": "table_slide",
  "content": {
    "title": "Target Market Comparison",
    "takeaway": "Sunbelt markets offer best risk-adjusted returns",
    "headers": ["Market", "Population Growth", "Job Growth", "Rent Growth"],
    "data": [
      ["Phoenix", "2.1%", "3.2%", "5.4%"],
      ["Tampa", "1.8%", "2.9%", "4.8%"],
      ["Austin", "2.4%", "3.8%", "6.1%"]
    ]
  }
}
```

**data_chart:**
```json
{
  "slide_type": "data_chart",
  "content": {
    "title": "SFR Rent Growth vs. Apartments",
    "takeaway": "SFR consistently outperforms apartment rent growth",
    "chart_data": {
      "type": "column",
      "categories": ["2020", "2021", "2022", "2023", "2024"],
      "series": [
        {"name": "SFR", "values": [3.2, 8.5, 12.1, 5.2, 4.1]},
        {"name": "Apartments", "values": [2.1, 6.2, 9.8, 3.1, 2.5]}
      ]
    }
  }
}
```

**section_divider:**
```json
{
  "slide_type": "section_divider",
  "content": {
    "title": "Market Opportunity",
    "subtitle": "Understanding the SFR Landscape"
  }
}
```

## Recommended Section Structure for Investment Pitch:

1. **Executive Summary** (2-3 slides)
   - title_slide, key_metrics, title_content (thesis summary)

2. **Market Opportunity** (4-6 slides)
   - section_divider, market size/growth, demand drivers, demographic trends

3. **Investment Strategy** (3-5 slides)
   - section_divider, target criteria, geographic focus, acquisition approach

4. **Target Markets** (3-5 slides)
   - Market comparison table, top market deep-dives

5. **Competitive Positioning** (2-3 slides)
   - two_column comparing advantages, competitive landscape

6. **Financial Projections** (3-4 slides)
   - key_metrics for returns, table for assumptions, charts for projections

7. **Risk Factors & Mitigation** (2-3 slides)
   - two_column with risks and mitigations

8. **ESG / Impact** (1-2 slides) - if applicable

9. **Conclusion** (2-3 slides)
   - Investment highlights, next steps, contact

## Your Task:

Please reorganize my research into this JSON format. Requirements:

1. **Create 35-45 slides** organized into logical sections
2. **Every slide must have a "takeaway"** - the one key insight
3. **Include specific data points** - percentages, dollar amounts, growth rates
4. **Use the right slide type** for each content:
   - Comparisons → two_column
   - 3-4 key stats → key_metrics
   - Trend data → data_chart
   - Multi-row data → table_slide
   - Narrative/bullets → title_content
5. **Cite sources** where possible (will be added to footnotes)
6. **Keep titles under 60 characters**
7. **Maximum 6-7 bullets per slide** (concise, not verbose)

Here is my research to reorganize:

[PASTE YOUR RESEARCH HERE]
```

---

## Usage Notes

1. Replace `[TOPIC]` with your specific topic (e.g., "Single Family Built-for-Rent (SFR/BFR) communities")
2. Paste your deep research at the end where indicated
3. The AI will output a JSON file you can save as `your_topic_outline.json`
4. Run the generator with: `python generate_formatted_v30.py` (after updating paths)

## Chart Types Available

- `column` - Vertical bar chart (most common)
- `bar` - Horizontal bar chart
- `line` - Line chart for trends over time
- `pie` - Pie chart for proportions

## Tips for Better Results

- Ask the AI to "include specific numbers and percentages from the research"
- Request "no more than 6 bullets per slide"
- Ask for "a takeaway sentence for every slide"
- If the output is too long, ask to "consolidate into fewer, denser slides"
