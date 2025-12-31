# Claude Code Prompt: Light Industrial Investment Thesis Presentation

## Task Description

Generate a professional investor pitch deck for a light industrial real estate portfolio acquisition strategy. This presentation will be used to pitch a joint venture structure to institutional investors (US public pension fund and international sovereign wealth fund).

## Input Files

1. **Markdown Outline:** `light_industrial_thesis_outline.md` - Structured slide-by-slide content
2. **Research Document:** `light_industrial_research.md` - Comprehensive investment thesis with market data and citations

## Generation Instructions

```bash
# Activate environment
conda activate pptx-design

# Navigate to project
cd C:\Users\wpcol\claudecode\pptx-design

# Option 1: Generate from markdown outline
python -m pptx_generator generate \
  --request "Create an institutional investor pitch deck for a light industrial real estate joint venture. Use the markdown outline at light_industrial_thesis_outline.md. Style: professional, data-driven, suitable for pension fund and sovereign wealth fund audience." \
  --tone sales_pitch \
  --verbosity detailed

# Option 2: If using the markdown parser directly
python -c "
from pptx_generator.modules.markdown_parser import markdown_to_outline
from pptx_generator.modules.orchestrator import Orchestrator

# Parse the markdown outline
with open('light_industrial_thesis_outline.md', 'r') as f:
    markdown_content = f.read()

outline = markdown_to_outline(markdown_content)

# Generate presentation
orchestrator = Orchestrator()
result = orchestrator.generate(
    outline=outline,
    template='consulting_toolkit',
    options={
        'auto_layout': True,
        'evaluation': True,
        'section_headers': True,
        'tone': 'sales_pitch',
        'verbosity': 'detailed'
    }
)

result.presentation.save('output/light_industrial_thesis.pptx')
print(f'Generated: output/light_industrial_thesis.pptx')
print(f'Quality Grade: {result.evaluation.grade}')
"
```

## Slide Type Mapping

Use these slide types from the pptx_generator catalog:

| Outline Section | Recommended Slide Types |
|-----------------|------------------------|
| Title slide | `title_slide` |
| Section headers | `section_divider` |
| Key metrics (4+ numbers) | `key_metrics` |
| Data tables | `table_slide` |
| Charts with data | `data_chart` |
| Bullet content | `title_content` |
| Comparisons | `two_column` |
| Market deep dives | `title_content` or `two_column` |

## Chart Specifications

The outline includes embedded chart data. Extract and render these charts:

### Supply Pipeline Chart
```json
{
  "chart_type": "column",
  "title": "US Industrial Construction Pipeline",
  "categories": ["2021", "2022", "2023", "2024", "2025"],
  "series": [{"name": "Under Construction (M SF)", "values": [550, 1000, 850, 550, 383]}]
}
```

### Cap Rate Trend Chart
```json
{
  "chart_type": "line",
  "title": "Industrial Cap Rate Trend",
  "categories": ["Q2 2022", "Q4 2022", "Q2 2023", "Q4 2023", "Q2 2024", "Q4 2024"],
  "series": [{"name": "Industrial Cap Rate", "values": [5.22, 5.75, 6.10, 6.40, 6.51, 6.29]}]
}
```

### Manufacturing Construction Chart
```json
{
  "chart_type": "column",
  "title": "US Manufacturing Construction Spending",
  "categories": ["2020", "2021", "2022", "2023", "2024"],
  "series": [{"name": "Annual Spending ($B)", "values": [80, 95, 128, 195, 237]}]
}
```

### Returns Comparison Chart
```json
{
  "chart_type": "bar",
  "title": "10-Year Annualized Returns by Property Type",
  "categories": ["Industrial", "Multifamily", "Retail", "Office"],
  "series": [{"name": "Annualized Return", "values": [12.4, 9.8, 8.2, 6.5]}]
}
```

## Presentation Metadata

```json
{
  "presentation_type": "investment_pitch",
  "title": "Light Industrial Investment Thesis",
  "subtitle": "US Portfolio Acquisition Strategy | Institutional Joint Venture",
  "template": "consulting_toolkit",
  "author": "[GP Sponsor Name]",
  "date": "December 2025",
  "confidentiality": "Confidential - For Discussion Purposes Only"
}
```

## Section Structure (10 Sections, ~35-40 Slides)

1. **Executive Summary** (2-3 slides)
   - Investment opportunity overview
   - Why light industrial, why now

2. **Market Fundamentals** (4-5 slides)
   - US industrial market overview
   - Light industrial vs bulk logistics comparison
   - Supply pipeline contraction (chart)
   - Cap rate evolution (chart)

3. **Structural Demand Drivers** (4 slides)
   - Three megatrends overview
   - E-commerce demand quantified
   - Manufacturing renaissance (chart)
   - Interest rate sensitivity advantage

4. **Target Market Analysis** (7-8 slides)
   - Market selection framework
   - Tier 1 markets overview + Nashville/Tampa deep dives
   - Tier 2 markets overview + DFW deep dive
   - Tier 3 markets overview

5. **Investment Strategy** (4 slides)
   - Portfolio construction approach
   - Return expectations by strategy
   - Value creation levers
   - Historical performance (chart)

6. **Risk Factors & Mitigants** (3 slides)
   - Risk/mitigation matrix
   - Supply risk heat map
   - Defensive characteristics

7. **ESG & Sustainability** (4 slides)
   - ESG competitive advantages
   - Green building economics
   - Solar & EV infrastructure opportunity
   - ESG integration strategy

8. **JV Structure & Governance** (4 slides)
   - Capital structure
   - GP responsibilities
   - Governance framework
   - Fee structure

9. **Conclusion** (3 slides)
   - Investment highlights
   - Target portfolio summary
   - Next steps

10. **Appendix** (optional, as needed)
    - Contact information
    - Additional market data
    - Disclosures

## Style Guidelines

- **Color Palette:** Professional blues, grays; accent with green for positive metrics
- **Typography:** Clean sans-serif; hierarchy clear between titles/subtitles/body
- **Charts:** Minimal gridlines, clear labels, source citations
- **Tables:** Alternating row shading, bold headers
- **Imagery:** If using ComponentLibrary, search for: industrial, warehouse, logistics, map, growth icons

## Quality Targets

- Target PPTEval grade: **B+ or higher**
- All slides should pass layout validation
- Charts should include data labels where appropriate
- Tables should have consistent formatting
- Section dividers between major sections

## Output

Save the generated presentation to:
```
pptx_generator/output/light_industrial_thesis_YYYYMMDD.pptx
```

## Verification

After generation, verify:
1. All 10 sections are present with dividers
2. Charts render correctly with data
3. Tables are properly formatted
4. No placeholder text remains
5. Slide count is within 35-45 range
6. Run PPTEval and confirm grade meets target

---

## Alternative: JSON Outline Format

If the markdown parser encounters issues, here is the equivalent JSON structure for the first section to use as a template:

```json
{
  "presentation_type": "investment_pitch",
  "title": "Light Industrial Investment Thesis",
  "subtitle": "US Portfolio Acquisition Strategy | Institutional Joint Venture",
  "sections": [
    {
      "name": "Executive Summary",
      "slides": [
        {
          "slide_type": "title_slide",
          "content": {
            "title": "Light Industrial Investment Thesis",
            "subtitle": "US Portfolio Acquisition Strategy | Institutional Joint Venture"
          }
        },
        {
          "slide_type": "key_metrics",
          "content": {
            "title": "Investment Opportunity Overview",
            "metrics": [
              {"label": "Structure", "value": "49/49/2 JV"},
              {"label": "Target IRR", "value": "10-15%"},
              {"label": "Small-Bay Vacancy", "value": "3.4%"},
              {"label": "10-Yr Returns", "value": "12.4%"}
            ]
          }
        },
        {
          "slide_type": "title_content",
          "content": {
            "title": "Why Light Industrial, Why Now",
            "bullets": [
              "12.4% ten-year annualized returns—highest among all property types",
              "Small-bay vacancy: 3.4% nationally vs 7.1% overall industrial",
              "Construction pipeline at decade lows (0.3% of stock)",
              "Cap rates repriced 130 bps from 2022 trough—attractive entry",
              "Structural demand from e-commerce, nearshoring, obsolescence"
            ]
          }
        }
      ]
    }
  ]
}
```

Continue this pattern for remaining sections using the markdown outline as the content source.
