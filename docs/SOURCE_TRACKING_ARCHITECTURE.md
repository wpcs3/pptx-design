# Source Tracking Architecture for SEC Compliance

This document outlines the recommended architecture for tracking data sources throughout the presentation generation pipeline to meet SEC Marketing Rule requirements.

## Current Limitation

The current pipeline does not maintain source attribution throughout the data flow:
1. Research is conducted and sources are listed at the end
2. Data points are extracted and used in outlines
3. Slides are generated without knowing which data came from which source
4. Footnotes must be manually mapped after the fact

## Proposed Architecture

### Phase 1: Research Output Format

The research module should output structured data with inline source attribution:

```json
{
  "research_output": {
    "title": "Light Industrial Investment Thesis",
    "sections": [
      {
        "name": "Market Fundamentals",
        "data_points": [
          {
            "claim": "National vacancy reached 7.1% in Q3 2025",
            "value": "7.1%",
            "metric_type": "vacancy_rate",
            "source": {
              "id": "commercialcafe_2025_12",
              "name": "CommercialCafe National Industrial Report",
              "date": "December 2025",
              "url": "https://...",
              "page": null,
              "accessed_date": "2025-12-29"
            }
          },
          {
            "claim": "Small-bay vacancy at 3.4%",
            "value": "3.4%",
            "metric_type": "vacancy_rate",
            "segment": "small_bay",
            "source": {
              "id": "commercialcafe_2025_12",
              "name": "CommercialCafe National Industrial Report",
              "date": "December 2025"
            }
          }
        ]
      }
    ],
    "sources_registry": {
      "commercialcafe_2025_12": {
        "full_citation": "CommercialCafe National Industrial Report, December 2025",
        "short_citation": "CommercialCafe (Dec 2025)",
        "url": "https://...",
        "accessed_date": "2025-12-29",
        "type": "market_report"
      }
    }
  }
}
```

### Phase 2: Outline Schema Enhancement

The presentation outline should carry source references:

```json
{
  "slides": [
    {
      "slide_type": "table_slide",
      "content": {
        "title": "US Industrial Market Overview",
        "headers": ["Metric", "National", "Small-Bay"],
        "data": [
          {
            "row": ["Vacancy Rate", "7.1%", "3.4%"],
            "sources": ["commercialcafe_2025_12"]
          },
          {
            "row": ["Asking Rent", "$10.10/SF", "$11.25/SF"],
            "sources": ["cushman_2025_q2"]
          }
        ],
        "footnote_sources": ["commercialcafe_2025_12", "cushman_2025_q2"]
      }
    }
  ],
  "sources_registry": { ... }
}
```

### Phase 3: Template Renderer Integration

The `template_renderer.py` module should:

1. Accept a `sources_registry` parameter
2. For each slide, collect the `footnote_sources` from content
3. Automatically add formatted footnotes to the footer placeholder

```python
class TemplateRenderer:
    def render_slide(self, slide_data, sources_registry):
        # ... render content ...

        # Add footnotes if sources specified
        if 'footnote_sources' in slide_data.get('content', {}):
            source_ids = slide_data['content']['footnote_sources']
            footnote_text = self._format_footnotes(source_ids, sources_registry)
            self._add_footnote(slide, footnote_text)

    def _format_footnotes(self, source_ids, registry):
        citations = [registry[sid]['short_citation'] for sid in source_ids]
        return "Sources: " + "; ".join(citations) + "."
```

### Phase 4: SEC Compliance Features

#### 4.1 Source Validation
- Verify all data points have source attribution
- Flag any "orphan" claims without sources
- Generate compliance report

#### 4.2 Citation Formats
Support multiple citation formats:
- Short form: "CBRE (H1 2024)"
- Full form: "CBRE Cap Rate Survey, H1 2024"
- Academic: "CBRE. (2024). Cap Rate Survey, H1 2024."

#### 4.3 Disclosure Generation
Auto-generate required disclosures:
- Performance data disclaimers
- Forward-looking statement warnings
- Source list appendix slide

### Implementation Roadmap

| Phase | Description | Effort |
|-------|-------------|--------|
| 1 | Structured research output format | Medium |
| 2 | Outline schema enhancement | Low |
| 3 | Renderer integration | Medium |
| 4 | Compliance features | High |

### Example Footnote Formats

**Standard (current implementation):**
```
Sources: CBRE Cap Rate Survey, H1 2024; CommercialCafe National Industrial Report, December 2025.
```

**Abbreviated:**
```
Sources: CBRE (H1 2024); CommercialCafe (Dec 2025).
```

**Numbered (for complex slides):**
```
¹CBRE Cap Rate Survey, H1 2024. ²CommercialCafe National Industrial Report, December 2025.
```

### Configuration

Add to `pptx_generator/config/compliance_config.json`:

```json
{
  "footnotes": {
    "enabled": true,
    "format": "standard",
    "font_size": 6,
    "alignment": "right",
    "position": "bottom"
  },
  "disclosures": {
    "performance_disclaimer": true,
    "forward_looking_warning": true,
    "source_appendix": true
  },
  "validation": {
    "require_sources": true,
    "max_data_age_days": 365,
    "allowed_source_types": ["market_report", "regulatory_filing", "academic"]
  }
}
```

## SEC Marketing Rule Considerations

Per SEC Rule 206(4)-1 (Marketing Rule):

1. **Fair and balanced presentation**: All material facts must be disclosed
2. **Source disclosure**: Third-party data must be attributed
3. **Performance data**: Must include appropriate time periods and disclosures
4. **Hypothetical projections**: Must include clear disclaimers

This architecture enables:
- Automatic source tracking from research to final slide
- Consistent footnote formatting across all presentations
- Audit trail for compliance review
- Easy updates when source data changes
