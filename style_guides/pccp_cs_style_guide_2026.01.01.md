# PCCP Investor Presentation Style Guide

**Reference Document:** `Light_Industrial_Thesis_vFinal.pptx`
**Last Updated:** 2026-01-01
**Version:** 2.0
**Style Guide ID:** `pccp_cs_style_guide_2026.01.01`

---

## Overview

This style guide codifies the visual language, structure, and formatting standards for PCCP investor presentations. All generated presentations should match this specification for brand consistency. This document is used by the `pptx_generator` for automated style compliance checking.

---

## 1. Presentation Dimensions

### 1.1 Slide Size

| Property | Value |
|----------|-------|
| **Width** | 11.0 inches |
| **Height** | 8.5 inches |
| **Aspect Ratio** | Landscape (approximately 1.29:1) |
| **Orientation** | Landscape |

---

## 2. Master Slide Layouts

### 2.1 Available Layouts

| Layout Name | Usage | Slide Count (Reference) |
|-------------|-------|-------------------------|
| **Frontpage** | Title slide with full-bleed photo | 1 |
| **Bullet Content** | Standard content with bullets/metrics | 16 |
| **Section Title Slide** | Section dividers with full-bleed photo | 9 |
| **Table** | Data tables with headers | 12 |
| **Chart** | Charts with narrative | 4 |
| **Side by Side List** | Two-column comparison | 1 |
| **Contact** | Contact information | 1 |
| **Disclaimers** | Legal disclosures | 1 |
| **End** | Closing slide with logo | 1 |

### 2.2 Layout Selection Rules

| Content Type | Required Layout |
|--------------|-----------------|
| Opening slide with title | `Frontpage` |
| New section introduction | `Section Title Slide` |
| Bullet points (3-6 items) | `Bullet Content` |
| Key metrics (3-4 numbers) | `Bullet Content` (with metric boxes) |
| Data table | `Table` |
| Bar/Line/Column chart | `Chart` |
| Two-column comparison | `Side by Side List` |
| Contact/Thank you | `Contact` |
| Legal text | `Disclaimers` |
| Final closing | `End` |

---

## 3. Color Palette

### 3.1 Primary Colors

| Name | Hex | RGB | Usage |
|------|-----|-----|-------|
| **PCCP Navy** | `#051C2C` | 5, 28, 44 | Table headers, primary branding |
| **Body Text Dark** | `#061F32` | 6, 31, 50 | Body text, chart subtitles |
| **White** | `#FFFFFF` | 255, 255, 255 | Backgrounds, text on dark |
| **Light Gray** | `#F5F5F5` | 245, 245, 245 | Alternating table rows |
| **Medium Gray** | `#A6A6A6` | 166, 166, 166 | Section labels, footer text |
| **Gridline Gray** | `#D9D9D9` | 217, 217, 217 | Chart gridlines |

### 3.2 Chart Colors

| Name | Hex | Usage |
|------|-----|-------|
| **Chart Primary** | `#051C2C` | Primary bar/column fill |
| **Chart Secondary** | `#4A90A4` | Secondary series |
| **Chart Accent** | `#7FB3D5` | Tertiary series |

### 3.3 Metric Box Colors

| Name | Hex | Usage |
|------|-----|-------|
| **Metric Box Background** | `#051C2C` | Rounded rectangle fill |
| **Metric Box Text** | `#FFFFFF` | Value and label text |

---

## 4. Typography

### 4.1 Font Stack

| Element | Font | Size (pt) | Weight | Color |
|---------|------|-----------|--------|-------|
| **Slide Title** | Arial | 32 | Regular | Inherit from master |
| **Slide Subtitle** | Arial | 14-18 | Regular | `#061F32` |
| **Metric Value** | Arial | 28 | Bold | `#FFFFFF` |
| **Metric Label** | Arial | 14 | Regular | `#FFFFFF` |
| **Section Label** | Arial | 9 | Regular | `#A6A6A6` |
| **Body Bullets** | Arial | 12-14 | Regular | `#061F32` |
| **Table Header** | Arial | 10-11 | Bold | `#FFFFFF` |
| **Table Body** | Arial | 10 | Regular | `#061F32` |
| **Source Citation** | Arial | 6 | Regular | `#A6A6A6` |
| **Footer** | Arial | 6-8 | Regular | `#A6A6A6` |

### 4.2 Text Formatting Rules

- **Bullet points**: Use filled circle bullets, Slate color
- **Category headers in bullets**: ALL CAPS followed by em-dash (—)
- **Numbers**: Right-aligned in tables, centered in metric boxes
- **Percentages**: Include % symbol, no decimal if whole number
- **Currency**: USD format with $ prefix, comma separators

---

## 5. Slide-by-Slide Layout Specifications

### 5.1 Frontpage (Title Slide)

```
┌──────────────────────────────────────────────────────────────┐
│  [LOGO - white version]                                       │
│  Position: (0.4", 0.4")                                       │
│  Size: 2.5" x 1.796" (aspect ratio ~1.39:1)                   │
│                                                               │
│  ┌────────────────────────────────────────┐                  │
│  │ TITLE                                  │ Position: (0.4", 3.2")
│  │ 32pt Bold White                        │ Size: 8.0" x 1.8"
│  │                                        │                   │
│  │ Subtitle | Context                     │ Position: (0.4", 5.2")
│  │ 18pt Regular White                     │ Size: 8.0" x 1.1"
│  │ Date                                   │                   │
│  └────────────────────────────────────────┘                  │
│                                                               │
│  [FULL-BLEED PHOTOGRAPHY - Position: (0", 0"), Size: 11" x 8.5"]
└──────────────────────────────────────────────────────────────┘
```

**Specifications:**
| Element | Position | Size | Notes |
|---------|----------|------|-------|
| Background Image | (0", 0") | 11" x 8.5" | Full-bleed |
| Logo Placeholder | (0.4", 0.4") | 2.5" x 1.796" | White version |
| Title | (0.4", 3.2") | 8.0" x 1.8" | Bold, white |
| Subtitle | (0.4", 5.2") | 8.0" x 1.1" | Regular, white |

### 5.2 Section Title Slide

```
┌──────────────────────────────────────────────────────────────┐
│                                                               │
│  [FULL-BLEED PHOTOGRAPHY]                                     │
│  Position: (0", 0"), Size: 11" x 8.5"                         │
│                                                               │
│                                                               │
│                                                               │
│  ┌────────────────────────────────────────┐                  │
│  │ Section Title                          │ Position: (0.4", 5.5")
│  │ 32pt Bold White on dark overlay        │ Size: 10.2" x 1.0"
│  └────────────────────────────────────────┘                  │
│                                                               │
└──────────────────────────────────────────────────────────────┘
```

**Specifications:**
| Element | Position | Size | Notes |
|---------|----------|------|-------|
| Background Image | (0", 0") | 11" x 8.5" | Full-bleed |
| Title | (0.4", 5.5") | 10.2" x 1.0" | Bold, white, bottom-left |

### 5.3 Bullet Content Slide

```
┌──────────────────────────────────────────────────────────────┐
│  Title Area                              Section Label        │
│  Position: (0.4", 0.4")                  Position: (7.1", 0.2")
│  Size: 10.199" x 1.0"                    Size: 3.5" x 0.2"   │
│  32pt                                    9pt, #A6A6A6        │
│─────────────────────────────────────────────────────────────│
│  Subtitle/Thesis Statement                                    │
│  Position: (0.4", 1.7")                                       │
│  Size: 10.2" x 0.8"                                           │
│  14-18pt, #061F32                                             │
│                                                               │
│  Content Area                                                 │
│  Position: (0.399", 2.8")                                     │
│  Size: 10.201" x 4.7"                                         │
│                                                               │
│  • CATEGORY HEADER — Body text continues with details        │
│  • CATEGORY HEADER — Another point with supporting metrics   │
│                                                               │
│                                                               │
│  Source Citation                         CONFIDENTIAL PCCP...│
│  Position: (0.399", 7.5")                                     │
│  Size: 10.2" x 0.399"                                         │
│  6pt, #A6A6A6                                                 │
└──────────────────────────────────────────────────────────────┘
```

**Specifications:**
| Element | Position | Size | Font |
|---------|----------|------|------|
| Title | (0.4", 0.4") | 10.199" x 1.0" | 32pt |
| Section Label | (7.1", 0.2") | 3.5" x 0.2" | 9pt, #A6A6A6 |
| Subtitle | (0.4", 1.7") | 10.2" x 0.8" | 14-18pt |
| Content | (0.399", 2.8") | 10.201" x 4.7" | 12-14pt |
| Footer | (0.399", 7.5") | 10.2" x 0.399" | 6pt, #A6A6A6 |

### 5.4 Key Metrics Slide (Using Bullet Content Layout)

```
┌──────────────────────────────────────────────────────────────┐
│  Title                                   Section Label        │
│                                                               │
│  Subtitle/Thesis Statement                                    │
│                                                               │
│  ┌─────────────┐ ┌─────────────┐ ┌─────────────┐ ┌─────────────┐
│  │   VALUE     │ │   VALUE     │ │   VALUE     │ │   VALUE     │
│  │   28pt Bold │ │   28pt Bold │ │   28pt Bold │ │   28pt Bold │
│  │   Label     │ │   Label     │ │   Label     │ │   Label     │
│  │   14pt      │ │   14pt      │ │   14pt      │ │   14pt      │
│  └─────────────┘ └─────────────┘ └─────────────┘ └─────────────┘
│                                                               │
│  Metric Box Specifications:                                   │
│  - Shape: Rounded Rectangle                                   │
│  - Fill: #051C2C (PCCP Navy)                                 │
│  - Size: 2.4" x 1.2" each                                    │
│  - Spacing: ~0.2" between boxes                              │
│  - Positions: (0.399", 4.2"), (3.0", 4.2"),                 │
│               (5.6", 4.2"), (8.2", 4.2")                     │
│                                                               │
│  Source Citation                                              │
└──────────────────────────────────────────────────────────────┘
```

**Metric Box Specifications:**
| Property | Value |
|----------|-------|
| Shape | Rounded Rectangle (AUTO_SHAPE) |
| Fill Color | `#051C2C` |
| Width | 2.4 inches |
| Height | 1.2 inches |
| Horizontal Spacing | 0.2 inches gap |
| Value Font | Arial, 28pt, Bold, White |
| Label Font | Arial, 14pt, Regular, White |

**Metric Box Positions (4-column layout):**
| Box # | Left Position | Top Position |
|-------|---------------|--------------|
| 1 | 0.399" | 4.2" |
| 2 | 3.0" | 4.2" |
| 3 | 5.6" | 4.2" |
| 4 | 8.2" | 4.2" |

### 5.5 Table Slide

```
┌──────────────────────────────────────────────────────────────┐
│  Title                                   Section Label        │
│                                                               │
│  Subtitle/Thesis Statement                                    │
│                                                               │
│  ┌───────────────────────────────────────────────────────────┐
│  │ Header 1  │ Header 2  │ Header 3  │ Header 4  │ Header 5 │
│  │ #051C2C bg, White text, Bold                              │
│  ├───────────────────────────────────────────────────────────┤
│  │ Data      │ Data      │ Data      │ Data      │ Data     │ #FFFFFF
│  │ Data      │ Data      │ Data      │ Data      │ Data     │ #F5F5F5
│  │ Data      │ Data      │ Data      │ Data      │ Data     │ #FFFFFF
│  │ Data      │ Data      │ Data      │ Data      │ Data     │ #F5F5F5
│  └───────────────────────────────────────────────────────────┘
│                                                               │
│  Source Citation                                              │
└──────────────────────────────────────────────────────────────┘
```

**Table Specifications:**
| Property | Value |
|----------|-------|
| Header Row Background | `#051C2C` |
| Header Row Text | White, Bold |
| Data Row 1 (odd) | `#FFFFFF` |
| Data Row 2 (even) | `#F5F5F5` |
| Cell Margin Left | 0.1 inches |
| Cell Margin Right | 0.1 inches |
| Cell Margin Top | 0.05 inches |
| Cell Margin Bottom | 0.05 inches |
| Border Style | None (no visible borders) |
| Text Alignment (text columns) | Left |
| Text Alignment (number columns) | Right |

### 5.6 Chart Slide

```
┌──────────────────────────────────────────────────────────────┐
│  Title                                   Section Label        │
│  Position: (0.4", 0.4")                  9pt, #A6A6A6        │
│  Size: 10.199" x 1.0"                                         │
│─────────────────────────────────────────────────────────────│
│  Subtitle/Thesis Statement                                    │
│  Position: (0.4", 1.7")                                       │
│  Size: 10.2" x 0.8"                                           │
│                                                               │
│  ┌───────────────────────────────────────────────────────────┐
│  │                                                           │
│  │                    CHART AREA                             │
│  │  Position: (0.399", 2.7")                                 │
│  │  Size: 10.2" x 4.0"                                       │
│  │                                                           │
│  │  - Major horizontal gridlines: ON                         │
│  │  - Gridline width: 0.5pt                                  │
│  │  - Gridline color: #D9D9D9                                │
│  │  - Value axis tick marks: NONE                            │
│  │  - Category axis tick marks: NONE                         │
│  │                                                           │
│  └───────────────────────────────────────────────────────────┘
│                                                               │
│  Chart Narrative/Interpretation                               │
│  Position: (0.399", 6.8")                                     │
│  Size: 10.201" x 0.7"                                         │
│  14pt, #061F32                                                │
│                                                               │
│  Source Citation                                              │
│  Position: (0.399", 7.5")                                     │
│  6pt, #A6A6A6                                                 │
└──────────────────────────────────────────────────────────────┘
```

**Chart Specifications:**
| Property | Value |
|----------|-------|
| Chart Position | (0.399", 2.7") |
| Chart Size | 10.2" x 4.0" |
| Major Horizontal Gridlines | **ON** |
| Gridline Width | **0.5pt** |
| Gridline Color | **#D9D9D9** |
| Value Axis Major Tick Mark | **NONE** |
| Value Axis Minor Tick Mark | **NONE** |
| Category Axis Major Tick Mark | **NONE** |
| Category Axis Minor Tick Mark | **NONE** |
| Data Labels | As needed, 10pt |
| Legend Position | Below chart or right |

### 5.7 Side by Side List (Two-Column)

```
┌──────────────────────────────────────────────────────────────┐
│  Title                                   Section Label        │
│                                                               │
│  Subtitle/Thesis Statement                                    │
│                                                               │
│  ┌────────────────────────┐  ┌────────────────────────┐      │
│  │ Column A Title         │  │ Column B Title         │      │
│  │ Bold, 16pt             │  │ Bold, 16pt             │      │
│  │                        │  │                        │      │
│  │ • CATEGORY — Detail    │  │ • CATEGORY — Detail    │      │
│  │ • CATEGORY — Detail    │  │ • CATEGORY — Detail    │      │
│  │ • CATEGORY — Detail    │  │ • CATEGORY — Detail    │      │
│  └────────────────────────┘  └────────────────────────┘      │
│                                                               │
│  Source Citation                                              │
└──────────────────────────────────────────────────────────────┘
```

### 5.8 Contact Slide

```
┌──────────────────────────────────────────────────────────────┐
│  [LOGO - top left]                                            │
│                                                               │
│                    Thank You                                  │
│                    32pt Bold, centered                        │
│                                                               │
│                    Contact Information | Investor Relations   │
│                    18pt, centered                             │
│                                                               │
│                    [Contact Details]                          │
│                                                               │
└──────────────────────────────────────────────────────────────┘
```

### 5.9 End Slide (Logo Close)

```
┌──────────────────────────────────────────────────────────────┐
│                                                               │
│  [FULL-BLEED PHOTOGRAPHY]                                     │
│                                                               │
│                                                               │
│                    [PCCP LOGO - CENTERED]                     │
│                    White version, 150px height               │
│                                                               │
│                                                               │
│                                                               │
└──────────────────────────────────────────────────────────────┘
```

---

## 6. Chart Formatting Details

### 6.1 Gridline Specifications

| Property | Required Value | Tolerance |
|----------|----------------|-----------|
| Major Horizontal Gridlines | ON | Required |
| Gridline Width | 0.5pt | +/- 0.1pt |
| Gridline Color | `#D9D9D9` | Exact match |
| Minor Gridlines | OFF | Required |

### 6.2 Axis Tick Marks

| Axis | Major Tick | Minor Tick |
|------|------------|------------|
| Value (Y) Axis | **NONE** | NONE |
| Category (X) Axis | **NONE** | NONE |

### 6.3 Data Labels

| Chart Type | Data Labels | Position | Format |
|------------|-------------|----------|--------|
| Column/Bar | Optional | Outside End | Match number format |
| Line | Optional | Above | Match number format |
| Pie/Donut | Required | Outside | Percentage |

### 6.4 Number Formatting

| Data Type | Format Example | Format Code |
|-----------|----------------|-------------|
| Whole numbers | 1,234 | `#,##0` |
| Percentages | 12.5% | `0.0%` or `0%` |
| Currency | $1,234 | `$#,##0` |
| Millions | $1.2M | `$#,##0.0,,"M"` |
| Billions | $1.2B | `$#,##0.0,,,"B"` |

---

## 7. Table Formatting Details

### 7.1 Cell Styling

| Property | Header Row | Data Rows (Odd) | Data Rows (Even) |
|----------|------------|-----------------|------------------|
| Background | `#051C2C` | `#FFFFFF` | `#F5F5F5` |
| Text Color | `#FFFFFF` | `#061F32` | `#061F32` |
| Font Weight | Bold | Regular | Regular |
| Font Size | 10-11pt | 10pt | 10pt |

### 7.2 Cell Margins

| Margin | Value (inches) | Value (EMU) |
|--------|----------------|-------------|
| Left | 0.1" | 91440 |
| Right | 0.1" | 91440 |
| Top | 0.05" | 45720 |
| Bottom | 0.05" | 45720 |

### 7.3 Column Alignment Rules

| Column Content | Alignment |
|----------------|-----------|
| Text labels | Left |
| Numbers/percentages | Right |
| Mixed (header) | Left |
| Ranking/Index | Center |

### 7.4 Border Styling

| Property | Value |
|----------|-------|
| Border Width | None (0) |
| Border Color | N/A |
| Cell Separation | Via alternating row colors |

---

## 8. Standard Elements

### 8.1 Footer Specifications

| Element | Position | Content | Font |
|---------|----------|---------|------|
| Source Citation | Left-aligned, (0.399", 7.5") | "Sources: [Publisher] [Report] [Date]." | 6pt, #A6A6A6 |
| Confidential | Right-aligned | "CONFIDENTIAL PCCP, LLC [Page#]" | 6pt, #A6A6A6 |

### 8.2 Section Label

| Property | Value |
|----------|-------|
| Position | (7.1", 0.2") |
| Size | 3.5" x 0.2" |
| Font | Arial, 9pt, Regular |
| Color | `#A6A6A6` |
| Content | Section name (e.g., "Market Fundamentals") |

### 8.3 Logo Usage

| Slide Type | Logo Version | Position | Size |
|------------|--------------|----------|------|
| Frontpage | White | (0.4", 0.4") | 2.5" x ~1.8" |
| End | White, centered | Center | ~2.5" width |
| Content slides | None | N/A | N/A |

---

## 9. Content Guidelines

### 9.1 Bullet Point Format

**Required Format:**
```
• CATEGORY HEADER — Body text continues with supporting detail and specific metrics where relevant
```

**Rules:**
- Category header: ALL CAPS
- Separator: Em-dash (—), not hyphen (-)
- Body text: Sentence case
- Include specific data points where available
- Maximum 2-3 lines per bullet
- Maximum 5-7 bullets per slide

### 9.2 Source Citation Format

**Standard Format:**
```
Sources: [Publisher] [Report/Index Name] [Period/Year].
```

**Examples:**
```
Sources: CBRE Industrial & Logistics Report Q4 2025.
Sources: CoStar Market Analytics December 2025; RCLCO ODCE/NPI Q3 2025.
Sources: PCCP Management Estimates.
```

### 9.3 Thesis Statement

Every content slide should have a thesis statement in the subtitle area that:
- Summarizes the slide's key takeaway
- Is a complete sentence
- Is action-oriented or insight-driven
- Maximum 2 lines

**Good Examples:**
- "Small-bay industrial outperforms across every key metric: tighter vacancy, stronger rent growth, and minimal new supply."
- "Construction down 62% from peak with small-bay at a decade low."

**Bad Examples:**
- "Market Overview" (too vague)
- "Data" (not actionable)

---

## 10. Presentation Structure

### 10.1 Standard Section Flow

```
1. TITLE SLIDE (Frontpage)
2. EXECUTIVE SUMMARY (1-2 slides)
3. [SECTION DIVIDER] → Content Slides (repeat per section)
   - Market Fundamentals (4-6 slides)
   - Target Markets (2-3 slides)
   - Demand Drivers (3-4 slides)
   - Investment Strategy (6-8 slides)
   - Competitive Positioning (2 slides)
   - Risk Management (4-5 slides)
   - ESG Strategy (2 slides)
   - JV Structure (3 slides)
   - Conclusion (2-3 slides)
4. CONTACT SLIDE
5. DISCLOSURES
6. END SLIDE (Logo Close)
```

### 10.2 Section Divider Placement

Insert a Section Title Slide before each major section:
- Market Fundamentals
- Target Markets
- Demand Drivers
- Investment Strategy
- Competitive Landscape
- Risk Management
- Investment Structure
- Conclusion/Summary

---

## 11. Validation Rules

### 11.1 Required Checks

| Check | Rule | Priority |
|-------|------|----------|
| Slide dimensions | 11" x 8.5" | HIGH |
| Chart gridlines | 0.5pt, #D9D9D9 | HIGH |
| Chart tick marks | NONE on both axes | HIGH |
| Table header color | #051C2C | HIGH |
| Table alternating rows | #FFFFFF / #F5F5F5 | HIGH |
| Table cell margins | 0.1" L/R, 0.05" T/B | MEDIUM |
| Bullet format | ALL CAPS + em-dash | HIGH |
| Source citations | Present on data slides | MEDIUM |
| Footer | Present on content slides | MEDIUM |
| Section labels | Present on content slides | LOW |

### 11.2 Color Tolerance

| Color Property | Allowed Deviation |
|----------------|-------------------|
| Primary colors | Exact match required |
| Gridline color | RGB variance ≤ 5 |
| Text colors | RGB variance ≤ 10 |

### 11.3 Position Tolerance

| Element Type | Allowed Deviation |
|--------------|-------------------|
| Content areas | ± 0.1 inches |
| Footer elements | ± 0.05 inches |
| Chart area | ± 0.1 inches |

---

## 12. File Naming Convention

```
[Company]_[Topic]_[Version].pptx

Examples:
PCCP_Light_Industrial_Thesis_vFinal.pptx
PCCP_Q4_2025_Fund_Update_v2.pptx
PCCP_Nashville_Deal_Memo_v1.pptx
```

---

## 13. Checklist for Automated Validation

### Pre-Generation
- [ ] Presentation type defined
- [ ] Section structure defined
- [ ] Slide count estimated

### Post-Generation Validation
- [ ] Slide dimensions: 11" x 8.5"
- [ ] All charts have 0.5pt #D9D9D9 gridlines
- [ ] All charts have NONE tick marks
- [ ] All tables have #051C2C headers
- [ ] All tables have alternating row colors
- [ ] All tables have 0.1"/0.05" cell margins
- [ ] All bullets use ALL CAPS + em-dash format
- [ ] All data slides have source citations
- [ ] All content slides have footers
- [ ] Logo on Frontpage and End slides

---

## Appendix A: EMU Conversions

| Measurement | Inches | EMU |
|-------------|--------|-----|
| Slide width | 11.0" | 10058400 |
| Slide height | 8.5" | 7772400 |
| Standard margin | 0.4" | 365760 |
| Cell margin (L/R) | 0.1" | 91440 |
| Cell margin (T/B) | 0.05" | 45720 |
| Gridline width | 0.5pt | 6350 |

---

## Appendix B: python-pptx Constants

```python
from pptx.enum.chart import XL_TICK_MARK

# Tick mark values
XL_TICK_MARK.NONE = -4142
XL_TICK_MARK.OUTSIDE = 3
XL_TICK_MARK.INSIDE = 2
XL_TICK_MARK.CROSS = 4
```

---

*This style guide is automatically used by the pptx_generator presentation review system for compliance checking and gap analysis.*
