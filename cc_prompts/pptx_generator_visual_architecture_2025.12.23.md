# PowerPoint Generator: Visual Architecture

## System Overview

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                           CONFIGURATION LAYER                                │
│  ┌──────────────────┐  ┌──────────────────┐  ┌──────────────────────────┐   │
│  │  style_guide.json │  │ slide_catalog.json│  │ content_patterns.json   │   │
│  │  ────────────────│  │ ─────────────────│  │ ─────────────────────── │   │
│  │  • Colors        │  │  • Slide types   │  │  • Presentation types   │   │
│  │  • Fonts         │  │  • Layouts       │  │  • Section patterns     │   │
│  │  • Spacing       │  │  • Element specs │  │  • Reusable sections    │   │
│  │  • Master slides │  │  • Examples      │  │  • Research mappings    │   │
│  └──────────────────┘  └──────────────────┘  └──────────────────────────┘   │
└─────────────────────────────────────────────────────────────────────────────┘
                                    │
                    ┌───────────────┴───────────────┐
                    ▼                               ▼
┌─────────────────────────────────┐   ┌─────────────────────────────────┐
│       TEMPLATE SOURCES          │   │        SLIDE LIBRARY            │
│  ────────────────────────────── │   │  ────────────────────────────── │
│  template_market_analysis.pptx  │──▶│  Indexed & searchable slides    │
│  template_business_case.pptx    │   │  Copy/paste between decks       │
│  template_consulting.pptx       │   │  Reusable section retrieval     │
│  template_due_diligence.pptx    │   │                                 │
└─────────────────────────────────┘   └─────────────────────────────────┘
```

## Workflow Pipeline

```
    USER REQUEST
    "Create pitch for $200M logistics fund"
         │
         ▼
┌─────────────────────────────────────────────────────────────────┐
│                    OUTLINE GENERATOR                             │
│  ┌───────────────────────────────────────────────────────────┐  │
│  │ Input:  User request + content_patterns.json              │  │
│  │ Output: Structured outline with sections & slide types    │  │
│  │                                                           │  │
│  │ Logic:                                                    │  │
│  │ 1. Classify presentation type (pitch, analysis, DD...)   │  │
│  │ 2. Map to section template from patterns                 │  │
│  │ 3. Tag each section: reusable | research | user_input    │  │
│  │ 4. Estimate slide count per section                      │  │
│  └───────────────────────────────────────────────────────────┘  │
└─────────────────────────────────────────────────────────────────┘
         │
         ▼
    ┌─────────────┐
    │ USER REVIEW │ ◄─── Edit outline, add/remove sections
    └─────────────┘
         │ Approved
         ▼
┌─────────────────────────────────────────────────────────────────┐
│                    CONTENT ASSEMBLER                             │
│                                                                  │
│  For each section in outline:                                    │
│                                                                  │
│  ┌─────────────────────┐    ┌─────────────────────┐             │
│  │ If "reusable"       │    │ If "research"       │             │
│  │ ─────────────────── │    │ ─────────────────── │             │
│  │                     │    │                     │             │
│  │  ┌───────────────┐  │    │  ┌───────────────┐  │             │
│  │  │ Slide Library │  │    │  │Research Agent │  │             │
│  │  │───────────────│  │    │  │───────────────│  │             │
│  │  │ Copy slides   │  │    │  │ Deep research │  │             │
│  │  │ from template │  │    │  │ Web search    │  │             │
│  │  │               │  │    │  │ Format data   │  │             │
│  │  └───────────────┘  │    │  └───────────────┘  │             │
│  └─────────────────────┘    └─────────────────────┘             │
│                                                                  │
└─────────────────────────────────────────────────────────────────┘
         │
         ▼
┌─────────────────────────────────────────────────────────────────┐
│                    SLIDE RENDERER                                │
│  ┌───────────────────────────────────────────────────────────┐  │
│  │ For each slide in content:                                │  │
│  │                                                           │  │
│  │   1. Get slide_type from catalog                          │  │
│  │   2. Select appropriate master layout                     │  │
│  │   3. Create shapes per type specification                 │  │
│  │   4. Apply style_guide formatting                         │  │
│  │   5. Insert content (text, charts, images)                │  │
│  │                                                           │  │
│  │ Output: Populated Presentation object                     │  │
│  └───────────────────────────────────────────────────────────┘  │
└─────────────────────────────────────────────────────────────────┘
         │
         ▼
    ┌─────────────┐
    │ DRAFT PPTX  │
    └─────────────┘
         │
         ▼
    ┌─────────────┐
    │ USER REVIEW │ ◄─── "Simplify slide 12", "Add risk section"
    └─────────────┘
         │
         ▼
┌─────────────────────────────────────────────────────────────────┐
│                    REFINEMENT LOOP                               │
│  ┌───────────────────────────────────────────────────────────┐  │
│  │ Parse user feedback:                                      │  │
│  │  • Content edit → re-render specific slides               │  │
│  │  • Structure edit → modify outline, regenerate            │  │
│  │  • Style override → apply local formatting                │  │
│  │                                                           │  │
│  │ Re-export PPTX after each iteration                       │  │
│  └───────────────────────────────────────────────────────────┘  │
└─────────────────────────────────────────────────────────────────┘
         │
         ▼
    ┌─────────────┐
    │ FINAL PPTX  │
    └─────────────┘
```

## Module Dependency Graph

```
                    orchestrator.py
                          │
         ┌────────────────┼────────────────┐
         │                │                │
         ▼                ▼                ▼
outline_generator.py  slide_library.py  research_agent.py
         │                │                │
         └────────────────┼────────────────┘
                          │
                          ▼
                  slide_renderer.py
                          │
                          ▼
                  ┌───────────────┐
                  │  python-pptx  │
                  └───────────────┘

All modules read from:
  • config/style_guide.json
  • config/slide_catalog.json
  • config/content_patterns.json
```

## Data Flow for Slide Types

```
┌─────────────────────────────────────────────────────────────────┐
│                    SLIDE CATALOG LOOKUP                          │
│                                                                  │
│  slide_type: "two_column_comparison"                             │
│         │                                                        │
│         ▼                                                        │
│  ┌───────────────────────────────────────────────────────────┐  │
│  │ catalog entry:                                            │  │
│  │   master_layout: "Two Content"                            │  │
│  │   elements:                                               │  │
│  │     - title (top center)                                  │  │
│  │     - left_column (text or image)                         │  │
│  │     - right_column (text or image)                        │  │
│  │     - footer (optional)                                   │  │
│  └───────────────────────────────────────────────────────────┘  │
│         │                                                        │
│         ▼                                                        │
│  ┌───────────────────────────────────────────────────────────┐  │
│  │ style_guide application:                                  │  │
│  │   title → fonts.title (28pt, bold, #333333)              │  │
│  │   columns → fonts.body (16pt, #666666)                   │  │
│  │   background → colors.backgrounds[0]                      │  │
│  └───────────────────────────────────────────────────────────┘  │
│         │                                                        │
│         ▼                                                        │
│  ┌───────────────────────────────────────────────────────────┐  │
│  │ content insertion:                                        │  │
│  │   content: {                                              │  │
│  │     "title": "Strategy A vs Strategy B",                  │  │
│  │     "left": {"header": "Strategy A", "bullets": [...]},   │  │
│  │     "right": {"header": "Strategy B", "bullets": [...]}   │  │
│  │   }                                                       │  │
│  └───────────────────────────────────────────────────────────┘  │
│         │                                                        │
│         ▼                                                        │
│      Rendered Slide                                              │
└─────────────────────────────────────────────────────────────────┘
```

## Reusable vs Bespoke Content Decision

```
┌─────────────────────────────────────────────────────────────────┐
│                                                                  │
│   Section from outline                                           │
│         │                                                        │
│         ▼                                                        │
│   ┌─────────────┐                                               │
│   │ Is section  │                                               │
│   │ reusable?   │                                               │
│   └──────┬──────┘                                               │
│          │                                                       │
│    ┌─────┴─────┐                                                │
│    ▼           ▼                                                │
│   YES          NO                                               │
│    │           │                                                │
│    ▼           ▼                                                │
│ ┌──────────┐ ┌──────────────────────────────────┐              │
│ │ Slide    │ │ Content source?                  │              │
│ │ Library  │ │                                  │              │
│ │ ──────── │ │  ┌──────────┐   ┌──────────┐    │              │
│ │ Copy     │ │  │ research │   │user_input│    │              │
│ │ slides   │ │  └────┬─────┘   └────┬─────┘    │              │
│ │ from     │ │       │              │          │              │
│ │ template │ │       ▼              ▼          │              │
│ │          │ │  ┌──────────┐   ┌──────────┐    │              │
│ └──────────┘ │  │ Research │   │ Prompt   │    │              │
│              │  │ Agent    │   │ user for │    │              │
│              │  │ ──────── │   │ content  │    │              │
│              │  │ Web srch │   │          │    │              │
│              │  │ Format   │   │          │    │              │
│              │  └──────────┘   └──────────┘    │              │
│              └──────────────────────────────────┘              │
│                                                                  │
└─────────────────────────────────────────────────────────────────┘
```

## Presentation Type Templates

```
MARKET ANALYSIS                    INVESTMENT PITCH
─────────────────                  ─────────────────
┌───────────────┐                  ┌───────────────┐
│ Title         │                  │ Title         │
├───────────────┤                  ├───────────────┤
│ Exec Summary  │ ◄─ research      │ Exec Summary  │ ◄─ research
├───────────────┤                  ├───────────────┤
│ Macro Econ    │ ◄─ research      │ Company       │ ◄─ REUSABLE
│ • GDP         │                  │ Overview      │
│ • Inflation   │                  ├───────────────┤
│ • Rates       │                  │ Track Record  │ ◄─ REUSABLE
├───────────────┤                  ├───────────────┤
│ RE Market     │ ◄─ research      │ Market Opp    │ ◄─ research
│ • Cap rates   │                  ├───────────────┤
│ • Volume      │                  │ Strategy      │ ◄─ user_input
├───────────────┤                  ├───────────────┤
│ Outlook       │ ◄─ research      │ Pipeline      │ ◄─ user_input
├───────────────┤                  ├───────────────┤
│ Appendix      │                  │ Contact       │ ◄─ REUSABLE
└───────────────┘                  └───────────────┘


DUE DILIGENCE                      BUSINESS CASE
─────────────────                  ─────────────────
┌───────────────┐                  ┌───────────────┐
│ Title         │                  │ Title         │
├───────────────┤                  ├───────────────┤
│ Property      │ ◄─ user_input    │ Problem       │ ◄─ user_input
│ Overview      │                  │ Statement     │
├───────────────┤                  ├───────────────┤
│ Market        │ ◄─ research      │ Analysis      │ ◄─ research
│ Analysis      │                  ├───────────────┤
├───────────────┤                  │ Solution      │ ◄─ user_input
│ Financial     │ ◄─ user_input    ├───────────────┤
│ Analysis      │                  │ Financials    │ ◄─ user_input
├───────────────┤                  ├───────────────┤
│ Risk Factors  │ ◄─ research      │ Implementation│ ◄─ user_input
├───────────────┤                  ├───────────────┤
│ Recommendation│                  │ Recommendation│
└───────────────┘                  └───────────────┘
```
