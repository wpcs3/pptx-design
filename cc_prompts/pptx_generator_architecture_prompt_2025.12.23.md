# PowerPoint Presentation Generator: Architecture & Build Prompt

## Project Overview

Build an AI-powered presentation generation system for an investment firm that creates investor pitch decks. The system should:

1. **Codify company standards** from 200+ slide example templates
2. **Generate presentation outlines** based on user requirements
3. **Assemble content** from reusable slides OR via deep research
4. **Produce polished PPTX files** matching company format exactly

**Template Source Directory**: `C:\Users\wpcol\claudecode\pptx-design\pptx_templates`

**Available Templates** (each 200+ slides, showing format + content patterns):
- `template_market_analysis.pptx` → Market/economic analysis presentations
- `template_business_consulting_toolkit.pptx` → Consulting frameworks
- `template_business_case.pptx` → Investment cases/pitches  
- `pptx_template_due_diligence.pptx` → Due diligence reports

---

## Phase 1: Template Analysis & Codification

### 1.1 Create Style Guide Extractor

Build a script that analyzes all templates and extracts a unified `style_guide.json`:

```python
# Extract from each template:
{
  "colors": {
    "primary": "#XXXXXX",      # Main brand color
    "secondary": "#XXXXXX",    
    "accent": ["#...", "#..."],
    "text": {
      "title": "#XXXXXX",
      "body": "#XXXXXX",
      "subtle": "#XXXXXX"
    },
    "backgrounds": ["#...", "#..."]
  },
  "fonts": {
    "title": {"name": "...", "size_pt": 44, "bold": true},
    "subtitle": {"name": "...", "size_pt": 28, "bold": false},
    "body": {"name": "...", "size_pt": 18},
    "caption": {"name": "...", "size_pt": 12}
  },
  "spacing": {
    "margins": {"left": 0.5, "right": 0.5, "top": 0.5, "bottom": 0.5},
    "line_spacing": 1.15,
    "paragraph_spacing_pt": 12
  },
  "master_slides": {
    # Map of slide layout names → usage patterns
  }
}
```

**Tasks:**
1. Use `python-pptx` to iterate all slides in all templates
2. Extract unique color values, fonts, sizes from shape properties
3. Identify the most common values as "standard"
4. Document master slide layouts and their purposes
5. Save to `config/style_guide.json`

### 1.2 Create Slide Type Catalog

Analyze templates to build a `slide_catalog.json` documenting reusable patterns:

```python
{
  "slide_types": [
    {
      "id": "title_slide",
      "name": "Title Slide",
      "description": "Opening slide with presentation title and date",
      "master_layout": "Title Slide",
      "elements": [
        {"type": "title", "purpose": "Presentation name"},
        {"type": "subtitle", "purpose": "Date, presenter, or tagline"}
      ],
      "usage": "First slide of every presentation",
      "examples": [
        {"template": "template_market_analysis.pptx", "slide_index": 0}
      ]
    },
    {
      "id": "section_divider",
      "name": "Section Divider",
      "description": "Full-bleed section header",
      "master_layout": "Section Header",
      "elements": [...],
      "usage": "Between major presentation sections"
    },
    {
      "id": "two_column_comparison",
      "name": "Two-Column Comparison",
      "description": "Side-by-side comparison layout",
      ...
    },
    {
      "id": "data_chart",
      "name": "Data Visualization Slide",
      "description": "Chart with supporting narrative",
      ...
    },
    {
      "id": "key_metrics",
      "name": "Key Metrics Dashboard",
      "description": "3-5 KPI callout boxes",
      ...
    },
    # ... more types discovered from analysis
  ]
}
```

**Tasks:**
1. Cluster slides by layout similarity (shape positions, element counts)
2. Assign semantic names to each cluster
3. Document which master layout each type uses
4. Record example indices from source templates
5. Save to `config/slide_catalog.json`

### 1.3 Create Content Pattern Library

Document what content appears in each presentation type:

```python
{
  "presentation_types": {
    "market_analysis": {
      "description": "Economic and real estate market overview",
      "typical_sections": [
        {
          "name": "Executive Summary",
          "slide_types": ["title_slide", "key_takeaways"],
          "content_sources": ["internal_data", "research"]
        },
        {
          "name": "Macro Economic Overview", 
          "slide_types": ["section_divider", "data_chart", "bullet_points"],
          "content_sources": ["research"],
          "research_topics": ["GDP growth", "inflation", "employment", "interest rates"]
        },
        {
          "name": "Real Estate Market",
          "slide_types": ["data_chart", "two_column_comparison", "map_visual"],
          "content_sources": ["research", "internal_data"],
          "research_topics": ["cap rates", "transaction volume", "rent growth"]
        },
        ...
      ],
      "typical_slide_count": {"min": 15, "max": 30}
    },
    "investment_pitch": {
      "description": "Pitch for specific investment strategy",
      "typical_sections": [
        {"name": "Company Overview", "reusable": true},
        {"name": "Track Record", "reusable": true},
        {"name": "Investment Thesis", "content_sources": ["research", "user_input"]},
        {"name": "Deal Pipeline", "content_sources": ["internal_data"]},
        ...
      ]
    },
    "due_diligence": {...},
    "business_case": {...}
  },
  
  "reusable_sections": {
    "company_overview": {
      "description": "Standard 3-5 slides about firm history, team, AUM",
      "update_frequency": "quarterly",
      "source_template": "template_business_case.pptx",
      "source_slides": [2, 3, 4, 5, 6]
    },
    "track_record": {
      "description": "Historical performance metrics",
      "update_frequency": "quarterly",
      ...
    }
  }
}
```

**Tasks:**
1. Manually review templates to identify section patterns
2. Tag which sections are "reusable" (periodic update) vs "bespoke" (per-presentation)
3. Map research topics to each bespoke section
4. Save to `config/content_patterns.json`

---

## Phase 2: Build Core Components

### 2.1 Slide Renderer Module

`modules/slide_renderer.py` - Generates individual slides from specifications:

```python
class SlideRenderer:
    def __init__(self, style_guide: dict, template_path: str):
        """Initialize with style rules and base template."""
        
    def create_slide(self, slide_type: str, content: dict) -> Slide:
        """
        Create a slide of given type with provided content.
        
        Args:
            slide_type: ID from slide_catalog.json
            content: {
                "title": "...",
                "body": "...",
                "bullets": [...],
                "chart_data": {...},
                "images": [...]
            }
        """
        
    def apply_style(self, shape, style_type: str):
        """Apply standard formatting to a shape."""
        
    def insert_chart(self, slide, chart_spec: dict):
        """Create chart from specification."""
```

### 2.2 Reusable Slide Library

`modules/slide_library.py` - Manages reusable slide retrieval:

```python
class SlideLibrary:
    def __init__(self, templates_dir: str, catalog: dict):
        """Index all slides from template files."""
        
    def search(self, query: str, slide_types: list = None) -> list:
        """
        Semantic search for slides matching query.
        Returns list of (template, slide_index, relevance_score).
        """
        
    def get_section(self, section_id: str) -> list[Slide]:
        """Retrieve pre-defined reusable section."""
        
    def copy_slide(self, source_template: str, slide_index: int, 
                   target_presentation: Presentation) -> Slide:
        """Copy slide from source template to target."""
```

### 2.3 Research Agent

`modules/research_agent.py` - Deep research for bespoke content:

```python
class ResearchAgent:
    def __init__(self, content_patterns: dict):
        """Initialize with content pattern definitions."""
        
    async def research_section(self, section_name: str, 
                                context: dict) -> dict:
        """
        Perform deep research for a presentation section.
        
        Args:
            section_name: e.g., "Macro Economic Overview"
            context: User-provided context (strategy, geography, etc.)
            
        Returns:
            Structured content ready for slide generation:
            {
                "slides": [
                    {
                        "slide_type": "data_chart",
                        "title": "GDP Growth Outlook",
                        "content": {...},
                        "sources": [...]
                    },
                    ...
                ]
            }
        """
        
    def format_for_slides(self, research_results: dict, 
                          slide_types: list) -> list[dict]:
        """Transform research into slide-ready content."""
```

### 2.4 Outline Generator

`modules/outline_generator.py` - Creates presentation structure:

```python
class OutlineGenerator:
    def __init__(self, content_patterns: dict, slide_catalog: dict):
        """Initialize with pattern libraries."""
        
    def generate_outline(self, user_request: str) -> dict:
        """
        Generate presentation outline from user description.
        
        Returns:
            {
                "presentation_type": "investment_pitch",
                "title": "...",
                "sections": [
                    {
                        "name": "Executive Summary",
                        "slides": [
                            {"slide_type": "title_slide", "content_source": "user_input"},
                            {"slide_type": "key_takeaways", "content_source": "generated"}
                        ]
                    },
                    {
                        "name": "Company Overview",
                        "slides": [...],
                        "source": "reusable:company_overview"
                    },
                    {
                        "name": "Market Opportunity",
                        "slides": [...],
                        "content_source": "research",
                        "research_topics": ["industrial real estate trends", "...]
                    },
                    ...
                ],
                "estimated_slide_count": 25
            }
        """
        
    def refine_outline(self, outline: dict, user_feedback: str) -> dict:
        """Adjust outline based on user edits."""
```

### 2.5 Presentation Orchestrator

`modules/orchestrator.py` - Main workflow controller:

```python
class PresentationOrchestrator:
    def __init__(self, config_dir: str, templates_dir: str):
        """Load all configs and initialize sub-modules."""
        self.style_guide = load_json("config/style_guide.json")
        self.slide_catalog = load_json("config/slide_catalog.json")
        self.content_patterns = load_json("config/content_patterns.json")
        
        self.outline_gen = OutlineGenerator(...)
        self.slide_lib = SlideLibrary(...)
        self.research = ResearchAgent(...)
        self.renderer = SlideRenderer(...)
        
    async def create_presentation(self, user_request: str) -> Workflow:
        """
        Full workflow:
        1. Generate outline → await user approval
        2. Assemble content (reusable + research)
        3. Generate PPTX
        4. Support iterative refinement
        """
        
    def export_pptx(self, presentation: Presentation, path: str):
        """Save final PPTX file."""
```

---

## Phase 3: Create Claude Skill

Create a Claude Skill at `~/.claude/skills/pptx-generator/SKILL.md`:

```markdown
# PowerPoint Presentation Generator Skill

## Purpose
Generate professional investor presentations following company standards.

## Available Commands

### 1. Analyze Templates
Analyze source templates to update style guide and catalogs.
```bash
python -m pptx_generator.analyze --templates-dir <path>
```

### 2. Generate Outline
Create presentation outline from user description:
```bash
python -m pptx_generator.outline --request "Create a pitch for our new logistics fund targeting $500M"
```

### 3. Build Presentation
Generate full PPTX from approved outline:
```bash
python -m pptx_generator.build --outline outline.json --output presentation.pptx
```

### 4. Refine Presentation
Make iterative edits to existing draft:
```bash
python -m pptx_generator.refine --presentation draft.pptx --feedback "Make slide 5 more concise"
```

## Style Guide Location
`config/style_guide.json` - Colors, fonts, spacing standards

## Slide Catalog Location  
`config/slide_catalog.json` - Available slide types and layouts

## Content Patterns Location
`config/content_patterns.json` - Section templates and research mappings

## Workflow
1. User describes presentation need
2. Generate outline with `outline` command
3. User reviews/edits outline
4. Build presentation with `build` command
5. User reviews draft, provides feedback
6. Refine with `refine` command (repeat as needed)
7. Export final PPTX
```

---

## Phase 4: Directory Structure

Create this project structure:

```
pptx-generator/
├── config/
│   ├── style_guide.json          # Extracted from templates
│   ├── slide_catalog.json        # Slide type definitions
│   └── content_patterns.json     # Section/content mappings
│
├── modules/
│   ├── __init__.py
│   ├── template_analyzer.py      # Phase 1 extraction
│   ├── slide_renderer.py         # Individual slide creation
│   ├── slide_library.py          # Reusable slide management
│   ├── research_agent.py         # Deep research for content
│   ├── outline_generator.py      # Outline creation
│   └── orchestrator.py           # Main workflow
│
├── templates/                    # Symlink to source templates
│   └── → C:\Users\wpcol\claudecode\pptx-design\pptx_templates
│
├── output/                       # Generated presentations
│
├── cache/
│   ├── slide_index/             # Indexed slides for search
│   └── research/                # Cached research results
│
├── tests/
│   ├── test_renderer.py
│   ├── test_library.py
│   └── test_orchestrator.py
│
├── __main__.py                  # CLI entry point
└── requirements.txt
```

---

## Phase 5: Implementation Order

### Step 1: Template Analysis (Foundation)
1. Build `template_analyzer.py`
2. Run against all 4 templates
3. Generate initial `style_guide.json`
4. Generate initial `slide_catalog.json`
5. **Manual review**: Refine catalog with semantic names

### Step 2: Slide Library
1. Build `slide_library.py`
2. Index all slides from templates
3. Implement slide copying between presentations
4. Test: Copy specific slides to new presentation

### Step 3: Slide Renderer
1. Build `slide_renderer.py`
2. Implement each slide type from catalog
3. Test: Generate each slide type with sample content
4. Visual comparison with templates (use LibreOffice rendering)

### Step 4: Content Patterns
1. **Manual creation**: Build `content_patterns.json` by reviewing templates
2. Document section patterns for each presentation type
3. Tag reusable vs. bespoke sections

### Step 5: Outline Generator
1. Build `outline_generator.py`
2. Use content patterns to structure outlines
3. Test: Generate outlines for various request types

### Step 6: Research Agent
1. Build `research_agent.py`
2. Implement deep research for each content category
3. Format research output for slide consumption
4. Test: Research → slide content pipeline

### Step 7: Orchestrator
1. Build `orchestrator.py`
2. Wire all components together
3. Implement full workflow with user checkpoints
4. Test: End-to-end presentation generation

### Step 8: Claude Skill
1. Create SKILL.md
2. Test skill invocation from Claude Code
3. Document common workflows

---

## Key Technical Decisions

### Template Preservation
- Always copy the master template as base for new presentations
- This preserves master slides, theme colors, fonts
- Only modify slide content, never master layouts

### Slide Type Detection
Use clustering based on:
- Number of shapes
- Shape positions (normalized to slide dimensions)
- Text frame presence and properties
- Chart/table/image presence

### Research Integration
The research agent should:
- Use web search for current market data
- Cache results with timestamps
- Format data for visualization (chart-ready)
- Cite sources on slides or in notes

### Iterative Refinement
Support these edit types:
- Content changes ("make this more concise")
- Structural changes ("add a slide about X")
- Style overrides ("use blue background here")
- Slide reordering

---

## Example User Interaction

```
User: Create a pitch deck for our new $200M industrial logistics fund 
      targeting last-mile distribution centers in secondary markets.

[Outline Generator]
→ Generates outline with 22 slides across 6 sections:
  1. Title Slide
  2. Executive Summary (2 slides) 
  3. Company Overview (4 slides) ← reusable
  4. Track Record (3 slides) ← reusable
  5. Market Opportunity (5 slides) ← research: industrial RE trends
  6. Investment Strategy (4 slides) ← research + user input
  7. Pipeline/Targets (2 slides) ← user input
  8. Closing/Contact (1 slide) ← reusable

User: Looks good, but add a section on risk factors after Strategy.

[Refine Outline]
→ Adds section 7: Risk Factors (2 slides)

User: Approved, generate the deck.

[Content Assembly]
→ Copies reusable sections from templates
→ Deep research on industrial logistics, cap rates, rent growth
→ Generates bespoke slides from research

[Slide Generation]
→ Renders all slides using style guide
→ Exports to presentation.pptx

User: Slide 12 chart is too busy, simplify it.

[Iterative Refinement]
→ Modifies slide 12, re-exports
```

---

## Acceptance Criteria

1. **Style Fidelity**: Generated slides match template formatting exactly
2. **Content Quality**: Research content is accurate, properly sourced
3. **Reusability**: Standard sections copy cleanly from templates
4. **Workflow**: User can review/edit at each checkpoint
5. **Output**: Final PPTX opens correctly in PowerPoint, uses correct fonts/colors

---

## Getting Started

Begin with Phase 1, Step 1:

```bash
# Create project structure
mkdir -p pptx-generator/{config,modules,output,cache,tests}

# Analyze first template
python -c "
from pptx import Presentation
prs = Presentation('templates/template_market_analysis.pptx')
print(f'Slides: {len(prs.slides)}')
print(f'Layouts: {[l.name for l in prs.slide_layouts]}')
"
```

Then build the `template_analyzer.py` to extract style and catalog data.
