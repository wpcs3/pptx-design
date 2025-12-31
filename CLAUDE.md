# CLAUDE.md - PPTX Design System

Project-specific instructions for Claude Code when working in this repository.

**Repository**: https://github.com/wpcs3/pptx-design

## Project Overview

A Python system for:
1. **Template Analysis** - Analyzing PowerPoint templates and recreating them programmatically
2. **Presentation Generation** - Generating investor pitch decks from structured outlines using the `pptx_generator` module
3. **Component Library** - Reusing extracted charts, tables, images, and styles from templates
4. **Unified API** - Simple `pptx_design` module for creating presentations with fluent interface

## Quick Start (Unified API)

```python
from pptx_design import PresentationBuilder

builder = PresentationBuilder("consulting_toolkit")
builder.add_title_slide("Q4 Review", "2025 Analysis")
builder.add_agenda(["Overview", "Analysis", "Recommendations"])
builder.add_content_slide("Key Findings", bullets=["Finding 1", "Finding 2"])
builder.save("presentation.pptx")
```

See `docs/API_REFERENCE.md` and `docs/TUTORIAL.md` for full documentation.

## Development Environment

```bash
# Always activate the conda environment first
conda activate pptx-design

# Working directory
cd C:\Users\wpcol\claudecode\pptx-design
```

### Dependencies

- **Conda packages**: poppler, pdf2image, pillow, scikit-image, numpy
- **Pip packages**: python-pptx, anthropic, click, rich

### External Tools

- **LibreOffice**: Required for PPTX → PDF conversion (headless mode)
- **Poppler**: Required for PDF → PNG conversion (via pdf2image)
- **GitHub CLI**: Installed at `C:\Program Files\GitHub CLI\gh.exe` for repository operations

---

## PPTX Generator (Primary System)

The `pptx_generator` module creates professional presentations from JSON outlines.

### Key Commands

```bash
# Generate outline from natural language request
python -m pptx_generator outline --request "Create a pitch deck for a $150M industrial fund"

# Build presentation from outline
python -m pptx_generator build --outline outline.json --output presentation.pptx

# Full generation workflow (outline → review → build)
python -m pptx_generator generate --request "Create investor pitch deck" --auto-approve

# List available slide types
python -m pptx_generator list-types

# List presentation patterns
python -m pptx_generator list-patterns

# Test rendering with sample slides
python -m pptx_generator test-render

# Browse classified library content
python -m pptx_generator browse-library

# Run content classification
python -m pptx_generator classify

# Find images by category
python -m pptx_generator find-images -c logo
python -m pptx_generator find-images -c photo -l 5
python -m pptx_generator find-images -c icon

# Re-index slide pool with enhanced classification
python -m pptx_generator.modules.reindex_slides --templates-dir pptx_templates --output cache/slide_pool_index.json
```

### Module Architecture

| Module | Purpose |
|--------|---------|
| `orchestrator.py` | Main workflow controller with Phase 2 evaluation integration |
| `template_renderer.py` | Renders slides using template layouts + ComponentLibrary |
| `outline_generator.py` | Creates presentation outlines from requests |
| `component_library.py` | Access to extracted charts, tables, images |
| `library_enhancer.py` | Domain tagging and smart component matching |
| `content_classifier.py` | Image/icon/diagram classification |
| `slide_library.py` | Reusable slide section management |
| `research_agent.py` | Content research for sections |
| `layout_cascade.py` | Phase 2: Automatic layout selection from content analysis |
| `llm_provider.py` | Phase 1: Multi-provider LLM integration via LiteLLM |
| `slide_pool.py` | Phase 4: PPTAgent-inspired slide indexing and clone-edit workflow |
| `iterative_refiner.py` | Phase 4: Quality-based outline refinement loop |
| `reindex_slides.py` | Phase 4: Enhanced slide type classification (95% accuracy) |

### ComponentLibrary Integration

The system includes a comprehensive library of extracted components:

| Component Type | Count | Usage |
|---------------|-------|-------|
| **Basic Components** | | |
| Charts | 357 | Auto-matched by type + structure |
| Tables | 180 | Auto-matched by dimensions |
| Images | 259 | Available for manual selection |
| Shapes | 19,376 | Icons, backgrounds, separators |
| Diagrams | 358 | Process flows, org charts |
| **Styles** | | |
| Color Palettes | 4 | Template-sourced color schemes |
| Typography Presets | 43 | Font/size/weight combinations |
| **Layouts** | | |
| Layout Blueprints | 1,131 | Slide zone definitions |
| Grid Systems | 21 | Column/row arrangements |
| **Content Patterns** | | |
| Text Patterns | 123 | Bullet/title/callout styles |
| Deck Templates | 4 | Full presentation structures |
| Sequences | 18 | Slide flow patterns |

### LibraryEnhancer Features

The `LibraryEnhancer` adds intelligent component matching:

- **Domain tagging**: Components tagged with real estate terms (market_analysis, financial_metrics, property_types, etc.)
- **Purpose classification**: Shapes classified as icons, separators, backgrounds, callouts
- **Smart matching**: Charts matched by structure (series count × category count)
- **Style loading**: Color palettes and typography presets from library

### Slide Types

| Type | Description |
|------|-------------|
| `title_slide` | Title + subtitle |
| `title_content` | Title + bullet points |
| `two_column` | Side-by-side comparison |
| `data_chart` | Chart with title and narrative |
| `table_slide` | Data table with headers |
| `key_metrics` | 3-4 metric boxes |
| `section_divider` | Section break slide |

### Outline JSON Schema

```json
{
  "presentation_type": "investment_pitch",
  "title": "Fund Name",
  "sections": [
    {
      "name": "Section Name",
      "slides": [
        {
          "slide_type": "data_chart",
          "content": {
            "title": "Chart Title",
            "chart_data": {
              "type": "column",
              "categories": ["A", "B", "C"],
              "series": [{"name": "Series 1", "values": [1, 2, 3]}]
            },
            "library_component_id": "optional_specific_chart_id"
          }
        }
      ]
    }
  ]
}
```

### Configuration Files

| File | Purpose |
|------|---------|
| `config/style_guide.json` | Colors, fonts, spacing |
| `config/slide_catalog.json` | 279 cataloged slide types |
| `config/content_patterns.json` | Presentation templates + outline schema |

---

## Component Library CLI

The `pptx_extractor.library_cli` module provides commands for extracting and browsing components.

### Extraction Commands

```bash
# Run unified extraction on all templates (extracts everything)
python -m pptx_extractor.library_cli extract-all -t pptx_templates

# Extract single template
python -m pptx_extractor.library_cli extract --template template.pptx
```

### Browse Commands

```bash
# View full library statistics
python -m pptx_extractor.library_cli full-stats

# Browse color palettes and typography
python -m pptx_extractor.library_cli styles

# Browse layout blueprints and grids
python -m pptx_extractor.library_cli layouts

# Find layouts matching content requirements
python -m pptx_extractor.library_cli find-layout --charts 1 --tables 1

# Browse text patterns (bullets, titles, callouts)
python -m pptx_extractor.library_cli text-patterns

# Browse slide sequences and deck templates
python -m pptx_extractor.library_cli sequences

# Generate a deck outline from structure type
python -m pptx_extractor.library_cli generate-deck data_heavy --topic "Q4 Analysis"
python -m pptx_extractor.library_cli generate-deck executive_presentation --topic "Strategy Review"

# Browse chart styles
python -m pptx_extractor.library_cli chart-styles

# Browse diagram templates
python -m pptx_extractor.library_cli diagram-templates
```

### Extractor Modules

| Module | Purpose |
|--------|---------|
| `style_extractor.py` | Color palettes, typography, gradients, shadows, 3D effects |
| `chart_style_extractor.py` | Chart formatting profiles (axis, legend, series) |
| `layout_blueprint_extractor.py` | Grid systems, zones, content areas |
| `diagram_template_extractor.py` | Shape combinations as reusable patterns |
| `text_template_extractor.py` | Bullet patterns, title styles, callouts |
| `sequence_extractor.py` | Slide sequences, deck templates |
| `unified_extractor.py` | Master extractor that runs all above |

---

## Template Analyzer (pptx_extractor)

The `pptx_extractor` module for template analysis and recreation.

### Commands

```bash
# Render template to PNG
python -m pptx_extractor.analyzer render --template "template_name.pptx"

# Analyze a template slide
python -m pptx_extractor.analyzer analyze --template "template_name.pptx" --slide 0

# Extract theme (colors, fonts)
python -m pptx_extractor.analyzer theme --template "template_name.pptx"

# Extract slide masters
python -m pptx_extractor.analyzer masters --template "template_name.pptx"

# Recreate from description
python -m pptx_extractor.recreator recreate --description "description_name"
```

---

## Presentation Builder

The `create_presentation.py` script demonstrates building presentations using extracted components.

### Usage

```bash
# Generate a sample presentation using library components
python create_presentation.py
```

### Output
- `outputs/strategic_business_review.pptx` - 12-slide executive presentation
- Uses color palette from `template_business_case`
- Uses typography presets (title, heading, body, etc.)
- Includes chart from library

### PresentationBuilder Class

```python
from create_presentation import PresentationBuilder

builder = PresentationBuilder()
builder.add_title_slide("My Title", "Subtitle")
builder.add_section_slide("Section Name")
builder.add_content_slide("Title", ["Bullet 1", "Bullet 2"])
builder.add_data_slide("Chart Title")  # Uses library chart
builder.add_comparison_slide("Compare", [{"title": "A", "points": [...]}])
builder.add_timeline_slide("Roadmap", [{"date": "Q1", "description": "..."}])
builder.add_summary_slide("Takeaways", ["Point 1", "Point 2"])
builder.add_contact_slide("Questions")
builder.save("output.pptx")
```

---

## Code Style

- Use type hints for function signatures
- Use `pathlib.Path` for all file paths
- Use `rich` for CLI output formatting
- Use `logging` module for debug/info messages

---

## Current Status (2025-12-31)

### Light Industrial Thesis Presentation (2025-12-29 to 2025-12-31)
- ✅ **46-slide investor pitch deck** generated from JSON outline
- ✅ **Template-based generation**: Uses `Light_Industrial_Thesis_v27_CS_edits.pptx` as base template
- ✅ **Custom formatting scripts**: `generate_formatted_v27.py` through `generate_formatted_v30.py`
- ✅ **Template format config**: `pptx_generator/config/template_format_v27.json` with layout indices and styling
- ✅ **Chart formatting**: Slides 8, 9, 16, 25 with proper number formats (comma, percentage, USD)
- ✅ **Section images**: Gemini-generated backgrounds for section dividers
- ✅ **Logo handling**: White logo with correct aspect ratio (2.382)
- ✅ **Table styling**: Cell margins, column alignment by majority, alternating row colors
- ⚠️ **Known issue**: Table borders still show as white in PowerPoint (PDF renders correctly)
  - Attempted fixes: noFill, zero-width borders, color matching
  - May require manual adjustment in PowerPoint or alternative library (Spire.Presentation)

### Phase 4: Advanced Features (2025-12-29)
- ✅ **Markdown Input**: Parse Markdown/Marp to presentation outlines (`pptx_generator/modules/markdown_parser.py`)
  - YAML frontmatter, charts, metrics, tables, two-column layouts
  - Exported: `MarkdownParser`, `markdown_to_outline`, `parse_marp_file`
- ✅ **ML Layout Classification**: LayoutParser for semantic zone detection (`pptx_extractor/layout_classifier.py`)
  - Detects title, text, figure, table regions
  - **Three-tier fallback**: Detectron2 → PaddleOCR → Basic heuristics
  - PaddleOCR added as Windows-friendly alternative (2025-12-30)
  - Exported: `SlideLayoutClassifier`, `classify_slide`, `classify_presentation_slides`
- ✅ **LayoutLMv3 Semantic Analyzer**: Deep document understanding (`pptx_extractor/semantic_analyzer.py`)
  - Content purpose classification, template suggestions
  - Exported: `SemanticSlideAnalyzer`, `analyze_slide`, `extract_text_boxes_from_pptx`
- ✅ **SlidePool (PPTAgent-inspired)**: Reference slide matching and cloning (`pptx_generator/modules/slide_pool.py`)
  - Indexes 1131 slides from templates by functional type
  - **95% classification accuracy** (reduced from 30% unknown to 5%)
  - Clone-and-edit workflow for better visual fidelity
  - Edit actions: replace_title, update_bullets, swap_chart_data, update_table
- ✅ **Enhanced Re-indexing**: Improved slide type classification (`pptx_generator/modules/reindex_slides.py`)
  - Layout name patterns, title patterns, visual heuristics
  - Properly classifies content_visual, metrics, timeline slides
- ✅ **Iterative Refinement**: Quality-based outline improvement (`pptx_generator/modules/iterative_refiner.py`)
  - Analyzes evaluation issues and generates refinement actions
  - Multi-pass refinement loop until target grade reached
  - Optional LLM-powered enhancement
- ✅ **Public API Integration**: All Phase 4 features available via `from pptx_design import ...`
- ✅ **Graceful Fallbacks**: All ML features work without heavy dependencies via heuristic mode

### Phase 3: Agent Integration (2025-12-29)
- ✅ **MCP Server**: 15 tools for Claude/agent access to PresentationBuilder
- ✅ **Agent-Native Interface**: Structured tool definitions for OpenAI/Anthropic function calling
- ✅ **Image Search Integration**: Pexels/Unsplash API for automatic image sourcing
- ✅ **Tool Schema Export**: JSON Schema format for LLM integration

### Phase 2: Quality & Evaluation (2025-12-29)
- ✅ **PPTEval Framework**: Content/Design/Coherence evaluation (PPTAgent-inspired)
- ✅ **Cascading Layout Handlers**: Automatic layout selection based on content analysis
- ✅ **Round-trip Testing**: Extraction → generation fidelity verification
- ✅ **Orchestrator Integration**: Evaluation runs automatically after generation
- ✅ **Functional Slide Insertion**: Auto-insert section headers, TOC, ending slides

### Phase 1: LLM Foundation (2025-12-29)
- ✅ **LiteLLM Integration**: Unified interface for 100+ LLM models
- ✅ **Multi-Provider Support**: Anthropic, OpenAI, Google, Ollama
- ✅ **Tone Controls**: 6 options (professional, casual, sales_pitch, educational, executive, default)
- ✅ **Verbosity Controls**: 3 levels (concise, standard, detailed)
- ✅ **Async Generation**: LiteLLM-powered async content generation
- ✅ **LLM Config**: Configuration file for model and style settings

### Unified API (pptx_design module)
- ✅ **PresentationBuilder**: Fluent API for creating presentations
- ✅ **TemplateRegistry**: 4 templates indexed with 27-31 layouts each
- ✅ **ContentPipeline**: Separates content from layout concerns
- ✅ **VisualTester**: SSIM-based automated visual testing
- ✅ **Documentation**: API reference and tutorial in `docs/`

### Master Layout Integration
- ✅ **TemplateGenerator**: Uses actual master layouts (100% position accuracy)
- ✅ **Placeholder-based rendering**: No LLM approximation needed
- ✅ **Automatic footers/headers**: From master layouts
- ✅ **Zero API cost**: For layout positioning

### PPTX Generator
- ✅ Full presentation generation from outlines
- ✅ ComponentLibrary integration (22,792 components)
- ✅ LibraryEnhancer with domain tagging (479 tagged)
- ✅ Smart chart/table matching by structure
- ✅ Color palette and typography loading from library
- ✅ LLM-powered content generation with tone/verbosity

### Content Classification
- ✅ **Image Classification**: 259 images categorized
- ✅ **Icon Extraction**: 479 icons extracted and tagged
- ✅ **Diagram Templates**: 53 diagram patterns

### Template Analyzer
- ✅ PPTX → PNG rendering via LibreOffice
- ✅ Theme and master extraction
- ✅ Multi-slide deck combining
- ✅ Line styles and inline formatting

---

## Completed Features

### Phase 4: Advanced Features (2025-12-29)
- [x] **Markdown Parser**: Full Markdown-to-outline conversion (`pptx_generator/modules/markdown_parser.py`)
- [x] **Marp Compatibility**: Support for Marp-style frontmatter and slide separators
- [x] **ML Layout Classification**: LayoutParser + PubLayNet for semantic region detection (`pptx_extractor/layout_classifier.py`)
- [x] **LayoutLMv3 Semantic Analysis**: Microsoft's multimodal transformer for slide understanding (`pptx_extractor/semantic_analyzer.py`)
- [x] **Heuristic Fallbacks**: All ML features work without heavy dependencies
- [x] **Code Block Support**: Parse charts, metrics, and tables from Markdown code blocks
- [x] **Two-Column Syntax**: Custom syntax for two-column layouts in Markdown
- [x] **Content Purpose Classification**: Auto-classify slide purpose (informative, persuasive, etc.)

### Phase 3: Agent Integration (2025-12-29)
- [x] **MCP Server**: Model Context Protocol server with 15 tools (`pptx_design/mcp_server.py`)
- [x] **Agent Tools Interface**: Structured tool definitions for LLMs (`pptx_design/agent_tools.py`)
- [x] **OpenAI Format**: `get_openai_tools()` for function calling
- [x] **Anthropic Format**: `get_anthropic_tools()` for tool use
- [x] **Image Search**: Pexels/Unsplash integration (`pptx_generator/modules/image_search.py`)
- [x] **Tool Execution**: AgentInterface class with validated execution
- [x] **Tool Schema Export**: Export to JSON for external agent integration

### Phase 2: Quality & Evaluation (2025-12-29)
- [x] **PPTEval Framework**: Content/Design/Coherence evaluation (`pptx_design/evaluation.py`)
- [x] **Cascading Layout Handlers**: Auto layout selection from content (`pptx_generator/modules/layout_cascade.py`)
- [x] **Round-trip Testing**: Extract → generate fidelity verification (`pptx_design/roundtrip.py`)
- [x] **Orchestrator Integration**: Evaluation and layout cascade in workflow
- [x] **GenerationOptions**: Configurable auto-layout, evaluation, section headers
- [x] **GenerationResult**: Dataclass with presentation, evaluation, layout decisions
- [x] **Quality Grading**: A-F grades with recommendations
- [x] **Functional Slides**: Auto-insert section headers, TOC, ending slides

### Phase 1: LLM Foundation (2025-12-29)
- [x] **LiteLLM Integration**: Unified multi-provider LLM access (`pptx_generator/modules/llm_provider.py`)
- [x] **Multi-Provider Support**: Anthropic Claude, OpenAI GPT, Google Gemini, Ollama local models
- [x] **Tone Controls**: 6 content tone options (professional, casual, sales_pitch, etc.)
- [x] **Verbosity Controls**: 3 verbosity levels (concise, standard, detailed)
- [x] **GenerationConfig**: Dataclass for managing generation settings
- [x] **Async Generation**: LiteLLM-powered async content generation support
- [x] **LLM Config File**: Configuration for models and settings (`pptx_generator/config/llm_config.json`)
- [x] **ContentGenerator Updates**: Integrated tone/verbosity into content generation

### Unified API & Improvements (2025-12-27)
- [x] **PresentationBuilder**: Fluent API for creating presentations (`pptx_design/builder.py`)
- [x] **TemplateRegistry**: Catalog of templates with metadata (`pptx_design/registry.py`)
- [x] **ContentPipeline**: Separates content from layout concerns (`pptx_design/pipeline.py`)
- [x] **VisualTester**: Automated SSIM-based visual testing (`pptx_design/testing.py`)
- [x] **Documentation**: API reference and tutorial (`docs/API_REFERENCE.md`, `docs/TUTORIAL.md`)

### Master Layout Integration (2025-12-27)
- [x] **TemplateGenerator**: New class that uses actual master layouts instead of LLM position approximation
- [x] **Master Extractor**: Extracts all layouts with exact placeholder positions and formatting
- [x] **TemplateRenderer Update**: Now fills placeholders directly from master layouts
- [x] **Slide Clearing**: Automatically removes template slides, keeping only layouts
- [x] **Placeholder Mapping**: Maps content keys (title, subtitle, body) to placeholder types
- [x] **Skip Title Parameter**: Render methods accept `skip_title` to avoid duplicate content

### Component Library Extraction (2025-12-25)
- [x] **Unified Extractor**: Master extractor running all 7 component extractors
- [x] **Style Extractor**: Color palettes, typography presets, gradients, shadows, 3D effects
- [x] **Chart Style Extractor**: Chart formatting profiles (axis, legend, series colors)
- [x] **Layout Blueprint Extractor**: Grid systems, zone detection, content-aware matching
- [x] **Diagram Template Extractor**: Shape combinations as reusable patterns
- [x] **Text Template Extractor**: Bullet patterns, title styles, callouts
- [x] **Sequence Extractor**: Slide sequences, deck templates, narrative flows
- [x] **Library CLI**: 10+ new commands (extract-all, styles, layouts, find-layout, generate-deck, etc.)
- [x] **ComponentLibrary Integration**: All new component types accessible via API
- [x] **Presentation Builder**: Sample script generating presentations from extracted components

### Content Classification (2025-12-26)
- [x] **ContentClassifier module**: Automated classification of library content
- [x] **Image classification**: 259 images into 8 categories (logo, icon, photo, screenshot, background, chart_image, decorative, unknown)
- [x] **Icon extraction**: 479 icons from images and shapes
- [x] **Diagram templates**: 53 process flow patterns with layout detection
- [x] **LibraryEnhancer integration**: find_logo(), find_background_image(), find_icons(), find_diagram_template()
- [x] **CLI commands**: browse-library, classify, find-images

### PPTX Generator (2025-12-24/25)
- [x] **Orchestrator workflow**: Outline → Content → Render → Export
- [x] **Template-based rendering**: Uses actual template layouts
- [x] **ComponentLibrary**: Access to extracted charts, tables, images
- [x] **LibraryEnhancer**: Domain tagging and smart matching
- [x] **Chart styling**: Data labels, legends, color fills
- [x] **Table styling**: Headers, alternating rows
- [x] **Style loading**: Color palettes and typography from library
- [x] **Empty presentation fix**: Template slides removed, only layouts kept

### Template Analyzer (2025-12-23)
- [x] PPTX → PNG rendering pipeline
- [x] Multi-slide deck combiner
- [x] Slide master/layout extraction
- [x] Theme extraction
- [x] Line styles (dashed, dotted, etc.)
- [x] Inline text formatting

---

## File Structure

```
pptx-design/
├── pptx_design/              # ⭐ Unified API module
│   ├── __init__.py           # Package exports
│   ├── builder.py            # PresentationBuilder (fluent API)
│   ├── registry.py           # TemplateRegistry (template catalog)
│   ├── pipeline.py           # ContentPipeline (content/layout separation)
│   ├── testing.py            # VisualTester (SSIM comparison)
│   ├── evaluation.py         # Phase 2: PPTEval framework
│   ├── roundtrip.py          # Phase 2: Round-trip testing
│   ├── mcp_server.py         # ⭐ Phase 3: MCP server (15 tools)
│   └── agent_tools.py        # ⭐ Phase 3: Agent-native interface
├── pptx_generator/           # Presentation generation system
│   ├── __main__.py           # CLI entry point
│   ├── config/               # Configuration files
│   │   ├── style_guide.json
│   │   ├── slide_catalog.json
│   │   ├── content_patterns.json
│   │   └── llm_config.json   # Phase 1: LLM configuration
│   ├── modules/              # Core modules
│   │   ├── orchestrator.py   # Updated: Phase 2 evaluation integration
│   │   ├── template_renderer.py  # Updated: uses master layout placeholders
│   │   ├── component_library.py
│   │   ├── library_enhancer.py
│   │   ├── content_classifier.py
│   │   ├── layout_cascade.py # Phase 2: Cascading layout handlers
│   │   ├── llm_provider.py   # Phase 1: Multi-provider LLM
│   │   ├── image_search.py   # Phase 3: Pexels/Unsplash integration
│   │   ├── markdown_parser.py # ⭐ Phase 4: Markdown-to-outline parser
│   │   └── ...
│   └── output/               # Generated presentations
├── pptx_extractor/           # Extractors and analyzers
│   ├── template_generator.py # Master layout-based generation
│   ├── layout_classifier.py  # ⭐ Phase 4: ML layout classification
│   ├── semantic_analyzer.py  # ⭐ Phase 4: LayoutLMv3 integration
│   ├── master_extractor.py   # ⭐ NEW: Layout extraction
│   ├── library_cli.py        # Library CLI commands
│   ├── unified_extractor.py  # Master extractor
│   ├── analyzer.py           # Template analysis CLI
│   └── ...
├── pptx_component_library/   # Extracted components index
│   ├── master_index.json     # Unified index
│   ├── charts/, tables/, images/, shapes/
│   ├── styles/, layouts/, text_templates/, sequences/
│   └── ...
├── pptx_templates/           # Source PowerPoint templates (4 templates)
├── config/                   # Project configuration
│   └── template_registry.json  # ⭐ NEW: Template metadata
├── docs/                     # ⭐ NEW: Documentation
│   ├── API_REFERENCE.md      # Full API documentation
│   ├── TUTORIAL.md           # Getting started guide
│   └── IMPROVEMENTS.md       # Roadmap and recommendations
├── outputs/                  # Generated outputs
└── CLAUDE.md                 # This file
```

---

## Next Steps

### All Phases Complete!
The improvement roadmap (Phases 1-4) is now fully implemented.

### Completed Phases
- ✅ **Phase 4: Advanced Features** - Fully integrated into `pptx_design` API
  - Markdown parser: `MarkdownParser`, `markdown_to_outline`, `parse_marp_file`
  - Layout classifier: `SlideLayoutClassifier`, `classify_slide` (heuristic mode on Windows)
  - Semantic analyzer: `SemanticSlideAnalyzer`, `analyze_slide` (LayoutLMv3)
  - **SlidePool**: 1131 indexed slides, clone-and-edit workflow (PPTAgent-inspired)
  - **IterativeRefiner**: Multi-pass quality refinement until target grade
- ✅ **Phase 3: Agent Integration** (MCP Server, Agent Tools, Image Search)
- ✅ **Phase 2: Quality & Evaluation** (PPTEval, layout cascade, round-trip testing)
- ✅ **Phase 1: LLM Foundation** (LiteLLM, multi-provider, tone/verbosity)
- ✅ **Unified API** (`pptx_design` module v1.4.0)
- ✅ **Master layout integration** (100% position accuracy)
- ✅ **Template registry** (4 templates, 117 layouts)

### Platform Notes
- **Windows**: Detectron2 not available (no pre-built wheels), use PaddleOCR as ML fallback:
  ```bash
  pip install paddlepaddle paddleocr
  ```
- **Linux/macOS**: Full ML support with `pip install layoutparser[detectron2]`

### Future Ideas
- Web UI for outline editing and component browsing
- Real-time preview during generation
- Content population from external data sources
- AI-powered component recommendation based on slide content
- Fine-tuned LayoutLMv3 model for slide-specific understanding
- Interactive Markdown editor with live preview
