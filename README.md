# PPTX Design System

A Python system for PowerPoint template analysis and professional presentation generation.

## Features

- **Template Analysis** - Analyze PowerPoint templates and extract reusable components
- **Presentation Generation** - Generate presentations from natural language or structured outlines
- **Component Library** - 22,000+ extracted charts, tables, images, and styles
- **Unified API** - Simple fluent interface for creating presentations
- **Multi-LLM Support** - Works with Anthropic, OpenAI, Google, and Ollama models
- **Agent Integration** - MCP server and tool definitions for AI agent access

## Quick Start

```python
from pptx_design import PresentationBuilder

builder = PresentationBuilder("consulting_toolkit")
builder.add_title_slide("Q4 Review", "2025 Analysis")
builder.add_agenda(["Overview", "Analysis", "Recommendations"])
builder.add_content_slide("Key Findings", bullets=["Finding 1", "Finding 2"])
builder.save("presentation.pptx")
```

## Installation

```bash
# Clone the repository
git clone https://github.com/wpcs3/pptx-design.git
cd pptx-design

# Create conda environment
conda create -n pptx-design python=3.11
conda activate pptx-design

# Install conda packages
conda install -c conda-forge poppler pdf2image pillow scikit-image numpy

# Install pip packages
pip install python-pptx anthropic click rich litellm
```

### External Dependencies

- **LibreOffice** - Required for PPTX to PDF conversion ([download](https://www.libreoffice.org/download/))
- **Poppler** - Required for PDF to PNG conversion (installed via conda)

## Usage

### Generate Presentations from Natural Language

```bash
# Generate outline from request
python -m pptx_generator outline --request "Create a pitch deck for a $150M industrial fund"

# Build presentation from outline
python -m pptx_generator build --outline outline.json --output presentation.pptx

# Full workflow (outline → review → build)
python -m pptx_generator generate --request "Create investor pitch deck" --auto-approve
```

### Use the Python API

```python
from pptx_design import PresentationBuilder

# Create a presentation using a template
builder = PresentationBuilder("consulting_toolkit")

# Add slides
builder.add_title_slide("Strategic Review", "Q4 2025")
builder.add_section_slide("Market Analysis")
builder.add_content_slide("Key Findings", bullets=[
    "Revenue grew 15% YoY",
    "Market share increased to 23%",
    "Customer satisfaction at 4.5/5"
])
builder.add_chart_slide("Revenue Trend", chart_type="line", data={
    "categories": ["Q1", "Q2", "Q3", "Q4"],
    "series": [{"name": "Revenue", "values": [100, 120, 135, 150]}]
})

builder.save("quarterly_review.pptx")
```

### Browse Component Library

```bash
# View library statistics
python -m pptx_extractor.library_cli full-stats

# Browse color palettes and typography
python -m pptx_extractor.library_cli styles

# Find layouts matching content requirements
python -m pptx_extractor.library_cli find-layout --charts 1 --tables 1

# Find images by category
python -m pptx_generator find-images -c logo
python -m pptx_generator find-images -c photo -l 5
```

### Analyze Templates

```bash
# Render template to PNG
python -m pptx_extractor.analyzer render --template "template_name.pptx"

# Extract theme (colors, fonts)
python -m pptx_extractor.analyzer theme --template "template_name.pptx"

# Extract slide masters
python -m pptx_extractor.analyzer masters --template "template_name.pptx"
```

## Component Library

The system includes a comprehensive library of extracted components:

| Component Type | Count |
|---------------|-------|
| Charts | 357 |
| Tables | 180 |
| Images | 259 |
| Shapes | 19,376 |
| Diagrams | 358 |
| Layout Blueprints | 1,131 |
| Typography Presets | 43 |
| Color Palettes | 4 |

## Project Structure

```
pptx-design/
├── pptx_design/              # Unified API module
│   ├── builder.py            # PresentationBuilder (fluent API)
│   ├── registry.py           # TemplateRegistry (template catalog)
│   ├── pipeline.py           # ContentPipeline
│   ├── evaluation.py         # PPTEval framework
│   ├── mcp_server.py         # MCP server for AI agents
│   └── agent_tools.py        # Agent tool definitions
├── pptx_generator/           # Presentation generation system
│   ├── modules/
│   │   ├── orchestrator.py   # Main workflow controller
│   │   ├── template_renderer.py
│   │   ├── component_library.py
│   │   ├── llm_provider.py   # Multi-provider LLM support
│   │   └── markdown_parser.py
│   └── config/               # Configuration files
├── pptx_extractor/           # Template analysis and extraction
│   ├── analyzer.py           # Template analysis CLI
│   ├── library_cli.py        # Library browser CLI
│   └── unified_extractor.py  # Component extraction
├── pptx_component_library/   # Extracted components index
├── pptx_templates/           # Source PowerPoint templates
├── docs/                     # Documentation
│   ├── API_REFERENCE.md
│   └── TUTORIAL.md
└── outputs/                  # Generated presentations
```

## Key Modules

| Module | Purpose |
|--------|---------|
| `pptx_design` | Unified API with PresentationBuilder, evaluation, and agent tools |
| `pptx_generator` | Outline-based presentation generation with LLM support |
| `pptx_extractor` | Template analysis and component extraction |

## Documentation

- [API Reference](docs/API_REFERENCE.md)
- [Tutorial](docs/TUTORIAL.md)
- [Claude Code Instructions](CLAUDE.md)

## License

MIT
