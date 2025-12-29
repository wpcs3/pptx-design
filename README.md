# PPTX Design System

A Python-based system for analyzing PowerPoint templates and recreating them programmatically. Uses an iterative vision-based feedback loop to extract precise design specifications and generate matching slides.

## Overview

This system:
1. **Analyzes** existing PowerPoint templates to extract their visual design
2. **Generates** natural language descriptions of slide layouts, colors, fonts, and elements
3. **Recreates** templates programmatically using `python-pptx`
4. **Validates** recreations through visual comparison until they match the original

## Prerequisites

### Required Software

1. **Python 3.10+** (via Conda/Miniforge recommended)
   - Download: https://conda-forge.org/download/

2. **LibreOffice** (for PPTX to PDF conversion)
   - Download: https://www.libreoffice.org/download/download-libreoffice/
   - Default path: `C:\Program Files\LibreOffice\program\soffice.exe`

### Python Environment

This project uses a Conda environment named `pptx-design` with Poppler installed via conda-forge.

## Installation

### 1. Create and activate the Conda environment

```bash
# Create environment with Python 3.11
conda create -n pptx-design python=3.11
conda activate pptx-design
```

### 2. Install Conda packages

```bash
# Install image processing and PDF tools
conda install -c conda-forge poppler pdf2image pillow scikit-image numpy
```

### 3. Install Pip packages

```bash
# Install remaining dependencies
pip install -r requirements.txt
```

Or install manually:
```bash
pip install python-pptx anthropic click rich
```

### 4. Verify installation

```bash
# Check all dependencies
python -m src.analyzer check

# Verify Poppler
pdftoppm -h

# Verify LibreOffice
"C:\Program Files\LibreOffice\program\soffice.exe" --version
```

## Project Structure

```
C:\Users\wpcol\claudecode\pptx-design\
├── pptx_templates/              # Original PowerPoint templates
├── src/
│   ├── __init__.py
│   ├── renderer.py              # PPTX → PNG conversion
│   ├── comparator.py            # Visual comparison (SSIM)
│   ├── descriptor.py            # Generate descriptions
│   ├── generator.py             # Create PPTX from descriptions
│   ├── analyzer.py              # Main: analyze templates
│   ├── recreator.py             # Main: recreate from descriptions
│   └── prompts/
│       ├── description_prompt.txt
│       ├── comparison_prompt.txt
│       └── refinement_prompt.txt
├── descriptions/                # Generated descriptions (JSON + MD)
├── outputs/                     # Generated PPTX files and renders
├── diffs/                       # Visual diff images
├── config.py                    # Configuration
├── requirements.txt
└── README.md
```

## Usage

**Important:** Always activate the Conda environment first!

```bash
conda activate pptx-design
cd C:\Users\wpcol\claudecode\pptx-design
```

### List Available Templates

```bash
python -m src.analyzer list
```

### Analyze a Template

Analyzes a template and generates a description through iterative refinement:

```bash
# Analyze slide 1 (index 0) of a template
python -m src.analyzer analyze --template "template_business_case.pptx" --slide 0

# Analyze all slides
python -m src.analyzer analyze --template "template_business_case.pptx" --all-slides

# Use Anthropic API directly (requires ANTHROPIC_API_KEY env var)
python -m src.analyzer analyze --template "template.pptx" --use-anthropic

# Non-interactive mode (returns data for external processing)
python -m src.analyzer analyze --template "template.pptx" --non-interactive
```

### Render Template to PNG

```bash
python -m src.analyzer render --template "template_business_case.pptx"
```

### List Saved Descriptions

```bash
python -m src.recreator list
```

### Recreate from Description

```bash
# Basic recreation
python -m src.recreator recreate --description "template_name_slide_1_final"

# Specify output path
python -m src.recreator recreate -d "template_name" -o "my_output.pptx"

# With validation against original
python -m src.recreator recreate -d "template_name" --validate
```

### Batch Recreate All Descriptions

```bash
python -m src.recreator batch --validate
```

### View a Description

```bash
python -m src.recreator view --description "template_name"
```

## How It Works

### Analysis Process

```
┌─────────────────────────────────────────────────────────────┐
│                    ITERATIVE ANALYSIS LOOP                  │
├─────────────────────────────────────────────────────────────┤
│                                                             │
│  1. Render original template slide to PNG                   │
│                    ↓                                        │
│  2. Generate initial description (vision analysis)          │
│                    ↓                                        │
│  3. ┌─────────── LOOP ──────────────────────────────────┐  │
│     │                                                    │  │
│     │  a. Generate PPTX from current description         │  │
│     │                    ↓                               │  │
│     │  b. Render generated PPTX to PNG                   │  │
│     │                    ↓                               │  │
│     │  c. Compare original vs generated (SSIM)           │  │
│     │                    ↓                               │  │
│     │  d. Similarity >= 95%? → EXIT (success!)           │  │
│     │                    ↓                               │  │
│     │  e. Get diff feedback, refine description          │  │
│     │                    ↓                               │  │
│     │     (repeat up to 10 iterations)                   │  │
│     │                                                    │  │
│     └────────────────────────────────────────────────────┘  │
│                    ↓                                        │
│  4. Save final description (JSON + Markdown)                │
│                                                             │
└─────────────────────────────────────────────────────────────┘
```

### Description Format

Descriptions are saved as JSON with this structure:

```json
{
    "slide_dimensions": {
        "width_inches": 13.333,
        "height_inches": 7.5,
        "aspect_ratio": "16:9"
    },
    "background": {
        "type": "solid",
        "color": "#1A365D"
    },
    "elements": [
        {
            "id": "title",
            "type": "textbox",
            "position": {
                "left_inches": 1.0,
                "top_inches": 2.5,
                "width_inches": 11.333,
                "height_inches": 1.5
            },
            "text_properties": {
                "placeholder_text": "PRESENTATION TITLE",
                "font_family": "Calibri Light",
                "font_size_pt": 44,
                "font_color": "#FFFFFF",
                "bold": true,
                "alignment": "center"
            }
        }
    ],
    "color_palette": ["#1A365D", "#FFFFFF", "#4299E1"],
    "design_notes": "Corporate style with dark blue background..."
}
```

## Configuration

Edit `config.py` to customize:

- **LibreOffice path**: `LIBREOFFICE_PATH`
- **Poppler path**: Auto-detected from conda environment
- **Max iterations**: `MAX_ITERATIONS` (default: 10)
- **Similarity threshold**: `SIMILARITY_THRESHOLD` (default: 0.95)
- **Render DPI**: `RENDER_DPI` (default: 150)

## Troubleshooting

### LibreOffice not found

Check the path in `config.py`:
```python
LIBREOFFICE_PATH = r"C:\Program Files\LibreOffice\program\soffice.exe"
```

Or for 32-bit:
```python
LIBREOFFICE_PATH = r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"
```

### Poppler not found

1. Ensure the conda environment is activated:
   ```bash
   conda activate pptx-design
   ```

2. Verify Poppler installation:
   ```bash
   where pdftoppm
   ```

3. If not found, reinstall:
   ```bash
   conda install -c conda-forge poppler
   ```

### Font substitution issues

If rendered slides don't match due to fonts:
1. Install the required fonts system-wide
2. Common fonts: Calibri, Calibri Light, Arial, Segoe UI

### Memory issues with large presentations

- Reduce `RENDER_DPI` in config.py
- Process slides individually instead of all at once

## API Reference

### renderer.py

```python
from src.renderer import render_slide, render_all_slides, get_slide_count

# Render single slide
image_path = render_slide(pptx_path, slide_index=0)

# Render all slides
image_paths = render_all_slides(pptx_path)

# Get slide count
count = get_slide_count(pptx_path)
```

### comparator.py

```python
from src.comparator import compute_similarity, generate_diff_image

# Compute SSIM similarity (0-1)
similarity = compute_similarity(image1_path, image2_path)

# Generate visual diff
diff_path = generate_diff_image(original_path, generated_path, mode="sidebyside")
```

### generator.py

```python
from src.generator import generate_slide_from_description

description = {
    "slide_dimensions": {"width_inches": 13.333, "height_inches": 7.5},
    "background": {"type": "solid", "color": "#1A365D"},
    "elements": [...]
}

pptx_path = generate_slide_from_description(description, output_path)
```

### descriptor.py

```python
from src.descriptor import describe_slide_design, save_description, load_description

# Generate description (returns prompt data for external processing)
prompt_data = describe_slide_design(image_path, use_anthropic=False)

# Or use Anthropic API directly
description = describe_slide_design(image_path, use_anthropic=True)

# Save/load descriptions
save_description(description, "template_name")
description = load_description("template_name")
```

## License

This project is for personal/educational use.

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Submit a pull request
