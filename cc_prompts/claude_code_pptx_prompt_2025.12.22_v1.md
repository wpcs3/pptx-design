# Claude Code Prompt: PowerPoint Template Analysis & Recreation System

## Project Overview

You are building a system that analyzes existing PowerPoint templates and learns to recreate their visual design programmatically. The system uses an iterative vision-based feedback loop: render a slide, compare it to the original, describe the differences, refine the code, and repeat until the output matches.

## Working Directory

```
C:\Users\wpcol\claudecode\pptx-design
```

## Template Location

```
C:\Users\wpcol\claudecode\pptx-design\pptx_templates
```

These templates contain formatting, layouts, placeholder boxes, and styling—but no actual content. They are professionally designed skeletons.

---

## Your Task

Build the complete infrastructure for this system. Create all scripts, utilities, and documentation needed to:

1. **Analyze any template** from the `pptx_templates` folder
2. **Extract a natural language description** of its design through iterative refinement
3. **Recreate the template programmatically** using python-pptx
4. **Validate the recreation** by visual comparison until it matches the original

---

## System Architecture

Create the following project structure:

```
C:\Users\wpcol\claudecode\pptx-design\
├── pptx_templates/              # (existing) Original templates
├── src/
│   ├── __init__.py
│   ├── renderer.py              # PPTX → PNG conversion via LibreOffice
│   ├── comparator.py            # Visual comparison logic
│   ├── descriptor.py            # Generate natural language descriptions
│   ├── generator.py             # Create PPTX from descriptions
│   ├── analyzer.py              # Main orchestration: analyze a template
│   └── recreator.py             # Main orchestration: recreate from description
├── descriptions/                # Generated template descriptions (JSON + markdown)
├── outputs/                     # Generated PPTX files and renders
├── diffs/                       # Visual diff images
├── config.py                    # Configuration (paths, iteration limits, etc.)
├── requirements.txt
├── setup.py                     # Optional: make it installable
└── README.md
```

---

## Detailed Component Specifications

### 1. `config.py`

```python
"""
Central configuration for the PPTX design system.
"""
from pathlib import Path

# Base paths
PROJECT_ROOT = Path(r"C:\Users\wpcol\claudecode\pptx-design")
TEMPLATE_DIR = PROJECT_ROOT / "pptx_templates"
OUTPUT_DIR = PROJECT_ROOT / "outputs"
DESCRIPTION_DIR = PROJECT_ROOT / "descriptions"
DIFF_DIR = PROJECT_ROOT / "diffs"

# Ensure directories exist
for dir_path in [OUTPUT_DIR, DESCRIPTION_DIR, DIFF_DIR]:
    dir_path.mkdir(parents=True, exist_ok=True)

# LibreOffice path (adjust if needed for Windows)
# Common locations:
LIBREOFFICE_PATH = r"C:\Program Files\LibreOffice\program\soffice.exe"
# Alternative: r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"

# Iteration settings
MAX_ITERATIONS = 10
SIMILARITY_THRESHOLD = 0.95  # 0-1 scale for programmatic comparison

# Rendering settings
RENDER_DPI = 150  # Higher = more detail but slower
```

---

### 2. `src/renderer.py`

Create a module that converts PPTX slides to PNG images using LibreOffice in headless mode.

**Requirements:**
- Function: `render_slide(pptx_path: Path, slide_index: int, output_path: Path) -> Path`
- Function: `render_all_slides(pptx_path: Path, output_dir: Path) -> list[Path]`
- Use LibreOffice headless mode: `soffice --headless --convert-to png --outdir <dir> <file>`
- Handle the fact that LibreOffice outputs to PDF first, then you may need ImageMagick or pdf2image to get PNGs
- Alternative approach: Convert to PDF, then use `pdf2image` (poppler) to extract PNGs
- Include error handling for missing LibreOffice installation
- Log the conversion process

**LibreOffice Command Reference:**
```bash
# Convert PPTX to PDF
soffice --headless --convert-to pdf --outdir ./output ./input.pptx

# Then use poppler/pdf2image to convert PDF pages to PNG
```

**Important Windows Considerations:**
- LibreOffice path may vary; make it configurable
- Use subprocess with shell=True on Windows if needed
- Handle spaces in paths properly

---

### 3. `src/comparator.py`

Create a module for comparing two images (original template render vs. generated output).

**Requirements:**
- Function: `compute_similarity(image1_path: Path, image2_path: Path) -> float`
  - Returns 0.0 (completely different) to 1.0 (identical)
  - Use structural similarity (SSIM) from scikit-image
  
- Function: `generate_diff_image(image1_path: Path, image2_path: Path, output_path: Path) -> Path`
  - Create a visual diff highlighting differences
  - Consider side-by-side, overlay, or heatmap approaches
  
- Function: `prepare_comparison_prompt(image1_path: Path, image2_path: Path) -> dict`
  - Prepare data structure for sending to vision model
  - Return base64-encoded images ready for API call

**Libraries to use:**
- `scikit-image` for SSIM
- `Pillow` for image manipulation
- `numpy` for array operations

---

### 4. `src/descriptor.py`

Create a module that generates natural language descriptions of slide designs.

**This is the core intelligence of the system.**

**Requirements:**

- Function: `describe_slide_design(image_path: Path) -> dict`
  - Analyze a single slide image
  - Return structured description including:
    ```python
    {
        "layout": {
            "type": "title_slide | section_header | content | two_column | etc",
            "orientation": "landscape",
            "margins": {"top": "...", "bottom": "...", "left": "...", "right": "..."}
        },
        "background": {
            "type": "solid | gradient | image",
            "primary_color": "#RRGGBB",
            "secondary_color": "#RRGGBB",  # if gradient
            "description": "Natural language description"
        },
        "elements": [
            {
                "type": "text_box | shape | image_placeholder | line | etc",
                "purpose": "title | subtitle | body | decoration | logo_placeholder",
                "position": {
                    "x": "percentage or inches from left",
                    "y": "percentage or inches from top",
                    "width": "...",
                    "height": "..."
                },
                "style": {
                    "font_family": "...",
                    "font_size": "...",
                    "font_color": "#RRGGBB",
                    "font_weight": "bold | normal",
                    "alignment": "left | center | right",
                    "fill_color": "...",
                    "border": "..."
                },
                "description": "Natural language description of this element"
            }
        ],
        "color_palette": ["#RRGGBB", ...],
        "overall_style": "Natural language summary of the design aesthetic",
        "python_pptx_hints": "Specific guidance for recreating with python-pptx"
    }
    ```

- Function: `refine_description(current_description: dict, diff_feedback: str) -> dict`
  - Take existing description and feedback about what doesn't match
  - Return improved description
  
- Function: `save_description(description: dict, template_name: str)`
  - Save to `descriptions/{template_name}.json`
  - Also save a human-readable markdown version

**Implementation Notes:**
- This module will call Claude's API (or use the orchestrating Claude instance)
- The prompt should ask for extremely precise measurements and colors
- Include few-shot examples of good descriptions in the prompt

---

### 5. `src/generator.py`

Create a module that generates PPTX files from natural language descriptions.

**Requirements:**

- Function: `generate_slide_from_description(description: dict, output_path: Path) -> Path`
  - Take a structured description (from descriptor.py)
  - Create a PPTX file using python-pptx
  - Return path to generated file

- Function: `apply_element(slide, element: dict)`
  - Add a single element to a slide based on its description
  - Handle: text boxes, shapes, lines, placeholders, images

- Function: `parse_color(color_str: str) -> RGBColor`
  - Convert various color formats to python-pptx RGBColor

- Function: `parse_measurement(measurement_str: str, slide_width: int) -> int`
  - Convert "2 inches", "50%", "150pt" to EMUs (English Metric Units)

**python-pptx Reference Patterns:**
```python
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

prs = Presentation()
prs.slide_width = Inches(13.333)  # Widescreen 16:9
prs.slide_height = Inches(7.5)

slide_layout = prs.slide_layouts[6]  # Blank layout
slide = prs.slides.add_slide(slide_layout)

# Add text box
from pptx.util import Inches, Pt
textbox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(5), Inches(1))
tf = textbox.text_frame
p = tf.paragraphs[0]
p.text = "Title"
p.font.size = Pt(44)
p.font.bold = True
p.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
```

---

### 6. `src/analyzer.py`

**Main orchestration script for analyzing an existing template.**

This implements the iterative loop for EXTRACTING a description from a template.

```python
"""
Analyze an existing PPTX template and generate a precise natural language description.

Usage:
    python -m src.analyzer --template "template_name.pptx" --slide 0
    python -m src.analyzer --template "template_name.pptx" --all-slides
"""

def analyze_template(template_path: Path, slide_index: int = 0, max_iterations: int = 10):
    """
    Iteratively analyze a template slide until the description is precise enough
    to recreate it accurately.
    
    Process:
    1. Render the original template slide to PNG
    2. Generate initial description using vision
    3. Loop:
        a. Generate PPTX from current description
        b. Render generated PPTX to PNG
        c. Compare original vs generated
        d. If similarity > threshold: done
        e. Otherwise: get diff feedback, refine description
    4. Save final description
    """
    pass
```

**The key insight:** We're not trying to recreate the template in this script—we're trying to create a DESCRIPTION that's detailed enough that the generator can recreate it. The iteration refines the description, not the code.

---

### 7. `src/recreator.py`

**Main orchestration script for recreating a template from a saved description.**

```python
"""
Recreate a PPTX slide from a saved description.

Usage:
    python -m src.recreator --description "template_name.json" --output "output.pptx"
"""

def recreate_from_description(description_path: Path, output_path: Path, validate: bool = True):
    """
    Generate a PPTX from a saved description.
    
    If validate=True, also render and compare to original (if available).
    """
    pass
```

---

### 8. Vision Comparison Prompt Template

Create a file `src/prompts/comparison_prompt.txt`:

```
You are analyzing two PowerPoint slide images to identify precise differences.

IMAGE 1: The original template (target design)
IMAGE 2: A programmatically generated recreation attempt

Your task is to describe EXACTLY what differs between them so the recreation can be improved.

Be extremely specific about:

1. **Position differences**: "The title text box is 0.5 inches too low" not "the title is in the wrong place"

2. **Size differences**: "The subtitle font appears to be 18pt but should be 24pt"

3. **Color differences**: Provide exact hex codes if possible. "The background is #1E3A5F but should be #1A365D"

4. **Alignment issues**: "The body text is left-aligned but should be centered"

5. **Missing elements**: "There is a decorative line below the title in the original that is missing"

6. **Extra elements**: "The recreation has a footer that doesn't exist in the original"

7. **Spacing/margins**: "The left margin appears to be 0.5 inches but should be 1 inch"

8. **Font issues**: "The font appears to be Arial but should be Calibri Light"

Format your response as a structured list of differences, ordered by visual prominence (most noticeable first).

If the images appear nearly identical, say "MATCH" and note any minor, negligible differences.
```

---

### 9. Description Generation Prompt Template

Create a file `src/prompts/description_prompt.txt`:

```
You are a PowerPoint design analyst. Your task is to describe this slide template with enough precision that it can be recreated programmatically using python-pptx.

Analyze this slide image and provide a complete specification.

## Required Output Format (JSON):

```json
{
    "slide_dimensions": {
        "width_inches": 13.333,
        "height_inches": 7.5,
        "aspect_ratio": "16:9"
    },
    "background": {
        "type": "solid|gradient|image",
        "color": "#RRGGBB",
        "gradient_start": "#RRGGBB",
        "gradient_end": "#RRGGBB", 
        "gradient_direction": "horizontal|vertical|diagonal"
    },
    "elements": [
        {
            "id": "element_1",
            "type": "textbox|shape|line|image_placeholder",
            "position": {
                "left_inches": 0.0,
                "top_inches": 0.0,
                "width_inches": 0.0,
                "height_inches": 0.0
            },
            "text_properties": {
                "placeholder_text": "TITLE GOES HERE",
                "font_family": "Calibri Light",
                "font_size_pt": 44,
                "font_color": "#RRGGBB",
                "bold": false,
                "italic": false,
                "alignment": "center|left|right",
                "vertical_alignment": "top|middle|bottom"
            },
            "shape_properties": {
                "fill_color": "#RRGGBB",
                "fill_transparency": 0,
                "border_color": "#RRGGBB",
                "border_width_pt": 1,
                "shape_type": "rectangle|rounded_rectangle|oval|line"
            }
        }
    ],
    "color_palette": ["#RRGGBB", "#RRGGBB"],
    "design_notes": "Free-form observations about the design style, spacing patterns, visual hierarchy, etc."
}
```

## Guidelines:

1. **Measure precisely**: Estimate positions and sizes as accurately as possible. Use the slide dimensions as reference (standard widescreen is 13.333" × 7.5").

2. **Identify ALL elements**: Include decorative shapes, lines, backgrounds, even subtle design elements.

3. **Note visual hierarchy**: Which elements are most prominent? This affects z-order.

4. **Observe spacing patterns**: Are elements evenly spaced? Following a grid? Margins consistent?

5. **Color accuracy**: Try to identify exact colors. Note if colors are from a brand palette.

6. **Font identification**: Identify fonts as precisely as possible. Common PowerPoint fonts: Calibri, Calibri Light, Arial, Century Gothic, Segoe UI.

Output ONLY the JSON, no additional commentary.
```

---

### 10. `requirements.txt`

```
python-pptx>=0.6.21
Pillow>=9.0.0
scikit-image>=0.19.0
numpy>=1.21.0
pdf2image>=1.16.0
anthropic>=0.18.0  # For Claude API calls if running standalone
click>=8.0.0       # For CLI interface
rich>=13.0.0       # For nice terminal output
```

---

### 11. `README.md`

Create comprehensive documentation including:

1. **Installation instructions**
   - Python 3.10+ required
   - Install LibreOffice (link to download)
   - Install poppler for pdf2image (Windows: download binaries, add to PATH)
   - `pip install -r requirements.txt`

2. **Configuration**
   - How to set LibreOffice path
   - Adjusting iteration limits and thresholds

3. **Usage examples**
   ```bash
   # Analyze a template and generate its description
   python -m src.analyzer --template "Corporate Blue.pptx"
   
   # Recreate a template from its description
   python -m src.recreator --description "Corporate Blue.json" --output "recreated.pptx"
   
   # Analyze all templates in the folder
   python -m src.analyzer --all
   ```

4. **How it works** (explain the iterative loop)

5. **Troubleshooting**
   - LibreOffice not found
   - Poppler not installed
   - Font substitution issues

---

## Implementation Order

Build the system in this order:

1. **First: `config.py` and `requirements.txt`** — Set up the foundation
2. **Second: `renderer.py`** — Get LibreOffice rendering working (this is the critical dependency)
3. **Third: `comparator.py`** — Image comparison utilities
4. **Fourth: `generator.py`** — PPTX generation from descriptions
5. **Fifth: `descriptor.py`** — Vision-based description generation
6. **Sixth: `analyzer.py`** — Main orchestration loop
7. **Seventh: `recreator.py`** — Simplified recreation from saved descriptions
8. **Finally: `README.md`** — Documentation

---

## Testing Strategy

After building each component, test it:

1. **Test renderer**: Convert a template PPTX to PNG, verify the image looks correct
2. **Test comparator**: Compare two similar images, verify similarity score makes sense
3. **Test generator**: Create a simple slide from a hardcoded description
4. **Test full loop**: Run analyzer on a simple template with clear geometric shapes

Start with the simplest possible template (solid color background, one centered title) before attempting complex designs.

---

## Critical Notes

1. **LibreOffice installation**: The user must have LibreOffice installed. Check for it at startup and give clear error messages.

2. **Poppler for pdf2image**: On Windows, this requires downloading poppler binaries and adding to PATH. Document this clearly.

3. **Font availability**: Fonts must be installed on the system for LibreOffice to render them correctly. If a template uses a font not installed, LibreOffice will substitute.

4. **EMU conversions**: python-pptx uses English Metric Units internally. 914400 EMUs = 1 inch. The `Inches()`, `Pt()`, and `Cm()` helpers handle conversions.

5. **Iteration limits**: Set reasonable limits (10 iterations max) to prevent infinite loops if perfect matching isn't achievable.

6. **API costs**: If using Claude API for vision comparison, be aware of costs. Consider caching descriptions.

---

## Begin Implementation

Start by creating the project structure and implementing `config.py`. Then move to `renderer.py` to verify LibreOffice integration works on the user's Windows system.

Ask the user to confirm LibreOffice is installed before proceeding with render tests.
