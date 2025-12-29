# IMPROVEMENT_ROADMAP.md

## GitHub Repository Analysis for pptx_design Improvements

**Generated**: 2025-12-28
**Based on**: Systematic review of 13 open-source PowerPoint generation/extraction projects

---

## Executive Summary

This roadmap consolidates findings from reviewing:
- **TIER 1**: Presenton, PPTAgent, SlideDeck AI (AI-powered generation)
- **TIER 2**: pptxtojson, airppt-parser, pptx-compose (extraction/parsing)
- **TIER 3**: pptx-automizer, pptx-template, PptxGenJS (template management)
- **TIER 4**: LayoutParser, DiT/LayoutLMv3 (ML-based analysis)
- **TIER 5**: Office-PowerPoint-MCP-Server, Slidev (emerging approaches)

**Top 3 Strategic Improvements:**
1. **MCP Server Integration** - Enable Claude/agent access to PresentationBuilder
2. **LLM Content Pipeline** - Add multi-provider LLM support for content generation
3. **PPTEval-style Quality Framework** - Implement content/design/coherence evaluation

---

## Repository Analyses

### TIER 1: AI-Powered Generation

---

## Presenton

**Repository**: [github.com/presenton/presenton](https://github.com/presenton/presenton)

### Key Findings

- **Multi-LLM Architecture**: Environment variable-based provider switching (OpenAI, Claude, Gemini, Ollama, Custom)
- **Tone Controls**: 6 tone options (default, casual, professional, funny, educational, sales_pitch)
- **Verbosity Settings**: concise, standard, text-heavy via `length` parameter
- **REST API Design**: Clean `/api/v1/ppt/presentation/generate` endpoint returning presentation_id + paths

### Recommended Improvements for pptx_design

#### Improvement 1: Multi-LLM Provider Support

- **What**: Add configurable LLM provider switching via environment variables
- **Why**: Currently our system has no LLM integration; this addresses gap #1 (No AI/LLM integration)
- **Where**: New module `pptx_generator/modules/llm_provider.py`
- **How**:
```python
# pptx_generator/modules/llm_provider.py
import os
from typing import Optional
from anthropic import Anthropic
import openai

class LLMProvider:
    PROVIDERS = {
        'anthropic': lambda: Anthropic(),
        'openai': lambda: openai.OpenAI(),
        'ollama': lambda: OllamaClient(os.getenv('OLLAMA_URL', 'http://localhost:11434')),
    }

    def __init__(self):
        self.provider = os.getenv('LLM_PROVIDER', 'anthropic')
        self.model = os.getenv('LLM_MODEL', 'claude-3-5-sonnet-20241022')
        self.client = self.PROVIDERS[self.provider]()

    def generate(self, prompt: str, system: str = None) -> str:
        if self.provider == 'anthropic':
            return self.client.messages.create(
                model=self.model,
                max_tokens=4096,
                system=system or "",
                messages=[{"role": "user", "content": prompt}]
            ).content[0].text
        # ... other providers
```
- **Effort**: Medium
- **Priority**: P0

#### Improvement 2: Tone and Verbosity Controls

- **What**: Add tone (professional, casual, etc.) and verbosity (concise, standard, detailed) parameters
- **Why**: Enables content customization without manual editing
- **Where**: `pptx_generator/modules/outline_generator.py`, `pptx_design/builder.py`
- **How**: Add parameters to generation methods, inject into LLM prompts
- **Effort**: Low
- **Priority**: P1

### Code Patterns to Adopt

```python
# Tone injection pattern from Presenton
TONE_PROMPTS = {
    "professional": "Use formal business language, data-driven statements",
    "casual": "Use conversational tone, relatable examples",
    "sales_pitch": "Emphasize benefits, use persuasive language, include CTAs",
}

def generate_content(topic: str, tone: str = "professional", verbosity: str = "standard"):
    tone_instruction = TONE_PROMPTS.get(tone, TONE_PROMPTS["professional"])
    verbosity_instruction = f"Content density: {verbosity}"
    # Inject into prompt...
```

---

## PPTAgent

**Repository**: [github.com/icip-cas/PPTAgent](https://github.com/icip-cas/PPTAgent)

### Key Findings

- **Two-Phase Architecture**: Stage I (analysis via `induct.py`) → Stage II (generation via `pptgen.py`)
- **Functional Slide Types**: Opening, Table of Contents, Section Header, Ending (auto-inserted)
- **PPTEval Framework**: Evaluates Content accuracy, Visual Design, Logical Coherence
- **Best Practices**:
  - Max 6 elements per slide for visual clarity
  - Text should occupy ~60% of element space
  - `length_factor` parameter (0.5-2.5) to adjust text density
- **Configuration**: `num_slides`, `length_factor`, `sim_bound` (similarity threshold)

### Recommended Improvements for pptx_design

#### Improvement 3: PPTEval-style Quality Framework

- **What**: Implement evaluation metrics for generated presentations
- **Why**: Addresses gap #4 (No quality evaluation framework)
- **Where**: New module `pptx_design/evaluation.py`
- **How**:
```python
# pptx_design/evaluation.py
from dataclasses import dataclass
from typing import List
from pptx import Presentation

@dataclass
class EvaluationResult:
    content_score: float      # 0-1: Topic coverage, accuracy
    design_score: float       # 0-1: Visual consistency, element count
    coherence_score: float    # 0-1: Logical flow, section structure
    overall_score: float
    issues: List[str]

class PresentationEvaluator:
    MAX_ELEMENTS_PER_SLIDE = 6
    TARGET_TEXT_DENSITY = 0.6

    def evaluate(self, pptx_path: str) -> EvaluationResult:
        prs = Presentation(pptx_path)
        content = self._evaluate_content(prs)
        design = self._evaluate_design(prs)
        coherence = self._evaluate_coherence(prs)
        return EvaluationResult(
            content_score=content,
            design_score=design,
            coherence_score=coherence,
            overall_score=(content + design + coherence) / 3,
            issues=self.issues
        )

    def _evaluate_design(self, prs: Presentation) -> float:
        issues = []
        for i, slide in enumerate(prs.slides):
            shape_count = len([s for s in slide.shapes if s.has_text_frame or s.has_table])
            if shape_count > self.MAX_ELEMENTS_PER_SLIDE:
                issues.append(f"Slide {i+1}: Too many elements ({shape_count})")
        return max(0, 1 - len(issues) * 0.1)
```
- **Effort**: Medium
- **Priority**: P1

#### Improvement 4: Functional Slide Type Auto-Insertion

- **What**: Automatically insert Opening, TOC, Section Headers, Ending slides
- **Why**: Ensures consistent presentation structure without manual specification
- **Where**: `pptx_generator/modules/orchestrator.py`
- **How**: Rule-based insertion after outline generation
- **Effort**: Low
- **Priority**: P2

#### Improvement 5: Length Factor Parameter

- **What**: Add `length_factor` (0.5-2.5) to control text density relative to templates
- **Why**: Enables content volume customization for different audiences
- **Where**: `pptx_generator/modules/template_renderer.py`
- **Effort**: Low
- **Priority**: P2

### Code Patterns to Adopt

```python
# From PPTAgent: Slide element limit enforcement
class SlideConstraints:
    MAX_ELEMENTS = 6
    MAX_BULLETS = 6
    MAX_TEXT_PER_ELEMENT = 50  # words

    @staticmethod
    def validate_slide(slide_content: dict) -> List[str]:
        warnings = []
        if len(slide_content.get('elements', [])) > SlideConstraints.MAX_ELEMENTS:
            warnings.append("Too many elements - consider splitting slide")
        return warnings
```

---

## SlideDeck AI

**Repository**: [github.com/barun-saha/slide-deck-ai](https://github.com/barun-saha/slide-deck-ai)

### Key Findings

- **LiteLLM Integration**: Unified API for 8+ LLM providers (OpenAI, Gemini, Claude, SambaNova, etc.)
- **JSON5 Parsing**: Tolerant JSON parsing with fallback error correction
- **Cascading Layout Handlers**: Icons → Tables → Two-Column → Step-by-Step → Default
- **Pexels Image Search**: Keyword extraction → image search → probabilistic insertion
- **Refinement Loop**: `revise()` method with 16-message chat history limit

### Recommended Improvements for pptx_design

#### Improvement 6: LiteLLM Integration

- **What**: Use LiteLLM for unified multi-provider LLM access
- **Why**: Simpler than custom provider code, supports 100+ models
- **Where**: `pptx_generator/modules/llm_provider.py`
- **How**:
```python
# pip install litellm
from litellm import completion

def generate_with_litellm(prompt: str, model: str = "claude-3-5-sonnet-20241022"):
    response = completion(
        model=model,
        messages=[{"role": "user", "content": prompt}],
    )
    return response.choices[0].message.content
```
- **Effort**: Low
- **Priority**: P0
- **Dependency**: `pip install litellm`

#### Improvement 7: Cascading Layout Handler Pattern

- **What**: Implement layout selection cascade based on content type detection
- **Why**: Automates optimal layout selection instead of manual specification
- **Where**: `pptx_generator/modules/template_renderer.py`
- **How**:
```python
class LayoutCascade:
    HANDLERS = [
        ('icons', lambda c: bool(c.get('icons'))),
        ('table', lambda c: bool(c.get('table_data'))),
        ('two_column', lambda c: len(c.get('columns', [])) == 2),
        ('step_by_step', lambda c: any(s.startswith('>>') for s in c.get('bullets', []))),
        ('default', lambda c: True),
    ]

    def select_layout(self, content: dict) -> str:
        for layout_name, matcher in self.HANDLERS:
            if matcher(content):
                return layout_name
        return 'default'
```
- **Effort**: Medium
- **Priority**: P1

#### Improvement 8: Image Search Integration

- **What**: Add Pexels/Unsplash API integration for automatic image sourcing
- **Why**: Enables automatic imagery without manual asset selection
- **Where**: New module `pptx_generator/modules/image_search.py`
- **Effort**: Medium
- **Priority**: P2

### Code Patterns to Adopt

```python
# From SlideDeck AI: JSON5 parsing with fallback
import json5

def parse_llm_json(response: str) -> dict:
    try:
        return json5.loads(response)
    except Exception:
        # Attempt to fix common JSON issues
        fixed = fix_malformed_json(response)
        return json5.loads(fixed)

def fix_malformed_json(text: str) -> str:
    # Remove trailing commas, fix quotes, etc.
    import re
    text = re.sub(r',\s*([}\]])', r'\1', text)
    return text
```

---

### TIER 2: Extraction & Parsing

---

## pptxtojson-english

**Repository**: [github.com/alastairapple/pptxtojson-english](https://github.com/alastairapple/pptxtojson-english)

### Key Findings

- **layoutElements Property**: Separates master slide elements from content elements
- **themeColors Array**: Extracts theme colors as hex array `['#4472C4', '#ED7D31', ...]`
- **Points Coordinate System**: All measurements in pt (points), not pixels
- **Rich Text as HTML**: Text content preserved with inline HTML formatting

### Recommended Improvements for pptx_design

#### Improvement 9: Standardize on Points Coordinate System

- **What**: Ensure all extracted/generated coordinates use points consistently
- **Why**: Eliminates EMU↔pixel conversion errors
- **Where**: `pptx_extractor/`, `pptx_component_library/`
- **How**: Add conversion utilities, update extraction schemas
- **Effort**: Medium
- **Priority**: P2

#### Improvement 10: Layout Elements Separation

- **What**: Distinguish inherited master elements from content elements in extraction
- **Why**: Prevents duplicate elements when using extracted templates
- **Where**: `pptx_extractor/master_extractor.py`
- **Effort**: Low
- **Priority**: P2

---

## airppt-parser

**Repository**: [github.com/airpptx/airppt-parser](https://github.com/airpptx/airppt-parser)

### Key Findings

- **PowerPointElement Interface**: Standardized schema with name, shapeType, positions, paragraph, shape, fontStyle, links, raw
- **Dual Position Tracking**: `elementPosition` (x/y) + `elementOffsetPosition` (cx/cy dimensions)
- **Opacity Support**: Shape properties include opacity values
- **Raw XML Preservation**: Stores unprocessed XML for debugging

### Recommended Improvements for pptx_design

#### Improvement 11: Standardized Element Interface

- **What**: Define a consistent `SlideElement` dataclass matching airppt-parser's schema
- **Why**: Enables better component library interoperability
- **Where**: New file `pptx_design/schemas.py`
- **How**:
```python
from dataclasses import dataclass
from typing import Optional, Dict, Any

@dataclass
class Position:
    x: float  # points
    y: float  # points

@dataclass
class Dimensions:
    width: float   # points
    height: float  # points

@dataclass
class SlideElement:
    name: str
    element_type: str  # text, image, shape, table, chart
    position: Position
    dimensions: Dimensions
    paragraph: Optional[Dict[str, Any]] = None
    shape_style: Optional[Dict[str, Any]] = None
    font_style: Optional[Dict[str, Any]] = None
    opacity: float = 1.0
    raw_xml: Optional[str] = None
```
- **Effort**: Medium
- **Priority**: P2

---

## pptx-compose

**Repository**: [github.com/shobhitsharma/pptx-compose](https://github.com/shobhitsharma/pptx-compose)

### Key Findings

- **Bidirectional Conversion**: `toJSON()` and `toPPTX()` methods
- **Round-trip Capability**: Parse → modify → regenerate workflow
- **CLI Tool**: `bin/convert` for command-line operations

### Recommended Improvements for pptx_design

#### Improvement 12: Round-trip Testing

- **What**: Add test suite verifying extraction → generation fidelity
- **Why**: Ensures our pipeline doesn't lose information
- **Where**: `tests/test_roundtrip.py`
- **How**: Extract template → regenerate → compare SSIM scores
- **Effort**: Medium
- **Priority**: P2

---

### TIER 3: Template Management & Automation

---

## pptx-automizer

**Repository**: [github.com/singerla/pptx-automizer](https://github.com/singerla/pptx-automizer)

### Key Findings

- **Template Labels**: Load templates with memorable identifiers (e.g., 'shapes', 'graph')
- **Auto Import Slide Masters**: `autoImportSlideMasters: true` preserves original formatting
- **xmldom Callbacks**: Deep XML customization via DOM callbacks
- **Multi-template Composition**: Merge slides from different templates

### Recommended Improvements for pptx_design

#### Improvement 13: Template Label System

- **What**: Enhance TemplateRegistry to support loading by memorable labels
- **Why**: Cleaner API than file paths
- **Where**: `pptx_design/registry.py`
- **How**:
```python
class TemplateRegistry:
    def __init__(self):
        self.templates = {}
        self.labels = {}  # label -> template_id mapping

    def register(self, template_path: str, label: str = None):
        template_id = self._extract_id(template_path)
        self.templates[template_id] = Template(template_path)
        if label:
            self.labels[label] = template_id

    def get(self, label_or_id: str) -> Template:
        template_id = self.labels.get(label_or_id, label_or_id)
        return self.templates[template_id]
```
- **Effort**: Low
- **Priority**: P2

#### Improvement 14: Multi-template Slide Merging

- **What**: Enable composing presentations from multiple template sources
- **Why**: Allows mixing slide styles from different templates
- **Where**: `pptx_design/builder.py`
- **Effort**: Medium
- **Priority**: P3

---

## pptx-template

**Repository**: [github.com/m3dev/pptx-template](https://github.com/m3dev/pptx-template)

### Key Findings

- **Bracket DSL**: Placeholders like `{sales.0.june.us}` with dot-notation paths
- **Array Indexing**: `.0`, `.1` for accessing list elements
- **JSON Config Binding**: Model data from JSON files or Python dicts
- **Selective Substitution**: Only identified placeholders modified

### Recommended Improvements for pptx_design

#### Improvement 15: Enhanced Placeholder DSL

- **What**: Support dot-notation JSON paths in placeholder syntax
- **Why**: Enables complex data binding without custom code
- **Where**: `pptx_generator/modules/template_renderer.py`
- **How**:
```python
import re
from functools import reduce

def resolve_path(data: dict, path: str) -> Any:
    """Resolve dot-notation path like 'sales.0.june.us' in data dict."""
    parts = path.split('.')
    def get_part(obj, key):
        if isinstance(obj, list):
            return obj[int(key)]
        return obj.get(key)
    return reduce(get_part, parts, data)

def substitute_placeholders(text: str, data: dict) -> str:
    pattern = r'\{([a-zA-Z0-9_.]+)\}'
    def replacer(match):
        path = match.group(1)
        value = resolve_path(data, path)
        return str(value) if value is not None else match.group(0)
    return re.sub(pattern, replacer, text)
```
- **Effort**: Low
- **Priority**: P2

---

## PptxGenJS

**Repository**: [github.com/gitbrent/PptxGenJS](https://github.com/gitbrent/PptxGenJS)

### Key Findings

- **Programmatic Slide Masters**: Define custom masters via API
- **tableToSlides()**: One-line HTML table → PowerPoint conversion
- **Rich API**: Text, tables, shapes, images, charts with full formatting control
- **4-line Presentations**: Minimal boilerplate for simple decks

### Recommended Improvements for pptx_design

#### Improvement 16: HTML Table Import

- **What**: Add method to convert HTML tables to PowerPoint tables
- **Why**: Enables easy import from web data
- **Where**: `pptx_design/builder.py`
- **How**: Parse HTML table with BeautifulSoup, map to python-pptx table
- **Effort**: Medium
- **Priority**: P3

---

### TIER 4: Vision-Based Analysis

---

## LayoutParser

**Repository**: [github.com/Layout-Parser/layout-parser](https://github.com/Layout-Parser/layout-parser)

### Key Findings

- **AutoLayoutModel**: Unified API for layout detection models
- **Region Filtering**: `layout.filter_by()` for spatial queries
- **OCR Integration**: TesseractAgent for text extraction
- **Detectron2 Backend**: EfficientDet/PubLayNet models available

### Recommended Improvements for pptx_design

#### Improvement 17: ML-Based Slide Layout Classification

- **What**: Use LayoutParser to classify slide layouts from rendered images
- **Why**: Improves semantic understanding of slide designs (addresses gap #3)
- **Where**: New module `pptx_extractor/layout_classifier.py`
- **How**:
```python
import layoutparser as lp
from PIL import Image

class SlideLayoutClassifier:
    def __init__(self):
        self.model = lp.AutoLayoutModel('lp://PubLayNet/faster_rcnn_R_50_FPN_3x/config')

    def classify(self, slide_image_path: str) -> dict:
        image = Image.open(slide_image_path)
        layout = self.model.detect(image)

        regions = {
            'title': [],
            'text': [],
            'figure': [],
            'table': [],
        }

        for block in layout:
            region_type = block.type.lower()
            if region_type in regions:
                regions[region_type].append({
                    'bbox': block.coordinates,
                    'score': block.score
                })

        return {
            'layout_type': self._infer_layout_type(regions),
            'regions': regions
        }

    def _infer_layout_type(self, regions: dict) -> str:
        if len(regions['title']) == 1 and not regions['text']:
            return 'title_slide'
        if regions['figure'] and not regions['table']:
            return 'image_slide'
        # ... more heuristics
        return 'content_slide'
```
- **Effort**: High
- **Priority**: P2
- **Dependency**: `pip install layoutparser[detectron2]`

---

## DiT / LayoutLMv3

**Repository**: [github.com/microsoft/unilm/tree/master/layoutlmv3](https://github.com/microsoft/unilm/tree/master/layoutlmv3)

### Key Findings

- **Multimodal Architecture**: Unified text + image + layout understanding
- **Document AI Tasks**: Form understanding (F1: 0.9059), layout analysis (mAP: 95.1%)
- **HuggingFace Models**: `microsoft/layoutlmv3-base`, `microsoft/layoutlmv3-large`
- **Pre-training Objectives**: Text masking, image masking, word-patch alignment

### Recommended Improvements for pptx_design

#### Improvement 18: LayoutLMv3 for Slide Understanding

- **What**: Integrate LayoutLMv3 for semantic slide content extraction
- **Why**: Enables understanding of what slides "mean" not just what they contain
- **Where**: New module `pptx_extractor/semantic_analyzer.py`
- **How**:
```python
from transformers import AutoProcessor, AutoModelForTokenClassification
import torch

class SemanticSlideAnalyzer:
    def __init__(self):
        self.processor = AutoProcessor.from_pretrained("microsoft/layoutlmv3-base")
        self.model = AutoModelForTokenClassification.from_pretrained(
            "microsoft/layoutlmv3-base",
            num_labels=7  # Custom slide element labels
        )

    def analyze(self, slide_image, text_boxes: list) -> dict:
        encoding = self.processor(
            slide_image,
            text_boxes,
            return_tensors="pt"
        )
        outputs = self.model(**encoding)
        predictions = torch.argmax(outputs.logits, dim=-1)
        return self._decode_predictions(predictions, text_boxes)
```
- **Effort**: High
- **Priority**: P3
- **Dependency**: `pip install transformers torch`

---

### TIER 5: Emerging Approaches

---

## Office-PowerPoint-MCP-Server

**Repository**: [github.com/GongRzhe/Office-PowerPoint-MCP-Server](https://github.com/GongRzhe/Office-PowerPoint-MCP-Server)

### Key Findings

- **34 MCP Tools**: Organized into 11 modules (presentation, content, template, structure, design, etc.)
- **Unified Operations**: `manage_text`, `manage_image`, `apply_professional_design`
- **Global State Tracking**: Maintains active presentation references
- **Extraction for Verification**: `extract_slide_text` enables read-back loops

### Recommended Improvements for pptx_design

#### Improvement 19: MCP Server for PresentationBuilder

- **What**: Expose PresentationBuilder as an MCP server for Claude/agent access
- **Why**: Addresses gap #5 (No MCP server integration)
- **Where**: New file `pptx_design/mcp_server.py`
- **How**:
```python
from mcp import Server
from pptx_design import PresentationBuilder

server = Server("pptx-design")

@server.tool("create_presentation")
async def create_presentation(template: str, title: str) -> dict:
    """Create a new presentation with the specified template."""
    builder = PresentationBuilder(template)
    builder.add_title_slide(title)
    return {"status": "created", "presentation_id": builder.id}

@server.tool("add_slide")
async def add_slide(presentation_id: str, slide_type: str, content: dict) -> dict:
    """Add a slide to an existing presentation."""
    builder = get_builder(presentation_id)
    if slide_type == "title_content":
        builder.add_content_slide(content["title"], bullets=content.get("bullets", []))
    elif slide_type == "data_chart":
        builder.add_chart_slide(content["title"], content["chart_data"])
    return {"status": "added", "slide_count": len(builder.presentation.slides)}

@server.tool("save_presentation")
async def save_presentation(presentation_id: str, output_path: str) -> dict:
    """Save the presentation to a file."""
    builder = get_builder(presentation_id)
    builder.save(output_path)
    return {"status": "saved", "path": output_path}

@server.tool("extract_slide_text")
async def extract_slide_text(presentation_id: str, slide_index: int) -> dict:
    """Extract text content from a specific slide for verification."""
    builder = get_builder(presentation_id)
    slide = builder.presentation.slides[slide_index]
    text_content = []
    for shape in slide.shapes:
        if shape.has_text_frame:
            text_content.append(shape.text_frame.text)
    return {"slide_index": slide_index, "text": text_content}
```
- **Effort**: Medium
- **Priority**: P0
- **Dependency**: `pip install mcp`

### Integration Points

- MCP server wraps existing `PresentationBuilder` API
- Enables agentic workflows: generate outline → create slides → review → refine
- Extraction tools enable verification loops (generate → extract → validate)

---

## Slidev

**Repository**: [github.com/slidevjs/slidev](https://github.com/slidevjs/slidev)

### Key Findings

- **Markdown-Based**: Plain text → interactive slides
- **UnoCSS Theming**: On-demand utility-first CSS
- **Multi-Format Export**: PDF, PNG, PPTX
- **Vue 3 Integration**: Embed interactive components
- **Developer-Centric**: Code highlighting, live coding, Mermaid diagrams

### Recommended Improvements for pptx_design

#### Improvement 20: Markdown Input Support

- **What**: Accept Markdown as alternative to JSON outlines
- **Why**: Markdown is more natural for content authoring
- **Where**: New module `pptx_generator/modules/markdown_parser.py`
- **How**:
```python
import re
from typing import List, Dict

def markdown_to_outline(markdown: str) -> Dict:
    """Convert Markdown to presentation outline JSON."""
    slides = []
    current_slide = None

    for line in markdown.split('\n'):
        if line.startswith('# '):  # New section/title
            if current_slide:
                slides.append(current_slide)
            current_slide = {
                "slide_type": "title_slide",
                "content": {"title": line[2:].strip()}
            }
        elif line.startswith('## '):  # Content slide title
            if current_slide:
                slides.append(current_slide)
            current_slide = {
                "slide_type": "title_content",
                "content": {"title": line[3:].strip(), "bullets": []}
            }
        elif line.startswith('- '):  # Bullet point
            if current_slide and "bullets" in current_slide["content"]:
                current_slide["content"]["bullets"].append(line[2:].strip())

    if current_slide:
        slides.append(current_slide)

    return {"slides": slides}
```
- **Effort**: Medium
- **Priority**: P2

---

## Consolidated Improvement Roadmap

### Quick Wins (Low Effort, High Impact)

| # | Improvement | Where | Dependencies | Priority |
|---|-------------|-------|--------------|----------|
| 2 | Tone/Verbosity Controls | `outline_generator.py` | None | P1 |
| 6 | LiteLLM Integration | New `llm_provider.py` | `litellm` | P0 |
| 4 | Functional Slide Auto-Insert | `orchestrator.py` | None | P2 |
| 5 | Length Factor Parameter | `template_renderer.py` | None | P2 |
| 10 | Layout Elements Separation | `master_extractor.py` | None | P2 |
| 13 | Template Label System | `registry.py` | None | P2 |
| 15 | Enhanced Placeholder DSL | `template_renderer.py` | None | P2 |

### Strategic Enhancements (Medium Effort, High Impact)

| # | Improvement | Where | Dependencies | Priority |
|---|-------------|-------|--------------|----------|
| 1 | Multi-LLM Provider Support | New `llm_provider.py` | `anthropic`, `openai` | P0 |
| 3 | PPTEval Quality Framework | New `evaluation.py` | None | P1 |
| 7 | Cascading Layout Handlers | `template_renderer.py` | None | P1 |
| 8 | Image Search Integration | New `image_search.py` | `requests` | P2 |
| 11 | Standardized Element Interface | New `schemas.py` | None | P2 |
| 12 | Round-trip Testing | `tests/` | None | P2 |
| 19 | MCP Server | New `mcp_server.py` | `mcp` | P0 |
| 20 | Markdown Input Support | New `markdown_parser.py` | None | P2 |

### Future Capabilities (High Effort, Transformational)

| # | Improvement | Where | Dependencies | Priority |
|---|-------------|-------|--------------|----------|
| 9 | Points Coordinate Standardization | Multiple extractors | None | P2 |
| 14 | Multi-template Slide Merging | `builder.py` | None | P3 |
| 16 | HTML Table Import | `builder.py` | `beautifulsoup4` | P3 |
| 17 | ML Layout Classification | New `layout_classifier.py` | `layoutparser` | P2 |
| 18 | LayoutLMv3 Integration | New `semantic_analyzer.py` | `transformers`, `torch` | P3 |

---

## Implementation Order Recommendation

### Phase 1: LLM Foundation (Week 1-2)
1. **LiteLLM Integration** (#6) - Enables all AI features
2. **Multi-LLM Provider Support** (#1) - Production flexibility
3. **Tone/Verbosity Controls** (#2) - Immediate content customization

### Phase 2: Quality & Evaluation (Week 3-4)
4. **PPTEval Framework** (#3) - Enables quality gates
5. **Cascading Layout Handlers** (#7) - Smarter layout selection
6. **Round-trip Testing** (#12) - Pipeline verification

### Phase 3: Agent Integration (Week 5-6)
7. **MCP Server** (#19) - Enable Claude/agent access
8. **Functional Slide Auto-Insert** (#4) - Consistent structure
9. **Image Search Integration** (#8) - Automatic imagery

### Phase 4: Advanced Features (Week 7+)
10. **Markdown Input** (#20) - Alternative authoring
11. **ML Layout Classification** (#17) - Semantic understanding
12. **LayoutLMv3 Integration** (#18) - Deep document understanding

---

## Dependency Summary

```bash
# Phase 1 - LLM
pip install litellm anthropic openai

# Phase 2 - Quality
# No new dependencies

# Phase 3 - Agent
pip install mcp

# Phase 4 - ML
pip install layoutparser[detectron2] transformers torch beautifulsoup4
```

---

## Architecture Impact

### New Files to Create

```
pptx_design/
├── evaluation.py          # PPTEval framework
├── mcp_server.py          # MCP server for agents
├── schemas.py             # Standardized element interfaces

pptx_generator/modules/
├── llm_provider.py        # Multi-LLM support
├── image_search.py        # Pexels/Unsplash integration
├── markdown_parser.py     # Markdown → outline

pptx_extractor/
├── layout_classifier.py   # ML-based layout detection
├── semantic_analyzer.py   # LayoutLMv3 integration

tests/
├── test_roundtrip.py      # Extraction → generation fidelity
├── test_evaluation.py     # PPTEval tests
```

### Files to Modify

```
pptx_design/
├── registry.py            # Add template labels
├── builder.py             # Add HTML import, multi-template

pptx_generator/modules/
├── orchestrator.py        # Add functional slide auto-insert
├── template_renderer.py   # Add cascading handlers, length_factor, DSL
├── outline_generator.py   # Add tone/verbosity

pptx_extractor/
├── master_extractor.py    # Layout elements separation
```

---

## Sources

- [Presenton](https://github.com/presenton/presenton)
- [PPTAgent](https://github.com/icip-cas/PPTAgent)
- [SlideDeck AI](https://github.com/barun-saha/slide-deck-ai)
- [pptxtojson-english](https://github.com/alastairapple/pptxtojson-english)
- [airppt-parser](https://github.com/airpptx/airppt-parser)
- [pptx-compose](https://github.com/shobhitsharma/pptx-compose)
- [pptx-automizer](https://github.com/singerla/pptx-automizer)
- [pptx-template](https://github.com/m3dev/pptx-template)
- [PptxGenJS](https://github.com/gitbrent/PptxGenJS)
- [LayoutParser](https://github.com/Layout-Parser/layout-parser)
- [LayoutLMv3](https://github.com/microsoft/unilm/tree/master/layoutlmv3)
- [Office-PowerPoint-MCP-Server](https://github.com/GongRzhe/Office-PowerPoint-MCP-Server)
- [Slidev](https://github.com/slidevjs/slidev)
