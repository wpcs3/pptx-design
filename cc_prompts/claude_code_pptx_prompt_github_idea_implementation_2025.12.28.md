# Claude Code: Systematic GitHub Repository Review for pptx_design Improvements

## Objective

Systematically review each GitHub repository from the "PowerPoint Extraction and Generation: The Open-Source Landscape" research report. For each repository, identify specific techniques, architectures, and features that could improve our `pptx_design` project. Produce actionable recommendations with code-level implementation guidance.

---

## Our Project Context

Our `pptx_design` project is a Python system for:
- **Template Analysis**: Analyzing PowerPoint templates and recreating them programmatically
- **Presentation Generation**: Generating investor pitch decks from structured JSON outlines
- **Component Library**: 22,792 extracted components (charts, tables, images, shapes, layouts, styles)
- **Unified API**: `PresentationBuilder` with fluent interface using actual master layouts

### Current Architecture
```
pptx_design/          → Unified API (builder, registry, pipeline, testing)
pptx_generator/       → Orchestrator, template_renderer, component_library, library_enhancer
pptx_extractor/       → template_generator, master_extractor, unified_extractor
pptx_component_library/ → Extracted components index (master_index.json)
```

### Current Capabilities
- Master layout-based rendering (100% position accuracy via placeholders)
- ComponentLibrary with domain tagging (479 tagged components)
- Smart chart/table matching by structure (series × categories)
- SSIM-based visual testing
- LibreOffice headless PPTX→PNG rendering

### Known Gaps (prioritize improvements here)
1. No AI/LLM integration for content generation
2. No reference presentation style learning
3. Limited semantic understanding of slide designs
4. No iterative refinement or quality evaluation framework
5. No MCP server integration for agent access

---

## Repositories to Review

### TIER 1: HIGH-PRIORITY (AI-Powered Generation)

#### 1. Presenton (github.com/presenton/presenton)
**Review focus**: Template upload → style extraction → brand-consistent generation

Investigate:
- [ ] How does Presenton extract design elements from uploaded PPTX files?
- [ ] What JSON schema do they use for template representation?
- [ ] How do they apply extracted styles to generated content?
- [ ] What's their architecture for multi-LLM support (OpenAI, Claude, Gemini, Ollama)?
- [ ] How do they implement tone controls (professional, casual, sales_pitch)?
- [ ] What's their REST API design for presentation generation?

Map to our project:
- Compare to our `pptx_extractor/unified_extractor.py`
- Compare to our `pptx_generator/modules/template_renderer.py`
- Identify features missing from our `PresentationBuilder`

#### 2. PPTAgent (github.com/icip-cas/PPTAgent)
**Review focus**: Two-phase architecture (analysis → generation) + PPTEval framework

Investigate:
- [ ] How does the analysis phase extract "functional slide types" and content schemas?
- [ ] How do they select reference slides as templates for new content?
- [ ] What's the PPTEval framework for evaluating content accuracy, visual design, and logical coherence?
- [ ] How do they implement iterative editing actions?
- [ ] What's their slide type classification system?

Map to our project:
- Compare to our `config/slide_catalog.json` (279 cataloged slide types)
- Compare to our `pptx_generator/modules/library_enhancer.py` (domain tagging)
- Identify evaluation metrics we should implement in `pptx_design/testing.py`

#### 3. SlideDeck AI (github.com/barun-saha/slide-deck-ai)
**Review focus**: LLM → JSON schema → python-pptx assembly pipeline

Investigate:
- [ ] What JSON schema do they use for slide structure?
- [ ] How do they integrate LiteLLM for multi-provider support?
- [ ] What's their keyword extraction → image search pipeline?
- [ ] How do they handle content-to-layout mapping?

Map to our project:
- Compare to our outline JSON schema in `pptx_generator/config/content_patterns.json`
- Identify improvements for our `pptx_generator/modules/outline_generator.py`

---

### TIER 2: EXTRACTION & PARSING

#### 4. pptxtojson (github.com/pptist/pptxtojson or github.com/alastairapple/pptxtojson-english)
**Review focus**: Comprehensive JSON extraction including master slide elements

Investigate:
- [ ] What properties do they extract that we don't? (Check their TypeScript definitions)
- [ ] How do they handle `layoutElements` from master slides?
- [ ] What's their coordinate system and measurement units?
- [ ] How do they extract text styling (fonts, colors, alignment)?
- [ ] How do they handle gradients, shadows, 3D effects?

Map to our project:
- Compare to our `pptx_extractor/` extractors
- Compare JSON output to our `pptx_component_library/master_index.json`
- Identify missing properties in our extraction

#### 5. airppt-parser (github.com/airpptx/airppt-parser)
**Review focus**: Standardized PowerPointElement interface

Investigate:
- [ ] What's their `PowerPointElement` interface schema?
- [ ] How do they normalize `elementPosition` and `elementOffsetPosition`?
- [ ] What font attributes do they capture (name, size, fillColor)?
- [ ] How do they handle shape properties (border, fill, opacity)?
- [ ] What's their paragraph alignment handling?

Map to our project:
- Compare to our component schemas in `pptx_component_library/`
- Could we adopt their interface for better downstream compatibility?

#### 6. pptx-compose (github.com/shobhitsharma/pptx-compose)
**Review focus**: Bidirectional PPTX↔JSON conversion

Investigate:
- [ ] How do they achieve round-trip fidelity?
- [ ] What OpenXML structures do they preserve that we might lose?
- [ ] How do they handle complex elements (SmartArt, embedded objects)?

Map to our project:
- Test round-trip conversion with our templates
- Identify fidelity gaps in our extraction → generation pipeline

---

### TIER 3: TEMPLATE MANAGEMENT & AUTOMATION

#### 7. pptx-automizer (github.com/singerla/pptx-automizer)
**Review focus**: Template library management with slide master preservation

Investigate:
- [ ] How do they load multiple templates by label?
- [ ] What's their slide master merging strategy?
- [ ] How do they use xmldom callbacks for customization?
- [ ] How do they preserve template styling while replacing content?

Map to our project:
- Compare to our `pptx_design/registry.py` (TemplateRegistry)
- Identify improvements for managing our 4 templates with 117 layouts

#### 8. pptx-template (github.com/m3dev/pptx-template)
**Review focus**: DSL placeholders + JSON data model

Investigate:
- [ ] What's their placeholder DSL syntax (e.g., `{sales.0.june.us}`)?
- [ ] How do they map JSON paths to slide elements?
- [ ] How do they preserve all non-placeholder formatting?

Map to our project:
- Compare to our placeholder-based rendering in `template_renderer.py`
- Could we adopt their DSL for more flexible content injection?

#### 9. PptxGenJS (github.com/gitbrent/PptxGenJS)
**Review focus**: Programmatic Slide Master creation

Investigate:
- [ ] How do they define custom Slide Masters programmatically?
- [ ] What's their API for setting master-level styles?
- [ ] How do they handle complex chart/table creation?

Map to our project:
- While we use python-pptx, their API design may inform improvements
- Check if any features are missing from python-pptx

---

### TIER 4: VISION-BASED ANALYSIS

#### 10. LayoutParser (github.com/Layout-Parser/layout-parser)
**Review focus**: ML-based document layout detection

Investigate:
- [ ] What models do they use (Detectron2, EfficientDet)?
- [ ] How could we apply this to slide layout classification?
- [ ] What's their API for detecting text blocks, images, tables?

Map to our project:
- Could enhance our `pptx_extractor/` with visual layout detection
- Useful for analyzing slides where XML structure is ambiguous

#### 11. DiT / LayoutLMv3 (Microsoft's document understanding models)
**Review focus**: Combined text + layout + image understanding

Investigate:
- [ ] How do they represent layout information?
- [ ] What tasks can they perform (classification, extraction, generation)?
- [ ] What's the inference cost and latency?

Map to our project:
- Potential enhancement for `content_classifier.py`
- Could improve semantic understanding of slide content

---

### TIER 5: EMERGING APPROACHES

#### 12. Office-PowerPoint-MCP-Server (github.com/GongRzhe/Office-PowerPoint-MCP-Server)
**Review focus**: MCP protocol for AI agent access to PowerPoint

Investigate:
- [ ] What tools do they expose via MCP?
- [ ] How do they handle presentation state management?
- [ ] What's their error handling and validation approach?

Map to our project:
- Could we expose our `PresentationBuilder` as an MCP server?
- This would enable direct Claude/agent integration

#### 13. Slidev (github.com/slidevjs/slidev)
**Review focus**: Slides-as-code with PPTX export

Investigate:
- [ ] How do they convert Markdown → slides?
- [ ] What's their theming system?
- [ ] How does their PPTX export work?

Map to our project:
- Could we support Markdown input in addition to JSON outlines?
- Their theming approach may inform our style system

---

## Output Format

For each repository, produce a structured analysis:

```markdown
## [Repository Name]

### Key Findings
- [Technique/feature 1]: Brief description
- [Technique/feature 2]: Brief description

### Recommended Improvements for pptx_design

#### Improvement 1: [Name]
- **What**: Description of the improvement
- **Why**: How it addresses our gaps
- **Where**: Which module(s) to modify
- **How**: Implementation approach (code snippets if helpful)
- **Effort**: Low/Medium/High
- **Priority**: P0/P1/P2

#### Improvement 2: [Name]
...

### Code Patterns to Adopt
```python
# Example code from the repository that we should emulate
```

### Integration Points
- How this connects to our existing architecture
```

---

## Final Deliverable

After reviewing all repositories, produce a consolidated **IMPROVEMENT_ROADMAP.md** with:

1. **Quick Wins** (Low effort, High impact)
2. **Strategic Enhancements** (Medium effort, High impact)
3. **Future Capabilities** (High effort, Transformational)

For each item, include:
- Specific files to modify
- Dependencies to add
- Estimated implementation time
- Code architecture recommendations

---

## Execution Instructions

1. Clone/fetch each repository (or browse via GitHub API)
2. Focus on: README, core source files, JSON schemas, API definitions
3. Skip: tests, CI configs, documentation boilerplate
4. Prioritize: Novel techniques not present in our codebase
5. Be specific: Reference exact file paths and function names

Start with TIER 1 repositories (Presenton, PPTAgent, SlideDeck AI) as they're most relevant to our AI-powered generation goals.
