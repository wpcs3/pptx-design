"""
PPTX Design System - Unified API

A Python system for creating professional PowerPoint presentations
using template-based generation with master layout support.

Example:
    from pptx_design import PresentationBuilder

    builder = PresentationBuilder("consulting_toolkit")
    builder.add_title_slide("Q4 Review", "Strategic Analysis 2025")
    builder.add_agenda(["Overview", "Analysis", "Recommendations"])
    builder.add_content_slide("Executive Summary", bullets=["Key finding 1", "Key finding 2"])
    builder.save("output.pptx")

Pipeline Example:
    from pptx_design import ContentPipeline, PresentationBuilder

    pipeline = ContentPipeline("consulting_toolkit")
    outline = pipeline.create_standard_outline(
        title="Quarterly Review",
        sections=["Overview", "Analysis", "Recommendations"]
    )
    slides = pipeline.prepare_for_rendering(outline)
    # Use slides with PresentationBuilder...

Evaluation Example (Phase 2):
    from pptx_design import evaluate_presentation, quick_score

    result = evaluate_presentation("presentation.pptx")
    print(result.summary())
    print(f"Grade: {result.grade}")

Markdown Input Example (Phase 4):
    from pptx_design import markdown_to_outline

    markdown = '''
    # Q4 Review
    Strategic Analysis

    ## Agenda
    - Overview
    - Analysis
    - Recommendations
    '''
    outline = markdown_to_outline(markdown)
    # Use outline with PresentationBuilder...

Layout Classification Example (Phase 4):
    from pptx_design import classify_slide, SlideLayoutClassifier

    # Quick classification
    result = classify_slide("slide_image.png")
    print(result.layout_type)    # "content_with_figure"
    print(result.confidence)      # 0.85
    print(result.region_summary)  # {"title": 1, "text": 2, "figure": 1}

    # Batch classification
    from pptx_design import classify_presentation_slides
    results = classify_presentation_slides("slides_dir/", pattern="*.png")

Semantic Analysis Example (Phase 4):
    from pptx_design import analyze_slide, extract_text_boxes_from_pptx

    # Extract text boxes from PPTX
    text_boxes = extract_text_boxes_from_pptx("presentation.pptx", slide_index=0)

    # Analyze semantics
    result = analyze_slide("slide_image.png", text_boxes)
    print(result.content_purpose)    # "data_presentation"
    print(result.suggested_template) # "market_analysis"
    print(result.title)              # "Q4 2025 Business Review"
"""

from .builder import PresentationBuilder, create_presentation
from .registry import TemplateRegistry, build_registry
from .pipeline import (
    ContentPipeline,
    ContentParser,
    LayoutMatcher,
    PresentationOutline,
    SlideContent,
    SlideType,
)
from .testing import VisualTester, TestResult, TestSuite, quick_compare
from .evaluation import (
    PresentationEvaluator,
    EvaluationResult,
    ContentScore,
    DesignScore,
    CoherenceScore,
    evaluate_presentation,
    quick_score,
)
from .roundtrip import (
    RoundtripTester,
    RoundtripResult,
    test_roundtrip,
    quick_fidelity_score,
)
from .agent_tools import (
    AgentInterface,
    ToolResult,
    get_tool_definitions,
    get_openai_tools,
    get_anthropic_tools,
    TOOL_DEFINITIONS,
)

# Phase 4 - Markdown Parser (import from pptx_generator)
try:
    from pptx_generator.modules.markdown_parser import (
        MarkdownParser,
        markdown_to_outline,
        parse_marp_file,
    )
    _MARKDOWN_AVAILABLE = True
except ImportError:
    _MARKDOWN_AVAILABLE = False
    MarkdownParser = None
    markdown_to_outline = None
    parse_marp_file = None

# Phase 4 - ML Layout Classifier (optional dependencies)
try:
    from pptx_extractor.layout_classifier import (
        SlideLayoutClassifier,
        LayoutClassification,
        DetectedRegion,
        classify_slide,
        classify_presentation_slides,
    )
    _LAYOUT_CLASSIFIER_AVAILABLE = True
except ImportError:
    _LAYOUT_CLASSIFIER_AVAILABLE = False
    SlideLayoutClassifier = None
    LayoutClassification = None
    DetectedRegion = None
    classify_slide = None
    classify_presentation_slides = None

# Phase 4 - LayoutLMv3 Semantic Analyzer (optional dependencies)
try:
    from pptx_extractor.semantic_analyzer import (
        SemanticSlideAnalyzer,
        SemanticAnalysisResult,
        SemanticLabel,
        TextBox,
        analyze_slide,
        extract_text_boxes_from_pptx,
    )
    _SEMANTIC_ANALYZER_AVAILABLE = True
except ImportError:
    _SEMANTIC_ANALYZER_AVAILABLE = False
    SemanticSlideAnalyzer = None
    SemanticAnalysisResult = None
    SemanticLabel = None
    TextBox = None
    analyze_slide = None
    extract_text_boxes_from_pptx = None

__version__ = "1.4.0"
__all__ = [
    # Builder
    "PresentationBuilder",
    "create_presentation",
    # Registry
    "TemplateRegistry",
    "build_registry",
    # Pipeline
    "ContentPipeline",
    "ContentParser",
    "LayoutMatcher",
    "PresentationOutline",
    "SlideContent",
    "SlideType",
    # Testing
    "VisualTester",
    "TestResult",
    "TestSuite",
    "quick_compare",
    # Evaluation (Phase 2)
    "PresentationEvaluator",
    "EvaluationResult",
    "ContentScore",
    "DesignScore",
    "CoherenceScore",
    "evaluate_presentation",
    "quick_score",
    # Round-trip (Phase 2)
    "RoundtripTester",
    "RoundtripResult",
    "test_roundtrip",
    "quick_fidelity_score",
    # Agent Tools (Phase 3)
    "AgentInterface",
    "ToolResult",
    "get_tool_definitions",
    "get_openai_tools",
    "get_anthropic_tools",
    "TOOL_DEFINITIONS",
    # Markdown Parser (Phase 4)
    "MarkdownParser",
    "markdown_to_outline",
    "parse_marp_file",
    # Layout Classifier (Phase 4)
    "SlideLayoutClassifier",
    "LayoutClassification",
    "DetectedRegion",
    "classify_slide",
    "classify_presentation_slides",
    # Semantic Analyzer (Phase 4)
    "SemanticSlideAnalyzer",
    "SemanticAnalysisResult",
    "SemanticLabel",
    "TextBox",
    "analyze_slide",
    "extract_text_boxes_from_pptx",
]
