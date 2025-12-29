"""
PowerPoint Presentation Generator CLI

Main entry point for the presentation generation system.
"""

import argparse
import asyncio
import json
import logging
import sys
from pathlib import Path

# Default paths
DEFAULT_CONFIG_DIR = Path(__file__).parent / "config"
DEFAULT_TEMPLATES_DIR = Path(__file__).parent.parent / "pptx_templates"
DEFAULT_OUTPUT_DIR = Path(__file__).parent / "output"


def setup_logging(verbose: bool = False):
    """Configure logging."""
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format="%(levelname)s: %(message)s"
    )


def cmd_analyze(args):
    """Run template analysis."""
    from .modules.template_analyzer import TemplateAnalyzer

    print(f"Analyzing templates in: {args.templates_dir}")
    analyzer = TemplateAnalyzer(args.templates_dir)
    style_path, catalog_path = analyzer.save_results(args.output_dir)

    print(f"\nStyle guide saved to: {style_path}")
    print(f"Slide catalog saved to: {catalog_path}")


def cmd_outline(args):
    """Generate presentation outline."""
    from .modules.outline_generator import OutlineGenerator

    # Load configs
    config_dir = Path(args.config_dir)

    with open(config_dir / "content_patterns.json", "r") as f:
        content_patterns = json.load(f)
    with open(config_dir / "slide_catalog.json", "r") as f:
        slide_catalog = json.load(f)

    generator = OutlineGenerator(content_patterns, slide_catalog)
    outline = generator.generate_outline(args.request)

    # Print outline
    print("\n" + "=" * 60)
    print(generator.outline_to_text(outline))
    print("=" * 60)

    # Save if output specified
    if args.output:
        generator.save_outline(outline, args.output)
        print(f"\nOutline saved to: {args.output}")


def cmd_build(args):
    """Build presentation from outline."""
    from .modules.orchestrator import PresentationOrchestrator

    async def run_build():
        orchestrator = PresentationOrchestrator(
            args.config_dir,
            args.templates_dir,
            args.output_dir
        )

        # Load outline
        with open(args.outline, "r") as f:
            outline = json.load(f)

        print("Assembling content...")
        enriched = await orchestrator.assemble_content(outline)

        print("Generating presentation...")
        prs = orchestrator.generate_pptx(enriched)

        print("Exporting...")
        output_path = orchestrator.export_pptx(prs, args.output)

        print(f"\nPresentation saved to: {output_path}")

    asyncio.run(run_build())


def cmd_generate(args):
    """Generate presentation from request (full workflow)."""
    from .modules.orchestrator import PresentationOrchestrator

    async def run_generate():
        orchestrator = PresentationOrchestrator(
            args.config_dir,
            args.templates_dir,
            args.output_dir
        )

        print(f"\nRequest: {args.request}")
        print("\n" + "=" * 60)
        print("Generating outline...")

        workflow = await orchestrator.create_presentation(args.request)

        print("\n" + workflow.get_outline_preview())
        print("=" * 60)

        if not args.auto_approve:
            response = input("\nApprove outline? (y/n/modify): ").strip().lower()
            if response == "n":
                print("Cancelled.")
                return
            elif response == "modify" or response == "m":
                feedback = input("Modification: ").strip()
                workflow.modify_outline(feedback)
                print("\n" + workflow.get_outline_preview())
                response = input("\nApprove modified outline? (y/n): ").strip().lower()
                if response != "y":
                    print("Cancelled.")
                    return

        workflow.approve_outline()
        print("\nAssembling content...")
        await workflow.assemble_content()

        print("Generating presentation...")
        workflow.generate_presentation()

        output_path = workflow.finalize(args.output)
        print(f"\nPresentation saved to: {output_path}")

    asyncio.run(run_generate())


def cmd_refine(args):
    """Refine an existing presentation."""
    from pptx import Presentation
    from .modules.orchestrator import PresentationOrchestrator

    orchestrator = PresentationOrchestrator(
        args.config_dir,
        args.templates_dir,
        args.output_dir
    )

    print(f"Loading: {args.presentation}")
    prs = Presentation(args.presentation)

    print(f"Applying feedback: {args.feedback}")
    prs = orchestrator.refine_presentation(prs, args.feedback, args.slide)

    output_path = args.output or args.presentation.replace(".pptx", "_refined.pptx")
    prs.save(output_path)
    print(f"Saved to: {output_path}")


def cmd_list_types(args):
    """List available slide types."""
    config_dir = Path(args.config_dir)
    catalog_path = config_dir / "slide_catalog.json"

    if not catalog_path.exists():
        print("Slide catalog not found. Run 'analyze' first.")
        return

    with open(catalog_path, "r") as f:
        catalog = json.load(f)

    print("\nAvailable Slide Types:")
    print("-" * 60)

    for slide_type in catalog.get("slide_types", [])[:20]:
        print(f"  {slide_type['id']}: {slide_type['name']}")
        print(f"      Layout: {slide_type['master_layout']}")
        print(f"      Occurrences: {slide_type.get('occurrence_count', 0)}")
        print()


def cmd_list_patterns(args):
    """List available presentation patterns."""
    config_dir = Path(args.config_dir)
    patterns_path = config_dir / "content_patterns.json"

    if not patterns_path.exists():
        print("Content patterns not found.")
        return

    with open(patterns_path, "r") as f:
        patterns = json.load(f)

    print("\nPresentation Types:")
    print("-" * 60)

    for ptype, config in patterns.get("presentation_types", {}).items():
        print(f"\n  {ptype}:")
        print(f"      {config.get('description', '')}")
        print(f"      Sections: {len(config.get('typical_sections', []))}")
        slide_range = config.get('typical_slide_count', {})
        print(f"      Slides: {slide_range.get('min', 0)}-{slide_range.get('max', 0)}")


def cmd_llm_status(args):
    """Show LLM provider status and available models."""
    from .modules.llm_provider import LLMManager

    manager = LLMManager()
    status = manager.get_status()

    print("\n" + "=" * 60)
    print("LLM PROVIDER STATUS")
    print("=" * 60)

    print(f"\nAPI Key Configuration (from .env file):")
    print(f"  Anthropic (Claude): {'[OK] Configured' if status['anthropic_configured'] else '[X] Not configured'}")
    print(f"  OpenAI (GPT):       {'[OK] Configured' if status['openai_configured'] else '[X] Not configured'}")
    print(f"  Google (Gemini):    {'[OK] Configured' if status['google_configured'] else '[X] Not configured'}")

    print(f"\nAvailable Models:")
    print("-" * 60)
    for model in status['models']:
        status_icon = "OK" if model['available'] else "X"
        current = " <- current" if model['current'] else ""
        print(f"  [{status_icon}] {model['name']}: {model['display_name']}{current}")
        print(f"       Provider: {model['provider']}")

    if not status['anthropic_configured'] and not status['openai_configured'] and not status['google_configured']:
        print("\n[!] No API keys configured!")
        print("    Add your API keys to the .env file in the project folder.")
        print("    Run: python -m pptx_generator llm-setup for instructions.")


def cmd_llm_setup(args):
    """Show instructions for setting up LLM API keys."""
    print("""
================================================================
           LLM API KEY SETUP INSTRUCTIONS
================================================================

API keys are stored in the .env file in the project folder:
  C:\\Users\\wpcol\\claudecode\\pptx-design\\.env

STEP 1: Get Your API Keys
-------------------------

ANTHROPIC (Claude 4.5):
  1. Go to: https://console.anthropic.com/
  2. Sign up or log in
  3. Click "API Keys" in the left sidebar
  4. Click "Create Key"
  5. Copy the key (starts with 'sk-ant-')

OPENAI (GPT-5):
  1. Go to: https://platform.openai.com/
  2. Sign up or log in
  3. Click your profile -> "View API keys"
  4. Click "Create new secret key"
  5. Copy the key (starts with 'sk-')

GOOGLE (Gemini 3):
  1. Go to: https://aistudio.google.com/apikey
  2. Sign up or log in with Google account
  3. Click "Create API Key"
  4. Copy the key

STEP 2: Add Keys to .env File
-----------------------------
Open the .env file and paste your keys:

  ANTHROPIC_API_KEY=sk-ant-your-key-here
  OPENAI_API_KEY=sk-your-key-here
  GOOGLE_API_KEY=your-google-key-here

STEP 3: Verify Setup
--------------------
  python -m pptx_generator llm-status

STEP 4: Test Generation
-----------------------
  python -m pptx_generator llm-test --model claude-sonnet-4.5
  python -m pptx_generator llm-test --model gpt-5
  python -m pptx_generator llm-test --model gemini-3-pro

AVAILABLE MODELS:
-----------------
Anthropic:  claude-opus-4.5, claude-sonnet-4.5, claude-haiku-4.5
OpenAI:     gpt-5.2, gpt-5.1, gpt-5, gpt-5-mini, gpt-5-nano
Google:     gemini-3-pro, gemini-3-flash

""")


def cmd_llm_test(args):
    """Test LLM generation with a simple prompt."""
    from .modules.llm_provider import LLMManager

    manager = LLMManager(args.model)

    if not manager.is_available():
        print(f"\n[X] Model '{args.model}' is not available.")
        print("    Check that the API key is set correctly.")
        print("    Run: python -m pptx_generator llm-setup")
        return

    print(f"\nTesting {manager.current_model.display_name}...")
    print("-" * 60)

    test_prompt = args.prompt or "Write a one-sentence summary of why industrial real estate is a good investment in 2025."

    print(f"Prompt: {test_prompt}\n")

    try:
        response = manager.generate(test_prompt, temperature=0.7)
        print(f"Response:\n{response.content}")
        print(f"\nTokens: {response.usage['input_tokens']} input, {response.usage['output_tokens']} output")
        print(f"Model: {response.model}")
    except Exception as e:
        print(f"[X] Error: {e}")


def cmd_generate_with_llm(args):
    """Generate presentation with LLM-powered content."""
    from .modules.content_generator import ContentGenerator
    import asyncio

    generator = ContentGenerator(model=args.model)

    if not generator.llm.is_available():
        print(f"\n[X] Model '{args.model}' is not available.")
        print("    Run: python -m pptx_generator llm-setup")
        return

    print(f"\nUsing model: {generator.llm.current_model.display_name}")
    print(f"Request: {args.request}")
    print("-" * 60)

    # Generate outline
    print("\nGenerating outline...")
    outline_result = generator.generate_outline(
        args.request,
        presentation_type=args.type
    )

    if "parse_error" in outline_result.content:
        print(f"Warning: Could not parse outline JSON")
        print(outline_result.raw_response)
        return

    outline = outline_result.content
    print(f"Outline generated: {outline.get('title', 'Untitled')}")
    print(f"Sections: {len(outline.get('sections', []))}")
    print(f"Estimated slides: {outline.get('estimated_slide_count', 'unknown')}")

    # Save outline
    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    outline_path = output_dir / f"{args.output or 'generated'}_outline.json"
    with open(outline_path, 'w') as f:
        json.dump(outline, f, indent=2)
    print(f"\nOutline saved to: {outline_path}")

    # Optionally enrich with slide content
    if args.enrich:
        print("\nEnriching slides with content...")
        enriched = generator.enrich_outline(outline)

        enriched_path = output_dir / f"{args.output or 'generated'}_enriched.json"
        with open(enriched_path, 'w') as f:
            json.dump(enriched, f, indent=2)
        print(f"Enriched outline saved to: {enriched_path}")

        stats = enriched.get('generation_stats', {})
        print(f"\nGeneration stats:")
        print(f"  Model: {stats.get('model', 'unknown')}")
        print(f"  Total tokens: {stats.get('total_input_tokens', 0)} input, {stats.get('total_output_tokens', 0)} output")


def cmd_classify(args):
    """Run content classification on library."""
    from .modules.content_classifier import ContentClassifier
    import json

    print("Running content classification...")

    library_path = Path(args.library_dir)
    index_path = library_path / "library_index.json"

    if not index_path.exists():
        print(f"Library index not found at: {index_path}")
        return

    with open(index_path, 'r', encoding='utf-8') as f:
        library_index = json.load(f)

    classifier = ContentClassifier(library_path)
    stats = classifier.classify_all(library_index)

    print("\nClassification Complete!")
    print()
    print("IMAGE CATEGORIES:")
    for cat, count in sorted(stats['images'].items(), key=lambda x: -x[1]):
        print(f"  {cat}: {count}")

    print()
    print("ICON CATEGORIES:")
    for cat, count in sorted(stats['icons'].items(), key=lambda x: -x[1]):
        print(f"  {cat}: {count}")

    print()
    print("DIAGRAM TYPES:")
    for dtype, count in sorted(stats['diagrams'].items(), key=lambda x: -x[1]):
        print(f"  {dtype}: {count}")


def cmd_browse_library(args):
    """Browse classified library content."""
    from .modules.component_library import ComponentLibrary
    from .modules.library_enhancer import LibraryEnhancer

    library = ComponentLibrary()
    enhancer = LibraryEnhancer(library)

    print("\n" + "=" * 60)
    print("LIBRARY BROWSER")
    print("=" * 60)

    # Show image categories
    print("\nIMAGE CATEGORIES:")
    print("-" * 40)
    categories = enhancer.get_image_categories()
    for cat, count in sorted(categories.items(), key=lambda x: -x[1]):
        print(f"  {cat}: {count}")

    # Show domain stats
    print("\nDOMAIN TAGS:")
    print("-" * 40)
    domain_stats = enhancer.get_domain_stats()
    for domain, count in sorted(domain_stats.items(), key=lambda x: -x[1]):
        print(f"  {domain}: {count}")

    # Show purpose stats
    print("\nCOMPONENT PURPOSES:")
    print("-" * 40)
    purpose_stats = enhancer.get_purpose_stats()
    for comp_type, purposes in purpose_stats.items():
        print(f"\n  {comp_type}:")
        for purpose, count in sorted(purposes.items(), key=lambda x: -x[1])[:5]:
            print(f"    {purpose}: {count}")

    # Show quick finds
    print("\n" + "=" * 60)
    print("QUICK FIND EXAMPLES")
    print("=" * 60)

    logo = enhancer.find_logo()
    print(f"\nLogo: {logo}")

    bg = enhancer.find_background_image()
    print(f"Background: {bg}")

    icons = enhancer.find_icons(4)
    print(f"Icons (4): {icons}")

    diagram = enhancer.find_diagram_template(4)
    if diagram:
        print(f"Diagram template: {diagram['name']} ({diagram['type']})")


def cmd_find_images(args):
    """Find images by category."""
    from .modules.content_classifier import ContentClassifier

    library_path = Path(args.library_dir)
    classifier = ContentClassifier(library_path)

    category = args.category
    images = classifier.get_images_by_category(category)

    print(f"\n{category.upper()} IMAGES ({len(images)} found):")
    print("-" * 60)

    for img in images[:args.limit]:
        print(f"\n  ID: {img.id}")
        print(f"  File: {img.filename}")
        print(f"  Size: {img.width_inches:.2f}\" x {img.height_inches:.2f}\"")
        print(f"  Aspect: {img.aspect_ratio}")
        print(f"  Tags: {', '.join(img.tags)}")
        print(f"  Use cases: {', '.join(img.use_cases)}")


def cmd_render_test(args):
    """Create a test presentation with all slide types."""
    from .modules.slide_renderer import SlideRenderer
    import json

    config_dir = Path(args.config_dir)

    with open(config_dir / "style_guide.json", "r") as f:
        style_guide = json.load(f)
    with open(config_dir / "slide_catalog.json", "r") as f:
        slide_catalog = json.load(f)

    renderer = SlideRenderer(style_guide, slide_catalog)
    prs = renderer.create_presentation()

    # Create sample slides
    test_slides = [
        ("title_slide", {
            "title": "Sample Presentation",
            "subtitle": "Generated by PPTX Generator"
        }),
        ("section_divider", {
            "title": "Section One"
        }),
        ("title_content", {
            "title": "Key Points",
            "bullets": ["First point", "Second point", "Third point"]
        }),
        ("two_column", {
            "title": "Comparison",
            "left_column": {
                "header": "Option A",
                "bullets": ["Pro 1", "Pro 2"]
            },
            "right_column": {
                "header": "Option B",
                "bullets": ["Pro 1", "Pro 2"]
            }
        }),
        ("key_metrics", {
            "title": "Key Metrics",
            "metrics": [
                {"label": "Revenue", "value": "$1M"},
                {"label": "Growth", "value": "25%"},
                {"label": "Users", "value": "10K"}
            ]
        }),
        ("table_slide", {
            "title": "Data Table",
            "headers": ["Name", "Value", "Status"],
            "data": [
                ["Item 1", "100", "Active"],
                ["Item 2", "200", "Pending"],
                ["Item 3", "300", "Complete"]
            ]
        })
    ]

    for slide_type, content in test_slides:
        renderer.create_slide(prs, slide_type, content)

    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / "test_render.pptx"
    prs.save(str(output_path))

    print(f"Test presentation saved to: {output_path}")


def main():
    """Main CLI entry point."""
    parser = argparse.ArgumentParser(
        description="PowerPoint Presentation Generator",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Analyze templates and create config files
  python -m pptx_generator analyze

  # Generate outline from request
  python -m pptx_generator outline --request "Create a pitch for our logistics fund"

  # Build presentation from outline
  python -m pptx_generator build --outline outline.json

  # Full generation workflow
  python -m pptx_generator generate --request "Create investor pitch deck"

  # List available slide types
  python -m pptx_generator list-types

  # Create test presentation
  python -m pptx_generator test-render
"""
    )

    parser.add_argument(
        "-v", "--verbose",
        action="store_true",
        help="Enable verbose logging"
    )
    parser.add_argument(
        "--config-dir",
        default=str(DEFAULT_CONFIG_DIR),
        help="Path to configuration directory"
    )
    parser.add_argument(
        "--templates-dir",
        default=str(DEFAULT_TEMPLATES_DIR),
        help="Path to PPTX templates directory"
    )
    parser.add_argument(
        "--output-dir",
        default=str(DEFAULT_OUTPUT_DIR),
        help="Path to output directory"
    )

    subparsers = parser.add_subparsers(dest="command", help="Available commands")

    # Analyze command
    analyze_parser = subparsers.add_parser(
        "analyze",
        help="Analyze templates and create configuration files"
    )
    analyze_parser.set_defaults(func=cmd_analyze)

    # Outline command
    outline_parser = subparsers.add_parser(
        "outline",
        help="Generate presentation outline from request"
    )
    outline_parser.add_argument(
        "--request", "-r",
        required=True,
        help="Description of the presentation needed"
    )
    outline_parser.add_argument(
        "--output", "-o",
        help="Output file for outline JSON"
    )
    outline_parser.set_defaults(func=cmd_outline)

    # Build command
    build_parser = subparsers.add_parser(
        "build",
        help="Build presentation from approved outline"
    )
    build_parser.add_argument(
        "--outline",
        required=True,
        help="Path to outline JSON file"
    )
    build_parser.add_argument(
        "--output", "-o",
        help="Output filename"
    )
    build_parser.set_defaults(func=cmd_build)

    # Generate command (full workflow)
    generate_parser = subparsers.add_parser(
        "generate",
        help="Full generation workflow from request to PPTX"
    )
    generate_parser.add_argument(
        "--request", "-r",
        required=True,
        help="Description of the presentation needed"
    )
    generate_parser.add_argument(
        "--output", "-o",
        help="Output filename"
    )
    generate_parser.add_argument(
        "--auto-approve",
        action="store_true",
        help="Auto-approve outline without prompting"
    )
    generate_parser.set_defaults(func=cmd_generate)

    # Refine command
    refine_parser = subparsers.add_parser(
        "refine",
        help="Refine an existing presentation"
    )
    refine_parser.add_argument(
        "--presentation", "-p",
        required=True,
        help="Path to PPTX file to refine"
    )
    refine_parser.add_argument(
        "--feedback", "-f",
        required=True,
        help="Refinement feedback"
    )
    refine_parser.add_argument(
        "--slide", "-s",
        type=int,
        help="Specific slide number to modify"
    )
    refine_parser.add_argument(
        "--output", "-o",
        help="Output filename"
    )
    refine_parser.set_defaults(func=cmd_refine)

    # List types command
    list_types_parser = subparsers.add_parser(
        "list-types",
        help="List available slide types"
    )
    list_types_parser.set_defaults(func=cmd_list_types)

    # List patterns command
    list_patterns_parser = subparsers.add_parser(
        "list-patterns",
        help="List available presentation patterns"
    )
    list_patterns_parser.set_defaults(func=cmd_list_patterns)

    # Test render command
    test_render_parser = subparsers.add_parser(
        "test-render",
        help="Create a test presentation with sample slides"
    )
    test_render_parser.set_defaults(func=cmd_render_test)

    # Classify command
    classify_parser = subparsers.add_parser(
        "classify",
        help="Run content classification on library"
    )
    classify_parser.add_argument(
        "--library-dir",
        default="pptx_component_library",
        help="Path to component library directory"
    )
    classify_parser.set_defaults(func=cmd_classify)

    # Browse library command
    browse_parser = subparsers.add_parser(
        "browse-library",
        help="Browse classified library content"
    )
    browse_parser.set_defaults(func=cmd_browse_library)

    # Find images command
    find_images_parser = subparsers.add_parser(
        "find-images",
        help="Find images by category"
    )
    find_images_parser.add_argument(
        "--category", "-c",
        required=True,
        choices=["logo", "icon", "photo", "screenshot", "background", "chart_image", "decorative", "unknown"],
        help="Image category to search"
    )
    find_images_parser.add_argument(
        "--limit", "-l",
        type=int,
        default=10,
        help="Maximum results to show"
    )
    find_images_parser.add_argument(
        "--library-dir",
        default="pptx_component_library",
        help="Path to component library directory"
    )
    find_images_parser.set_defaults(func=cmd_find_images)

    # LLM status command
    llm_status_parser = subparsers.add_parser(
        "llm-status",
        help="Show LLM provider status and available models"
    )
    llm_status_parser.set_defaults(func=cmd_llm_status)

    # LLM setup command
    llm_setup_parser = subparsers.add_parser(
        "llm-setup",
        help="Show instructions for setting up LLM API keys"
    )
    llm_setup_parser.set_defaults(func=cmd_llm_setup)

    # LLM test command
    llm_test_parser = subparsers.add_parser(
        "llm-test",
        help="Test LLM generation"
    )
    llm_test_parser.add_argument(
        "--model", "-m",
        default="claude-sonnet-4.5",
        choices=[
            "claude-opus-4.5", "claude-sonnet-4.5", "claude-haiku-4.5",
            "gpt-5.2", "gpt-5.1", "gpt-5", "gpt-5-mini", "gpt-5-nano",
            "gemini-3-pro", "gemini-3-flash"
        ],
        help="Model to test"
    )
    llm_test_parser.add_argument(
        "--prompt", "-p",
        help="Custom test prompt"
    )
    llm_test_parser.set_defaults(func=cmd_llm_test)

    # Generate with LLM command
    llm_generate_parser = subparsers.add_parser(
        "llm-generate",
        help="Generate presentation outline using LLM"
    )
    llm_generate_parser.add_argument(
        "--request", "-r",
        required=True,
        help="Description of the presentation needed"
    )
    llm_generate_parser.add_argument(
        "--model", "-m",
        default="claude-sonnet-4.5",
        choices=[
            "claude-opus-4.5", "claude-sonnet-4.5", "claude-haiku-4.5",
            "gpt-5.2", "gpt-5.1", "gpt-5", "gpt-5-mini", "gpt-5-nano",
            "gemini-3-pro", "gemini-3-flash"
        ],
        help="Model to use for generation"
    )
    llm_generate_parser.add_argument(
        "--type", "-t",
        choices=["investment_pitch", "market_analysis", "due_diligence", "business_case", "consulting_framework"],
        help="Presentation type"
    )
    llm_generate_parser.add_argument(
        "--output", "-o",
        help="Output filename prefix"
    )
    llm_generate_parser.add_argument(
        "--enrich",
        action="store_true",
        help="Also generate content for each slide"
    )
    llm_generate_parser.set_defaults(func=cmd_generate_with_llm)

    args = parser.parse_args()

    if not args.command:
        parser.print_help()
        sys.exit(1)

    setup_logging(args.verbose)
    args.func(args)


if __name__ == "__main__":
    main()
