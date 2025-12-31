"""Generate Light Industrial presentation v18 with takeaways."""

import json
from pathlib import Path


def main():
    """Generate the presentation from the v18 outline."""
    print("Generating Light Industrial presentation v18 with takeaways...")
    print("=" * 60)

    # Load the v18 outline (already has images and takeaways)
    outline_path = Path("pptx_generator/output/light_industrial_thesis_v18.json")

    if not outline_path.exists():
        print(f"Outline not found: {outline_path}")
        return

    with open(outline_path) as f:
        outline = json.load(f)

    print(f"\nLoaded outline: {outline.get('title')}")
    print(f"Sections: {len(outline.get('sections', []))}")

    # Count slides with takeaways
    takeaway_count = 0
    for section in outline.get("sections", []):
        for slide in section.get("slides", []):
            if slide.get("content", {}).get("takeaway"):
                takeaway_count += 1
    print(f"Slides with takeaways: {takeaway_count}")

    # Import and run the orchestrator
    from pptx_generator.modules.orchestrator import PresentationOrchestrator, GenerationOptions

    # Set up paths
    config_dir = Path("pptx_generator/config")
    templates_dir = Path("pptx_templates")
    output_dir = Path("pptx_generator/output")

    # Create orchestrator with options
    options = GenerationOptions(
        auto_layout=True,
        auto_section_headers=False,  # We already have section dividers
        evaluate_after=True,
        use_slide_pool=False  # Disable to use simpler rendering
    )

    orchestrator = PresentationOrchestrator(
        config_dir=str(config_dir),
        templates_dir=str(templates_dir),
        output_dir=str(output_dir),
        options=options
    )

    # Generate the presentation
    output_path = output_dir / "Light_Industrial_Thesis_v21.pptx"

    print("\nGenerating presentation...")
    result = orchestrator.generate_pptx_with_evaluation(
        outline=outline,
        context={"request": "Light Industrial Investment Thesis v20 with Bullet Formatting"}
    )

    # Save the presentation
    result.presentation.save(str(output_path))

    print(f"\n{'=' * 60}")
    print(f"Generated presentation: {output_path}")
    print(f"Slide count: {len(result.presentation.slides)}")

    if result.evaluation:
        print(f"Quality Grade: {result.evaluation.grade}")
        print(f"Overall Score: {result.evaluation.overall_score:.1f}")

    # Convert to PDF
    print("\nConverting to PDF...")
    try:
        import subprocess
        pdf_path = output_path.with_suffix('.pdf')

        # Use LibreOffice for conversion
        cmd = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            "--headless",
            "--convert-to", "pdf",
            "--outdir", str(output_dir),
            str(output_path)
        ]
        subprocess.run(cmd, check=True, capture_output=True)
        print(f"PDF saved: {pdf_path}")
    except Exception as e:
        print(f"PDF conversion failed: {e}")

    return output_path


if __name__ == "__main__":
    main()
