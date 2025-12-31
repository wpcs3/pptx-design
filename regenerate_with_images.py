"""Regenerate Light Industrial presentation with Gemini-generated images."""

import json
from pathlib import Path

# Image mapping: section name -> image file
SECTION_IMAGE_MAP = {
    "Executive Summary": "title_slide.png",  # For title slide
    "Market Fundamentals": "market_fundamentals.png",
    "Target Markets": "target_markets.png",
    "Demand Drivers": "demand_drivers.png",
    "Investment Strategy": "investment_strategy.png",
    "Competitive Positioning": "competitive_positioning.png",
    "Risk Management": "risk_management.png",
    "ESG Strategy": "esg_strategy.png",
    "JV Structure": "jv_structure.png",
    "Conclusion": "conclusion.png",
}

# Alternate names that map to the same images
SECTION_NAME_ALIASES = {
    "Structural Demand Drivers": "demand_drivers.png",
    "Target Market Analysis": "target_markets.png",
    "Risk Factors": "risk_management.png",
    "Risk Factors & Mitigants": "risk_management.png",
    "ESG & Sustainability": "esg_strategy.png",
    "JV Structure & Governance": "jv_structure.png",
}


def add_images_to_outline(outline: dict, image_dir: str = "cache/images/") -> dict:
    """Add background_image paths to title slides and section dividers."""
    image_path = Path(image_dir)

    for section in outline.get("sections", []):
        section_name = section.get("name", "")

        # Find the appropriate image for this section
        image_file = SECTION_IMAGE_MAP.get(section_name) or SECTION_NAME_ALIASES.get(section_name)

        for slide in section.get("slides", []):
            slide_type = slide.get("slide_type", "")
            content = slide.get("content", {})

            # Add image to title_slide (first section only)
            if slide_type == "title_slide" and section_name == "Executive Summary":
                title_image = image_path / "title_slide.png"
                if title_image.exists():
                    content["background_image"] = str(title_image.absolute())
                    print(f"  Added title_slide image")

            # Add image to section_divider
            elif slide_type == "section_divider" and image_file:
                section_image = image_path / image_file
                if section_image.exists():
                    content["background_image"] = str(section_image.absolute())
                    print(f"  Added {image_file} to '{section_name}' section divider")

    return outline


def main():
    """Regenerate the presentation with images."""
    print("Regenerating Light Industrial presentation with images...")
    print("=" * 60)

    # Load the v14 outline
    outline_path = Path("pptx_generator/output/light_industrial_thesis_v14.json")

    if not outline_path.exists():
        print(f"Outline not found: {outline_path}")
        return

    with open(outline_path) as f:
        outline = json.load(f)

    print(f"\nLoaded outline: {outline.get('title')}")
    print(f"Sections: {len(outline.get('sections', []))}")

    # Add images to outline
    print("\nAdding background images...")
    outline = add_images_to_outline(outline)

    # Save updated outline
    updated_outline_path = Path("pptx_generator/output/light_industrial_thesis_v17.json")
    with open(updated_outline_path, "w") as f:
        json.dump(outline, f, indent=2)
    print(f"\nSaved updated outline: {updated_outline_path}")

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
    output_path = output_dir / "Light_Industrial_Thesis_v17.pptx"

    print("\nGenerating presentation...")
    result = orchestrator.generate_pptx_with_evaluation(
        outline=outline,
        context={"request": "Light Industrial Investment Thesis with Images"}
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
