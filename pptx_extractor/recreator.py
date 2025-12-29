"""
Template Recreator - Generate PPTX from Saved Descriptions

Recreates a PowerPoint slide from a saved description file.

Usage:
    python -m pptx_extractor.recreator --description "template_name.json" --output "output.pptx"
    python -m pptx_extractor.recreator --description "template_name.json" --validate
"""
import json
import logging
import sys
from pathlib import Path
from typing import Optional

import click
from rich.console import Console
from rich.panel import Panel
from rich.table import Table

# Add parent directory to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent))

from config import DESCRIPTION_DIR, OUTPUT_DIR, TEMPLATE_DIR, SIMILARITY_THRESHOLD
from pptx_extractor.generator import (
    generate_slide_from_description,
    generate_from_json,
    generate_multi_slide_presentation,
    combine_json_descriptions
)
from pptx_extractor.renderer import render_slide, verify_dependencies
from pptx_extractor.comparator import compare_slides, generate_diff_image
from pptx_extractor.descriptor import load_description

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Rich console for nice output
console = Console()


class RecreatorError(Exception):
    """Exception raised when recreation fails."""
    pass


def list_descriptions() -> list[Path]:
    """
    List all available saved descriptions.

    Returns:
        List of paths to JSON description files
    """
    descriptions = []

    for item in DESCRIPTION_DIR.iterdir():
        if item.is_file() and item.suffix.lower() == '.json':
            descriptions.append(item)

    return sorted(descriptions)


def display_descriptions():
    """Display available descriptions in a nice table."""
    descriptions = list_descriptions()

    if not descriptions:
        console.print("[yellow]No saved descriptions found.[/yellow]")
        console.print(f"Run 'python -m pptx_extractor.analyzer' to analyze templates first.")
        return []

    table = Table(title="Available Descriptions")
    table.add_column("Index", style="cyan")
    table.add_column("Description Name", style="green")
    table.add_column("Elements", style="dim")

    for i, desc_path in enumerate(descriptions):
        try:
            with open(desc_path) as f:
                desc = json.load(f)
            element_count = len(desc.get("elements", []))
        except Exception:
            element_count = "?"

        table.add_row(str(i), desc_path.stem, str(element_count))

    console.print(table)
    return descriptions


def find_description(description_name: str) -> Path:
    """
    Find a description file by name.

    Args:
        description_name: Description name or path

    Returns:
        Full path to the description file

    Raises:
        RecreatorError: If description not found
    """
    # Try exact path first
    if Path(description_name).exists():
        return Path(description_name)

    # Try in description directory
    desc_path = DESCRIPTION_DIR / description_name
    if desc_path.exists():
        return desc_path

    # Try adding .json extension
    if not description_name.endswith('.json'):
        desc_path = DESCRIPTION_DIR / f"{description_name}.json"
        if desc_path.exists():
            return desc_path

    # Search for partial match
    for desc in list_descriptions():
        if description_name in desc.stem:
            return desc

    raise RecreatorError(f"Description not found: {description_name}")


def find_original_template(description_name: str) -> Optional[Path]:
    """
    Try to find the original template that a description was based on.

    Args:
        description_name: Name of the description

    Returns:
        Path to original template if found, None otherwise
    """
    # Extract template name from description name
    # Description names are like: template_name_slide_1_final
    parts = description_name.rsplit('_slide_', 1)
    if len(parts) >= 1:
        template_name = parts[0]

        # Search for the template
        for template in TEMPLATE_DIR.iterdir():
            if template.is_file() and template.suffix.lower() == '.pptx':
                if template_name in template.stem:
                    return template
            elif template.is_dir():
                for subitem in template.iterdir():
                    if subitem.suffix.lower() == '.pptx':
                        if template_name in subitem.stem:
                            return subitem

    return None


def recreate_from_description(
    description_path: Path,
    output_path: Optional[Path] = None,
    validate: bool = False
) -> dict:
    """
    Generate a PPTX from a saved description.

    Args:
        description_path: Path to the JSON description file
        output_path: Path for the output PPTX (auto-generated if None)
        validate: If True, render and compare to original if available

    Returns:
        Dict with recreation results
    """
    description_path = Path(description_path)

    if not description_path.exists():
        raise RecreatorError(f"Description file not found: {description_path}")

    # Load description
    with open(description_path) as f:
        description = json.load(f)

    description_name = description_path.stem

    console.print(Panel(
        f"[bold]Recreating from:[/bold] {description_name}\n"
        f"[bold]Elements:[/bold] {len(description.get('elements', []))}",
        title="Template Recreation"
    ))

    # Determine output path
    if output_path is None:
        output_path = OUTPUT_DIR / f"{description_name}_recreated.pptx"
    else:
        output_path = Path(output_path)

    # Generate the PPTX
    console.print("\n[bold cyan]Generating PPTX...[/bold cyan]")
    generated_pptx = generate_slide_from_description(description, output_path)
    console.print(f"  Created: {generated_pptx}")

    result = {
        "description_path": str(description_path),
        "output_path": str(generated_pptx),
        "validated": False,
        "similarity": None,
        "diff_path": None
    }

    # Validation
    if validate:
        console.print("\n[bold cyan]Validating against original...[/bold cyan]")

        # Try to find original template
        original_template = find_original_template(description_name)

        if original_template is None:
            console.print("  [yellow]Original template not found, skipping validation.[/yellow]")
        else:
            console.print(f"  Found original: {original_template}")

            # Extract slide index from description name
            slide_index = 0
            parts = description_name.rsplit('_slide_', 1)
            if len(parts) == 2:
                try:
                    # Format is like "1_final" or "1_v3"
                    slide_part = parts[1].split('_')[0]
                    slide_index = int(slide_part) - 1
                except ValueError:
                    pass

            # Render both
            output_dir = OUTPUT_DIR / "validation" / description_name
            output_dir.mkdir(parents=True, exist_ok=True)

            original_image = render_slide(
                original_template,
                slide_index,
                output_dir / "original.png"
            )

            generated_image = render_slide(
                generated_pptx,
                0,  # Generated PPTX has only one slide
                output_dir / "generated.png"
            )

            # Compare
            comparison = compare_slides(original_image, generated_image)

            result["validated"] = True
            result["similarity"] = comparison["similarity"]
            result["original_image"] = str(original_image)
            result["generated_image"] = str(generated_image)
            result["diff_path"] = str(comparison.get("diff_path"))

            # Display results
            similarity = comparison["similarity"]
            matches = similarity >= SIMILARITY_THRESHOLD

            if matches:
                console.print(f"\n  [bold green]MATCH![/bold green] Similarity: {similarity:.4f}")
            else:
                console.print(f"\n  [yellow]Not quite matching.[/yellow] Similarity: {similarity:.4f}")
                console.print(f"  Threshold: {SIMILARITY_THRESHOLD}")

            console.print(f"  Diff image: {comparison.get('diff_path')}")

    console.print(f"\n[bold green]Recreation complete![/bold green]")
    console.print(f"Output: {generated_pptx}")

    return result


def batch_recreate(validate: bool = False) -> list[dict]:
    """
    Recreate all saved descriptions.

    Args:
        validate: If True, validate each recreation

    Returns:
        List of recreation results
    """
    descriptions = list_descriptions()

    if not descriptions:
        console.print("[yellow]No saved descriptions found.[/yellow]")
        return []

    console.print(f"\n[bold]Recreating {len(descriptions)} descriptions...[/bold]\n")

    results = []
    for desc_path in descriptions:
        try:
            result = recreate_from_description(desc_path, validate=validate)
            results.append(result)
        except Exception as e:
            console.print(f"[red]Failed to recreate {desc_path.stem}: {e}[/red]")
            results.append({
                "description_path": str(desc_path),
                "error": str(e)
            })

    # Summary
    console.print(f"\n{'='*60}")
    console.print("[bold]Recreation Summary[/bold]")

    success = sum(1 for r in results if "error" not in r)
    console.print(f"  Successful: {success}/{len(results)}")

    if validate:
        validated = [r for r in results if r.get("validated")]
        if validated:
            avg_similarity = sum(r["similarity"] for r in validated) / len(validated)
            console.print(f"  Average similarity: {avg_similarity:.4f}")

    return results


def find_descriptions_for_template(template_name: str) -> list[Path]:
    """
    Find all slide descriptions for a given template, sorted by slide number.

    Args:
        template_name: Base name of the template

    Returns:
        List of description file paths, sorted by slide number
    """
    import re

    matching = []
    pattern = re.compile(rf"{re.escape(template_name)}_slide_(\d+)")

    for desc_path in list_descriptions():
        match = pattern.match(desc_path.stem)
        if match:
            slide_num = int(match.group(1))
            matching.append((slide_num, desc_path))

    # Sort by slide number and return just the paths
    matching.sort(key=lambda x: x[0])
    return [path for _, path in matching]


def combine_descriptions(
    description_paths: list[Path],
    output_path: Optional[Path] = None
) -> dict:
    """
    Combine multiple slide descriptions into a single multi-slide PPTX.

    Args:
        description_paths: List of description file paths (in slide order)
        output_path: Path for the output PPTX

    Returns:
        Dict with combination results
    """
    if not description_paths:
        raise RecreatorError("No descriptions provided for combination")

    console.print(Panel(
        f"[bold]Combining {len(description_paths)} slides[/bold]",
        title="Multi-Slide Combination"
    ))

    # Display slides being combined
    table = Table(title="Slides to Combine")
    table.add_column("Order", style="cyan")
    table.add_column("Description", style="green")

    for i, path in enumerate(description_paths, 1):
        table.add_row(str(i), path.stem)

    console.print(table)

    # Generate combined PPTX
    console.print("\n[bold cyan]Generating combined PPTX...[/bold cyan]")
    generated_pptx = combine_json_descriptions(description_paths, output_path)

    console.print(f"\n[bold green]Combination complete![/bold green]")
    console.print(f"Output: {generated_pptx}")
    console.print(f"Total slides: {len(description_paths)}")

    return {
        "output_path": str(generated_pptx),
        "slide_count": len(description_paths),
        "descriptions": [str(p) for p in description_paths]
    }


def combine_template_slides(
    template_name: str,
    output_path: Optional[Path] = None
) -> dict:
    """
    Combine all slides from a template into a single PPTX.

    Automatically finds all slide descriptions for the template.

    Args:
        template_name: Base name of the template
        output_path: Path for the output PPTX

    Returns:
        Dict with combination results
    """
    description_paths = find_descriptions_for_template(template_name)

    if not description_paths:
        raise RecreatorError(
            f"No slide descriptions found for template: {template_name}\n"
            f"Expected files like: {template_name}_slide_1_final.json"
        )

    console.print(f"Found {len(description_paths)} slides for template: {template_name}")

    if output_path is None:
        output_path = OUTPUT_DIR / f"{template_name}_combined.pptx"

    return combine_descriptions(description_paths, output_path)


# CLI Commands
@click.group()
def cli():
    """PPTX Template Recreator - Generate slides from descriptions."""
    pass


@cli.command()
def list():
    """List available descriptions."""
    display_descriptions()


@cli.command()
@click.option('--description', '-d', required=True, help='Description name or path')
@click.option('--output', '-o', default=None, help='Output PPTX path')
@click.option('--validate', '-v', is_flag=True, help='Validate against original')
def recreate(description, output, validate):
    """Recreate a template from its description."""
    try:
        # Verify dependencies if validating
        if validate:
            ok, missing = verify_dependencies()
            if not ok:
                console.print("[bold red]Missing dependencies for validation:[/bold red]")
                for m in missing:
                    console.print(f"  - {m}")
                console.print("\nContinuing without validation...")
                validate = False

        desc_path = find_description(description)
        output_path = Path(output) if output else None

        recreate_from_description(desc_path, output_path, validate=validate)

    except RecreatorError as e:
        console.print(f"[bold red]Error:[/bold red] {e}")
        sys.exit(1)
    except Exception as e:
        console.print(f"[bold red]Unexpected error:[/bold red] {e}")
        logger.exception("Recreation failed")
        sys.exit(1)


@cli.command()
@click.option('--validate', '-v', is_flag=True, help='Validate each recreation')
def batch(validate):
    """Recreate all saved descriptions."""
    try:
        if validate:
            ok, missing = verify_dependencies()
            if not ok:
                console.print("[bold red]Missing dependencies for validation:[/bold red]")
                for m in missing:
                    console.print(f"  - {m}")
                console.print("\nContinuing without validation...")
                validate = False

        batch_recreate(validate=validate)

    except Exception as e:
        console.print(f"[bold red]Error:[/bold red] {e}")
        sys.exit(1)


@cli.command()
@click.option('--description', '-d', required=True, help='Description to view')
def view(description):
    """View a description's contents."""
    try:
        desc_path = find_description(description)

        with open(desc_path) as f:
            desc = json.load(f)

        console.print_json(data=desc)

    except RecreatorError as e:
        console.print(f"[bold red]Error:[/bold red] {e}")
        sys.exit(1)


@cli.command()
@click.option('--template', '-t', default=None, help='Template name to combine all slides')
@click.option('--descriptions', '-d', multiple=True, help='Specific description files to combine')
@click.option('--output', '-o', default=None, help='Output PPTX path')
def combine(template, descriptions, output):
    """
    Combine multiple slide descriptions into a single multi-slide PPTX.

    Two modes:
    1. By template: --template "template_name" (auto-finds all slides)
    2. Manual: --descriptions file1.json --descriptions file2.json
    """
    try:
        output_path = Path(output) if output else None

        if template:
            # Auto-find all slides for this template
            combine_template_slides(template, output_path)

        elif descriptions:
            # Manual specification of files
            desc_paths = []
            for desc in descriptions:
                desc_paths.append(find_description(desc))
            combine_descriptions(desc_paths, output_path)

        else:
            console.print("[yellow]Please specify either --template or --descriptions[/yellow]")
            console.print("\nExamples:")
            console.print("  python -m pptx_extractor.recreator combine --template my_template")
            console.print("  python -m pptx_extractor.recreator combine -d slide1.json -d slide2.json")
            sys.exit(1)

    except RecreatorError as e:
        console.print(f"[bold red]Error:[/bold red] {e}")
        sys.exit(1)
    except Exception as e:
        console.print(f"[bold red]Unexpected error:[/bold red] {e}")
        logger.exception("Combination failed")
        sys.exit(1)


@cli.command()
def templates():
    """List templates that have slide descriptions available for combining."""
    import re

    descriptions = list_descriptions()

    if not descriptions:
        console.print("[yellow]No saved descriptions found.[/yellow]")
        return

    # Group by template name
    template_slides = {}
    pattern = re.compile(r"(.+)_slide_(\d+)")

    for desc_path in descriptions:
        match = pattern.match(desc_path.stem)
        if match:
            template_name = match.group(1)
            slide_num = int(match.group(2))
            if template_name not in template_slides:
                template_slides[template_name] = []
            template_slides[template_name].append(slide_num)

    if not template_slides:
        console.print("[yellow]No multi-slide templates found.[/yellow]")
        return

    table = Table(title="Templates Available for Combining")
    table.add_column("Template Name", style="green")
    table.add_column("Slides", style="cyan")
    table.add_column("Command", style="dim")

    for template_name, slides in sorted(template_slides.items()):
        slides.sort()
        slide_str = ", ".join(str(s) for s in slides)
        cmd = f"python -m pptx_extractor.recreator combine -t {template_name}"
        table.add_row(template_name, slide_str, cmd)

    console.print(table)


if __name__ == "__main__":
    cli()
