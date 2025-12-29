"""
Template Analyzer - Main Orchestration Script

Analyzes an existing PPTX template and generates a precise natural language
description through an iterative vision-based feedback loop.

Usage:
    python -m pptx_extractor.analyzer --template "template_name.pptx" --slide 0
    python -m pptx_extractor.analyzer --template "template_name.pptx" --all-slides
    python -m pptx_extractor.analyzer --list  # List available templates
"""
import json
import logging
import sys
from pathlib import Path
from typing import Optional

import click
from rich.console import Console
from rich.panel import Panel
from rich.progress import Progress, SpinnerColumn, TextColumn
from rich.table import Table

# Add parent directory to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent))

from config import (
    TEMPLATE_DIR,
    OUTPUT_DIR,
    DESCRIPTION_DIR,
    MAX_ITERATIONS,
    SIMILARITY_THRESHOLD
)
from pptx_extractor.renderer import render_slide, render_all_slides, get_slide_count, verify_dependencies
from pptx_extractor.comparator import compute_similarity, generate_diff_image, compare_slides
from pptx_extractor.generator import generate_slide_from_description
from pptx_extractor.descriptor import (
    describe_slide_design,
    compare_slides_for_differences,
    refine_description,
    save_description,
    load_description
)

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Rich console for nice output (force_terminal=True for Windows encoding compatibility)
console = Console(force_terminal=True, no_color=False, legacy_windows=True)


class AnalyzerError(Exception):
    """Exception raised when analysis fails."""
    pass


def list_templates() -> list[Path]:
    """
    List all available PPTX templates.

    Returns:
        List of paths to template files
    """
    templates = []

    for item in TEMPLATE_DIR.iterdir():
        if item.is_file() and item.suffix.lower() == '.pptx':
            templates.append(item)
        elif item.is_dir():
            # Check for PPTX files in subdirectories
            for subitem in item.iterdir():
                if subitem.is_file() and subitem.suffix.lower() == '.pptx':
                    templates.append(subitem)

    return sorted(templates)


def display_templates():
    """Display available templates in a nice table."""
    templates = list_templates()

    table = Table(title="Available Templates")
    table.add_column("Index", style="cyan")
    table.add_column("Template Name", style="green")
    table.add_column("Location", style="dim")

    for i, template in enumerate(templates):
        rel_path = template.relative_to(TEMPLATE_DIR)
        table.add_row(str(i), template.stem, str(rel_path.parent))

    console.print(table)
    return templates


def find_template(template_name: str) -> Path:
    """
    Find a template by name.

    Args:
        template_name: Template name or partial path

    Returns:
        Full path to the template

    Raises:
        AnalyzerError: If template not found
    """
    # Try exact path first
    if Path(template_name).exists():
        return Path(template_name)

    # Try in template directory
    template_path = TEMPLATE_DIR / template_name
    if template_path.exists():
        return template_path

    # Try adding .pptx extension
    if not template_name.endswith('.pptx'):
        template_path = TEMPLATE_DIR / f"{template_name}.pptx"
        if template_path.exists():
            return template_path

    # Search in subdirectories
    for template in list_templates():
        if template_name in template.name or template_name in str(template):
            return template

    raise AnalyzerError(f"Template not found: {template_name}")


def analyze_template(
    template_path: Path,
    slide_index: int = 0,
    max_iterations: int = None,
    use_anthropic: bool = False,
    interactive: bool = True,
    model: Optional[str] = None
) -> dict:
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

    Args:
        template_path: Path to the PPTX template
        slide_index: Zero-based index of the slide to analyze
        max_iterations: Maximum refinement iterations
        use_anthropic: If True, use Anthropic API directly (legacy)
        interactive: If True, prompt for user input during analysis
        model: LLM model to use (e.g., "claude-sonnet-4.5", "gpt-5", "gemini-3-pro")

    Returns:
        Final description dict
    """
    if max_iterations is None:
        max_iterations = MAX_ITERATIONS

    template_path = Path(template_path)
    template_name = template_path.stem

    console.print(Panel(
        f"[bold]Analyzing Template:[/bold] {template_name}\n"
        f"[bold]Slide:[/bold] {slide_index + 1}\n"
        f"[bold]Max Iterations:[/bold] {max_iterations}",
        title="Template Analysis"
    ))

    # Create output directory for this analysis
    analysis_dir = OUTPUT_DIR / template_name / f"slide_{slide_index + 1}"
    analysis_dir.mkdir(parents=True, exist_ok=True)

    # Step 1: Render the original template slide
    console.print("\n[bold cyan]Step 1:[/bold cyan] Rendering original template...")

    original_image = render_slide(
        template_path,
        slide_index,
        analysis_dir / "original.png"
    )
    console.print(f"  Original rendered: {original_image}")

    # Step 2: Generate initial description
    console.print("\n[bold cyan]Step 2:[/bold cyan] Generating initial description...")

    if model:
        # Use configurable LLM provider
        description = describe_slide_design(original_image, model=model)
        console.print(f"  Initial description generated via {model}")
    elif use_anthropic:
        description = describe_slide_design(original_image, use_anthropic=True)
        console.print("  Initial description generated via Anthropic API")
    else:
        # Return data for manual/external processing
        prompt_data = describe_slide_design(original_image, use_anthropic=False)

        if interactive:
            console.print("\n[yellow]Vision analysis required.[/yellow]")
            console.print("Please analyze the image and provide a JSON description.")
            console.print(f"Image path: {original_image}")
            console.print("\nPaste the JSON description (end with an empty line):")

            lines = []
            while True:
                line = input()
                if line == "":
                    break
                lines.append(line)

            try:
                description = json.loads("\n".join(lines))
            except json.JSONDecodeError as e:
                raise AnalyzerError(f"Invalid JSON: {e}")
        else:
            # Non-interactive mode: return prompt data for external processing
            console.print("  [yellow]Non-interactive mode: returning prompt data[/yellow]")
            return {
                "status": "needs_vision",
                "prompt_data": prompt_data,
                "original_image": str(original_image),
                "analysis_dir": str(analysis_dir)
            }

    # Save initial description
    save_description(description, f"{template_name}_slide_{slide_index + 1}_v0")

    # Step 3: Iterative refinement loop
    console.print("\n[bold cyan]Step 3:[/bold cyan] Iterative refinement loop...")

    for iteration in range(max_iterations):
        console.print(f"\n[bold]Iteration {iteration + 1}/{max_iterations}[/bold]")

        # 3a: Generate PPTX from current description
        generated_pptx = generate_slide_from_description(
            description,
            analysis_dir / f"generated_v{iteration + 1}.pptx"
        )

        # 3b: Render generated PPTX
        generated_image = render_slide(
            generated_pptx,
            0,  # Always slide 0 since we generated a single-slide PPTX
            analysis_dir / f"generated_v{iteration + 1}.png"
        )

        # 3c: Compare original vs generated
        result = compare_slides(original_image, generated_image)
        similarity = result["similarity"]

        console.print(f"  Similarity: {similarity:.4f} (threshold: {SIMILARITY_THRESHOLD})")

        # 3d: Check if we've reached the threshold
        if similarity >= SIMILARITY_THRESHOLD:
            console.print(f"\n[bold green]SUCCESS![/bold green] Similarity threshold reached.")
            break

        # 3e: Get diff feedback and refine
        if use_anthropic:
            diff_feedback = compare_slides_for_differences(
                original_image,
                generated_image,
                use_anthropic=True
            )
            console.print(f"  Differences identified")

            description = refine_description(
                description,
                diff_feedback,
                use_anthropic=True
            )
        else:
            if interactive:
                # Generate visual diff
                diff_image = generate_diff_image(
                    original_image,
                    generated_image,
                    analysis_dir / f"diff_v{iteration + 1}.png"
                )

                console.print(f"\n[yellow]Visual comparison needed.[/yellow]")
                console.print(f"Original: {original_image}")
                console.print(f"Generated: {generated_image}")
                console.print(f"Diff: {diff_image}")
                console.print("\nDescribe the differences (end with empty line):")

                lines = []
                while True:
                    line = input()
                    if line == "":
                        break
                    lines.append(line)

                diff_feedback = "\n".join(lines)

                if diff_feedback.strip().upper() == "MATCH":
                    console.print("\n[bold green]User confirmed match![/bold green]")
                    break

                console.print("\nProvide the refined JSON description (end with empty line):")
                lines = []
                while True:
                    line = input()
                    if line == "":
                        break
                    lines.append(line)

                try:
                    description = json.loads("\n".join(lines))
                except json.JSONDecodeError as e:
                    console.print(f"[red]Invalid JSON: {e}[/red]")
                    continue
            else:
                # Return data for external processing
                return {
                    "status": "needs_refinement",
                    "iteration": iteration + 1,
                    "current_description": description,
                    "similarity": similarity,
                    "original_image": str(original_image),
                    "generated_image": str(generated_image),
                    "analysis_dir": str(analysis_dir)
                }

        # Save intermediate description
        save_description(description, f"{template_name}_slide_{slide_index + 1}_v{iteration + 1}")

    # Save final description
    final_json, final_md = save_description(
        description,
        f"{template_name}_slide_{slide_index + 1}_final"
    )

    console.print(f"\n[bold green]Analysis complete![/bold green]")
    console.print(f"Final description: {final_json}")
    console.print(f"Markdown summary: {final_md}")

    return description


def analyze_all_slides(
    template_path: Path,
    max_iterations: int = None,
    use_anthropic: bool = False,
    interactive: bool = True,
    model: Optional[str] = None
) -> list[dict]:
    """
    Analyze all slides in a template.

    Args:
        template_path: Path to the PPTX template
        max_iterations: Maximum refinement iterations per slide
        use_anthropic: If True, use Anthropic API directly (legacy)
        interactive: If True, prompt for user input
        model: LLM model to use (e.g., "claude-sonnet-4.5", "gpt-5", "gemini-3-pro")

    Returns:
        List of description dicts, one per slide
    """
    template_path = Path(template_path)
    slide_count = get_slide_count(template_path)

    console.print(f"\n[bold]Analyzing {slide_count} slides in {template_path.name}[/bold]\n")

    descriptions = []
    for i in range(slide_count):
        console.print(f"\n{'='*60}")
        desc = analyze_template(
            template_path,
            slide_index=i,
            max_iterations=max_iterations,
            use_anthropic=use_anthropic,
            model=model,
            interactive=interactive
        )
        descriptions.append(desc)

    return descriptions


# CLI Commands
@click.group()
def cli():
    """PPTX Template Analyzer - Extract design specifications from templates."""
    pass


@cli.command()
def list():
    """List available templates."""
    display_templates()


@cli.command()
def check():
    """Check system dependencies."""
    console.print("[bold]Checking dependencies...[/bold]\n")

    ok, missing = verify_dependencies()

    if ok:
        console.print("[bold green]All dependencies are available![/bold green]")
    else:
        console.print("[bold red]Missing dependencies:[/bold red]")
        for m in missing:
            console.print(f"  - {m}")
        sys.exit(1)


@cli.command()
@click.option('--template', '-t', required=True, help='Template name or path')
@click.option('--slide', '-s', default=0, help='Slide index (0-based)')
@click.option('--all-slides', '-a', is_flag=True, help='Analyze all slides')
@click.option('--max-iterations', '-m', default=None, type=int, help='Max iterations')
@click.option('--use-anthropic', '-api', is_flag=True, help='Use Anthropic API directly (legacy)')
@click.option('--model', '-M', default=None,
              help='LLM model to use (e.g., claude-sonnet-4.5, gpt-5, gemini-3-pro)')
@click.option('--non-interactive', '-n', is_flag=True, help='Non-interactive mode')
def analyze(template, slide, all_slides, max_iterations, use_anthropic, model, non_interactive):
    """Analyze a template and generate its description."""
    try:
        # Verify dependencies first
        ok, missing = verify_dependencies()
        if not ok:
            console.print("[bold red]Missing dependencies:[/bold red]")
            for m in missing:
                console.print(f"  - {m}")
            console.print("\nPlease install missing dependencies and try again.")
            sys.exit(1)

        template_path = find_template(template)
        console.print(f"Found template: {template_path}")

        if all_slides:
            analyze_all_slides(
                template_path,
                max_iterations=max_iterations,
                use_anthropic=use_anthropic,
                interactive=not non_interactive,
                model=model
            )
        else:
            analyze_template(
                template_path,
                slide_index=slide,
                max_iterations=max_iterations,
                use_anthropic=use_anthropic,
                interactive=not non_interactive,
                model=model
            )

    except AnalyzerError as e:
        console.print(f"[bold red]Error:[/bold red] {e}")
        sys.exit(1)
    except Exception as e:
        console.print(f"[bold red]Unexpected error:[/bold red] {e}")
        logger.exception("Analysis failed")
        sys.exit(1)


@cli.command()
@click.option('--template', '-t', required=True, help='Template name or path')
def render(template):
    """Render all slides of a template to PNG."""
    try:
        ok, missing = verify_dependencies()
        if not ok:
            console.print("[bold red]Missing dependencies:[/bold red]")
            for m in missing:
                console.print(f"  - {m}")
            sys.exit(1)

        template_path = find_template(template)

        with Progress(
            SpinnerColumn(spinner_name="line"),  # ASCII-safe spinner for Windows
            TextColumn("[progress.description]{task.description}"),
            console=console
        ) as progress:
            task = progress.add_task(f"Rendering {template_path.name}...", total=None)
            output_paths = render_all_slides(template_path)
            progress.update(task, completed=True)

        console.print(f"\n[bold green]Rendered {len(output_paths)} slides:[/bold green]")
        for p in output_paths:
            console.print(f"  - {p}")

    except Exception as e:
        console.print(f"[bold red]Error:[/bold red] {e}")
        sys.exit(1)


@cli.command('batch')
@click.option('--render-only', '-r', is_flag=True, help='Only render, skip analysis')
@click.option('--extract-all', '-e', is_flag=True, help='Extract theme and masters too')
@click.option('--max-slides', '-m', default=None, type=int, help='Max slides per template')
def batch_analyze(render_only, extract_all, max_slides):
    """Analyze all templates in the templates folder."""
    try:
        templates = list_templates()

        if not templates:
            console.print("[yellow]No templates found in the templates folder.[/yellow]")
            return

        console.print(f"\n[bold]Found {len(templates)} templates to process[/bold]\n")

        # Summary table
        summary_table = Table(title="Batch Processing Summary")
        summary_table.add_column("Template", style="green")
        summary_table.add_column("Slides", style="cyan")
        summary_table.add_column("Status", style="dim")

        results = []

        for template_path in templates:
            console.print(f"\n{'='*60}")
            console.print(f"[bold]Processing:[/bold] {template_path.name}")

            try:
                slide_count = get_slide_count(template_path)
                slides_to_process = slide_count
                if max_slides:
                    slides_to_process = min(slide_count, max_slides)

                console.print(f"  Slides: {slide_count} (processing {slides_to_process})")

                # Render slides
                console.print("  Rendering slides...")
                output_paths = render_all_slides(template_path)
                console.print(f"  Rendered {len(output_paths)} slides")

                result = {
                    'template': template_path.name,
                    'slide_count': slide_count,
                    'rendered': len(output_paths),
                    'status': 'rendered'
                }

                # Extract theme and masters if requested
                if extract_all:
                    try:
                        from pptx_extractor.themes import extract_theme, save_theme_info
                        from pptx_extractor.masters import extract_all_masters, save_master_info

                        console.print("  Extracting theme...")
                        theme_info = extract_theme(template_path)
                        save_theme_info(theme_info, template_path.stem)

                        console.print("  Extracting masters...")
                        master_info = extract_all_masters(template_path)
                        save_master_info(master_info, template_path.stem)

                        result['theme_extracted'] = True
                        result['masters_extracted'] = True
                        result['status'] = 'extracted'

                    except Exception as e:
                        console.print(f"  [yellow]Warning: Could not extract theme/masters: {e}[/yellow]")
                        result['extraction_error'] = str(e)

                results.append(result)
                summary_table.add_row(
                    template_path.name,
                    str(slide_count),
                    result['status']
                )

            except Exception as e:
                console.print(f"  [red]Error: {e}[/red]")
                results.append({
                    'template': template_path.name,
                    'error': str(e),
                    'status': 'failed'
                })
                summary_table.add_row(
                    template_path.name,
                    "?",
                    f"[red]failed[/red]"
                )

        # Print summary
        console.print(f"\n{'='*60}")
        console.print(summary_table)

        # Stats
        successful = sum(1 for r in results if r.get('status') != 'failed')
        total_slides = sum(r.get('rendered', 0) for r in results)
        console.print(f"\n[bold]Summary:[/bold]")
        console.print(f"  Templates processed: {successful}/{len(templates)}")
        console.print(f"  Total slides rendered: {total_slides}")

    except Exception as e:
        console.print(f"[bold red]Error:[/bold red] {e}")
        logger.exception("Batch processing failed")
        sys.exit(1)


@cli.command('info')
@click.option('--template', '-t', required=True, help='Template name or path')
def template_info(template):
    """Show detailed information about a template."""
    try:
        from pptx_extractor.themes import extract_theme, extract_color_palette, extract_font_families
        from pptx_extractor.masters import extract_all_masters, extract_slide_layout_usage

        template_path = find_template(template)

        console.print(Panel(
            f"[bold]{template_path.name}[/bold]",
            title="Template Information"
        ))

        # Basic info
        slide_count = get_slide_count(template_path)
        console.print(f"\n[bold]Slides:[/bold] {slide_count}")

        # Layout usage
        usage = extract_slide_layout_usage(template_path)
        if usage:
            layout_table = Table(title="Slide Layouts")
            layout_table.add_column("Slide", style="cyan")
            layout_table.add_column("Layout", style="green")

            for u in usage:
                layout_table.add_row(str(u['slide_number']), u['layout_name'])

            console.print(layout_table)

        # Theme info
        theme_info = extract_theme(template_path)
        if theme_info.get('themes'):
            console.print("\n[bold]Theme:[/bold]")
            for t in theme_info['themes']:
                colors = t.get('colors', {})
                if colors:
                    accent_colors = [c for k, c in colors.items() if 'accent' in k and c]
                    if accent_colors:
                        console.print(f"  Accent colors: {', '.join(accent_colors[:4])}")

        # Fonts
        fonts = extract_font_families(template_path)
        if fonts.get('heading') or fonts.get('body'):
            console.print(f"\n[bold]Fonts:[/bold]")
            console.print(f"  Heading: {fonts.get('heading', 'N/A')}")
            console.print(f"  Body: {fonts.get('body', 'N/A')}")

        # Master info
        master_info = extract_all_masters(template_path)
        if master_info.get('slide_masters'):
            console.print(f"\n[bold]Slide Masters:[/bold] {len(master_info['slide_masters'])}")
            for master in master_info['slide_masters']:
                console.print(f"  - {master['name']}: {len(master['layouts'])} layouts")

    except AnalyzerError as e:
        console.print(f"[bold red]Error:[/bold red] {e}")
        sys.exit(1)
    except Exception as e:
        console.print(f"[bold red]Unexpected error:[/bold red] {e}")
        logger.exception("Info extraction failed")
        sys.exit(1)


@cli.command()
@click.option('--template', '-t', required=True, help='Template name or path')
def theme(template):
    """Extract theme (colors, fonts, effects) from a template."""
    try:
        from pptx_extractor.themes import extract_theme, save_theme_info, extract_color_palette, extract_font_families

        template_path = find_template(template)
        console.print(f"[bold]Extracting theme from:[/bold] {template_path.name}\n")

        # Extract theme info
        theme_info = extract_theme(template_path)

        for i, t in enumerate(theme_info.get('themes', [])):
            console.print(f"\n[bold cyan]Theme {i + 1}: {t.get('master_name', 'Unknown')}[/bold cyan]")

            # Color table
            colors = t.get('colors', {})
            if colors:
                color_table = Table(title="Color Scheme")
                color_table.add_column("Role", style="cyan")
                color_table.add_column("Color", style="green")

                for name, color in colors.items():
                    color_table.add_row(name, color or "N/A")

                console.print(color_table)

            # Font info
            fonts = t.get('fonts', {})
            if fonts:
                major = fonts.get('major', {})
                minor = fonts.get('minor', {})
                if major.get('latin'):
                    console.print(f"  Heading Font: {major['latin']}")
                if minor.get('latin'):
                    console.print(f"  Body Font: {minor['latin']}")

        # Show extracted palette
        palette = extract_color_palette(template_path)
        if palette:
            console.print(f"\n[bold]Color Palette:[/bold] {', '.join(palette)}")

        # Show fonts
        fonts = extract_font_families(template_path)
        if fonts.get('heading') or fonts.get('body'):
            console.print(f"[bold]Heading Font:[/bold] {fonts.get('heading', 'N/A')}")
            console.print(f"[bold]Body Font:[/bold] {fonts.get('body', 'N/A')}")

        # Save to file
        template_name = template_path.stem
        output_path = save_theme_info(theme_info, template_name)
        console.print(f"\n[bold green]Theme info saved:[/bold green] {output_path}")

    except AnalyzerError as e:
        console.print(f"[bold red]Error:[/bold red] {e}")
        sys.exit(1)
    except Exception as e:
        console.print(f"[bold red]Unexpected error:[/bold red] {e}")
        logger.exception("Theme extraction failed")
        sys.exit(1)


@cli.command()
@click.option('--template', '-t', required=True, help='Template name or path')
def masters(template):
    """Extract slide master and layout information from a template."""
    try:
        from pptx_extractor.masters import extract_all_masters, extract_slide_layout_usage, save_master_info

        template_path = find_template(template)
        console.print(f"[bold]Extracting masters from:[/bold] {template_path.name}\n")

        # Extract master info
        master_info = extract_all_masters(template_path)

        # Display summary
        console.print(f"[bold]Slide Dimensions:[/bold] {master_info['slide_dimensions']['width_inches']:.2f}\" x {master_info['slide_dimensions']['height_inches']:.2f}\"")
        console.print(f"[bold]Slide Masters:[/bold] {len(master_info['slide_masters'])}")

        for i, master in enumerate(master_info['slide_masters']):
            console.print(f"\n[bold cyan]Master {i + 1}: {master['name']}[/bold cyan]")
            console.print(f"  Placeholders: {len(master['placeholders'])}")
            console.print(f"  Shapes: {len(master['shapes'])}")
            console.print(f"  Layouts: {len(master['layouts'])}")

            # Show layouts in a table
            if master['layouts']:
                table = Table(title=f"Layouts in {master['name']}")
                table.add_column("Index", style="cyan")
                table.add_column("Name", style="green")
                table.add_column("Placeholders", style="dim")

                for j, layout in enumerate(master['layouts']):
                    table.add_row(
                        str(j),
                        layout['name'],
                        str(len(layout['placeholders']))
                    )

                console.print(table)

        # Show slide usage
        usage = extract_slide_layout_usage(template_path)
        if usage:
            console.print("\n[bold]Slide Layout Usage:[/bold]")
            for u in usage:
                console.print(f"  Slide {u['slide_number']}: {u['layout_name']}")

        # Save to file
        template_name = template_path.stem
        output_path = save_master_info(master_info, template_name)
        console.print(f"\n[bold green]Master info saved:[/bold green] {output_path}")

    except AnalyzerError as e:
        console.print(f"[bold red]Error:[/bold red] {e}")
        sys.exit(1)
    except Exception as e:
        console.print(f"[bold red]Unexpected error:[/bold red] {e}")
        logger.exception("Master extraction failed")
        sys.exit(1)


@cli.command('enhanced-analyze')
@click.option('--image', '-i', required=True, help='Path to slide image')
@click.option('--mode', '-m', default='standard',
              type=click.Choice(['fast', 'standard', 'thorough', 'ultra']),
              help='Analysis mode')
@click.option('--provider', '-p', default='anthropic',
              type=click.Choice(['anthropic', 'openai', 'google']),
              help='LLM provider')
@click.option('--no-cache', is_flag=True, help='Disable caching')
@click.option('--no-preprocess', is_flag=True, help='Disable image preprocessing')
@click.option('--output', '-o', default=None, help='Output JSON file path')
def enhanced_analyze(image, mode, provider, no_cache, no_preprocess, output):
    """
    Analyze a slide image using enhanced LLM vision analysis.

    Modes:
      fast      - Quick analysis with basic model
      standard  - Balanced accuracy and speed (default)
      thorough  - Multi-pass with validation
      ultra     - Extended thinking for complex slides
    """
    try:
        from pptx_extractor.enhanced_analyzer import (
            EnhancedAnalyzer, AnalysisConfig, AnalysisMode
        )

        image_path = Path(image)
        if not image_path.exists():
            console.print(f"[bold red]Error:[/bold red] Image not found: {image}")
            sys.exit(1)

        console.print(Panel(
            f"[bold]Image:[/bold] {image_path.name}\n"
            f"[bold]Mode:[/bold] {mode}\n"
            f"[bold]Provider:[/bold] {provider}",
            title="Enhanced Analysis"
        ))

        config = AnalysisConfig(
            mode=AnalysisMode(mode),
            preferred_provider=provider,
            use_cache=not no_cache,
            preprocess_image=not no_preprocess
        )

        analyzer = EnhancedAnalyzer(config)

        with Progress(
            SpinnerColumn(spinner_name="line"),
            TextColumn("[progress.description]{task.description}"),
            console=console
        ) as progress:
            task = progress.add_task(f"Analyzing with {mode} mode...", total=None)
            result = analyzer.analyze(image_path)
            progress.update(task, completed=True)

        # Display results
        console.print(f"\n[bold green]Analysis Complete[/bold green]")
        console.print(f"  Confidence: {result.confidence:.2%}")
        console.print(f"  Model: {result.model_used}")
        console.print(f"  Tokens: {result.tokens_used}")
        console.print(f"  Time: {result.processing_time_seconds:.2f}s")
        console.print(f"  Passes: {result.passes_completed}")
        console.print(f"  Cache hit: {result.cache_hit}")

        if result.corrections_applied:
            console.print(f"\n[yellow]Corrections applied:[/yellow]")
            for c in result.corrections_applied[:5]:
                console.print(f"  - {c}")

        # Save output
        if output:
            output_path = Path(output)
        else:
            output_path = image_path.parent / f"{image_path.stem}_description.json"

        with open(output_path, 'w') as f:
            json.dump(result.description, f, indent=2)

        console.print(f"\n[bold green]Saved to:[/bold green] {output_path}")

    except Exception as e:
        console.print(f"[bold red]Error:[/bold red] {e}")
        logger.exception("Enhanced analysis failed")
        sys.exit(1)


@cli.command('batch-enhanced')
@click.option('--dir', '-d', required=True, help='Directory containing slide images')
@click.option('--mode', '-m', default='standard',
              type=click.Choice(['fast', 'standard', 'thorough']),
              help='Analysis mode')
@click.option('--provider', '-p', default='anthropic',
              type=click.Choice(['anthropic', 'openai', 'google']),
              help='LLM provider')
@click.option('--parallel', '-j', default=3, help='Max parallel analyses')
@click.option('--output-dir', '-o', default=None, help='Output directory')
def batch_enhanced_analyze(dir, mode, provider, parallel, output_dir):
    """
    Analyze multiple slide images in parallel using enhanced analysis.
    """
    try:
        from pptx_extractor.enhanced_analyzer import (
            EnhancedAnalyzer, AnalysisConfig, AnalysisMode
        )

        input_dir = Path(dir)
        if not input_dir.exists():
            console.print(f"[bold red]Error:[/bold red] Directory not found: {dir}")
            sys.exit(1)

        # Find all PNG images
        images = list(input_dir.glob("*.png"))
        if not images:
            console.print("[yellow]No PNG images found in directory[/yellow]")
            return

        console.print(f"[bold]Found {len(images)} images to analyze[/bold]\n")

        config = AnalysisConfig(
            mode=AnalysisMode(mode),
            preferred_provider=provider,
            max_parallel=parallel
        )

        analyzer = EnhancedAnalyzer(config)

        # Analyze in batch
        console.print(f"Analyzing with {parallel} parallel workers...")
        results = analyzer.analyze_batch(images)

        # Output directory
        out_dir = Path(output_dir) if output_dir else input_dir / "analysis_results"
        out_dir.mkdir(parents=True, exist_ok=True)

        # Summary table
        summary_table = Table(title="Batch Analysis Results")
        summary_table.add_column("Image", style="cyan")
        summary_table.add_column("Confidence", style="green")
        summary_table.add_column("Time", style="dim")
        summary_table.add_column("Status", style="dim")

        for img, result in zip(images, results):
            # Save result
            out_file = out_dir / f"{img.stem}_description.json"
            with open(out_file, 'w') as f:
                json.dump(result.description, f, indent=2)

            status = "OK" if result.confidence > 0.7 else "[yellow]Low[/yellow]"
            if "error" in result.description:
                status = "[red]Error[/red]"

            summary_table.add_row(
                img.name,
                f"{result.confidence:.1%}",
                f"{result.processing_time_seconds:.1f}s",
                status
            )

        console.print(summary_table)
        console.print(f"\n[bold green]Results saved to:[/bold green] {out_dir}")

    except Exception as e:
        console.print(f"[bold red]Error:[/bold red] {e}")
        logger.exception("Batch analysis failed")
        sys.exit(1)


@cli.command('preprocess')
@click.option('--image', '-i', required=True, help='Path to slide image')
@click.option('--output', '-o', default=None, help='Output path')
@click.option('--no-grid', is_flag=True, help='Disable measurement grid')
@click.option('--no-rulers', is_flag=True, help='Disable rulers')
def preprocess_image(image, output, no_grid, no_rulers):
    """
    Preprocess a slide image for analysis (add grid, rulers, normalize).
    """
    try:
        from pptx_extractor.image_preprocessor import (
            ImagePreprocessor, PreprocessingConfig
        )

        image_path = Path(image)
        if not image_path.exists():
            console.print(f"[bold red]Error:[/bold red] Image not found: {image}")
            sys.exit(1)

        config = PreprocessingConfig(
            add_grid=not no_grid,
            add_rulers=not no_rulers,
            normalize_contrast=True,
            normalize_brightness=True
        )

        preprocessor = ImagePreprocessor(config)

        output_path = Path(output) if output else None
        result_path, metadata = preprocessor.preprocess(image_path, output_path)

        console.print(f"[bold green]Preprocessed image saved:[/bold green] {result_path}")
        console.print(f"  Original size: {metadata['original_size']}")
        console.print(f"  Output size: {metadata['output_size']}")
        console.print(f"  Applied: {', '.join(metadata['preprocessing_applied'])}")

    except Exception as e:
        console.print(f"[bold red]Error:[/bold red] {e}")
        sys.exit(1)


@cli.command('create-calibration')
@click.option('--output', '-o', default='calibration_reference.png', help='Output path')
def create_calibration(output):
    """
    Create a calibration reference image for verifying measurement accuracy.
    """
    try:
        from pptx_extractor.image_preprocessor import create_calibration_image

        output_path = Path(output)
        result = create_calibration_image(output_path)

        console.print(f"[bold green]Calibration image created:[/bold green] {result}")
        console.print("Use this to verify your LLM's measurement accuracy.")

    except Exception as e:
        console.print(f"[bold red]Error:[/bold red] {e}")
        sys.exit(1)


@cli.command('clear-cache')
def clear_cache():
    """Clear the analysis cache."""
    try:
        from pptx_extractor.enhanced_analyzer import AnalysisCache

        cache = AnalysisCache()
        cache.clear()
        console.print("[bold green]Cache cleared successfully[/bold green]")

    except Exception as e:
        console.print(f"[bold red]Error:[/bold red] {e}")
        sys.exit(1)


# ============================================================================
# OPTIMIZED EXTRACTION COMMANDS
# ============================================================================

@cli.command('estimate-cost')
@click.option('--slides', '-n', required=True, type=int, help='Number of slides')
@click.option('--detailed', is_flag=True, help='Show detailed breakdown')
def estimate_cost(slides, detailed):
    """
    Estimate the cost of extracting a template.

    Shows cost comparison between optimized and full extraction.
    """
    try:
        from pptx_extractor.optimized_extractor import quick_estimate
        from pptx_extractor.cached_vision import estimate_batch_cost

        estimate = quick_estimate(slides)

        console.print(f"\n[bold]Cost Estimate for {slides} slides[/bold]\n")

        table = Table(show_header=True)
        table.add_column("Method", style="cyan")
        table.add_column("Cost", style="green")
        table.add_column("Notes", style="dim")

        table.add_row(
            "Optimized (2-phase)",
            f"${estimate['optimized_total']:.2f}",
            f"~{estimate['estimated_unique_patterns']} unique patterns"
        )
        table.add_row(
            "Full extraction",
            f"${estimate['full_extraction_cost']:.2f}",
            "All slides with Sonnet"
        )
        table.add_row(
            "[bold]Savings[/bold]",
            f"[bold green]${estimate['savings_usd']:.2f}[/bold green]",
            f"[bold]{estimate['savings_percentage']}%[/bold]"
        )

        console.print(table)

        if detailed:
            console.print(f"\n[bold]Breakdown:[/bold]")
            console.print(f"  Phase 1 (categorize): ${estimate['phase1_cost']:.2f}")
            console.print(f"  Phase 2 (extract):    ${estimate['phase2_cost']:.2f}")

            # Show model costs
            console.print(f"\n[bold]Model comparison for full extraction:[/bold]")
            for model in ["claude-haiku-4.5", "claude-sonnet-4.5", "claude-opus-4.5"]:
                cost = estimate_batch_cost(slides, "detailed", model, False, False)
                console.print(f"  {model}: ${cost['final_cost_usd']:.2f}")

    except Exception as e:
        console.print(f"[bold red]Error:[/bold red] {e}")
        sys.exit(1)


@cli.command('extract-optimized')
@click.option('--template', '-t', required=True, help='Template PPTX file')
@click.option('--output-dir', '-o', default='descriptions', help='Output directory')
@click.option('--skip-render', is_flag=True, help='Skip rendering (use existing images)')
@click.option('--categorize-model', default='gemini-3-flash', help='Model for phase 1 (fast categorization)')
@click.option('--extract-model', default='gemini-3-flash', help='Model for phase 2 (detailed extraction)')
@click.option('--max-per-category', default=5, type=int, help='Max templates per category')
def extract_optimized(template, output_dir, skip_render, categorize_model, extract_model, max_per_category):
    """
    Extract template with optimized two-phase strategy.

    Phase 1: Fast categorization of all slides (Haiku)
    Phase 2: Detailed extraction of unique patterns (Sonnet)

    Saves 70-80% compared to full extraction.
    """
    try:
        from pptx_extractor.optimized_extractor import OptimizedExtractor, ExtractionConfig

        template_path = Path(template)
        if not template_path.exists():
            # Search in templates directory
            templates_dir = Path("pptx_templates")
            matches = list(templates_dir.rglob(f"*{template}*"))
            if matches:
                template_path = matches[0]
            else:
                console.print(f"[bold red]Error:[/bold red] Template not found: {template}")
                sys.exit(1)

        config = ExtractionConfig(
            categorize_model=categorize_model,
            extract_model=extract_model,
            output_dir=Path(output_dir),
            max_templates_per_category=max_per_category
        )

        extractor = OptimizedExtractor(config)
        result = extractor.process_template(template_path, render_slides=not skip_render)

        console.print(f"\n[bold green]Extraction complete![/bold green]")
        console.print(f"Library saved to: {result['library_path']}")

    except Exception as e:
        console.print(f"[bold red]Error:[/bold red] {e}")
        logger.exception("Optimized extraction failed")
        sys.exit(1)


@cli.command('build-library')
@click.option('--input-dir', '-i', required=True, help='Directory with extraction JSONs')
@click.option('--output', '-o', default='template_library.json', help='Output library file')
@click.option('--source-template', '-s', default=None, help='Source template name')
def build_library(input_dir, output, source_template):
    """
    Build a template library from existing extractions.

    Useful when you have already extracted slides and want to build/rebuild the library.
    """
    try:
        from pptx_extractor.template_index import TemplateIndexBuilder

        input_path = Path(input_dir)
        if not input_path.exists():
            console.print(f"[bold red]Error:[/bold red] Directory not found: {input_dir}")
            sys.exit(1)

        # Find all extraction JSONs (support various naming patterns)
        json_files = (
            list(input_path.glob("*_detailed.json")) +
            list(input_path.glob("*_description.json")) +
            list(input_path.glob("*_final.json"))
        )

        if not json_files:
            console.print("[yellow]No extraction files found[/yellow]")
            sys.exit(1)

        console.print(f"[bold]Found {len(json_files)} extraction files[/bold]\n")

        builder = TemplateIndexBuilder(output_dir=input_path)

        for json_file in json_files:
            try:
                with open(json_file, 'r') as f:
                    extraction = json.load(f)

                # Extract slide index from filename
                name = json_file.stem
                slide_idx = 0
                if '_slide_' in name:
                    try:
                        slide_idx = int(name.split('_slide_')[1].split('_')[0])
                    except (ValueError, IndexError):
                        pass

                builder.add_extraction(
                    source_file=source_template or "unknown.pptx",
                    slide_index=slide_idx,
                    extraction=extraction,
                    description_path=str(json_file)
                )
                console.print(f"  Added: {json_file.name}")

            except Exception as e:
                console.print(f"  [yellow]Skipped {json_file.name}: {e}[/yellow]")

        library_path = builder.save(output)
        library = builder.build_index()

        console.print(f"\n[bold green]Library built successfully![/bold green]")
        console.print(f"  Templates: {library.total_templates}")
        console.print(f"  Unique patterns: {library.unique_patterns}")
        console.print(f"  Categories: {list(library.categories.keys())}")
        console.print(f"  Saved to: {library_path}")

    except Exception as e:
        console.print(f"[bold red]Error:[/bold red] {e}")
        logger.exception("Library build failed")
        sys.exit(1)


@cli.command('search-templates')
@click.option('--library', '-l', default='descriptions/template_library.json', help='Library file')
@click.option('--type', '-t', 'slide_type', default=None, help='Slide type filter')
@click.option('--category', '-c', default=None, help='Template category filter')
@click.option('--style', '-s', default=None, help='Visual style filter')
@click.option('--limit', '-n', default=10, type=int, help='Max results')
def search_templates(library, slide_type, category, style, limit):
    """
    Search the template library for matching templates.
    """
    try:
        from pptx_extractor.template_index import TemplateIndexBuilder, TemplateSearcher

        library_path = Path(library)
        if not library_path.exists():
            console.print(f"[bold red]Error:[/bold red] Library not found: {library}")
            sys.exit(1)

        builder = TemplateIndexBuilder(output_dir=library_path.parent)
        lib = builder.load(library_path.name)
        searcher = TemplateSearcher(lib)

        # Parse style into components
        color_scheme = None
        visual_style = None
        if style:
            if '_' in style:
                color_scheme, visual_style = style.split('_', 1)
            else:
                visual_style = style

        results = searcher.find_templates(
            slide_type=slide_type,
            template_category=category,
            color_scheme=color_scheme,
            visual_style=visual_style,
            limit=limit
        )

        if not results:
            console.print("[yellow]No matching templates found[/yellow]")

            # Show available options
            summary = searcher.get_style_summary()
            console.print(f"\nAvailable categories: {summary['categories']}")
            console.print(f"Available styles: {summary['styles']}")
            return

        table = Table(title=f"Found {len(results)} templates")
        table.add_column("ID", style="cyan")
        table.add_column("Type", style="green")
        table.add_column("Category", style="yellow")
        table.add_column("Style", style="dim")
        table.add_column("Score", style="magenta")

        for template, score in results:
            table.add_row(
                template.id,
                template.slide_type,
                template.template_category,
                f"{template.color_scheme}/{template.visual_style}",
                f"{score:.1f}"
            )

        console.print(table)

    except Exception as e:
        console.print(f"[bold red]Error:[/bold red] {e}")
        sys.exit(1)


@cli.command('cache-stats')
def cache_stats():
    """
    Show prompt caching statistics and estimated savings.
    """
    try:
        from pptx_extractor.cached_vision import CachedVisionClient

        client = CachedVisionClient()
        stats = client.get_stats()

        console.print("\n[bold]Cache Statistics[/bold]\n")

        table = Table(show_header=False)
        table.add_column("Metric", style="cyan")
        table.add_column("Value", style="white")

        table.add_row("Cache writes", str(stats['cache_writes']))
        table.add_row("Cache reads", str(stats['cache_reads']))
        table.add_row("Cache hit rate", f"{stats['cache_hit_rate']:.1%}")
        table.add_row("Total input tokens", f"{stats['total_input_tokens']:,}")
        table.add_row("Total output tokens", f"{stats['total_output_tokens']:,}")
        table.add_row("Estimated savings", f"[green]${stats['estimated_savings_usd']:.4f}[/green]")

        console.print(table)

        console.print("\n[dim]Note: Statistics are session-based and reset on restart.[/dim]")

    except Exception as e:
        console.print(f"[bold red]Error:[/bold red] {e}")
        sys.exit(1)


if __name__ == "__main__":
    cli()
