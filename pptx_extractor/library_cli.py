"""
Library CLI - Search and browse the PowerPoint component library.

Commands:
  - stats: Show library statistics
  - search: Search for components by type, category, tags
  - info: Get detailed info about a specific component
  - categories: List categories by type
  - templates: List all templates
  - export: Export components to JSON/CSV
  - charts: List charts
  - images: List images

Extended Commands:
  - extract-all: Run unified extraction on all templates
  - styles: Browse color palettes, typography, and effects
  - layouts: Browse layout blueprints and grids
  - find-layout: Find layouts matching content requirements
  - diagram-templates: Browse diagram templates
  - text-patterns: Browse text/bullet patterns
  - sequences: Browse slide sequences
  - generate-deck: Generate deck outline from structure type
  - apply-style: Apply a style preset to a presentation
"""

import json
import click
from pathlib import Path
from rich.console import Console
from rich.table import Table
from rich.panel import Panel
from rich.tree import Tree
from rich.progress import Progress, SpinnerColumn, TextColumn
from rich import box

console = Console()

# Default library path
DEFAULT_LIBRARY_PATH = Path(__file__).parent.parent / "pptx_component_library"


def load_index(library_path: Path) -> dict:
    """Load the library index."""
    index_path = library_path / "library_index.json"
    if not index_path.exists():
        console.print(f"[yellow]Library index not found at: {index_path}[/yellow]")
        console.print("[dim]Run 'extract-all' to create the library first.[/dim]")
        return {'components': {}, 'templates': {}, 'metadata': {}}

    with open(index_path, 'r', encoding='utf-8') as f:
        return json.load(f)


def load_style_index(library_path: Path) -> dict:
    """Load the style index."""
    index_path = library_path / "styles" / "style_index.json"
    if index_path.exists():
        with open(index_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}


def load_layout_index(library_path: Path) -> dict:
    """Load the layout index."""
    index_path = library_path / "layouts" / "layout_index.json"
    if index_path.exists():
        with open(index_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}


def load_diagram_index(library_path: Path) -> dict:
    """Load the diagram template index."""
    index_path = library_path / "diagrams" / "diagram_template_index.json"
    if index_path.exists():
        with open(index_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}


def load_text_index(library_path: Path) -> dict:
    """Load the text template index."""
    index_path = library_path / "text_templates" / "text_template_index.json"
    if index_path.exists():
        with open(index_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}


def load_sequence_index(library_path: Path) -> dict:
    """Load the sequence index."""
    index_path = library_path / "sequences" / "sequence_index.json"
    if index_path.exists():
        with open(index_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}


def load_chart_style_index(library_path: Path) -> dict:
    """Load the chart style index."""
    index_path = library_path / "styles" / "chart_style_index.json"
    if index_path.exists():
        with open(index_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}


@click.group()
@click.option('--library', '-l', type=click.Path(exists=True),
              default=str(DEFAULT_LIBRARY_PATH),
              help='Path to the component library')
@click.pass_context
def cli(ctx, library):
    """PowerPoint Component Library CLI

    Search, browse, and manage your extracted PowerPoint components.
    """
    ctx.ensure_object(dict)
    ctx.obj['library_path'] = Path(library)
    ctx.obj['index'] = load_index(Path(library))


@cli.command()
@click.pass_context
def stats(ctx):
    """Show library statistics."""
    index = ctx.obj['index']

    # Header
    console.print(Panel.fit(
        "[bold blue]PowerPoint Component Library Statistics[/bold blue]",
        border_style="blue"
    ))

    # Metadata
    metadata = index.get('metadata', {})
    console.print(f"\n[dim]Created: {metadata.get('created', 'Unknown')}[/dim]")
    console.print(f"[dim]Last Updated: {metadata.get('last_updated', 'Unknown')}[/dim]")

    # Templates table
    templates = index.get('templates', {})
    if templates:
        table = Table(title="\nTemplates", box=box.ROUNDED)
        table.add_column("Template", style="cyan")
        table.add_column("Slides", justify="right")
        table.add_column("Images", justify="right")
        table.add_column("Charts", justify="right")
        table.add_column("Tables", justify="right")
        table.add_column("Diagrams", justify="right")

        for name, info in templates.items():
            comps = info.get('components', {})
            table.add_row(
                name,
                str(info.get('slide_count', 0)),
                str(comps.get('images', 0)),
                str(comps.get('charts', 0)),
                str(comps.get('tables', 0)),
                str(comps.get('diagrams', 0)),
            )

        console.print(table)

    # Component totals
    components = index.get('components', {})
    table = Table(title="\nComponent Totals", box=box.ROUNDED)
    table.add_column("Type", style="green")
    table.add_column("Count", justify="right", style="bold")

    total = 0
    for comp_type, items in components.items():
        count = len(items)
        total += count
        table.add_row(comp_type.title(), str(count))

    table.add_row("[bold]TOTAL[/bold]", f"[bold]{total}[/bold]")
    console.print(table)

    # Categories breakdown
    console.print("\n[bold]Categories by Type:[/bold]")
    for comp_type, items in components.items():
        if not items:
            continue

        categories = {}
        for item in items:
            cat = item.get('category', 'uncategorized')
            categories[cat] = categories.get(cat, 0) + 1

        if categories:
            tree = Tree(f"[cyan]{comp_type.title()}[/cyan]")
            for cat, count in sorted(categories.items(), key=lambda x: -x[1]):
                tree.add(f"{cat}: {count}")
            console.print(tree)


@cli.command()
@click.option('--type', '-t', 'comp_type',
              type=click.Choice(['images', 'charts', 'tables', 'shapes', 'diagrams', 'layouts']),
              help='Component type to search')
@click.option('--category', '-c', help='Filter by category')
@click.option('--tag', '-g', multiple=True, help='Filter by tag (can specify multiple)')
@click.option('--template', '-T', help='Filter by source template')
@click.option('--limit', '-n', default=20, help='Maximum results to show')
@click.option('--format', '-f', 'output_format',
              type=click.Choice(['table', 'json', 'brief']),
              default='table', help='Output format')
@click.argument('query', required=False)
@click.pass_context
def search(ctx, comp_type, category, tag, template, limit, output_format, query):
    """Search for components in the library.

    QUERY: Optional text to search in component metadata

    Examples:

      library search --type charts --category column_charts

      library search --type tables --template "market_analysis"

      library search -t images -n 50

      library search --tag RECTANGLE
    """
    index = ctx.obj['index']
    results = []

    # Determine which types to search
    types_to_search = [comp_type] if comp_type else list(index['components'].keys())

    for ctype in types_to_search:
        items = index['components'].get(ctype, [])

        for item in items:
            # Apply filters
            if category and item.get('category') != category:
                continue

            if tag:
                item_tags = item.get('tags', [])
                if not any(t.lower() in [it.lower() for it in item_tags] for t in tag):
                    continue

            if template:
                refs = item.get('references', [])
                if not any(template.lower() in r.get('template', '').lower() for r in refs):
                    continue

            if query:
                # Search in various fields
                searchable = json.dumps(item).lower()
                if query.lower() not in searchable:
                    continue

            item['_type'] = ctype
            results.append(item)

    # Limit results
    results = results[:limit]

    if not results:
        console.print("[yellow]No matching components found.[/yellow]")
        return

    console.print(f"\n[green]Found {len(results)} component(s)[/green]\n")

    if output_format == 'json':
        console.print_json(json.dumps(results, indent=2))

    elif output_format == 'brief':
        for item in results:
            console.print(f"[cyan]{item['_type']}[/cyan] | {item['id']} | {item.get('category', 'N/A')}")

    else:  # table
        table = Table(box=box.ROUNDED)
        table.add_column("Type", style="cyan", width=10)
        table.add_column("ID", style="dim", width=14)
        table.add_column("Category", width=18)
        table.add_column("Details", width=35)
        table.add_column("References", width=15)

        for item in results:
            # Build details string
            details = []
            if item.get('chart_type'):
                details.append(f"Chart: {item['chart_type']}")
            if item.get('rows') and item.get('cols'):
                details.append(f"Size: {item['rows']}x{item['cols']}")
            if item.get('format'):
                details.append(f"Format: {item['format']}")
            if item.get('shape_type'):
                details.append(f"Shape: {item['shape_type']}")
            if item.get('width_inches') and item.get('height_inches'):
                details.append(f"{item['width_inches']:.1f}\"x{item['height_inches']:.1f}\"")

            # References
            refs = item.get('references', [])
            ref_str = f"{len(refs)} template(s)"
            if len(refs) == 1:
                ref_str = f"{refs[0]['template'][:12]}..."

            table.add_row(
                item['_type'],
                item['id'],
                item.get('category', 'N/A'),
                ' | '.join(details[:2]) if details else 'N/A',
                ref_str
            )

        console.print(table)


@cli.command()
@click.argument('component_id')
@click.pass_context
def info(ctx, component_id):
    """Get detailed information about a component.

    COMPONENT_ID: The unique ID of the component (e.g., 5dd0375f74d5)
    """
    index = ctx.obj['index']
    library_path = ctx.obj['library_path']

    # Find the component
    found = None
    comp_type = None

    for ctype, items in index['components'].items():
        for item in items:
            if item['id'] == component_id:
                found = item
                comp_type = ctype
                break
        if found:
            break

    if not found:
        console.print(f"[red]Component not found: {component_id}[/red]")
        return

    # Display info
    console.print(Panel.fit(
        f"[bold blue]Component: {component_id}[/bold blue]",
        border_style="blue"
    ))

    table = Table(box=box.SIMPLE, show_header=False)
    table.add_column("Property", style="cyan", width=20)
    table.add_column("Value")

    table.add_row("Type", comp_type)
    table.add_row("Category", found.get('category', 'N/A'))
    table.add_row("Filename", found.get('filename', 'N/A'))

    if found.get('width_inches') and found.get('height_inches'):
        table.add_row("Dimensions", f"{found['width_inches']:.2f}\" x {found['height_inches']:.2f}\"")

    if found.get('chart_type'):
        table.add_row("Chart Type", found['chart_type'])

    if found.get('rows') and found.get('cols'):
        table.add_row("Table Size", f"{found['rows']} rows x {found['cols']} cols")

    if found.get('shape_type'):
        table.add_row("Shape Type", found['shape_type'])

    if found.get('format'):
        table.add_row("Format", found['format'])

    if found.get('size_bytes'):
        size_kb = found['size_bytes'] / 1024
        table.add_row("File Size", f"{size_kb:.1f} KB")

    if found.get('tags'):
        table.add_row("Tags", ', '.join(found['tags']))

    console.print(table)

    # References
    refs = found.get('references', [])
    if refs:
        console.print("\n[bold]References:[/bold]")
        ref_table = Table(box=box.SIMPLE)
        ref_table.add_column("Template", style="green")
        ref_table.add_column("Slide", justify="right")

        for ref in refs:
            ref_table.add_row(ref.get('template', 'Unknown'), str(ref.get('slide', 'N/A')))

        console.print(ref_table)

    # File path
    if found.get('filename'):
        file_path = library_path / comp_type / found['filename']
        console.print(f"\n[dim]File: {file_path}[/dim]")

        # If it's a JSON file, offer to show contents
        if file_path.suffix == '.json' and file_path.exists():
            console.print("\n[bold]Content Preview:[/bold]")
            with open(file_path, 'r', encoding='utf-8') as f:
                content = json.load(f)
            console.print_json(json.dumps(content, indent=2)[:1000])


@cli.command()
@click.option('--type', '-t', 'comp_type',
              type=click.Choice(['images', 'charts', 'tables', 'shapes', 'diagrams', 'layouts']),
              help='Component type')
@click.pass_context
def categories(ctx, comp_type):
    """List all categories in the library."""
    index = ctx.obj['index']

    types_to_show = [comp_type] if comp_type else list(index['components'].keys())

    for ctype in types_to_show:
        items = index['components'].get(ctype, [])
        if not items:
            continue

        # Count categories
        cat_counts = {}
        for item in items:
            cat = item.get('category', 'uncategorized')
            cat_counts[cat] = cat_counts.get(cat, 0) + 1

        table = Table(title=f"\n{ctype.title()} Categories", box=box.ROUNDED)
        table.add_column("Category", style="cyan")
        table.add_column("Count", justify="right", style="bold")

        for cat, count in sorted(cat_counts.items(), key=lambda x: -x[1]):
            table.add_row(cat, str(count))

        console.print(table)


@cli.command()
@click.pass_context
def templates(ctx):
    """List all templates in the library."""
    index = ctx.obj['index']
    templates = index.get('templates', {})

    if not templates:
        console.print("[yellow]No templates found in library.[/yellow]")
        return

    table = Table(title="Templates in Library", box=box.ROUNDED)
    table.add_column("Name", style="cyan")
    table.add_column("Slides", justify="right")
    table.add_column("Total Components", justify="right")
    table.add_column("Source", style="dim", max_width=50)

    for name, info in templates.items():
        comps = info.get('components', {})
        total = sum(comps.values())
        source = info.get('source_file', 'Unknown')

        table.add_row(
            name,
            str(info.get('slide_count', 0)),
            str(total),
            source[-50:] if len(source) > 50 else source
        )

    console.print(table)


@cli.command()
@click.option('--type', '-t', 'comp_type',
              type=click.Choice(['images', 'charts', 'tables', 'shapes', 'diagrams', 'layouts']),
              required=True, help='Component type to export')
@click.option('--category', '-c', help='Filter by category')
@click.option('--output', '-o', type=click.Path(), help='Output file path')
@click.option('--format', '-f', 'output_format',
              type=click.Choice(['json', 'csv']),
              default='json', help='Export format')
@click.pass_context
def export(ctx, comp_type, category, output, output_format):
    """Export components to JSON or CSV file.

    Examples:

      library export -t charts -o charts.json

      library export -t tables -c comparison_matrix -f csv -o matrices.csv
    """
    index = ctx.obj['index']

    items = index['components'].get(comp_type, [])

    if category:
        items = [i for i in items if i.get('category') == category]

    if not items:
        console.print("[yellow]No matching components to export.[/yellow]")
        return

    # Default output path
    if not output:
        output = f"{comp_type}_{category or 'all'}.{output_format}"

    output_path = Path(output)

    if output_format == 'json':
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(items, f, indent=2, ensure_ascii=False)

    elif output_format == 'csv':
        import csv

        # Flatten the data for CSV
        if items:
            # Get all possible keys
            all_keys = set()
            for item in items:
                all_keys.update(item.keys())

            # Remove complex nested fields
            simple_keys = [k for k in all_keys if k not in ('references', 'position', 'tags')]

            with open(output_path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.DictWriter(f, fieldnames=sorted(simple_keys))
                writer.writeheader()

                for item in items:
                    row = {k: item.get(k, '') for k in simple_keys}
                    writer.writerow(row)

    console.print(f"[green]Exported {len(items)} components to: {output_path}[/green]")


@cli.command()
@click.argument('chart_type', required=False)
@click.pass_context
def charts(ctx, chart_type):
    """List charts by type.

    CHART_TYPE: Optional filter (e.g., column_charts, bar_charts, line_charts)
    """
    index = ctx.obj['index']
    items = index['components'].get('charts', [])

    if chart_type:
        items = [i for i in items if i.get('category') == chart_type or i.get('chart_type', '').lower() == chart_type.lower()]

    if not items:
        console.print("[yellow]No matching charts found.[/yellow]")
        return

    table = Table(title=f"Charts ({len(items)} total)", box=box.ROUNDED)
    table.add_column("ID", style="dim", width=14)
    table.add_column("Chart Type", style="cyan")
    table.add_column("Category")
    table.add_column("Series", justify="right")
    table.add_column("Categories", justify="right")
    table.add_column("Templates", justify="right")

    for item in items[:50]:  # Limit display
        table.add_row(
            item['id'],
            item.get('chart_type', 'Unknown'),
            item.get('category', 'N/A'),
            str(item.get('series_count', 0)),
            str(item.get('category_count', 0)),
            str(len(item.get('references', [])))
        )

    console.print(table)

    if len(items) > 50:
        console.print(f"[dim]Showing first 50 of {len(items)} charts. Use --limit to see more.[/dim]")


@cli.command()
@click.pass_context
def images(ctx):
    """List all images in the library."""
    index = ctx.obj['index']
    items = index['components'].get('images', [])

    if not items:
        console.print("[yellow]No images found.[/yellow]")
        return

    # Group by format
    by_format = {}
    for item in items:
        fmt = item.get('format', 'unknown')
        by_format[fmt] = by_format.get(fmt, 0) + 1

    console.print(f"\n[bold]Images by Format:[/bold]")
    for fmt, count in sorted(by_format.items(), key=lambda x: -x[1]):
        console.print(f"  {fmt}: {count}")

    table = Table(title=f"\nImages ({len(items)} total)", box=box.ROUNDED)
    table.add_column("ID", style="dim", width=14)
    table.add_column("Format", style="cyan")
    table.add_column("Size", justify="right")
    table.add_column("Dimensions")
    table.add_column("Templates", justify="right")

    for item in items[:30]:
        size_kb = item.get('size_bytes', 0) / 1024
        dims = f"{item.get('width_inches', 0):.1f}\" x {item.get('height_inches', 0):.1f}\""

        table.add_row(
            item['id'],
            item.get('format', '?'),
            f"{size_kb:.1f} KB",
            dims,
            str(len(item.get('references', [])))
        )

    console.print(table)


# =============================================================================
# Extended Commands - New Feature Extractors
# =============================================================================

@cli.command('extract-all')
@click.option('--templates-dir', '-t', type=click.Path(exists=True),
              default=str(Path(__file__).parent.parent / "pptx_templates"),
              help='Directory containing PPTX templates')
@click.pass_context
def extract_all(ctx, templates_dir):
    """Run unified extraction on all templates.

    Extracts all component types:
    - Basic components (images, charts, tables, shapes, diagrams)
    - Styles (colors, typography, effects)
    - Chart styles
    - Layout blueprints
    - Diagram templates
    - Text patterns
    - Slide sequences

    Examples:

      library extract-all

      library extract-all -t /path/to/templates
    """
    from .unified_extractor import UnifiedLibraryExtractor

    library_path = ctx.obj['library_path']
    templates_path = Path(templates_dir)

    console.print(Panel.fit(
        "[bold blue]Unified Library Extraction[/bold blue]",
        border_style="blue"
    ))

    console.print(f"\n[dim]Templates: {templates_path}[/dim]")
    console.print(f"[dim]Output: {library_path}[/dim]\n")

    extractor = UnifiedLibraryExtractor(library_path)

    with Progress(
        SpinnerColumn(),
        TextColumn("[progress.description]{task.description}"),
        console=console
    ) as progress:
        task = progress.add_task("Extracting...", total=None)
        results = extractor.extract_all_templates(templates_path)
        progress.update(task, completed=True)

    # Show summary
    summary = extractor.get_full_summary()

    console.print("\n[bold green]Extraction Complete![/bold green]\n")

    table = Table(title="Extraction Summary", box=box.ROUNDED)
    table.add_column("Category", style="cyan")
    table.add_column("Count", justify="right", style="bold")

    table.add_row("Templates", str(summary.get('total_templates', 0)))

    components = summary.get('components', {})
    if isinstance(components, dict):
        for comp_type, count in components.get('total_components', {}).items():
            table.add_row(f"  {comp_type}", str(count))

    for key in ['styles', 'chart_styles', 'layouts', 'diagrams', 'text_templates', 'sequences']:
        section = summary.get(key, {})
        if isinstance(section, dict):
            total = sum(v for v in section.values() if isinstance(v, int))
            if total > 0:
                table.add_row(key.replace('_', ' ').title(), str(total))

    console.print(table)


@cli.command()
@click.option('--type', '-t', 'style_type',
              type=click.Choice(['colors', 'typography', 'gradients', 'shadows', 'effects', 'all']),
              default='all', help='Style type to show')
@click.option('--template', '-T', help='Filter by template')
@click.option('--limit', '-n', default=20, help='Maximum results')
@click.pass_context
def styles(ctx, style_type, template, limit):
    """Browse style presets (colors, typography, effects).

    Examples:

      library styles --type colors

      library styles -t typography --template market_analysis

      library styles -t gradients
    """
    library_path = ctx.obj['library_path']
    index = load_style_index(library_path)

    if not index:
        console.print("[yellow]No styles found. Run 'extract-all' first.[/yellow]")
        return

    console.print(Panel.fit(
        "[bold blue]Style Library[/bold blue]",
        border_style="blue"
    ))

    # Color palettes
    if style_type in ('colors', 'all'):
        palettes = index.get('color_palettes', [])
        if template:
            palettes = [p for p in palettes if p.get('template') == template]

        if palettes:
            table = Table(title="\nColor Palettes", box=box.ROUNDED)
            table.add_column("Template", style="cyan")
            table.add_column("Primary")
            table.add_column("Accent 1")
            table.add_column("Accent 2")
            table.add_column("Background")

            for p in palettes[:limit]:
                table.add_row(
                    p.get('template', 'Unknown'),
                    p.get('primary', 'N/A'),
                    p.get('accent1', 'N/A'),
                    p.get('accent2', 'N/A'),
                    p.get('background', 'N/A'),
                )

            console.print(table)

    # Typography
    if style_type in ('typography', 'all'):
        typography = index.get('typography_presets', [])
        if template:
            typography = [t for t in typography if t.get('template') == template]

        if typography:
            table = Table(title="\nTypography Presets", box=box.ROUNDED)
            table.add_column("Type", style="cyan")
            table.add_column("Font")
            table.add_column("Size", justify="right")
            table.add_column("Bold")
            table.add_column("Usage", justify="right")

            for t in typography[:limit]:
                table.add_row(
                    t.get('preset_type', 'N/A'),
                    t.get('font_family', 'N/A'),
                    f"{t.get('font_size_pt', 0):.0f}pt",
                    "Yes" if t.get('bold') else "No",
                    str(t.get('usage_count', 0)),
                )

            console.print(table)

    # Gradients
    if style_type in ('gradients', 'all'):
        gradients = index.get('gradient_presets', [])
        if gradients:
            table = Table(title="\nGradient Presets", box=box.ROUNDED)
            table.add_column("ID", style="dim")
            table.add_column("Type")
            table.add_column("Angle", justify="right")
            table.add_column("Stops", justify="right")
            table.add_column("Template")

            for g in gradients[:limit]:
                table.add_row(
                    g.get('id', 'N/A')[:12],
                    g.get('type', 'linear'),
                    f"{g.get('angle', 0):.0f}°",
                    str(len(g.get('stops', []))),
                    g.get('template', 'N/A'),
                )

            console.print(table)

    # Shadows
    if style_type in ('shadows', 'all'):
        shadows = index.get('shadow_presets', [])
        if shadows:
            table = Table(title="\nShadow Presets", box=box.ROUNDED)
            table.add_column("ID", style="dim")
            table.add_column("Type")
            table.add_column("Blur", justify="right")
            table.add_column("Distance", justify="right")
            table.add_column("Template")

            for s in shadows[:limit]:
                table.add_row(
                    s.get('id', 'N/A')[:12],
                    s.get('type', 'outer'),
                    f"{s.get('blur_radius', 0):.1f}",
                    f"{s.get('distance', 0):.1f}",
                    s.get('template', 'N/A'),
                )

            console.print(table)


@cli.command()
@click.option('--category', '-c', help='Filter by category')
@click.option('--template', '-T', help='Filter by template')
@click.option('--columns', type=int, help='Filter by column count')
@click.option('--rows', type=int, help='Filter by row count')
@click.option('--limit', '-n', default=20, help='Maximum results')
@click.pass_context
def layouts(ctx, category, template, columns, rows, limit):
    """Browse layout blueprints and grid systems.

    Examples:

      library layouts --category two_column

      library layouts --columns 2

      library layouts --template market_analysis
    """
    library_path = ctx.obj['library_path']
    index = load_layout_index(library_path)

    if not index:
        console.print("[yellow]No layouts found. Run 'extract-all' first.[/yellow]")
        return

    console.print(Panel.fit(
        "[bold blue]Layout Library[/bold blue]",
        border_style="blue"
    ))

    # Blueprints
    blueprints = index.get('blueprints', [])
    if category:
        blueprints = [b for b in blueprints if b.get('category') == category]
    if template:
        blueprints = [b for b in blueprints if b.get('template') == template]

    if blueprints:
        table = Table(title="\nLayout Blueprints", box=box.ROUNDED)
        table.add_column("ID", style="dim", width=12)
        table.add_column("Category", style="cyan")
        table.add_column("Elements", justify="right")
        table.add_column("Zones", justify="right")
        table.add_column("Template")

        for b in blueprints[:limit]:
            table.add_row(
                b.get('id', 'N/A')[:12],
                b.get('category', 'N/A'),
                str(b.get('element_count', 0)),
                str(len(b.get('zones', []))),
                b.get('template', 'N/A'),
            )

        console.print(table)

    # Grids
    grids = index.get('grids', [])
    if columns:
        grids = [g for g in grids if g.get('columns') == columns]
    if rows:
        grids = [g for g in grids if g.get('rows') == rows]

    if grids:
        table = Table(title="\nGrid Systems", box=box.ROUNDED)
        table.add_column("ID", style="dim", width=12)
        table.add_column("Columns", justify="right")
        table.add_column("Rows", justify="right")
        table.add_column("Margins (L/R/T/B)")
        table.add_column("Template")

        for g in grids[:limit]:
            margins = g.get('margins', {})
            margin_str = f"{margins.get('left', 0):.2f}/{margins.get('right', 0):.2f}/{margins.get('top', 0):.2f}/{margins.get('bottom', 0):.2f}"

            table.add_row(
                g.get('id', 'N/A')[:12],
                str(g.get('columns', 0)),
                str(g.get('rows', 0)),
                margin_str,
                g.get('template', 'N/A'),
            )

        console.print(table)

    # Patterns
    patterns = index.get('patterns', [])
    if patterns:
        console.print("\n[bold]Common Patterns:[/bold]")
        for p in patterns[:5]:
            console.print(f"  • {p.get('category', 'N/A')}: {p.get('occurrence_count', 0)} occurrences")


@cli.command('find-layout')
@click.option('--charts', type=int, default=0, help='Number of charts needed')
@click.option('--tables', type=int, default=0, help='Number of tables needed')
@click.option('--images', type=int, default=0, help='Number of images needed')
@click.option('--text', type=int, default=0, help='Number of text blocks needed')
@click.option('--limit', '-n', default=10, help='Maximum results')
@click.pass_context
def find_layout(ctx, charts, tables, images, text, limit):
    """Find layouts matching content requirements.

    Examples:

      library find-layout --charts 2

      library find-layout --charts 1 --tables 1

      library find-layout --text 3 --images 1
    """
    library_path = ctx.obj['library_path']
    index = load_layout_index(library_path)

    if not index:
        console.print("[yellow]No layouts found. Run 'extract-all' first.[/yellow]")
        return

    # Build requirements
    requirements = {}
    if charts > 0:
        requirements['chart'] = charts
    if tables > 0:
        requirements['table'] = tables
    if images > 0:
        requirements['image'] = images
    if text > 0:
        requirements['text_block'] = text

    if not requirements:
        console.print("[yellow]Please specify at least one content requirement.[/yellow]")
        return

    console.print(f"\n[bold]Finding layouts for:[/bold] {requirements}\n")

    # Search blueprints
    blueprints = index.get('blueprints', [])
    matches = []

    for bp in blueprints:
        content_types = bp.get('content_types', {})

        # Check if blueprint can accommodate requirements
        can_fit = True
        excess = 0
        for content_type, count in requirements.items():
            available = content_types.get(content_type, 0)
            if available < count:
                can_fit = False
                break
            excess += available - count

        if can_fit:
            matches.append((bp, excess))

    # Sort by closest match (least excess)
    matches.sort(key=lambda x: x[1])
    matches = matches[:limit]

    if not matches:
        console.print("[yellow]No matching layouts found.[/yellow]")
        return

    table = Table(title="Matching Layouts", box=box.ROUNDED)
    table.add_column("ID", style="dim", width=12)
    table.add_column("Category", style="cyan")
    table.add_column("Content Types")
    table.add_column("Fit Score", justify="right")
    table.add_column("Template")

    for bp, excess in matches:
        content = bp.get('content_types', {})
        content_str = ', '.join(f"{k}:{v}" for k, v in content.items() if v > 0)

        table.add_row(
            bp.get('id', 'N/A')[:12],
            bp.get('category', 'N/A'),
            content_str[:30],
            f"{100 - excess * 10}%",  # Simple fit score
            bp.get('template', 'N/A'),
        )

    console.print(table)


@cli.command('diagram-templates')
@click.option('--category', '-c',
              type=click.Choice(['process_flows', 'matrices', 'hierarchies', 'timelines', 'cycles', 'custom']),
              help='Diagram category')
@click.option('--template', '-T', help='Filter by template')
@click.option('--min-shapes', type=int, help='Minimum shape count')
@click.option('--limit', '-n', default=20, help='Maximum results')
@click.pass_context
def diagram_templates(ctx, category, template, min_shapes, limit):
    """Browse diagram templates (shape combinations).

    Examples:

      library diagram-templates --category process_flows

      library diagram-templates -c matrices

      library diagram-templates --min-shapes 5
    """
    library_path = ctx.obj['library_path']
    index = load_diagram_index(library_path)

    if not index:
        console.print("[yellow]No diagram templates found. Run 'extract-all' first.[/yellow]")
        return

    console.print(Panel.fit(
        "[bold blue]Diagram Template Library[/bold blue]",
        border_style="blue"
    ))

    templates = index.get('templates', [])

    if category:
        templates = [t for t in templates if t.get('category') == category]
    if template:
        templates = [t for t in templates if t.get('template') == template]
    if min_shapes:
        templates = [t for t in templates if t.get('shape_count', 0) >= min_shapes]

    if not templates:
        console.print("[yellow]No matching diagram templates found.[/yellow]")
        return

    # Show category summary
    categories = index.get('categories', {})
    if categories:
        console.print("\n[bold]Categories:[/bold]")
        for cat, ids in categories.items():
            console.print(f"  • {cat}: {len(ids)} templates")

    table = Table(title="\nDiagram Templates", box=box.ROUNDED)
    table.add_column("ID", style="dim", width=12)
    table.add_column("Category", style="cyan")
    table.add_column("Shapes", justify="right")
    table.add_column("Connectors", justify="right")
    table.add_column("Text Items", justify="right")
    table.add_column("Template")

    for t in templates[:limit]:
        table.add_row(
            t.get('id', 'N/A')[:12],
            t.get('category', 'N/A'),
            str(t.get('shape_count', 0)),
            str(t.get('connector_count', 0)),
            str(len(t.get('text_placeholders', []))),
            t.get('template', 'N/A'),
        )

    console.print(table)


@cli.command('text-patterns')
@click.option('--type', '-t', 'pattern_type',
              type=click.Choice(['bullets', 'titles', 'blocks', 'callouts', 'all']),
              default='all', help='Pattern type')
@click.option('--template', '-T', help='Filter by template')
@click.option('--limit', '-n', default=20, help='Maximum results')
@click.pass_context
def text_patterns(ctx, pattern_type, template, limit):
    """Browse text content patterns.

    Examples:

      library text-patterns --type bullets

      library text-patterns -t titles

      library text-patterns --template market_analysis
    """
    library_path = ctx.obj['library_path']
    index = load_text_index(library_path)

    if not index:
        console.print("[yellow]No text patterns found. Run 'extract-all' first.[/yellow]")
        return

    console.print(Panel.fit(
        "[bold blue]Text Pattern Library[/bold blue]",
        border_style="blue"
    ))

    # Bullet patterns
    if pattern_type in ('bullets', 'all'):
        bullets = index.get('bullet_patterns', [])
        if template:
            bullets = [b for b in bullets if b.get('template') == template]

        if bullets:
            table = Table(title="\nBullet Patterns", box=box.ROUNDED)
            table.add_column("ID", style="dim", width=12)
            table.add_column("Pattern", style="cyan")
            table.add_column("Items", justify="right")
            table.add_column("Depth", justify="right")
            table.add_column("Avg Words", justify="right")

            for b in bullets[:limit]:
                style = b.get('bullet_style', {})
                table.add_row(
                    b.get('id', 'N/A')[:12],
                    style.get('pattern', 'N/A'),
                    str(style.get('total_items', 0)),
                    str(style.get('max_depth', 0)),
                    f"{style.get('avg_words_per_item', 0):.1f}",
                )

            console.print(table)

    # Title patterns
    if pattern_type in ('titles', 'all'):
        titles = index.get('title_patterns', [])
        if template:
            titles = [t for t in titles if t.get('template') == template]

        if titles:
            table = Table(title="\nTitle Patterns", box=box.ROUNDED)
            table.add_column("ID", style="dim", width=12)
            table.add_column("Format", style="cyan")
            table.add_column("Case")
            table.add_column("Font Size", justify="right")
            table.add_column("Bold")

            for t in titles[:limit]:
                style = t.get('title_style', {})
                table.add_row(
                    t.get('id', 'N/A')[:12],
                    style.get('format', 'N/A'),
                    style.get('case', 'N/A'),
                    f"{style.get('font_size', 0) or 0:.0f}pt",
                    "Yes" if style.get('bold') else "No",
                )

            console.print(table)

    # Callouts
    if pattern_type in ('callouts', 'all'):
        callouts = index.get('callouts', [])
        if template:
            callouts = [c for c in callouts if c.get('template') == template]

        if callouts:
            table = Table(title="\nCallout Patterns", box=box.ROUNDED)
            table.add_column("ID", style="dim", width=12)
            table.add_column("Type", style="cyan")
            table.add_column("Sample Text")

            for c in callouts[:limit]:
                table.add_row(
                    c.get('id', 'N/A')[:12],
                    c.get('callout_type', 'N/A'),
                    c.get('sample_text', 'N/A')[:40],
                )

            console.print(table)


@cli.command()
@click.option('--template', '-T', help='Filter by template')
@click.option('--type', '-t', 'seq_type', help='Sequence type filter')
@click.option('--limit', '-n', default=20, help='Maximum results')
@click.pass_context
def sequences(ctx, template, seq_type, limit):
    """Browse slide sequence patterns.

    Examples:

      library sequences

      library sequences --template market_analysis

      library sequences -t section_content
    """
    library_path = ctx.obj['library_path']
    index = load_sequence_index(library_path)

    if not index:
        console.print("[yellow]No sequences found. Run 'extract-all' first.[/yellow]")
        return

    console.print(Panel.fit(
        "[bold blue]Slide Sequence Library[/bold blue]",
        border_style="blue"
    ))

    # Deck templates
    deck_templates = index.get('deck_templates', [])
    if template:
        deck_templates = [d for d in deck_templates if d.get('template') == template]

    if deck_templates:
        table = Table(title="\nDeck Templates", box=box.ROUNDED)
        table.add_column("Template", style="cyan")
        table.add_column("Slides", justify="right")
        table.add_column("Sections", justify="right")
        table.add_column("Structure Type")

        for d in deck_templates[:limit]:
            structure = d.get('structure', {})
            table.add_row(
                d.get('template', 'N/A'),
                str(d.get('slide_count', 0)),
                str(structure.get('section_count', 0)),
                structure.get('structure_type', 'N/A'),
            )

        console.print(table)

    # Sequences
    seqs = index.get('sequences', [])
    if template:
        seqs = [s for s in seqs if s.get('template') == template]
    if seq_type:
        seqs = [s for s in seqs if s.get('type') == seq_type]

    if seqs:
        table = Table(title="\nSlide Sequences", box=box.ROUNDED)
        table.add_column("ID", style="dim", width=12)
        table.add_column("Type", style="cyan")
        table.add_column("Length", justify="right")
        table.add_column("Slide Types")
        table.add_column("Template")

        for s in seqs[:limit]:
            slides = s.get('slides', [])
            slide_types = ', '.join(sl.get('type', '?') for sl in slides[:3])
            if len(slides) > 3:
                slide_types += '...'

            table.add_row(
                s.get('id', 'N/A')[:12],
                s.get('type', 'N/A'),
                str(s.get('length', 0)),
                slide_types[:30],
                s.get('template', 'N/A'),
            )

        console.print(table)

    # Common patterns
    patterns = index.get('common_patterns', [])
    if patterns:
        console.print("\n[bold]Common Patterns:[/bold]")
        for p in patterns[:5]:
            pattern_str = ' → '.join(p.get('pattern', []))
            console.print(f"  • {pattern_str} ({p.get('occurrence_count', 0)}x)")


@cli.command('generate-deck')
@click.argument('structure_type', type=click.Choice([
    'executive_presentation', 'data_heavy', 'brief', 'multi_section'
]))
@click.option('--topic', '-t', default='Presentation', help='Topic/title for the deck')
@click.option('--output', '-o', type=click.Path(), help='Output JSON file')
@click.pass_context
def generate_deck(ctx, structure_type, topic, output):
    """Generate a deck outline from a structure type.

    STRUCTURE_TYPE: Type of presentation structure

    Structure types:
      - executive_presentation: Full executive presentation with all sections
      - data_heavy: Data-focused presentation with many charts
      - brief: Short, concise presentation
      - multi_section: Multi-section detailed presentation

    Examples:

      library generate-deck executive_presentation --topic "Q4 Review"

      library generate-deck data_heavy -t "Sales Analysis" -o outline.json

      library generate-deck brief --topic "Project Update"
    """
    from .sequence_extractor import SequenceExtractor

    library_path = ctx.obj['library_path']

    extractor = SequenceExtractor(library_path)
    outline = extractor.generate_deck_outline(structure_type, topic)

    console.print(Panel.fit(
        f"[bold blue]Generated Deck Outline: {topic}[/bold blue]",
        border_style="blue"
    ))

    console.print(f"\n[dim]Structure: {structure_type}[/dim]")
    console.print(f"[dim]Slides: {outline['slide_count']}[/dim]\n")

    table = Table(title="Slide Outline", box=box.ROUNDED)
    table.add_column("#", style="dim", width=3)
    table.add_column("Type", style="cyan", width=12)
    table.add_column("Title")

    for i, slide in enumerate(outline['slides'], 1):
        table.add_row(
            str(i),
            slide.get('type', 'content'),
            slide.get('title', 'Untitled'),
        )

    console.print(table)

    if output:
        output_path = Path(output)
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(outline, f, indent=2)
        console.print(f"\n[green]Saved to: {output_path}[/green]")


@cli.command('chart-styles')
@click.option('--chart-type', '-c', help='Filter by chart type')
@click.option('--template', '-T', help='Filter by template')
@click.option('--limit', '-n', default=20, help='Maximum results')
@click.pass_context
def chart_styles(ctx, chart_type, template, limit):
    """Browse chart style presets.

    Examples:

      library chart-styles

      library chart-styles --chart-type column

      library chart-styles --template market_analysis
    """
    library_path = ctx.obj['library_path']
    index = load_chart_style_index(library_path)

    if not index:
        console.print("[yellow]No chart styles found. Run 'extract-all' first.[/yellow]")
        return

    console.print(Panel.fit(
        "[bold blue]Chart Style Library[/bold blue]",
        border_style="blue"
    ))

    styles = index.get('chart_styles', [])

    if chart_type:
        styles = [s for s in styles if chart_type.lower() in s.get('chart_type', '').lower()]
    if template:
        styles = [s for s in styles if s.get('template') == template]

    if not styles:
        console.print("[yellow]No matching chart styles found.[/yellow]")
        return

    table = Table(title="Chart Styles", box=box.ROUNDED)
    table.add_column("ID", style="dim", width=12)
    table.add_column("Chart Type", style="cyan")
    table.add_column("Colors", justify="right")
    table.add_column("Legend")
    table.add_column("Data Labels")
    table.add_column("Template")

    for s in styles[:limit]:
        colors = s.get('series_colors', [])
        color_count = len([c for c in colors if c])
        legend = s.get('legend', {})
        labels = s.get('data_labels', {})

        table.add_row(
            s.get('id', 'N/A')[:12],
            s.get('chart_type', 'N/A'),
            str(color_count),
            legend.get('position', 'none') if legend.get('visible') else 'none',
            'Yes' if labels.get('visible') else 'No',
            s.get('template', 'N/A'),
        )

    console.print(table)


@cli.command('full-stats')
@click.pass_context
def full_stats(ctx):
    """Show comprehensive statistics for all library components."""
    library_path = ctx.obj['library_path']

    console.print(Panel.fit(
        "[bold blue]Full Library Statistics[/bold blue]",
        border_style="blue"
    ))

    # Basic components
    index = ctx.obj['index']
    if index.get('components'):
        table = Table(title="\nBasic Components", box=box.ROUNDED)
        table.add_column("Type", style="cyan")
        table.add_column("Count", justify="right", style="bold")

        total = 0
        for comp_type, items in index.get('components', {}).items():
            count = len(items)
            total += count
            table.add_row(comp_type.title(), str(count))

        table.add_row("[bold]Total[/bold]", f"[bold]{total}[/bold]")
        console.print(table)

    # Styles
    style_index = load_style_index(library_path)
    if style_index:
        table = Table(title="\nStyles", box=box.ROUNDED)
        table.add_column("Type", style="cyan")
        table.add_column("Count", justify="right", style="bold")

        for key in ['color_palettes', 'typography_presets', 'gradient_presets', 'shadow_presets', 'effect_presets']:
            count = len(style_index.get(key, []))
            if count > 0:
                table.add_row(key.replace('_', ' ').title(), str(count))

        console.print(table)

    # Chart styles
    chart_index = load_chart_style_index(library_path)
    if chart_index.get('chart_styles'):
        console.print(f"\n[cyan]Chart Styles:[/cyan] {len(chart_index['chart_styles'])}")

    # Layouts
    layout_index = load_layout_index(library_path)
    if layout_index:
        console.print(f"[cyan]Layout Blueprints:[/cyan] {len(layout_index.get('blueprints', []))}")
        console.print(f"[cyan]Grid Systems:[/cyan] {len(layout_index.get('grids', []))}")

    # Diagrams
    diagram_index = load_diagram_index(library_path)
    if diagram_index.get('templates'):
        console.print(f"[cyan]Diagram Templates:[/cyan] {len(diagram_index['templates'])}")

    # Text patterns
    text_index = load_text_index(library_path)
    if text_index:
        total_text = sum(len(text_index.get(k, [])) for k in ['bullet_patterns', 'title_patterns', 'text_blocks', 'callouts'])
        console.print(f"[cyan]Text Patterns:[/cyan] {total_text}")

    # Sequences
    seq_index = load_sequence_index(library_path)
    if seq_index:
        console.print(f"[cyan]Deck Templates:[/cyan] {len(seq_index.get('deck_templates', []))}")
        console.print(f"[cyan]Sequences:[/cyan] {len(seq_index.get('sequences', []))}")


if __name__ == '__main__':
    cli()
