"""
Theme Extraction Module

Extracts color themes, font schemes, and effects from PPTX files.
"""
import json
import logging
from pathlib import Path
from typing import Optional
from xml.etree import ElementTree as ET

from pptx import Presentation
from pptx.oxml.ns import qn

# Add parent directory to path for config import
import sys
sys.path.insert(0, str(Path(__file__).parent.parent))

from config import DESCRIPTION_DIR

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class ThemeError(Exception):
    """Exception raised when theme extraction fails."""
    pass


# XML namespaces used in Office Open XML
NAMESPACES = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
}


def rgb_from_element(elem) -> Optional[str]:
    """
    Extract RGB hex color from an XML element.

    Args:
        elem: XML element that may contain color info

    Returns:
        Hex color string or None
    """
    if elem is None:
        return None

    # Direct srgbClr
    srgb = elem.find('.//a:srgbClr', NAMESPACES)
    if srgb is not None:
        return f"#{srgb.get('val', '000000')}"

    # System color
    sys_clr = elem.find('.//a:sysClr', NAMESPACES)
    if sys_clr is not None:
        last_clr = sys_clr.get('lastClr')
        if last_clr:
            return f"#{last_clr}"

    return None


def extract_color_scheme(theme_part) -> dict:
    """
    Extract the color scheme from a theme.

    Args:
        theme_part: The theme XML part

    Returns:
        Dict with color scheme information
    """
    colors = {}

    try:
        tree = ET.parse(theme_part.blob)
        root = tree.getroot()
    except Exception:
        # Try parsing as bytes
        try:
            root = ET.fromstring(theme_part.blob)
        except Exception as e:
            logger.warning(f"Could not parse theme XML: {e}")
            return colors

    # Find the color scheme element
    clr_scheme = root.find('.//a:clrScheme', NAMESPACES)
    if clr_scheme is None:
        return colors

    # Color mapping
    color_names = [
        ('dk1', 'dark1'),
        ('lt1', 'light1'),
        ('dk2', 'dark2'),
        ('lt2', 'light2'),
        ('accent1', 'accent1'),
        ('accent2', 'accent2'),
        ('accent3', 'accent3'),
        ('accent4', 'accent4'),
        ('accent5', 'accent5'),
        ('accent6', 'accent6'),
        ('hlink', 'hyperlink'),
        ('folHlink', 'followed_hyperlink'),
    ]

    for xml_name, friendly_name in color_names:
        elem = clr_scheme.find(f'a:{xml_name}', NAMESPACES)
        if elem is not None:
            color = rgb_from_element(elem)
            if color:
                colors[friendly_name] = color

    return colors


def extract_font_scheme(theme_part) -> dict:
    """
    Extract the font scheme from a theme.

    Args:
        theme_part: The theme XML part

    Returns:
        Dict with font scheme information
    """
    fonts = {
        'major': {},
        'minor': {}
    }

    try:
        root = ET.fromstring(theme_part.blob)
    except Exception as e:
        logger.warning(f"Could not parse theme XML: {e}")
        return fonts

    # Find the font scheme element
    font_scheme = root.find('.//a:fontScheme', NAMESPACES)
    if font_scheme is None:
        return fonts

    # Major fonts (headings)
    major_font = font_scheme.find('a:majorFont', NAMESPACES)
    if major_font is not None:
        latin = major_font.find('a:latin', NAMESPACES)
        if latin is not None:
            fonts['major']['latin'] = latin.get('typeface')
        ea = major_font.find('a:ea', NAMESPACES)
        if ea is not None:
            fonts['major']['east_asian'] = ea.get('typeface')
        cs = major_font.find('a:cs', NAMESPACES)
        if cs is not None:
            fonts['major']['complex_script'] = cs.get('typeface')

    # Minor fonts (body)
    minor_font = font_scheme.find('a:minorFont', NAMESPACES)
    if minor_font is not None:
        latin = minor_font.find('a:latin', NAMESPACES)
        if latin is not None:
            fonts['minor']['latin'] = latin.get('typeface')
        ea = minor_font.find('a:ea', NAMESPACES)
        if ea is not None:
            fonts['minor']['east_asian'] = ea.get('typeface')
        cs = minor_font.find('a:cs', NAMESPACES)
        if cs is not None:
            fonts['minor']['complex_script'] = cs.get('typeface')

    return fonts


def extract_format_scheme(theme_part) -> dict:
    """
    Extract format scheme (fill, line, effect styles) from a theme.

    Args:
        theme_part: The theme XML part

    Returns:
        Dict with format scheme information
    """
    formats = {
        'fill_styles': [],
        'line_styles': [],
        'effect_styles': [],
        'background_fill_styles': []
    }

    try:
        root = ET.fromstring(theme_part.blob)
    except Exception as e:
        logger.warning(f"Could not parse theme XML: {e}")
        return formats

    # Find format scheme
    fmt_scheme = root.find('.//a:fmtScheme', NAMESPACES)
    if fmt_scheme is None:
        return formats

    # Fill styles
    fill_style_lst = fmt_scheme.find('a:fillStyleLst', NAMESPACES)
    if fill_style_lst is not None:
        for i, fill in enumerate(fill_style_lst):
            style = {'index': i + 1, 'type': fill.tag.split('}')[-1]}
            color = rgb_from_element(fill)
            if color:
                style['color'] = color
            formats['fill_styles'].append(style)

    # Line styles
    ln_style_lst = fmt_scheme.find('a:lnStyleLst', NAMESPACES)
    if ln_style_lst is not None:
        for i, ln in enumerate(ln_style_lst):
            style = {'index': i + 1}
            width = ln.get('w')
            if width:
                style['width_emu'] = int(width)
                style['width_pt'] = int(width) / 12700  # EMU to points
            formats['line_styles'].append(style)

    # Background fill styles
    bg_fill_lst = fmt_scheme.find('a:bgFillStyleLst', NAMESPACES)
    if bg_fill_lst is not None:
        for i, fill in enumerate(bg_fill_lst):
            style = {'index': i + 1, 'type': fill.tag.split('}')[-1]}
            formats['background_fill_styles'].append(style)

    return formats


def extract_theme(pptx_path: Path) -> dict:
    """
    Extract complete theme information from a PPTX file.

    Args:
        pptx_path: Path to the PPTX file

    Returns:
        Dict with complete theme information
    """
    import zipfile

    pptx_path = Path(pptx_path)
    if not pptx_path.exists():
        raise ThemeError(f"PPTX file not found: {pptx_path}")

    result = {
        'source_file': pptx_path.name,
        'themes': []
    }

    # Extract themes directly from the PPTX zip archive
    try:
        with zipfile.ZipFile(pptx_path, 'r') as zf:
            theme_files = sorted([f for f in zf.namelist() if 'theme' in f.lower() and f.endswith('.xml')])

            for i, theme_file in enumerate(theme_files):
                theme_info = {
                    'master_index': i,
                    'master_name': f"Theme {i + 1}",
                    'source_file': theme_file,
                    'colors': {},
                    'fonts': {},
                    'formats': {}
                }

                try:
                    with zf.open(theme_file) as f:
                        xml_content = f.read()

                    # Parse the XML
                    root = ET.fromstring(xml_content)

                    # Extract color scheme
                    theme_info['colors'] = _extract_colors_from_xml(root)

                    # Extract font scheme
                    theme_info['fonts'] = _extract_fonts_from_xml(root)

                except Exception as e:
                    logger.warning(f"Could not parse theme {theme_file}: {e}")

                result['themes'].append(theme_info)

    except zipfile.BadZipFile:
        raise ThemeError(f"Invalid PPTX file: {pptx_path}")

    logger.info(f"Extracted {len(result['themes'])} theme(s) from {pptx_path.name}")

    return result


def _extract_colors_from_xml(root) -> dict:
    """Extract color scheme from parsed theme XML."""
    colors = {}

    # Find the color scheme element
    clr_scheme = root.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}clrScheme')
    if clr_scheme is None:
        return colors

    # Color mapping
    color_names = [
        ('dk1', 'dark1'),
        ('lt1', 'light1'),
        ('dk2', 'dark2'),
        ('lt2', 'light2'),
        ('accent1', 'accent1'),
        ('accent2', 'accent2'),
        ('accent3', 'accent3'),
        ('accent4', 'accent4'),
        ('accent5', 'accent5'),
        ('accent6', 'accent6'),
        ('hlink', 'hyperlink'),
        ('folHlink', 'followed_hyperlink'),
    ]

    ns = '{http://schemas.openxmlformats.org/drawingml/2006/main}'

    for xml_name, friendly_name in color_names:
        elem = clr_scheme.find(f'{ns}{xml_name}')
        if elem is not None:
            # Try srgbClr first
            srgb = elem.find(f'{ns}srgbClr')
            if srgb is not None:
                colors[friendly_name] = f"#{srgb.get('val', '000000')}"
                continue

            # Try sysClr
            sys_clr = elem.find(f'{ns}sysClr')
            if sys_clr is not None:
                last_clr = sys_clr.get('lastClr')
                if last_clr:
                    colors[friendly_name] = f"#{last_clr}"

    return colors


def _extract_fonts_from_xml(root) -> dict:
    """Extract font scheme from parsed theme XML."""
    fonts = {'major': {}, 'minor': {}}

    ns = '{http://schemas.openxmlformats.org/drawingml/2006/main}'

    # Find the font scheme element
    font_scheme = root.find(f'.//{ns}fontScheme')
    if font_scheme is None:
        return fonts

    # Major fonts (headings)
    major_font = font_scheme.find(f'{ns}majorFont')
    if major_font is not None:
        latin = major_font.find(f'{ns}latin')
        if latin is not None:
            fonts['major']['latin'] = latin.get('typeface')

    # Minor fonts (body)
    minor_font = font_scheme.find(f'{ns}minorFont')
    if minor_font is not None:
        latin = minor_font.find(f'{ns}latin')
        if latin is not None:
            fonts['minor']['latin'] = latin.get('typeface')

    return fonts


def extract_color_palette(pptx_path: Path) -> list[str]:
    """
    Extract just the color palette from a PPTX.

    Args:
        pptx_path: Path to the PPTX file

    Returns:
        List of hex color strings
    """
    theme_info = extract_theme(pptx_path)

    colors = []
    for theme in theme_info.get('themes', []):
        for color in theme.get('colors', {}).values():
            if color and color not in colors:
                colors.append(color)

    return colors


def extract_font_families(pptx_path: Path) -> dict:
    """
    Extract font families from a PPTX.

    Args:
        pptx_path: Path to the PPTX file

    Returns:
        Dict with 'heading' and 'body' font names
    """
    theme_info = extract_theme(pptx_path)

    fonts = {
        'heading': None,
        'body': None
    }

    for theme in theme_info.get('themes', []):
        theme_fonts = theme.get('fonts', {})
        if theme_fonts.get('major', {}).get('latin'):
            fonts['heading'] = theme_fonts['major']['latin']
        if theme_fonts.get('minor', {}).get('latin'):
            fonts['body'] = theme_fonts['minor']['latin']
        if fonts['heading'] and fonts['body']:
            break

    return fonts


def save_theme_info(
    theme_info: dict,
    template_name: str,
    output_dir: Optional[Path] = None
) -> Path:
    """
    Save extracted theme information to JSON.

    Args:
        theme_info: Dict from extract_theme()
        template_name: Name for the output file
        output_dir: Output directory (defaults to DESCRIPTION_DIR)

    Returns:
        Path to saved file
    """
    if output_dir is None:
        output_dir = DESCRIPTION_DIR
    else:
        output_dir = Path(output_dir)

    output_dir.mkdir(parents=True, exist_ok=True)

    output_path = output_dir / f"{template_name}_theme.json"

    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(theme_info, f, indent=2, ensure_ascii=False)

    logger.info(f"Theme info saved: {output_path}")
    return output_path


def generate_theme_css(theme_info: dict) -> str:
    """
    Generate CSS variables from theme colors.

    Args:
        theme_info: Dict from extract_theme()

    Returns:
        CSS string with color variables
    """
    css_lines = [":root {"]

    for theme in theme_info.get('themes', []):
        for name, color in theme.get('colors', {}).items():
            if color:
                css_name = name.replace('_', '-')
                css_lines.append(f"  --theme-{css_name}: {color};")

    css_lines.append("}")

    return "\n".join(css_lines)


def generate_theme_summary(theme_info: dict) -> str:
    """
    Generate a human-readable summary of the theme.

    Args:
        theme_info: Dict from extract_theme()

    Returns:
        Markdown formatted summary
    """
    lines = [f"# Theme Summary: {theme_info.get('source_file', 'Unknown')}\n"]

    for i, theme in enumerate(theme_info.get('themes', [])):
        lines.append(f"## Theme {i + 1}: {theme.get('master_name', 'Unknown')}\n")

        # Colors
        colors = theme.get('colors', {})
        if colors:
            lines.append("### Color Scheme\n")
            lines.append("| Role | Color |")
            lines.append("|------|-------|")
            for name, color in colors.items():
                lines.append(f"| {name} | {color} |")
            lines.append("")

        # Fonts
        fonts = theme.get('fonts', {})
        if fonts:
            lines.append("### Font Scheme\n")
            major = fonts.get('major', {})
            minor = fonts.get('minor', {})
            if major.get('latin'):
                lines.append(f"- **Headings (Major):** {major['latin']}")
            if minor.get('latin'):
                lines.append(f"- **Body (Minor):** {minor['latin']}")
            lines.append("")

    return "\n".join(lines)


if __name__ == "__main__":
    import sys
    from rich.console import Console
    from rich.table import Table
    from rich.panel import Panel

    console = Console()

    if len(sys.argv) < 2:
        console.print("Usage: python -m src.themes <pptx_file>")
        console.print("\nExtracts theme (colors, fonts, effects) from a PPTX file.")
        sys.exit(1)

    pptx_path = Path(sys.argv[1])

    console.print(f"\n[bold]Extracting theme from:[/bold] {pptx_path.name}\n")

    try:
        theme_info = extract_theme(pptx_path)

        for i, theme in enumerate(theme_info.get('themes', [])):
            console.print(Panel(
                f"[bold]{theme.get('master_name', 'Unknown')}[/bold]",
                title=f"Theme {i + 1}"
            ))

            # Color table
            colors = theme.get('colors', {})
            if colors:
                color_table = Table(title="Color Scheme")
                color_table.add_column("Role", style="cyan")
                color_table.add_column("Color", style="green")
                color_table.add_column("Sample", style="dim")

                for name, color in colors.items():
                    # Create a colored block for sample
                    sample = f"[on {color}]      [/]" if color.startswith('#') else ""
                    color_table.add_row(name, color or "N/A", sample)

                console.print(color_table)

            # Font table
            fonts = theme.get('fonts', {})
            if fonts:
                font_table = Table(title="Font Scheme")
                font_table.add_column("Type", style="cyan")
                font_table.add_column("Font Family", style="green")

                major = fonts.get('major', {})
                minor = fonts.get('minor', {})

                if major.get('latin'):
                    font_table.add_row("Headings (Major)", major['latin'])
                if minor.get('latin'):
                    font_table.add_row("Body (Minor)", minor['latin'])

                console.print(font_table)

            console.print()

        # Save to file
        template_name = pptx_path.stem
        output_path = save_theme_info(theme_info, template_name)
        console.print(f"[bold green]Theme info saved:[/bold green] {output_path}")

        # Also show extracted palette
        palette = extract_color_palette(pptx_path)
        if palette:
            console.print(f"\n[bold]Color Palette:[/bold] {', '.join(palette)}")

        # Show fonts
        fonts = extract_font_families(pptx_path)
        console.print(f"[bold]Heading Font:[/bold] {fonts.get('heading', 'N/A')}")
        console.print(f"[bold]Body Font:[/bold] {fonts.get('body', 'N/A')}")

    except Exception as e:
        console.print(f"[bold red]Error:[/bold red] {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
