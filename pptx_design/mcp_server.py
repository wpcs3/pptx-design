"""
MCP Server for PPTX Design System

Exposes PresentationBuilder as an MCP (Model Context Protocol) server
for Claude and other AI agents to create professional presentations.

Phase 3 Enhancement (2025-12-29):
- 15 MCP tools for presentation creation and manipulation
- State management for active presentations
- Template listing and selection
- Text extraction for verification loops

Usage:
    # Run as MCP server (stdio transport)
    python -m pptx_design.mcp_server

    # Or use with uv
    uv run python -m pptx_design.mcp_server

Requirements:
    pip install "mcp[cli]"
"""

import logging
import uuid
from pathlib import Path
from typing import Any, Dict, List, Optional

from mcp.server.fastmcp import FastMCP

# Configure logging (never use print() in MCP servers)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    handlers=[logging.FileHandler("pptx_mcp_server.log")]
)
logger = logging.getLogger(__name__)

# Initialize FastMCP server
mcp = FastMCP("pptx-design")

# Global state for active presentations
_active_presentations: Dict[str, Any] = {}


# =============================================================================
# Helper Functions
# =============================================================================

def _get_builder(presentation_id: str):
    """Get a PresentationBuilder by ID."""
    if presentation_id not in _active_presentations:
        raise ValueError(f"Presentation not found: {presentation_id}")
    return _active_presentations[presentation_id]


def _generate_id() -> str:
    """Generate a unique presentation ID."""
    return str(uuid.uuid4())[:8]


def _ensure_imports():
    """Lazily import PresentationBuilder to avoid circular imports."""
    from pptx_design import PresentationBuilder, TemplateRegistry
    return PresentationBuilder, TemplateRegistry


# =============================================================================
# Presentation Management Tools
# =============================================================================

@mcp.tool()
def create_presentation(template: str = "consulting_toolkit", title: str = "") -> Dict[str, Any]:
    """
    Create a new presentation with the specified template.

    Args:
        template: Template name to use. Options: consulting_toolkit, business_case,
                 market_analysis, default. Default is 'consulting_toolkit'.
        title: Optional title for the presentation (creates title slide if provided).

    Returns:
        Dictionary with presentation_id and status.
    """
    PresentationBuilder, _ = _ensure_imports()

    try:
        presentation_id = _generate_id()
        builder = PresentationBuilder(template)

        if title:
            builder.add_title_slide(title)

        _active_presentations[presentation_id] = builder
        logger.info(f"Created presentation {presentation_id} with template {template}")

        return {
            "status": "created",
            "presentation_id": presentation_id,
            "template": template,
            "slide_count": len(builder._slides)
        }
    except Exception as e:
        logger.error(f"Error creating presentation: {e}")
        return {"status": "error", "message": str(e)}


@mcp.tool()
def list_presentations() -> Dict[str, Any]:
    """
    List all active presentations in the current session.

    Returns:
        Dictionary with list of active presentation IDs and their slide counts.
    """
    presentations = []
    for pid, builder in _active_presentations.items():
        presentations.append({
            "presentation_id": pid,
            "slide_count": len(builder._slides)
        })

    return {
        "status": "success",
        "count": len(presentations),
        "presentations": presentations
    }


@mcp.tool()
def close_presentation(presentation_id: str) -> Dict[str, Any]:
    """
    Close and remove a presentation from active sessions.

    Args:
        presentation_id: The ID of the presentation to close.

    Returns:
        Dictionary with status.
    """
    if presentation_id in _active_presentations:
        del _active_presentations[presentation_id]
        logger.info(f"Closed presentation {presentation_id}")
        return {"status": "closed", "presentation_id": presentation_id}
    else:
        return {"status": "error", "message": f"Presentation not found: {presentation_id}"}


# =============================================================================
# Slide Addition Tools
# =============================================================================

@mcp.tool()
def add_title_slide(
    presentation_id: str,
    title: str,
    subtitle: str = ""
) -> Dict[str, Any]:
    """
    Add a title slide to the presentation.

    Args:
        presentation_id: The ID of the presentation.
        title: Main title text.
        subtitle: Optional subtitle text.

    Returns:
        Dictionary with status and slide index.
    """
    try:
        builder = _get_builder(presentation_id)
        builder.add_title_slide(title, subtitle)

        slide_index = len(builder._slides) - 1
        logger.info(f"Added title slide to {presentation_id}")

        return {
            "status": "added",
            "slide_type": "title_slide",
            "slide_index": slide_index,
            "slide_count": len(builder._slides)
        }
    except Exception as e:
        logger.error(f"Error adding title slide: {e}")
        return {"status": "error", "message": str(e)}


@mcp.tool()
def add_content_slide(
    presentation_id: str,
    title: str,
    bullets: List[str]
) -> Dict[str, Any]:
    """
    Add a content slide with title and bullet points.

    Args:
        presentation_id: The ID of the presentation.
        title: Slide title.
        bullets: List of bullet point strings.

    Returns:
        Dictionary with status and slide index.
    """
    try:
        builder = _get_builder(presentation_id)
        builder.add_content_slide(title, bullets=bullets)

        slide_index = len(builder._slides) - 1
        logger.info(f"Added content slide to {presentation_id}")

        return {
            "status": "added",
            "slide_type": "title_content",
            "slide_index": slide_index,
            "slide_count": len(builder._slides)
        }
    except Exception as e:
        logger.error(f"Error adding content slide: {e}")
        return {"status": "error", "message": str(e)}


@mcp.tool()
def add_section_slide(
    presentation_id: str,
    title: str,
    subtitle: str = ""
) -> Dict[str, Any]:
    """
    Add a section divider slide.

    Args:
        presentation_id: The ID of the presentation.
        title: Section title.
        subtitle: Optional section subtitle.

    Returns:
        Dictionary with status and slide index.
    """
    try:
        builder = _get_builder(presentation_id)
        builder.add_section_divider(title, subtitle)

        slide_index = len(builder._slides) - 1
        logger.info(f"Added section slide to {presentation_id}")

        return {
            "status": "added",
            "slide_type": "section_divider",
            "slide_index": slide_index,
            "slide_count": len(builder._slides)
        }
    except Exception as e:
        logger.error(f"Error adding section slide: {e}")
        return {"status": "error", "message": str(e)}


@mcp.tool()
def add_agenda_slide(
    presentation_id: str,
    items: List[str],
    title: str = "Agenda"
) -> Dict[str, Any]:
    """
    Add an agenda slide with list of topics.

    Args:
        presentation_id: The ID of the presentation.
        items: List of agenda items.
        title: Title for the agenda slide (default: "Agenda").

    Returns:
        Dictionary with status and slide index.
    """
    try:
        builder = _get_builder(presentation_id)
        builder.add_agenda(items, title=title)

        slide_index = len(builder._slides) - 1
        logger.info(f"Added agenda slide to {presentation_id}")

        return {
            "status": "added",
            "slide_type": "agenda",
            "slide_index": slide_index,
            "slide_count": len(builder._slides)
        }
    except Exception as e:
        logger.error(f"Error adding agenda slide: {e}")
        return {"status": "error", "message": str(e)}


@mcp.tool()
def add_two_column_slide(
    presentation_id: str,
    title: str,
    left_content: Dict[str, Any],
    right_content: Dict[str, Any]
) -> Dict[str, Any]:
    """
    Add a two-column comparison slide.

    Args:
        presentation_id: The ID of the presentation.
        title: Slide title.
        left_content: Dict with 'header' and 'bullets' for left column.
        right_content: Dict with 'header' and 'bullets' for right column.

    Returns:
        Dictionary with status and slide index.
    """
    try:
        builder = _get_builder(presentation_id)
        builder.add_two_column(title, left_content, right_content)

        slide_index = len(builder._slides) - 1
        logger.info(f"Added two-column slide to {presentation_id}")

        return {
            "status": "added",
            "slide_type": "two_column",
            "slide_index": slide_index,
            "slide_count": len(builder._slides)
        }
    except Exception as e:
        logger.error(f"Error adding two-column slide: {e}")
        return {"status": "error", "message": str(e)}


@mcp.tool()
def add_metrics_slide(
    presentation_id: str,
    title: str,
    metrics: List[Dict[str, str]]
) -> Dict[str, Any]:
    """
    Add a key metrics slide with KPI boxes.

    Args:
        presentation_id: The ID of the presentation.
        title: Slide title.
        metrics: List of dicts with 'label' and 'value' keys.
                 Example: [{"label": "Revenue", "value": "$1.2M"}, {"label": "Growth", "value": "25%"}]

    Returns:
        Dictionary with status and slide index.
    """
    try:
        builder = _get_builder(presentation_id)
        builder.add_metrics(title, metrics)

        slide_index = len(builder._slides) - 1
        logger.info(f"Added metrics slide to {presentation_id}")

        return {
            "status": "added",
            "slide_type": "key_metrics",
            "slide_index": slide_index,
            "slide_count": len(builder._slides)
        }
    except Exception as e:
        logger.error(f"Error adding metrics slide: {e}")
        return {"status": "error", "message": str(e)}


@mcp.tool()
def add_chart_slide(
    presentation_id: str,
    title: str,
    chart_type: str,
    categories: List[str],
    series: List[Dict[str, Any]]
) -> Dict[str, Any]:
    """
    Add a slide with a data chart.

    Args:
        presentation_id: The ID of the presentation.
        title: Slide title.
        chart_type: Type of chart (column, bar, line, pie).
        categories: List of category labels (x-axis).
        series: List of series dicts with 'name' and 'values' keys.
                Example: [{"name": "Revenue", "values": [100, 120, 150]}]

    Returns:
        Dictionary with status and slide index.
    """
    try:
        builder = _get_builder(presentation_id)

        chart_data = {
            "type": chart_type,
            "categories": categories,
            "series": series
        }
        builder.add_chart(title, chart_data)

        slide_index = len(builder._slides) - 1
        logger.info(f"Added chart slide to {presentation_id}")

        return {
            "status": "added",
            "slide_type": "data_chart",
            "slide_index": slide_index,
            "slide_count": len(builder._slides)
        }
    except Exception as e:
        logger.error(f"Error adding chart slide: {e}")
        return {"status": "error", "message": str(e)}


@mcp.tool()
def add_table_slide(
    presentation_id: str,
    title: str,
    headers: List[str],
    rows: List[List[str]]
) -> Dict[str, Any]:
    """
    Add a slide with a data table.

    Args:
        presentation_id: The ID of the presentation.
        title: Slide title.
        headers: List of column header strings.
        rows: List of row data (each row is a list of cell values).

    Returns:
        Dictionary with status and slide index.
    """
    try:
        builder = _get_builder(presentation_id)

        table_data = {
            "headers": headers,
            "rows": rows
        }
        builder.add_table(title, table_data)

        slide_index = len(builder._slides) - 1
        logger.info(f"Added table slide to {presentation_id}")

        return {
            "status": "added",
            "slide_type": "table_slide",
            "slide_index": slide_index,
            "slide_count": len(builder._slides)
        }
    except Exception as e:
        logger.error(f"Error adding table slide: {e}")
        return {"status": "error", "message": str(e)}


# =============================================================================
# Export and Extraction Tools
# =============================================================================

@mcp.tool()
def save_presentation(
    presentation_id: str,
    output_path: str
) -> Dict[str, Any]:
    """
    Save the presentation to a file.

    Args:
        presentation_id: The ID of the presentation.
        output_path: Path where the .pptx file should be saved.

    Returns:
        Dictionary with status and file path.
    """
    try:
        builder = _get_builder(presentation_id)

        # Ensure .pptx extension
        if not output_path.endswith('.pptx'):
            output_path += '.pptx'

        builder.save(output_path)
        logger.info(f"Saved presentation {presentation_id} to {output_path}")

        return {
            "status": "saved",
            "presentation_id": presentation_id,
            "path": output_path,
            "slide_count": len(builder._slides)
        }
    except Exception as e:
        logger.error(f"Error saving presentation: {e}")
        return {"status": "error", "message": str(e)}


@mcp.tool()
def extract_slide_text(
    presentation_id: str,
    slide_index: int
) -> Dict[str, Any]:
    """
    Extract text content from a specific slide for verification.

    Args:
        presentation_id: The ID of the presentation.
        slide_index: Zero-based index of the slide.

    Returns:
        Dictionary with slide text content.
    """
    try:
        builder = _get_builder(presentation_id)

        if slide_index < 0 or slide_index >= len(builder._slides):
            return {
                "status": "error",
                "message": f"Invalid slide index: {slide_index}. "
                          f"Presentation has {len(builder._slides)} slides."
            }

        # Extract text from slide spec (before building)
        slide_spec = builder._slides[slide_index]
        content = slide_spec.get("content", {})
        text_content = []

        # Extract text from content fields
        for key in ["title", "subtitle", "body"]:
            if key in content:
                val = content[key]
                if isinstance(val, str) and val:
                    text_content.append(val)

        # Extract bullets
        if "bullets" in content:
            text_content.extend(content["bullets"])

        return {
            "status": "success",
            "slide_index": slide_index,
            "text_count": len(text_content),
            "text": text_content
        }
    except Exception as e:
        logger.error(f"Error extracting slide text: {e}")
        return {"status": "error", "message": str(e)}


@mcp.tool()
def get_slide_count(presentation_id: str) -> Dict[str, Any]:
    """
    Get the number of slides in a presentation.

    Args:
        presentation_id: The ID of the presentation.

    Returns:
        Dictionary with slide count.
    """
    try:
        builder = _get_builder(presentation_id)
        return {
            "status": "success",
            "presentation_id": presentation_id,
            "slide_count": len(builder._slides)
        }
    except Exception as e:
        logger.error(f"Error getting slide count: {e}")
        return {"status": "error", "message": str(e)}


# =============================================================================
# Template Tools
# =============================================================================

@mcp.tool()
def list_templates() -> Dict[str, Any]:
    """
    List all available presentation templates.

    Returns:
        Dictionary with available templates and their descriptions.
    """
    _, TemplateRegistry = _ensure_imports()

    try:
        registry = TemplateRegistry()
        templates = []

        for template_id, template in registry.templates.items():
            templates.append({
                "id": template_id,
                "name": template.metadata.get("name", template_id),
                "description": template.metadata.get("description", ""),
                "layout_count": len(template.layouts),
                "recommended_for": template.metadata.get("recommended_for", [])
            })

        return {
            "status": "success",
            "count": len(templates),
            "templates": templates
        }
    except Exception as e:
        logger.error(f"Error listing templates: {e}")
        return {"status": "error", "message": str(e)}


@mcp.tool()
def get_template_layouts(template: str) -> Dict[str, Any]:
    """
    Get available layouts for a specific template.

    Args:
        template: Template name (e.g., 'consulting_toolkit').

    Returns:
        Dictionary with available layout types.
    """
    _, TemplateRegistry = _ensure_imports()

    try:
        registry = TemplateRegistry()

        if template not in registry.templates:
            return {
                "status": "error",
                "message": f"Template not found: {template}. "
                          f"Available: {list(registry.templates.keys())}"
            }

        template_obj = registry.templates[template]
        layouts = list(template_obj.layouts.keys())

        return {
            "status": "success",
            "template": template,
            "layout_count": len(layouts),
            "layouts": layouts
        }
    except Exception as e:
        logger.error(f"Error getting template layouts: {e}")
        return {"status": "error", "message": str(e)}


# =============================================================================
# Server Entry Point
# =============================================================================

def main():
    """Run the MCP server."""
    logger.info("Starting PPTX Design MCP Server")
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
