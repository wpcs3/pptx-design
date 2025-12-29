"""
Agent-Native Interface for PPTX Design System

Provides structured tool definitions for AI agents to create presentations.
Compatible with OpenAI function calling, Anthropic tool use, and other agent frameworks.

Phase 3 Enhancement (2025-12-29):
- Structured tool schemas for all presentation operations
- Function calling interface for LLM agents
- Direct execution methods with validation
- JSON Schema definitions for parameters

Usage:
    from pptx_design.agent_tools import AgentInterface, get_tool_definitions

    # Get tool definitions for LLM function calling
    tools = get_tool_definitions()

    # Create interface and execute tools
    agent = AgentInterface()
    result = agent.execute("create_presentation", {"template": "consulting_toolkit"})
"""

import json
import logging
from dataclasses import dataclass, field
from typing import Any, Callable, Dict, List, Optional, Union

logger = logging.getLogger(__name__)


# =============================================================================
# Tool Definitions (JSON Schema format for LLM function calling)
# =============================================================================

TOOL_DEFINITIONS = [
    {
        "name": "create_presentation",
        "description": "Create a new PowerPoint presentation with the specified template. Returns a presentation ID for subsequent operations.",
        "parameters": {
            "type": "object",
            "properties": {
                "template": {
                    "type": "string",
                    "description": "Template name to use. Options: consulting_toolkit, business_case, market_analysis, default",
                    "enum": ["consulting_toolkit", "business_case", "market_analysis", "default"],
                    "default": "consulting_toolkit"
                },
                "title": {
                    "type": "string",
                    "description": "Optional title for the presentation. Creates a title slide if provided."
                }
            },
            "required": []
        }
    },
    {
        "name": "add_title_slide",
        "description": "Add a title slide with main title and optional subtitle.",
        "parameters": {
            "type": "object",
            "properties": {
                "presentation_id": {
                    "type": "string",
                    "description": "The ID of the presentation to modify."
                },
                "title": {
                    "type": "string",
                    "description": "Main title text for the slide."
                },
                "subtitle": {
                    "type": "string",
                    "description": "Optional subtitle text."
                }
            },
            "required": ["presentation_id", "title"]
        }
    },
    {
        "name": "add_content_slide",
        "description": "Add a content slide with a title and bullet points.",
        "parameters": {
            "type": "object",
            "properties": {
                "presentation_id": {
                    "type": "string",
                    "description": "The ID of the presentation to modify."
                },
                "title": {
                    "type": "string",
                    "description": "Slide title."
                },
                "bullets": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "List of bullet point strings."
                }
            },
            "required": ["presentation_id", "title", "bullets"]
        }
    },
    {
        "name": "add_section_slide",
        "description": "Add a section divider slide to separate parts of the presentation.",
        "parameters": {
            "type": "object",
            "properties": {
                "presentation_id": {
                    "type": "string",
                    "description": "The ID of the presentation to modify."
                },
                "title": {
                    "type": "string",
                    "description": "Section title."
                },
                "subtitle": {
                    "type": "string",
                    "description": "Optional section subtitle."
                }
            },
            "required": ["presentation_id", "title"]
        }
    },
    {
        "name": "add_agenda_slide",
        "description": "Add an agenda slide with a list of topics.",
        "parameters": {
            "type": "object",
            "properties": {
                "presentation_id": {
                    "type": "string",
                    "description": "The ID of the presentation to modify."
                },
                "items": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "List of agenda items."
                },
                "title": {
                    "type": "string",
                    "description": "Title for the agenda slide.",
                    "default": "Agenda"
                }
            },
            "required": ["presentation_id", "items"]
        }
    },
    {
        "name": "add_two_column_slide",
        "description": "Add a two-column comparison slide for side-by-side content.",
        "parameters": {
            "type": "object",
            "properties": {
                "presentation_id": {
                    "type": "string",
                    "description": "The ID of the presentation to modify."
                },
                "title": {
                    "type": "string",
                    "description": "Slide title."
                },
                "left_content": {
                    "type": "object",
                    "properties": {
                        "header": {"type": "string"},
                        "bullets": {"type": "array", "items": {"type": "string"}}
                    },
                    "description": "Content for the left column."
                },
                "right_content": {
                    "type": "object",
                    "properties": {
                        "header": {"type": "string"},
                        "bullets": {"type": "array", "items": {"type": "string"}}
                    },
                    "description": "Content for the right column."
                }
            },
            "required": ["presentation_id", "title", "left_content", "right_content"]
        }
    },
    {
        "name": "add_metrics_slide",
        "description": "Add a key metrics slide with KPI boxes displaying important numbers.",
        "parameters": {
            "type": "object",
            "properties": {
                "presentation_id": {
                    "type": "string",
                    "description": "The ID of the presentation to modify."
                },
                "title": {
                    "type": "string",
                    "description": "Slide title."
                },
                "metrics": {
                    "type": "array",
                    "items": {
                        "type": "object",
                        "properties": {
                            "label": {"type": "string"},
                            "value": {"type": "string"}
                        },
                        "required": ["label", "value"]
                    },
                    "description": "List of metrics with label and value pairs. Example: [{\"label\": \"Revenue\", \"value\": \"$1.2M\"}]"
                }
            },
            "required": ["presentation_id", "title", "metrics"]
        }
    },
    {
        "name": "add_chart_slide",
        "description": "Add a slide with a data chart (column, bar, line, or pie chart).",
        "parameters": {
            "type": "object",
            "properties": {
                "presentation_id": {
                    "type": "string",
                    "description": "The ID of the presentation to modify."
                },
                "title": {
                    "type": "string",
                    "description": "Slide title."
                },
                "chart_type": {
                    "type": "string",
                    "enum": ["column", "bar", "line", "pie"],
                    "description": "Type of chart to create."
                },
                "categories": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "Category labels (x-axis values)."
                },
                "series": {
                    "type": "array",
                    "items": {
                        "type": "object",
                        "properties": {
                            "name": {"type": "string"},
                            "values": {"type": "array", "items": {"type": "number"}}
                        },
                        "required": ["name", "values"]
                    },
                    "description": "Data series for the chart."
                }
            },
            "required": ["presentation_id", "title", "chart_type", "categories", "series"]
        }
    },
    {
        "name": "add_table_slide",
        "description": "Add a slide with a data table.",
        "parameters": {
            "type": "object",
            "properties": {
                "presentation_id": {
                    "type": "string",
                    "description": "The ID of the presentation to modify."
                },
                "title": {
                    "type": "string",
                    "description": "Slide title."
                },
                "headers": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "Column header strings."
                },
                "rows": {
                    "type": "array",
                    "items": {
                        "type": "array",
                        "items": {"type": "string"}
                    },
                    "description": "Table rows, each row is an array of cell values."
                }
            },
            "required": ["presentation_id", "title", "headers", "rows"]
        }
    },
    {
        "name": "save_presentation",
        "description": "Save the presentation to a PowerPoint file.",
        "parameters": {
            "type": "object",
            "properties": {
                "presentation_id": {
                    "type": "string",
                    "description": "The ID of the presentation to save."
                },
                "output_path": {
                    "type": "string",
                    "description": "Path where the .pptx file should be saved."
                }
            },
            "required": ["presentation_id", "output_path"]
        }
    },
    {
        "name": "extract_slide_text",
        "description": "Extract text content from a specific slide for verification.",
        "parameters": {
            "type": "object",
            "properties": {
                "presentation_id": {
                    "type": "string",
                    "description": "The ID of the presentation."
                },
                "slide_index": {
                    "type": "integer",
                    "description": "Zero-based index of the slide."
                }
            },
            "required": ["presentation_id", "slide_index"]
        }
    },
    {
        "name": "get_slide_count",
        "description": "Get the number of slides in a presentation.",
        "parameters": {
            "type": "object",
            "properties": {
                "presentation_id": {
                    "type": "string",
                    "description": "The ID of the presentation."
                }
            },
            "required": ["presentation_id"]
        }
    },
    {
        "name": "list_templates",
        "description": "List all available presentation templates.",
        "parameters": {
            "type": "object",
            "properties": {},
            "required": []
        }
    },
    {
        "name": "search_images",
        "description": "Search for images to use in presentations. Requires PEXELS_API_KEY or UNSPLASH_ACCESS_KEY environment variable.",
        "parameters": {
            "type": "object",
            "properties": {
                "query": {
                    "type": "string",
                    "description": "Search keywords for the image."
                },
                "count": {
                    "type": "integer",
                    "description": "Number of images to return.",
                    "default": 3
                },
                "orientation": {
                    "type": "string",
                    "enum": ["landscape", "portrait", "square"],
                    "description": "Preferred image orientation."
                }
            },
            "required": ["query"]
        }
    }
]


# =============================================================================
# Agent Interface
# =============================================================================

@dataclass
class ToolResult:
    """Result from executing a tool."""
    success: bool
    data: Dict[str, Any] = field(default_factory=dict)
    error: Optional[str] = None

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary."""
        if self.success:
            return {"status": "success", **self.data}
        else:
            return {"status": "error", "message": self.error}


class AgentInterface:
    """
    Agent-native interface for presentation creation.

    Provides validated tool execution for AI agents.
    """

    def __init__(self):
        """Initialize the agent interface."""
        self._presentations: Dict[str, Any] = {}
        self._tool_handlers: Dict[str, Callable] = {}
        self._register_handlers()

    def _register_handlers(self):
        """Register tool execution handlers."""
        self._tool_handlers = {
            "create_presentation": self._create_presentation,
            "add_title_slide": self._add_title_slide,
            "add_content_slide": self._add_content_slide,
            "add_section_slide": self._add_section_slide,
            "add_agenda_slide": self._add_agenda_slide,
            "add_two_column_slide": self._add_two_column_slide,
            "add_metrics_slide": self._add_metrics_slide,
            "add_chart_slide": self._add_chart_slide,
            "add_table_slide": self._add_table_slide,
            "save_presentation": self._save_presentation,
            "extract_slide_text": self._extract_slide_text,
            "get_slide_count": self._get_slide_count,
            "list_templates": self._list_templates,
            "search_images": self._search_images,
        }

    def execute(self, tool_name: str, parameters: Dict[str, Any]) -> ToolResult:
        """
        Execute a tool with the given parameters.

        Args:
            tool_name: Name of the tool to execute
            parameters: Tool parameters as a dictionary

        Returns:
            ToolResult with success/failure and data
        """
        if tool_name not in self._tool_handlers:
            return ToolResult(
                success=False,
                error=f"Unknown tool: {tool_name}. Available: {list(self._tool_handlers.keys())}"
            )

        try:
            handler = self._tool_handlers[tool_name]
            result = handler(**parameters)
            return ToolResult(success=True, data=result)
        except Exception as e:
            logger.error(f"Tool execution error ({tool_name}): {e}")
            return ToolResult(success=False, error=str(e))

    def _get_builder(self, presentation_id: str):
        """Get a PresentationBuilder by ID."""
        if presentation_id not in self._presentations:
            raise ValueError(f"Presentation not found: {presentation_id}")
        return self._presentations[presentation_id]

    def _generate_id(self) -> str:
        """Generate a unique presentation ID."""
        import uuid
        return str(uuid.uuid4())[:8]

    # Tool handlers
    def _create_presentation(self, template: str = "consulting_toolkit", title: str = "") -> Dict:
        from pptx_design import PresentationBuilder

        presentation_id = self._generate_id()
        builder = PresentationBuilder(template)

        if title:
            builder.add_title_slide(title)

        self._presentations[presentation_id] = builder

        return {
            "presentation_id": presentation_id,
            "template": template,
            "slide_count": len(builder._slides)
        }

    def _add_title_slide(self, presentation_id: str, title: str, subtitle: str = "") -> Dict:
        builder = self._get_builder(presentation_id)
        builder.add_title_slide(title, subtitle)
        return {
            "slide_type": "title_slide",
            "slide_index": len(builder._slides) - 1
        }

    def _add_content_slide(self, presentation_id: str, title: str, bullets: List[str]) -> Dict:
        builder = self._get_builder(presentation_id)
        builder.add_content_slide(title, bullets=bullets)
        return {
            "slide_type": "title_content",
            "slide_index": len(builder._slides) - 1
        }

    def _add_section_slide(self, presentation_id: str, title: str, subtitle: str = "") -> Dict:
        builder = self._get_builder(presentation_id)
        builder.add_section_divider(title, subtitle)
        return {
            "slide_type": "section_divider",
            "slide_index": len(builder._slides) - 1
        }

    def _add_agenda_slide(self, presentation_id: str, items: List[str], title: str = "Agenda") -> Dict:
        builder = self._get_builder(presentation_id)
        builder.add_agenda(items, title=title)
        return {
            "slide_type": "agenda",
            "slide_index": len(builder._slides) - 1
        }

    def _add_two_column_slide(
        self,
        presentation_id: str,
        title: str,
        left_content: Dict,
        right_content: Dict
    ) -> Dict:
        builder = self._get_builder(presentation_id)
        builder.add_two_column(title, left_content, right_content)
        return {
            "slide_type": "two_column",
            "slide_index": len(builder._slides) - 1
        }

    def _add_metrics_slide(
        self,
        presentation_id: str,
        title: str,
        metrics: List[Dict[str, str]]
    ) -> Dict:
        builder = self._get_builder(presentation_id)
        builder.add_metrics(title, metrics)
        return {
            "slide_type": "key_metrics",
            "slide_index": len(builder._slides) - 1
        }

    def _add_chart_slide(
        self,
        presentation_id: str,
        title: str,
        chart_type: str,
        categories: List[str],
        series: List[Dict[str, Any]]
    ) -> Dict:
        builder = self._get_builder(presentation_id)
        chart_data = {
            "type": chart_type,
            "categories": categories,
            "series": series
        }
        builder.add_chart(title, chart_data)
        return {
            "slide_type": "data_chart",
            "slide_index": len(builder._slides) - 1
        }

    def _add_table_slide(
        self,
        presentation_id: str,
        title: str,
        headers: List[str],
        rows: List[List[str]]
    ) -> Dict:
        builder = self._get_builder(presentation_id)
        table_data = {"headers": headers, "rows": rows}
        builder.add_table(title, table_data)
        return {
            "slide_type": "table_slide",
            "slide_index": len(builder._slides) - 1
        }

    def _save_presentation(self, presentation_id: str, output_path: str) -> Dict:
        builder = self._get_builder(presentation_id)

        if not output_path.endswith('.pptx'):
            output_path += '.pptx'

        saved_path = builder.save(output_path)
        return {
            "path": str(saved_path),
            "slide_count": len(builder._slides)
        }

    def _extract_slide_text(self, presentation_id: str, slide_index: int) -> Dict:
        builder = self._get_builder(presentation_id)

        if slide_index < 0 or slide_index >= len(builder._slides):
            raise ValueError(f"Invalid slide index: {slide_index}")

        # Get text from the slide spec (before building)
        slide_spec = builder._slides[slide_index]
        content = slide_spec.get("content", {})
        text_content = []

        # Extract text from content
        for key in ["title", "subtitle", "body"]:
            if key in content:
                val = content[key]
                if isinstance(val, str) and val:
                    text_content.append(val)

        # Extract bullets
        if "bullets" in content:
            text_content.extend(content["bullets"])

        return {
            "slide_index": slide_index,
            "text": text_content
        }

    def _get_slide_count(self, presentation_id: str) -> Dict:
        builder = self._get_builder(presentation_id)
        return {"slide_count": len(builder._slides)}

    def _list_templates(self) -> Dict:
        from pptx_design import TemplateRegistry

        registry = TemplateRegistry()
        templates = []

        for template_id, template in registry.templates.items():
            templates.append({
                "id": template_id,
                "name": template.metadata.get("name", template_id),
                "layout_count": len(template.layouts)
            })

        return {"templates": templates}

    def _search_images(
        self,
        query: str,
        count: int = 3,
        orientation: str = None
    ) -> Dict:
        from pptx_generator.modules.image_search import ImageSearch

        search = ImageSearch()

        if not search.is_available:
            return {
                "images": [],
                "warning": "No image API keys configured"
            }

        results = search.search(query, count=count, orientation=orientation)
        images = [
            {
                "id": img.id,
                "source": img.source,
                "thumbnail_url": img.thumbnail_url,
                "photographer": img.photographer,
                "width": img.width,
                "height": img.height
            }
            for img in results
        ]

        return {"images": images, "count": len(images)}


# =============================================================================
# Convenience Functions
# =============================================================================

def get_tool_definitions() -> List[Dict[str, Any]]:
    """
    Get tool definitions in JSON Schema format for LLM function calling.

    Returns:
        List of tool definition dictionaries
    """
    return TOOL_DEFINITIONS


def get_openai_tools() -> List[Dict[str, Any]]:
    """
    Get tool definitions in OpenAI function calling format.

    Returns:
        List of tool definitions for OpenAI API
    """
    return [
        {"type": "function", "function": tool}
        for tool in TOOL_DEFINITIONS
    ]


def get_anthropic_tools() -> List[Dict[str, Any]]:
    """
    Get tool definitions in Anthropic tool use format.

    Returns:
        List of tool definitions for Anthropic API
    """
    return [
        {
            "name": tool["name"],
            "description": tool["description"],
            "input_schema": tool["parameters"]
        }
        for tool in TOOL_DEFINITIONS
    ]


def export_tool_schema(output_path: str = "pptx_tools.json"):
    """
    Export tool definitions to a JSON file.

    Args:
        output_path: Path to save the JSON file
    """
    with open(output_path, 'w') as f:
        json.dump(TOOL_DEFINITIONS, f, indent=2)
    logger.info(f"Exported tool schema to: {output_path}")


# =============================================================================
# CLI
# =============================================================================

def main():
    """CLI for agent tools."""
    import argparse

    parser = argparse.ArgumentParser(description="Agent Tools Interface")
    parser.add_argument("--export", "-e", action="store_true", help="Export tool schema to JSON")
    parser.add_argument("--output", "-o", default="pptx_tools.json", help="Output file for export")
    parser.add_argument("--list", "-l", action="store_true", help="List available tools")
    parser.add_argument("--format", "-f", choices=["default", "openai", "anthropic"],
                        default="default", help="Output format")

    args = parser.parse_args()

    if args.export:
        export_tool_schema(args.output)
        print(f"Exported tool schema to: {args.output}")

    elif args.list:
        if args.format == "openai":
            tools = get_openai_tools()
        elif args.format == "anthropic":
            tools = get_anthropic_tools()
        else:
            tools = get_tool_definitions()

        print(json.dumps(tools, indent=2))

    else:
        print("PPTX Design Agent Tools")
        print("=" * 40)
        print(f"Available tools: {len(TOOL_DEFINITIONS)}")
        for tool in TOOL_DEFINITIONS:
            print(f"  - {tool['name']}: {tool['description'][:60]}...")


if __name__ == "__main__":
    main()
