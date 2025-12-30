"""
Template-Based Slide Renderer

Generates slides using actual template master layouts and styling.
Integrates with ComponentLibrary for reusable charts, tables, and diagrams.

Fixed version: Uses actual placeholder positions and proper auto-sizing.
"""

import json
import logging
from copy import deepcopy
from pathlib import Path
from typing import Any, Dict, Optional, Tuple

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.slide import Slide
from pptx.util import Emu, Inches, Pt

# Import ComponentLibrary and LibraryEnhancer for reusable components
try:
    from .component_library import ComponentLibrary
    from .library_enhancer import LibraryEnhancer
    LIBRARY_AVAILABLE = True
except ImportError:
    LIBRARY_AVAILABLE = False
    ComponentLibrary = None
    LibraryEnhancer = None

logger = logging.getLogger(__name__)


class TemplateRenderer:
    """Renders slides using template master layouts and styling."""

    # Layout mappings for different slide types
    # Using template master layouts for professional styling
    # "Default" layout has title at top=0.39" (ABOVE the horizontal line)
    # "Top left title" has title at top=1.87" (BELOW the line) - avoid for standard content
    LAYOUT_MAP = {
        "title_slide": "Frontpage",
        "frontpage": "Frontpage",
        "section_divider": "Section breaker",
        "section_breaker": "Section breaker",
        "title_content": "Default",  # Title above horizontal line (top=0.39")
        "default": "Default",
        "top_left": "Default",  # Use Default for proper title placement
        "content": "Default",
        "two_column": "1/2 grey",
        "comparison": "1/2 grey",
        "sidebar_left": "1/3 grey",
        "sidebar_right": "2/3 grey",
        "agenda": "Agenda",
        "end_slide": "End",
        "blank": "Blank",
        "data_chart": "Default",  # Title above line
        "table_slide": "Default",  # Title above line
        "key_metrics": "Default",  # Title above line
    }

    # Slide dimensions (standard 16:9)
    SLIDE_WIDTH = 13.333  # inches
    SLIDE_HEIGHT = 7.5    # inches

    # Margins
    MARGIN_LEFT = 0.61
    MARGIN_RIGHT = 0.6
    MARGIN_TOP = 0.39
    MARGIN_BOTTOM = 0.5

    # Bullet character for lists
    BULLET_CHAR = "•"

    # Colors extracted from templates
    COLORS = {
        "primary": RGBColor(0x3C, 0x96, 0xB4),      # Teal
        "secondary": RGBColor(0xE5, 0x54, 0x6C),    # Coral/Red
        "accent1": RGBColor(0x05, 0x1C, 0x2C),      # Dark blue
        "accent2": RGBColor(0x00, 0xB0, 0x50),      # Green
        "text_dark": RGBColor(0x06, 0x1F, 0x32),    # Dark navy
        "text_body": RGBColor(0x33, 0x33, 0x33),    # Dark gray
        "text_light": RGBColor(0x66, 0x66, 0x66),   # Medium gray
        "white": RGBColor(0xFF, 0xFF, 0xFF),
        "light_gray": RGBColor(0xE6, 0xE6, 0xE6),
        "background": RGBColor(0xF2, 0xF2, 0xF2),
    }

    # Font settings - adjusted sizes for better fit
    FONTS = {
        "title": {"name": "Arial", "size": Pt(24), "bold": True, "color": "text_dark"},
        "subtitle": {"name": "Arial", "size": Pt(14), "bold": False, "color": "text_body"},
        "section": {"name": "Arial", "size": Pt(32), "bold": True, "color": "text_dark"},
        "heading": {"name": "Arial", "size": Pt(14), "bold": True, "color": "text_dark"},
        "body": {"name": "Arial", "size": Pt(12), "bold": False, "color": "text_body"},
        "body_large": {"name": "Arial", "size": Pt(13), "bold": False, "color": "text_body"},
        "bullet": {"name": "Arial", "size": Pt(12), "bold": False, "color": "text_body"},
        "caption": {"name": "Arial", "size": Pt(9), "bold": False, "color": "text_light"},
        "metric_value": {"name": "Arial", "size": Pt(28), "bold": True, "color": "white"},
        "metric_label": {"name": "Arial", "size": Pt(11), "bold": False, "color": "white"},
        "table_header": {"name": "Arial", "size": Pt(10), "bold": True, "color": "white"},
        "table_cell": {"name": "Arial", "size": Pt(10), "bold": False, "color": "text_body"},
        "chart_label": {"name": "Arial", "size": Pt(10), "bold": False, "color": "text_body"},
    }

    def __init__(self, template_path: str, use_library: bool = True):
        """
        Initialize renderer with a template.

        Args:
            template_path: Path to the PPTX template file
            use_library: Whether to use ComponentLibrary for charts/tables
        """
        self.template_path = Path(template_path)
        if not self.template_path.exists():
            raise FileNotFoundError(f"Template not found: {template_path}")

        # Load template to get layouts
        self._template_prs = Presentation(str(self.template_path))
        self._build_layout_index()
        self._extract_layout_dimensions()

        # Initialize component library and enhancer
        self.library = None
        self.enhancer = None
        if use_library and LIBRARY_AVAILABLE:
            try:
                self.library = ComponentLibrary()
                if self.library.is_available:
                    stats = self.library.get_stats()
                    total = sum(stats['components'].values())
                    logger.info(f"ComponentLibrary loaded with {total} components")

                    # Initialize enhancer for smart matching
                    if LibraryEnhancer:
                        self.enhancer = LibraryEnhancer(self.library)
                else:
                    self.library = None
            except Exception as e:
                logger.warning(f"Could not initialize ComponentLibrary: {e}")
                self.library = None

    def _build_layout_index(self) -> None:
        """Build index of available layouts."""
        self.layouts = {}
        for layout in self._template_prs.slide_layouts:
            self.layouts[layout.name] = layout
        logger.info(f"Indexed {len(self.layouts)} layouts from template")

    def _extract_layout_dimensions(self) -> None:
        """Extract actual placeholder dimensions from each layout."""
        self.layout_dimensions = {}

        for layout in self._template_prs.slide_layouts:
            dims = {
                "title": None,
                "subtitle": None,
                "body": None,
                "content_top": 1.8,  # Default content start
            }

            max_placeholder_bottom = self.MARGIN_TOP

            for ph in layout.placeholders:
                ph_type = str(ph.placeholder_format.type)
                left = ph.left / 914400  # EMUs to inches
                top = ph.top / 914400
                width = ph.width / 914400
                height = ph.height / 914400
                bottom = top + height

                if "TITLE" in ph_type:
                    dims["title"] = {"left": left, "top": top, "width": width, "height": height}
                    max_placeholder_bottom = max(max_placeholder_bottom, bottom)
                elif "SUBTITLE" in ph_type:
                    dims["subtitle"] = {"left": left, "top": top, "width": width, "height": height}
                    max_placeholder_bottom = max(max_placeholder_bottom, bottom)
                elif "BODY" in ph_type and ph.placeholder_format.idx != 17:
                    # idx 17 is typically a small header body, skip it
                    dims["body"] = {"left": left, "top": top, "width": width, "height": height}

            # Content area starts after title/subtitle placeholders
            dims["content_top"] = max_placeholder_bottom + 0.2
            self.layout_dimensions[layout.name] = dims

    def _get_content_area(self, layout_name: str) -> Dict[str, float]:
        """Get the content area dimensions for a layout."""
        dims = self.layout_dimensions.get(layout_name, {})
        content_top = dims.get("content_top", 1.8)

        return {
            "left": self.MARGIN_LEFT,
            "top": content_top,
            "width": self.SLIDE_WIDTH - self.MARGIN_LEFT - self.MARGIN_RIGHT,
            "height": self.SLIDE_HEIGHT - content_top - self.MARGIN_BOTTOM,
        }

    def create_presentation(self) -> Presentation:
        """Create a new empty presentation with template layouts."""
        prs = Presentation(str(self.template_path))

        # Remove all existing slides from the template
        while len(prs.slides) > 0:
            slide = prs.slides[0]
            rId = prs.part.relate_to(slide.part, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide')
            sldIdLst = prs._element.get_or_add_sldIdLst()
            sldId = sldIdLst.sldId_lst[0]
            sldIdLst.remove(sldId)
            prs.part.drop_rel(sldId.rId)

        logger.info(f"Created empty presentation with {len(prs.slide_layouts)} layouts")
        return prs

    def get_layout(self, prs: Presentation, slide_type: str):
        """Get the appropriate layout for a slide type."""
        layout_name = self.LAYOUT_MAP.get(slide_type, "Default")

        for layout in prs.slide_layouts:
            if layout.name == layout_name:
                return layout

        # Fallback to Default
        for layout in prs.slide_layouts:
            if layout.name == "Default":
                return layout

        return prs.slide_layouts[1]

    def create_slide(
        self,
        prs: Presentation,
        slide_type: str,
        content: dict,
        use_placeholders: bool = True
    ) -> Slide:
        """
        Create a slide using template layout.

        Args:
            prs: Presentation object
            slide_type: Type of slide to create
            content: Content dictionary
            use_placeholders: If True, fill master layout placeholders directly

        Returns:
            Created Slide object
        """
        layout = self.get_layout(prs, slide_type)
        slide = prs.slides.add_slide(layout)
        layout_name = layout.name

        # Clear ALL placeholder content from the template
        # This prevents template default text from appearing
        self._clear_placeholder_content(slide)

        # Handle title - add manually for Blank layout
        title = content.get("title", "")
        if title:
            if slide.shapes.title:
                self._set_title(slide.shapes.title, title, slide_type)
            elif layout_name == "Blank":
                # Add title manually for Blank layout
                self._add_title_to_blank_slide(slide, title, slide_type)

        # Route to appropriate renderer
        renderer_method = getattr(self, f"_render_{slide_type}", None)
        if renderer_method:
            renderer_method(slide, content, layout_name)
        else:
            self._render_default(slide, content, layout_name)

        return slide

    def _add_title_to_blank_slide(self, slide: Slide, title: str, slide_type: str) -> None:
        """Add a title text box to a blank slide layout."""
        # Title position matching template style
        left = Inches(self.MARGIN_LEFT)
        top = Inches(0.39)
        width = Inches(self.SLIDE_WIDTH - self.MARGIN_LEFT - self.MARGIN_RIGHT)
        height = Inches(0.6)

        title_box = slide.shapes.add_textbox(left, top, width, height)
        tf = title_box.text_frame
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE

        p = tf.paragraphs[0]
        p.text = title
        self._apply_font(p, "title")

        # Add a horizontal line below the title
        line_top = Inches(1.05)
        line = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            left, line_top,
            width, Inches(0.02)
        )
        line.fill.solid()
        line.fill.fore_color.rgb = self.COLORS["text_dark"]
        line.line.fill.background()

    def _clear_placeholder_content(self, slide: Slide) -> None:
        """Clear all content shapes from a slide to remove template defaults.

        This AGGRESSIVELY removes ALL shapes except:
        - Title placeholder (we'll fill it)
        - Footer/slide number/date placeholders
        - Horizontal lines in footer area (decorative dividers)
        """
        shapes_to_remove = []

        # Get title shape reference before iterating
        title_shape = slide.shapes.title

        # Identify ALL shapes to remove
        for shape in slide.shapes:
            # Keep slide title placeholder only
            if shape == title_shape:
                continue

            # Check if it's a footer/slide number/date placeholder - keep these
            if shape.is_placeholder:
                ph_type = str(shape.placeholder_format.type)
                if any(keep in ph_type for keep in ["FOOTER", "SLIDE_NUMBER", "DATE"]):
                    continue

            # Check for footer divider lines (thin horizontal lines at bottom)
            try:
                shape_top = shape.top / 914400  # EMUs to inches
                shape_height = shape.height / 914400
                shape_width = shape.width / 914400

                # Keep only thin horizontal lines (height < 0.1") in footer area (bottom 1")
                footer_area = self.SLIDE_HEIGHT - 1.0
                is_horizontal_line = shape_height < 0.1 and shape_width > 5.0

                if shape_top > footer_area and is_horizontal_line:
                    continue
            except Exception:
                pass

            # REMOVE EVERYTHING ELSE - no exceptions
            shapes_to_remove.append(shape)

        # Remove all identified shapes
        for shape in shapes_to_remove:
            try:
                sp = shape._element
                sp.getparent().remove(sp)
            except Exception as e:
                logger.warning(f"Failed to remove shape: {e}")

    def _set_title(self, title_shape, text: str, slide_type: str) -> None:
        """Set title text with proper formatting."""
        title_shape.text = text
        tf = title_shape.text_frame

        # Enable word wrap and auto-size
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE

        # Apply font
        style = "section" if slide_type == "section_divider" else "title"
        for para in tf.paragraphs:
            self._apply_font(para, style)

    def _create_text_box(
        self,
        slide: Slide,
        left: float,
        top: float,
        width: float,
        height: float,
        auto_size: bool = True
    ) -> Any:
        """Create a text box with proper settings."""
        textbox = slide.shapes.add_textbox(
            Inches(left), Inches(top),
            Inches(width), Inches(height)
        )
        tf = textbox.text_frame
        tf.word_wrap = True

        if auto_size:
            tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE

        return textbox

    def _add_bullets(
        self,
        slide: Slide,
        bullets: list,
        area: Dict[str, float],
        style: str = "bullet"
    ) -> None:
        """Add bullet points to a slide."""
        if not bullets:
            return

        textbox = self._create_text_box(
            slide,
            area["left"],
            area["top"],
            area["width"],
            area["height"]
        )
        tf = textbox.text_frame
        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE

        for i, bullet in enumerate(bullets):
            if i == 0:
                p = tf.paragraphs[0]
            else:
                p = tf.add_paragraph()

            p.text = f"{self.BULLET_CHAR}  {bullet}"
            p.level = 0
            self._apply_font(p, style)
            p.space_before = Pt(6)
            p.space_after = Pt(4)

    # ==================== Slide Renderers ====================

    def _render_title_slide(self, slide: Slide, content: dict, layout_name: str) -> None:
        """Render a title/frontpage slide."""
        subtitle = content.get("subtitle", "")

        # Find and set subtitle placeholder
        for shape in slide.placeholders:
            ph_type = str(shape.placeholder_format.type)
            if "SUBTITLE" in ph_type:
                shape.text = subtitle
                tf = shape.text_frame
                tf.word_wrap = True
                for para in tf.paragraphs:
                    self._apply_font(para, "subtitle")
                break

    def _render_section_divider(self, slide: Slide, content: dict, layout_name: str) -> None:
        """Render a section divider slide."""
        # Title already set in create_slide
        pass

    def _render_default(self, slide: Slide, content: dict, layout_name: str) -> None:
        """Render a default content slide."""
        self._render_title_content(slide, content, layout_name)

    def _render_title_content(self, slide: Slide, content: dict, layout_name: str) -> None:
        """Render a title + content slide."""
        bullets = content.get("bullets", [])
        body = content.get("body", "")

        if not bullets and not body:
            return

        area = self._get_content_area(layout_name)

        if body:
            textbox = self._create_text_box(
                slide, area["left"], area["top"],
                area["width"], min(area["height"], 1.5)
            )
            tf = textbox.text_frame
            p = tf.paragraphs[0]
            p.text = body
            self._apply_font(p, "body_large")

            # Adjust area for bullets
            area["top"] += 1.6
            area["height"] -= 1.6

        if bullets:
            self._add_bullets(slide, bullets, area)

    def _render_two_column(self, slide: Slide, content: dict, layout_name: str) -> None:
        """Render a two-column comparison slide."""
        left_col = content.get("left_column", content.get("left", {}))
        right_col = content.get("right_column", content.get("right", {}))

        area = self._get_content_area(layout_name)
        col_width = (area["width"] - 0.4) / 2  # Gap between columns

        if left_col:
            left_area = {
                "left": area["left"],
                "top": area["top"],
                "width": col_width,
                "height": area["height"]
            }
            self._render_column(slide, left_col, left_area)

        if right_col:
            right_area = {
                "left": area["left"] + col_width + 0.4,
                "top": area["top"],
                "width": col_width,
                "height": area["height"]
            }
            self._render_column(slide, right_col, right_area)

    def _render_column(self, slide: Slide, col_content: dict, area: Dict[str, float]) -> None:
        """Render a single column."""
        header = col_content.get("header", col_content.get("heading", ""))
        bullets = col_content.get("bullets", [])

        textbox = self._create_text_box(
            slide, area["left"], area["top"],
            area["width"], area["height"]
        )
        tf = textbox.text_frame
        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE

        # Header
        if header:
            p = tf.paragraphs[0]
            p.text = header
            self._apply_font(p, "heading")
            p.space_after = Pt(10)

        # Bullets
        for i, bullet in enumerate(bullets):
            p = tf.add_paragraph() if (i > 0 or header) else tf.paragraphs[0]
            p.text = f"{self.BULLET_CHAR}  {bullet}"
            p.level = 0
            self._apply_font(p, "body")
            p.space_before = Pt(4)
            p.space_after = Pt(2)

    def _render_key_metrics(self, slide: Slide, content: dict, layout_name: str) -> None:
        """Render a key metrics slide with KPI boxes."""
        metrics = content.get("metrics", [])
        if not metrics:
            return

        area = self._get_content_area(layout_name)

        num_metrics = min(len(metrics), 5)
        box_margin = 0.15
        box_width = (area["width"] - (num_metrics - 1) * box_margin) / num_metrics
        box_height = 1.4

        # Center vertically in content area
        top = area["top"] + (area["height"] - box_height) / 2

        for i, metric in enumerate(metrics[:5]):
            left = area["left"] + i * (box_width + box_margin)

            # Create box shape
            shape = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                Inches(left), Inches(top),
                Inches(box_width), Inches(box_height)
            )

            # Style the box
            shape.fill.solid()
            shape.fill.fore_color.rgb = self.COLORS["primary"]
            shape.line.fill.background()

            # Add text
            tf = shape.text_frame
            tf.word_wrap = True
            tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE

            # Value
            p = tf.paragraphs[0]
            p.text = str(metric.get("value", ""))
            p.alignment = PP_ALIGN.CENTER
            self._apply_font(p, "metric_value")

            # Label
            p2 = tf.add_paragraph()
            p2.text = metric.get("label", "")
            p2.alignment = PP_ALIGN.CENTER
            self._apply_font(p2, "metric_label")

    def _render_table_slide(self, slide: Slide, content: dict, layout_name: str) -> None:
        """Render a slide with a data table."""
        headers = content.get("headers", [])
        data = content.get("data", [])

        if not data and not headers:
            return

        area = self._get_content_area(layout_name)

        rows = len(data) + (1 if headers else 0)
        cols = len(headers) if headers else (len(data[0]) if data else 0)

        if rows == 0 or cols == 0:
            return

        # Calculate table dimensions - fit to content area
        row_height = min(0.35, (area["height"] - 0.2) / rows)
        table_height = min(rows * row_height, area["height"] - 0.2)

        table_shape = slide.shapes.add_table(
            rows, cols,
            Inches(area["left"]), Inches(area["top"]),
            Inches(area["width"]), Inches(table_height)
        )
        table = table_shape.table

        # Style header row
        if headers:
            for i, header in enumerate(headers):
                cell = table.cell(0, i)
                cell.text = str(header)
                cell.fill.solid()
                cell.fill.fore_color.rgb = self.COLORS["primary"]
                self._style_table_cell(cell, is_header=True)

        # Add data rows
        start_row = 1 if headers else 0
        for i, row in enumerate(data):
            for j, value in enumerate(row):
                if j < cols:  # Ensure we don't exceed columns
                    cell = table.cell(start_row + i, j)
                    cell.text = str(value)

                    # Alternate row colors
                    if i % 2 == 0:
                        cell.fill.solid()
                        cell.fill.fore_color.rgb = self.COLORS["white"]
                    else:
                        cell.fill.solid()
                        cell.fill.fore_color.rgb = self.COLORS["background"]

                    self._style_table_cell(cell, is_header=False)

    def _style_table_cell(self, cell, is_header: bool = False) -> None:
        """Style a table cell with proper margins and fonts."""
        cell.margin_left = Inches(0.05)
        cell.margin_right = Inches(0.05)
        cell.margin_top = Inches(0.03)
        cell.margin_bottom = Inches(0.03)

        tf = cell.text_frame
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE

        for para in tf.paragraphs:
            para.font.name = "Arial"
            para.font.size = Pt(10)
            if is_header:
                para.font.bold = True
                para.font.color.rgb = self.COLORS["white"]
            else:
                para.font.bold = False
                para.font.color.rgb = self.COLORS["text_body"]

    def _render_data_chart(self, slide: Slide, content: dict, layout_name: str) -> None:
        """Render a slide with a chart."""
        chart_data = content.get("chart_data", {})
        narrative = content.get("narrative", "")

        area = self._get_content_area(layout_name)

        if chart_data:
            # Reserve space for narrative if present
            chart_height = area["height"] - 0.6 if narrative else area["height"]
            self._add_chart(slide, chart_data, {
                "left": area["left"],
                "top": area["top"],
                "width": area["width"],
                "height": chart_height
            })

        # Add narrative/source text
        if narrative:
            textbox = self._create_text_box(
                slide,
                area["left"],
                area["top"] + area["height"] - 0.4,
                area["width"],
                0.4,
                auto_size=False
            )
            p = textbox.text_frame.paragraphs[0]
            p.text = narrative
            self._apply_font(p, "caption")

    def _add_chart(self, slide: Slide, chart_data: dict, area: Dict[str, float]) -> None:
        """Add a chart to the slide."""
        from pptx.chart.data import CategoryChartData
        from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION

        chart_type = chart_data.get("type", "column")
        categories = chart_data.get("categories", [])
        series_list = chart_data.get("series", [])

        if not categories or not series_list:
            return

        # Map chart types
        chart_type_map = {
            "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
            "bar": XL_CHART_TYPE.BAR_CLUSTERED,
            "line": XL_CHART_TYPE.LINE,
            "pie": XL_CHART_TYPE.PIE,
        }
        xl_chart_type = chart_type_map.get(chart_type, XL_CHART_TYPE.COLUMN_CLUSTERED)

        # Create chart data
        data = CategoryChartData()
        data.categories = categories

        for series in series_list:
            data.add_series(series.get("name", "Series"), series.get("values", []))

        # Adjust size for pie charts
        chart_width = area["width"] * 0.75 if chart_type == "pie" else area["width"]

        chart_shape = slide.shapes.add_chart(
            xl_chart_type,
            Inches(area["left"]), Inches(area["top"]),
            Inches(chart_width), Inches(area["height"]),
            data
        )
        chart = chart_shape.chart

        # Style the chart
        if chart_type == "pie":
            self._style_pie_chart(chart)
        else:
            self._style_bar_chart(chart)

    def _style_bar_chart(self, chart) -> None:
        """Style a bar/column/line chart."""
        try:
            plot = chart.plots[0]
            if hasattr(plot, 'series'):
                for i, series in enumerate(plot.series):
                    series.format.fill.solid()
                    if i == 0:
                        series.format.fill.fore_color.rgb = self.COLORS["primary"]
                    else:
                        series.format.fill.fore_color.rgb = self.COLORS["secondary"]

            # Add data labels
            plot.has_data_labels = True
            data_labels = plot.data_labels
            data_labels.font.size = Pt(9)
            data_labels.font.color.rgb = self.COLORS["text_body"]
        except Exception:
            pass

    def _style_pie_chart(self, chart) -> None:
        """Style a pie chart."""
        from pptx.enum.chart import XL_LEGEND_POSITION

        pie_colors = [
            self.COLORS["primary"],
            self.COLORS["secondary"],
            self.COLORS["accent2"],
            RGBColor(0xFF, 0xB8, 0x00),  # Orange
            RGBColor(0x7C, 0x4D, 0xC4),  # Purple
            self.COLORS["accent1"],
        ]

        try:
            plot = chart.plots[0]
            if hasattr(plot, 'series') and len(plot.series) > 0:
                series = plot.series[0]
                for i, point in enumerate(series.points):
                    point.format.fill.solid()
                    point.format.fill.fore_color.rgb = pie_colors[i % len(pie_colors)]

            # Add legend
            chart.has_legend = True
            chart.legend.position = XL_LEGEND_POSITION.RIGHT
            chart.legend.include_in_layout = False
            chart.legend.font.size = Pt(9)
        except Exception:
            pass

    def _apply_font(self, paragraph, style_name: str) -> None:
        """Apply font styling to a paragraph."""
        style = self.FONTS.get(style_name, self.FONTS["body"])

        if paragraph.runs:
            run = paragraph.runs[0]
        else:
            run = paragraph.add_run()

        run.font.name = style["name"]
        run.font.size = style["size"]
        run.font.bold = style.get("bold", False)

        color_key = style.get("color", "text_body")
        run.font.color.rgb = self.COLORS.get(color_key, self.COLORS["text_body"])

    # ==================== Component Library Integration ====================

    def find_library_chart(
        self,
        chart_type: str = "column",
        category: Optional[str] = None,
        min_series: int = 1
    ) -> Optional[dict]:
        """Find a matching chart component from the library."""
        if not self.library:
            return None

        type_map = {
            "column": "COLUMN",
            "bar": "BAR",
            "line": "LINE",
            "pie": "PIE",
            "area": "AREA",
        }
        lib_type = type_map.get(chart_type.lower(), "COLUMN")

        results = self.library.search_charts(
            chart_type=lib_type,
            min_series=min_series,
            limit=5
        )

        return results[0] if results else None

    def find_library_table(
        self,
        rows: int = 3,
        cols: int = 3,
        category: Optional[str] = None
    ) -> Optional[dict]:
        """Find a matching table component from the library."""
        if not self.library:
            return None

        results = self.library.search_tables(
            category=category,
            min_rows=max(1, rows - 2),
            max_rows=rows + 5,
            min_cols=max(1, cols - 1),
            max_cols=cols + 2,
            limit=5
        )

        if results:
            results.sort(key=lambda t: abs(t.get('rows', 0) - rows) + abs(t.get('cols', 0) - cols))
            return results[0]

        return None


def main():
    """Test the template renderer."""
    import argparse

    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

    parser = argparse.ArgumentParser(description="Template Renderer Test")
    parser.add_argument(
        "--template",
        default="pptx_templates/pptx_template_business_consulting_toolkit/template_business_consulting_toolkit.pptx",
        help="Path to template"
    )
    parser.add_argument(
        "--output",
        default="pptx_generator/output/template_render_test.pptx",
        help="Output file"
    )

    args = parser.parse_args()

    renderer = TemplateRenderer(args.template)
    prs = renderer.create_presentation()

    # Create test slides
    renderer.create_slide(prs, "title_slide", {
        "title": "Test Presentation",
        "subtitle": "Template Renderer Demo"
    })

    renderer.create_slide(prs, "section_divider", {
        "title": "Section One"
    })

    renderer.create_slide(prs, "title_content", {
        "title": "Key Points",
        "bullets": ["First important point", "Second important point", "Third important point"]
    })

    renderer.create_slide(prs, "key_metrics", {
        "title": "Key Metrics",
        "metrics": [
            {"label": "Revenue", "value": "$1.2M"},
            {"label": "Growth", "value": "25%"},
            {"label": "Users", "value": "10K"},
            {"label": "NPS", "value": "72"}
        ]
    })

    renderer.create_slide(prs, "table_slide", {
        "title": "Data Table",
        "headers": ["Name", "Value", "Status"],
        "data": [
            ["Item 1", "100", "Active"],
            ["Item 2", "200", "Pending"],
            ["Item 3", "300", "Complete"]
        ]
    })

    prs.save(args.output)
    print(f"Saved to: {args.output}")


if __name__ == "__main__":
    main()
