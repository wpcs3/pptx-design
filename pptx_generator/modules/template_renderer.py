"""
Template-Based Slide Renderer

Generates slides using actual template master layouts and styling.
Integrates with ComponentLibrary for reusable charts, tables, and diagrams.
"""

import json
import logging
from copy import deepcopy
from pathlib import Path
from typing import Any, Optional

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
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

# Placeholder type mapping for content
PLACEHOLDER_CONTENT_MAP = {
    "TITLE": ["title", "heading", "header"],
    "SUBTITLE": ["subtitle", "subheading", "sub_title"],
    "BODY": ["body", "content", "text", "bullet", "bullets"],
    "FOOTER": ["footer", "company", "company_name"],
    "SLIDE_NUMBER": ["slide_number", "page", "page_number"],
    "DATE": ["date"],
}

logger = logging.getLogger(__name__)


class TemplateRenderer:
    """Renders slides using template master layouts and styling."""

    # Layout mappings for different slide types
    LAYOUT_MAP = {
        "title_slide": "Frontpage",
        "frontpage": "Frontpage",
        "section_divider": "Section breaker",
        "section_breaker": "Section breaker",
        "title_content": "Default",
        "default": "Default",
        "top_left": "Top left title",
        "content": "Top left title",
        "two_column": "1/2 grey",
        "comparison": "1/2 grey",
        "sidebar_left": "1/3 grey",
        "sidebar_right": "2/3 grey",
        "agenda": "Agenda",
        "end_slide": "End",
        "blank": "Blank",
        "data_chart": "Default",
        "table_slide": "Default",
        "key_metrics": "Default",
    }

    # Standard positioning (in inches) - adjusted for better vertical balance
    POSITIONS = {
        "title": {"left": 0.61, "top": 0.39, "width": 12.12, "height": 0.8},
        "subtitle": {"left": 0.61, "top": 1.39, "width": 12.12, "height": 0.5},
        "body": {"left": 0.61, "top": 1.89, "width": 12.12, "height": 5.0},
        "content_area": {"left": 0.61, "top": 1.75, "width": 12.12, "height": 4.8},
        "left_column": {"left": 0.61, "top": 1.75, "width": 5.8, "height": 4.8},
        "right_column": {"left": 6.6, "top": 1.75, "width": 6.13, "height": 4.8},
        "chart": {"left": 0.61, "top": 1.65, "width": 12.12, "height": 4.8},
        "table": {"left": 0.61, "top": 1.75, "width": 12.12, "height": 4.6},
        "metrics_row": {"left": 0.61, "top": 2.8, "width": 12.12, "height": 2.2},
        "section_title": {"left": 0.61, "top": 3.0, "width": 12.12, "height": 1.5},
        "footer": {"left": 0.61, "top": 6.8, "width": 12.12, "height": 0.4},
    }

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

    # Font settings
    FONTS = {
        "title": {"name": "Arial", "size": Pt(28), "bold": True, "color": "text_dark"},
        "subtitle": {"name": "Arial", "size": Pt(18), "bold": False, "color": "text_body"},
        "section": {"name": "Arial", "size": Pt(36), "bold": True, "color": "white"},
        "heading": {"name": "Arial", "size": Pt(18), "bold": True, "color": "text_dark"},
        "body": {"name": "Arial", "size": Pt(15), "bold": False, "color": "text_body"},
        "body_large": {"name": "Arial", "size": Pt(16), "bold": False, "color": "text_body"},
        "bullet": {"name": "Arial", "size": Pt(16), "bold": False, "color": "text_body"},
        "caption": {"name": "Arial", "size": Pt(10), "bold": False, "color": "text_light"},
        "metric_value": {"name": "Arial", "size": Pt(36), "bold": True, "color": "white"},
        "metric_label": {"name": "Arial", "size": Pt(13), "bold": False, "color": "white"},
        "table_header": {"name": "Arial", "size": Pt(11), "bold": True, "color": "white"},
        "table_cell": {"name": "Arial", "size": Pt(11), "bold": False, "color": "text_body"},
        "chart_label": {"name": "Arial", "size": Pt(11), "bold": False, "color": "text_body"},
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
                        domain_stats = self.enhancer.get_domain_stats()
                        logger.info(f"LibraryEnhancer: {sum(domain_stats.values())} domain-tagged components")

                    # Load color palette and typography from library
                    self._load_library_styles()
                else:
                    logger.warning("ComponentLibrary index not found")
                    self.library = None
            except Exception as e:
                logger.warning(f"Could not initialize ComponentLibrary: {e}")
                self.library = None

    def _load_library_styles(self) -> None:
        """Load color palettes and typography from the library."""
        if not self.library:
            return

        # Load color palettes
        palettes = self.library.get_color_palettes()
        if palettes:
            palette = palettes[0]  # Use first palette
            self._apply_palette_to_colors(palette)
            logger.info(f"Loaded color palette from library")

        # Load typography presets
        presets = self.library.get_typography_presets()
        if presets:
            self._apply_typography_presets(presets)
            logger.info(f"Loaded {len(presets)} typography presets from library")

    def _apply_palette_to_colors(self, palette: dict) -> None:
        """Apply a library color palette to the COLORS dict."""
        color_mapping = {
            'primary': 'primary',
            'secondary': 'secondary',
            'accent1': 'accent1',
            'accent2': 'accent2',
            'accent3': 'accent2',  # Map to accent2 if no accent3
            'text': 'text_dark',
            'background': 'background',
        }

        for palette_key, colors_key in color_mapping.items():
            hex_color = palette.get(palette_key)
            if hex_color:
                try:
                    hex_color = hex_color.lstrip('#')
                    rgb = RGBColor(
                        int(hex_color[0:2], 16),
                        int(hex_color[2:4], 16),
                        int(hex_color[4:6], 16)
                    )
                    self.COLORS[colors_key] = rgb
                except (ValueError, IndexError):
                    pass

    def _apply_typography_presets(self, presets: list) -> None:
        """Apply typography presets to FONTS dict."""
        preset_mapping = {
            'title': 'title',
            'heading': 'heading',
            'body': 'body',
            'caption': 'caption',
        }

        for preset in presets:
            preset_type = preset.get('preset_type', '').lower()
            if preset_type in preset_mapping:
                fonts_key = preset_mapping[preset_type]
                font_name = preset.get('font_name')
                font_size = preset.get('font_size')
                is_bold = preset.get('is_bold', False)

                if font_name:
                    self.FONTS[fonts_key]['name'] = font_name
                if font_size:
                    self.FONTS[fonts_key]['size'] = Pt(font_size)
                self.FONTS[fonts_key]['bold'] = is_bold

    def _build_layout_index(self) -> None:
        """Build index of available layouts."""
        self.layouts = {}
        for layout in self._template_prs.slide_layouts:
            self.layouts[layout.name] = layout
        logger.info(f"Indexed {len(self.layouts)} layouts from template")

    # ==================== Component Library Integration ====================

    def find_library_chart(
        self,
        chart_type: str = "column",
        category: Optional[str] = None,
        min_series: int = 1
    ) -> Optional[dict]:
        """
        Find a matching chart component from the library.

        Args:
            chart_type: Type of chart ('column', 'bar', 'line', 'pie')
            category: Optional category filter
            min_series: Minimum number of data series

        Returns:
            Chart component metadata or None
        """
        if not self.library:
            return None

        # Map simple types to library chart types
        type_map = {
            "column": "COLUMN",
            "bar": "BAR",
            "line": "LINE",
            "pie": "PIE",
            "area": "AREA",
        }
        lib_type = type_map.get(chart_type.lower(), "COLUMN")

        # Map to category names
        category_map = {
            "column": "column_charts",
            "bar": "bar_charts",
            "line": "line_charts",
            "pie": "pie_charts",
            "area": "area_charts",
        }
        lib_category = category or category_map.get(chart_type.lower())

        results = self.library.search_charts(
            chart_type=lib_type,
            category=lib_category,
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
        """
        Find a matching table component from the library.

        Args:
            rows: Approximate number of rows needed
            cols: Approximate number of columns needed
            category: Optional category ('data_table', 'comparison_matrix', etc.)

        Returns:
            Table component metadata or None
        """
        if not self.library:
            return None

        # Search for tables with similar dimensions
        results = self.library.search_tables(
            category=category,
            min_rows=max(1, rows - 2),
            max_rows=rows + 5,
            min_cols=max(1, cols - 1),
            max_cols=cols + 2,
            limit=5
        )

        # Prefer exact or close match
        if results:
            # Sort by closeness to requested dimensions
            results.sort(key=lambda t: abs(t.get('rows', 0) - rows) + abs(t.get('cols', 0) - cols))
            return results[0]

        return None

    def get_library_chart_data(self, component_id: str) -> Optional[dict]:
        """Get chart data from a library component."""
        if not self.library:
            return None
        return self.library.get_chart_data(component_id)

    def get_library_table_data(self, component_id: str) -> Optional[dict]:
        """Get table data from a library component."""
        if not self.library:
            return None
        return self.library.get_table_data(component_id)

    def add_library_chart(
        self,
        slide: Slide,
        component_id: str,
        left: float,
        top: float,
        width: float,
        height: float,
        custom_data: Optional[dict] = None
    ) -> Optional[Any]:
        """
        Add a chart from the library to a slide.

        Args:
            slide: Target slide
            component_id: Library component ID
            left, top, width, height: Position and size in inches
            custom_data: Optional custom data to override library data

        Returns:
            Chart shape or None
        """
        if not self.library:
            return None

        return self.library.add_chart_from_library(
            slide,
            component_id,
            left=left,
            top=top,
            width=width,
            height=height,
            custom_data=custom_data
        )

    def add_library_table(
        self,
        slide: Slide,
        component_id: str,
        left: float,
        top: float,
        width: float,
        height: float,
        custom_data: Optional[list] = None
    ) -> Optional[Any]:
        """
        Add a table from the library to a slide.

        Args:
            slide: Target slide
            component_id: Library component ID
            left, top, width, height: Position and size in inches
            custom_data: Optional custom data to override library data

        Returns:
            Table shape or None
        """
        if not self.library:
            return None

        return self.library.add_table_from_library(
            slide,
            component_id,
            left=left,
            top=top,
            width=width,
            height=height,
            custom_data=custom_data
        )

    def add_library_image(
        self,
        slide: Slide,
        component_id: str,
        left: float,
        top: float,
        width: Optional[float] = None,
        height: Optional[float] = None
    ) -> Optional[Any]:
        """
        Add an image from the library to a slide.

        Args:
            slide: Target slide
            component_id: Library component ID
            left, top: Position in inches
            width, height: Optional size in inches

        Returns:
            Picture shape or None
        """
        if not self.library:
            return None

        return self.library.add_image_from_library(
            slide,
            component_id,
            left=left,
            top=top,
            width=width,
            height=height
        )

    def create_presentation(self) -> Presentation:
        """Create a new empty presentation with template layouts."""
        prs = Presentation(str(self.template_path))

        # Remove all existing slides from the template
        # We only want the layouts/styling, not the content slides
        while len(prs.slides) > 0:
            slide = prs.slides[0]
            rId = prs.part.relate_to(slide.part, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide')
            # Get the sldIdLst element and remove the first slide
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

        # First, fill placeholders from the master layout
        if use_placeholders:
            self._fill_placeholders(slide, content)

        # Then route to appropriate renderer for additional elements (charts, tables, etc.)
        renderer_method = getattr(self, f"_render_{slide_type}", None)
        if renderer_method:
            renderer_method(slide, content, skip_title=use_placeholders)
        elif not use_placeholders:
            self._render_default(slide, content)

        return slide

    def _fill_placeholders(self, slide: Slide, content: dict) -> None:
        """
        Fill placeholders from the master layout with content.

        This uses the exact positions defined in the master layout,
        eliminating the need for hardcoded positions.
        """
        for ph in slide.placeholders:
            ph_type = str(ph.placeholder_format.type).split('.')[-1].strip('()')
            ph_idx = ph.placeholder_format.idx

            # Find matching content
            text_content = None

            # Try direct idx match first
            if str(ph_idx) in content:
                text_content = content[str(ph_idx)]

            # Try type-based match
            if text_content is None:
                for type_key, content_keys in PLACEHOLDER_CONTENT_MAP.items():
                    if type_key in ph_type.upper():
                        for key in content_keys:
                            if key in content:
                                text_content = content[key]
                                break
                    if text_content:
                        break

            # Apply content to placeholder
            if text_content and ph.has_text_frame:
                self._set_placeholder_content(ph, text_content)

    def _set_placeholder_content(self, placeholder, content) -> None:
        """Set text in a placeholder, handling different content types."""
        tf = placeholder.text_frame

        if isinstance(content, str):
            # Simple text
            tf.paragraphs[0].text = content
        elif isinstance(content, list):
            # Bullet points
            for i, item in enumerate(content):
                if i == 0:
                    tf.paragraphs[0].text = f"{self.BULLET_CHAR}  {item}"
                else:
                    p = tf.add_paragraph()
                    p.text = f"{self.BULLET_CHAR}  {item}"
                    p.level = 0
        elif isinstance(content, dict):
            # Detailed content with formatting
            text = content.get("text", "")
            tf.paragraphs[0].text = text

            # Apply formatting if specified
            if "font_size_pt" in content:
                for p in tf.paragraphs:
                    for run in p.runs:
                        run.font.size = Pt(content["font_size_pt"])
            if "bold" in content:
                for p in tf.paragraphs:
                    for run in p.runs:
                        run.font.bold = content["bold"]

    def _render_title_slide(self, slide: Slide, content: dict, skip_title: bool = False) -> None:
        """Render a title/frontpage slide."""
        if skip_title:
            return  # Placeholders already filled

        title = content.get("title", "")
        subtitle = content.get("subtitle", "")

        # Set title placeholder
        if slide.shapes.title and title:
            slide.shapes.title.text = title
            self._apply_font(slide.shapes.title.text_frame.paragraphs[0], "title")

        # Find and set subtitle placeholder
        for shape in slide.placeholders:
            if shape.placeholder_format.type == 4:  # SUBTITLE
                shape.text = subtitle
                self._apply_font(shape.text_frame.paragraphs[0], "subtitle")
                break

    def _render_section_divider(self, slide: Slide, content: dict, skip_title: bool = False) -> None:
        """Render a section divider slide."""
        if skip_title:
            return  # Placeholders already filled

        title = content.get("title", "")

        # Set title placeholder
        if slide.shapes.title and title:
            slide.shapes.title.text = title
            para = slide.shapes.title.text_frame.paragraphs[0]
            self._apply_font(para, "section")
            para.alignment = PP_ALIGN.LEFT

    def _render_default(self, slide: Slide, content: dict, skip_title: bool = False) -> None:
        """Render a default content slide."""
        if skip_title:
            return  # Placeholders already filled - no additional elements needed

        title = content.get("title", "")
        body = content.get("body", "")
        bullets = content.get("bullets", [])

        # Set title
        if slide.shapes.title and title:
            slide.shapes.title.text = title
            self._apply_font(slide.shapes.title.text_frame.paragraphs[0], "title")

        # Add body/bullets content
        if body or bullets:
            pos = self.POSITIONS["content_area"]
            textbox = slide.shapes.add_textbox(
                Inches(pos["left"]), Inches(pos["top"]),
                Inches(pos["width"]), Inches(pos["height"])
            )
            tf = textbox.text_frame
            tf.word_wrap = True

            if body:
                p = tf.paragraphs[0]
                p.text = body
                self._apply_font(p, "body_large")

            for i, bullet in enumerate(bullets):
                if i == 0 and not body:
                    p = tf.paragraphs[0]
                else:
                    p = tf.add_paragraph()
                # Use bullet character instead of space indent
                p.text = f"{self.BULLET_CHAR}  {bullet}"
                p.level = 0
                self._apply_font(p, "bullet")
                p.space_before = Pt(10)
                p.space_after = Pt(6)

    def _render_title_content(self, slide: Slide, content: dict, skip_title: bool = False) -> None:
        """Render a title + content slide."""
        self._render_default(slide, content, skip_title)

    def _render_two_column(self, slide: Slide, content: dict, skip_title: bool = False) -> None:
        """Render a two-column comparison slide."""
        title = content.get("title", "")
        left_col = content.get("left_column", content.get("left", {}))
        right_col = content.get("right_column", content.get("right", {}))

        # Set title (only if not using placeholders)
        if not skip_title and slide.shapes.title and title:
            slide.shapes.title.text = title
            self._apply_font(slide.shapes.title.text_frame.paragraphs[0], "title")

        # Add columns (these are typically additional content not in placeholders)
        if left_col:
            self._add_column(slide, left_col, is_left=True)
        if right_col:
            self._add_column(slide, right_col, is_left=False)

    def _add_column(self, slide: Slide, col_content: dict, is_left: bool) -> None:
        """Add a column to a two-column slide."""
        pos_key = "left_column" if is_left else "right_column"
        pos = self.POSITIONS[pos_key]

        header = col_content.get("header", "")
        bullets = col_content.get("bullets", [])

        textbox = slide.shapes.add_textbox(
            Inches(pos["left"]), Inches(pos["top"]),
            Inches(pos["width"]), Inches(pos["height"])
        )
        tf = textbox.text_frame
        tf.word_wrap = True

        # Header with underline effect
        if header:
            p = tf.paragraphs[0]
            p.text = header
            self._apply_font(p, "heading")
            p.space_after = Pt(14)

        # Bullets with bullet character
        for i, bullet in enumerate(bullets):
            p = tf.add_paragraph() if (i > 0 or header) else tf.paragraphs[0]
            p.text = f"{self.BULLET_CHAR}  {bullet}"
            p.level = 0
            self._apply_font(p, "body")
            p.space_before = Pt(8)
            p.space_after = Pt(4)

    def _render_key_metrics(self, slide: Slide, content: dict, skip_title: bool = False) -> None:
        """Render a key metrics slide with KPI boxes."""
        title = content.get("title", "")
        metrics = content.get("metrics", [])

        # Set title (only if not using placeholders)
        if not skip_title and slide.shapes.title and title:
            slide.shapes.title.text = title
            self._apply_font(slide.shapes.title.text_frame.paragraphs[0], "title")

        if not metrics:
            return

        # Create metric boxes
        num_metrics = min(len(metrics), 5)
        total_width = 12.12
        box_margin = 0.2
        box_width = (total_width - (num_metrics - 1) * box_margin) / num_metrics
        box_height = 1.8

        start_left = 0.61
        top = 2.5

        for i, metric in enumerate(metrics[:5]):
            left = start_left + i * (box_width + box_margin)

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
            tf.paragraphs[0].alignment = PP_ALIGN.CENTER

            # Value
            p = tf.paragraphs[0]
            p.text = str(metric.get("value", ""))
            self._apply_font(p, "metric_value")

            # Label
            p2 = tf.add_paragraph()
            p2.text = metric.get("label", "")
            p2.alignment = PP_ALIGN.CENTER
            self._apply_font(p2, "metric_label")

    def _render_table_slide(self, slide: Slide, content: dict, skip_title: bool = False) -> None:
        """Render a slide with a data table."""
        title = content.get("title", "")
        headers = content.get("headers", [])
        data = content.get("data", [])
        library_component_id = content.get("library_component_id")
        use_library = content.get("use_library", False)  # Default off for tables to preserve styling
        table_category = content.get("table_category")  # 'data_table', 'comparison_matrix', etc.

        # Set title (only if not using placeholders)
        if not skip_title and slide.shapes.title and title:
            slide.shapes.title.text = title
            self._apply_font(slide.shapes.title.text_frame.paragraphs[0], "title")

        if not data and not library_component_id:
            return

        # Calculate dimensions
        rows = len(data) + (1 if headers else 0)
        cols = len(headers) if headers else (len(data[0]) if data else 0)

        pos = self.POSITIONS["table"]
        table_added = False

        # Prepare table data for library components
        table_data = [headers] + data if headers else data

        # Option 1: Use specific library component
        if library_component_id and self.library:
            table_shape = self.add_library_table(
                slide,
                library_component_id,
                left=pos["left"],
                top=pos["top"],
                width=pos["width"],
                height=pos["height"],
                custom_data=table_data
            )
            if table_shape:
                table_added = True
                logger.info(f"Used library table: {library_component_id}")
                self._style_table(table_shape.table, headers is not None)

        # Option 2: Auto-find matching library component using enhanced matching
        if not table_added and use_library and self.library and rows > 0 and cols > 0:
            lib_table_id = None

            # Try enhanced matching first
            if self.enhancer:
                # Extract domain hints from title
                domains = self.enhancer._match_domains(
                    self.enhancer._extract_keywords(title)
                )
                # Determine purpose based on content
                purpose = 'comparison' if table_category == 'comparison_matrix' else 'data_display'
                lib_table_id = self.enhancer.find_best_table(
                    rows=rows,
                    cols=cols,
                    purpose=purpose,
                    domains=domains if domains else None
                )

            # Fallback to basic matching
            if not lib_table_id:
                lib_table = self.find_library_table(
                    rows=rows,
                    cols=cols,
                    category=table_category
                )
                if lib_table:
                    lib_table_id = lib_table["id"]

            if lib_table_id:
                table_shape = self.add_library_table(
                    slide,
                    lib_table_id,
                    left=pos["left"],
                    top=pos["top"],
                    width=pos["width"],
                    height=pos["height"],
                    custom_data=table_data
                )
                if table_shape:
                    table_added = True
                    logger.info(f"Auto-matched library table: {lib_table_id} (structure: {rows}x{cols})")
                    self._style_table(table_shape.table, headers is not None)

        # Option 3: Create table from scratch (default)
        if not table_added and rows > 0 and cols > 0:
            self._create_styled_table(slide, headers, data, pos)

    def _create_styled_table(
        self,
        slide: Slide,
        headers: list,
        data: list,
        pos: dict
    ) -> None:
        """Create a styled table from scratch."""
        rows = len(data) + (1 if headers else 0)
        cols = len(headers) if headers else (len(data[0]) if data else 0)

        if rows == 0 or cols == 0:
            return

        # Calculate appropriate table height based on row count
        row_height = 0.38
        table_height = min(rows * row_height, pos["height"])

        table_shape = slide.shapes.add_table(
            rows, cols,
            Inches(pos["left"]), Inches(pos["top"]),
            Inches(pos["width"]), Inches(table_height)
        )
        table = table_shape.table

        # Style header row
        if headers:
            for i, header in enumerate(headers):
                cell = table.cell(0, i)
                cell.text = str(header)
                cell.fill.solid()
                cell.fill.fore_color.rgb = self.COLORS["primary"]
                cell.margin_top = Inches(0.05)
                cell.margin_bottom = Inches(0.05)
                para = cell.text_frame.paragraphs[0]
                para.font.color.rgb = self.COLORS["white"]
                para.font.bold = True
                para.font.size = Pt(11)
                para.font.name = "Arial"

        # Add data rows with proper alternating colors
        start_row = 1 if headers else 0
        for i, row in enumerate(data):
            for j, value in enumerate(row):
                cell = table.cell(start_row + i, j)
                cell.text = str(value)
                cell.margin_top = Inches(0.05)
                cell.margin_bottom = Inches(0.05)
                para = cell.text_frame.paragraphs[0]
                para.font.size = Pt(11)
                para.font.name = "Arial"
                para.font.color.rgb = self.COLORS["text_body"]

                # Alternate row colors: even rows white, odd rows gray
                if i % 2 == 0:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = self.COLORS["white"]
                else:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = self.COLORS["background"]

    def _style_table(self, table, has_header: bool = True) -> None:
        """Apply consistent styling to a table (e.g., from library)."""
        for row_idx in range(len(table.rows)):
            for col_idx in range(len(table.columns)):
                cell = table.cell(row_idx, col_idx)
                cell.margin_top = Inches(0.05)
                cell.margin_bottom = Inches(0.05)
                para = cell.text_frame.paragraphs[0]
                para.font.name = "Arial"

                if row_idx == 0 and has_header:
                    # Header row
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = self.COLORS["primary"]
                    para.font.color.rgb = self.COLORS["white"]
                    para.font.bold = True
                    para.font.size = Pt(11)
                else:
                    # Data rows
                    data_row = row_idx - 1 if has_header else row_idx
                    if data_row % 2 == 0:
                        cell.fill.solid()
                        cell.fill.fore_color.rgb = self.COLORS["white"]
                    else:
                        cell.fill.solid()
                        cell.fill.fore_color.rgb = self.COLORS["background"]
                    para.font.color.rgb = self.COLORS["text_body"]
                    para.font.size = Pt(11)

    def _render_data_chart(self, slide: Slide, content: dict, skip_title: bool = False) -> None:
        """Render a slide with a chart."""
        title = content.get("title", "")
        chart_data = content.get("chart_data", {})
        narrative = content.get("narrative", "")
        library_component_id = content.get("library_component_id")
        use_library = content.get("use_library", True)

        # Set title (only if not using placeholders)
        if not skip_title and slide.shapes.title and title:
            slide.shapes.title.text = title
            self._apply_font(slide.shapes.title.text_frame.paragraphs[0], "title")

        pos = self.POSITIONS["chart"]
        chart_added = False

        # Option 1: Use specific library component
        if library_component_id and self.library:
            chart_shape = self.add_library_chart(
                slide,
                library_component_id,
                left=pos["left"],
                top=pos["top"],
                width=pos["width"],
                height=pos["height"],
                custom_data={
                    "categories": chart_data.get("categories", []),
                    "series": chart_data.get("series", [])
                } if chart_data else None
            )
            if chart_shape:
                chart_added = True
                # Apply styling to library chart
                self._style_library_chart(chart_shape.chart, chart_data)
                logger.info(f"Used library chart: {library_component_id}")

        # Option 2: Auto-find a matching library component using enhanced matching
        if not chart_added and use_library and self.library and chart_data:
            chart_type = chart_data.get("type", "column")
            series = chart_data.get("series", [])
            categories = chart_data.get("categories", [])
            series_count = len(series)
            category_count = len(categories)

            # Try enhanced matching first (considers structure + domain)
            lib_chart_id = None
            if self.enhancer:
                # Extract domain hints from title
                domains = self.enhancer._match_domains(
                    self.enhancer._extract_keywords(title)
                )
                lib_chart_id = self.enhancer.find_best_chart(
                    chart_type=chart_type,
                    series_count=series_count,
                    category_count=category_count,
                    domains=domains if domains else None
                )

            # Fallback to basic matching
            if not lib_chart_id:
                lib_chart = self.find_library_chart(
                    chart_type=chart_type,
                    min_series=series_count
                )
                if lib_chart:
                    lib_chart_id = lib_chart["id"]

            if lib_chart_id:
                chart_shape = self.add_library_chart(
                    slide,
                    lib_chart_id,
                    left=pos["left"],
                    top=pos["top"],
                    width=pos["width"],
                    height=pos["height"],
                    custom_data={
                        "categories": categories,
                        "series": series
                    }
                )
                if chart_shape:
                    chart_added = True
                    # Apply styling to library chart
                    self._style_library_chart(chart_shape.chart, chart_data)
                    logger.info(f"Auto-matched library chart: {lib_chart_id} (structure: {series_count}x{category_count})")

        # Option 3: Fall back to creating chart from scratch
        if not chart_added:
            if chart_data:
                self._add_chart(slide, chart_data)
            else:
                self._add_chart_placeholder(slide)

        # Add narrative/source text
        if narrative:
            pos = self.POSITIONS["footer"]
            textbox = slide.shapes.add_textbox(
                Inches(pos["left"]), Inches(6.2),
                Inches(pos["width"]), Inches(0.5)
            )
            p = textbox.text_frame.paragraphs[0]
            p.text = narrative
            self._apply_font(p, "caption")

    def _add_chart(self, slide: Slide, chart_data: dict) -> None:
        """Add a chart to the slide."""
        from pptx.chart.data import CategoryChartData
        from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION

        chart_type = chart_data.get("type", "column")
        categories = chart_data.get("categories", [])
        series_list = chart_data.get("series", [])

        if not categories or not series_list:
            self._add_chart_placeholder(slide)
            return

        # Handle waterfall as a simulated visual (stacked bar)
        if chart_type == "waterfall":
            self._add_waterfall_chart(slide, chart_data)
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

        # Add chart - adjust size for pie to leave room for legend
        pos = self.POSITIONS["chart"]
        if chart_type == "pie":
            chart_width = pos["width"] * 0.7  # Smaller to leave legend room
        else:
            chart_width = pos["width"]

        chart_shape = slide.shapes.add_chart(
            xl_chart_type,
            Inches(pos["left"]), Inches(pos["top"]),
            Inches(chart_width), Inches(pos["height"]),
            data
        )
        chart = chart_shape.chart

        # Style the chart based on type
        if chart_type == "pie":
            self._style_pie_chart(chart, categories)
        else:
            # Style bar/column charts
            plot = chart.plots[0]
            if hasattr(plot, 'series'):
                for i, series in enumerate(plot.series):
                    series.format.fill.solid()
                    if i == 0:
                        series.format.fill.fore_color.rgb = self.COLORS["primary"]
                    else:
                        series.format.fill.fore_color.rgb = self.COLORS["secondary"]

            # Add data labels to column/bar charts
            if chart_type in ["column", "bar"]:
                try:
                    plot.has_data_labels = True
                    data_labels = plot.data_labels
                    data_labels.font.size = Pt(10)
                    data_labels.font.color.rgb = self.COLORS["text_body"]
                    data_labels.number_format = '0.0'
                except Exception:
                    pass  # Data labels not supported for this chart type

    def _style_library_chart(self, chart, chart_data: dict) -> None:
        """Apply consistent styling to a chart from the library."""
        from pptx.enum.chart import XL_LEGEND_POSITION

        chart_type = chart_data.get("type", "column") if chart_data else "column"
        categories = chart_data.get("categories", []) if chart_data else []

        try:
            # Style based on chart type
            if chart_type == "pie":
                # Add legend for pie charts
                chart.has_legend = True
                chart.legend.position = XL_LEGEND_POSITION.RIGHT
                chart.legend.include_in_layout = False
                chart.legend.font.size = Pt(10)

                # Apply colors
                pie_colors = [
                    self.COLORS["primary"],
                    self.COLORS["secondary"],
                    self.COLORS["accent1"],
                    self.COLORS["accent2"],
                    RGBColor(0x80, 0x80, 0x80),  # Gray
                ]
                if chart.series:
                    series = chart.series[0]
                    for i, point in enumerate(series.points):
                        point.format.fill.solid()
                        point.format.fill.fore_color.rgb = pie_colors[i % len(pie_colors)]

            else:
                # Column/bar/line charts
                # Add data labels
                for plot in chart.plots:
                    try:
                        plot.has_data_labels = True
                        data_labels = plot.data_labels
                        data_labels.font.size = Pt(9)
                        data_labels.font.color.rgb = self.COLORS["text_body"]
                    except Exception:
                        pass

                # Apply primary color to first series
                if chart.series:
                    series = chart.series[0]
                    series.format.fill.solid()
                    series.format.fill.fore_color.rgb = self.COLORS["primary"]

        except Exception as e:
            logger.warning(f"Could not style library chart: {e}")

    def _style_pie_chart(self, chart, categories: list) -> None:
        """Style a pie chart with custom colors and legend."""
        from pptx.enum.chart import XL_LEGEND_POSITION

        pie_colors = [
            self.COLORS["primary"],
            self.COLORS["secondary"],
            RGBColor(0x00, 0xB0, 0x50),  # Green
            RGBColor(0xFF, 0xB8, 0x00),  # Orange
            RGBColor(0x7C, 0x4D, 0xC4),  # Purple
            RGBColor(0x05, 0x1C, 0x2C),  # Dark blue
        ]

        plot = chart.plots[0]
        if hasattr(plot, 'series') and len(plot.series) > 0:
            series = plot.series[0]
            for i, point in enumerate(series.points):
                point.format.fill.solid()
                point.format.fill.fore_color.rgb = pie_colors[i % len(pie_colors)]

        # Add legend to pie chart
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.RIGHT
        chart.legend.include_in_layout = False
        chart.legend.font.size = Pt(10)

    def _add_waterfall_chart(self, slide: Slide, chart_data: dict) -> None:
        """Add a waterfall chart visualization using shapes."""
        categories = chart_data.get("categories", [])
        series = chart_data.get("series", [{}])[0]
        values = series.get("values", [])

        if not categories or not values:
            self._add_chart_placeholder(slide)
            return

        pos = self.POSITIONS["chart"]
        left_start = pos["left"]
        top = pos["top"] + 0.3  # Add some top padding
        chart_width = pos["width"]
        chart_height = pos["height"] - 0.8  # Leave room for labels

        # Calculate dimensions
        num_bars = len(categories)
        bar_width = (chart_width * 0.75) / num_bars
        bar_gap = (chart_width * 0.25) / (num_bars + 1)

        # Find max value for scaling - use cumulative max for waterfall
        cumulative = 0
        max_cumulative = 0
        for val in values[:-1]:  # Exclude final total
            cumulative += val
            max_cumulative = max(max_cumulative, cumulative)
        max_cumulative = max(max_cumulative, values[-1])  # Compare with final total

        # Chart area dimensions
        chart_bottom = top + chart_height * 0.75
        chart_top_area = top + chart_height * 0.1
        usable_height = chart_bottom - chart_top_area

        # Scale factor
        scale = usable_height / max_cumulative if max_cumulative > 0 else 1

        # Draw baseline
        baseline = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(left_start), Inches(chart_bottom),
            Inches(chart_width), Inches(0.02)
        )
        baseline.fill.solid()
        baseline.fill.fore_color.rgb = self.COLORS["text_light"]
        baseline.line.fill.background()

        # Running position for waterfall effect
        current_top = chart_bottom

        for i, (cat, val) in enumerate(zip(categories, values)):
            x = left_start + bar_gap + i * (bar_width + bar_gap)
            bar_height = abs(val) * scale

            # For final bar (total), start from baseline
            if i == len(categories) - 1:
                bar_top = chart_bottom - bar_height
                bar_color = self.COLORS["accent1"]  # Dark for total
            else:
                # Waterfall: each bar stacks on previous
                bar_top = current_top - bar_height
                bar_color = self.COLORS["primary"] if val >= 0 else self.COLORS["secondary"]
                current_top = bar_top  # Update for next bar

            # Ensure minimum bar height for visibility
            bar_height = max(bar_height, 0.15)

            # Add bar
            shape = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                Inches(x), Inches(bar_top),
                Inches(bar_width), Inches(bar_height)
            )
            shape.fill.solid()
            shape.fill.fore_color.rgb = bar_color
            shape.line.fill.background()

            # Add value label above bar
            label_box = slide.shapes.add_textbox(
                Inches(x - 0.1), Inches(bar_top - 0.28),
                Inches(bar_width + 0.2), Inches(0.25)
            )
            tf = label_box.text_frame
            p = tf.paragraphs[0]
            if isinstance(val, float):
                p.text = f"{val:.1f}%"
            else:
                p.text = str(val)
            p.alignment = PP_ALIGN.CENTER
            p.font.size = Pt(11)
            p.font.bold = True
            p.font.color.rgb = self.COLORS["text_dark"]

            # Add category label below chart
            cat_box = slide.shapes.add_textbox(
                Inches(x - 0.15), Inches(chart_bottom + 0.08),
                Inches(bar_width + 0.3), Inches(0.6)
            )
            tf2 = cat_box.text_frame
            tf2.word_wrap = True
            p2 = tf2.paragraphs[0]
            p2.text = cat  # Full category name
            p2.alignment = PP_ALIGN.CENTER
            p2.font.size = Pt(9)
            p2.font.color.rgb = self.COLORS["text_body"]

        # Add connector lines between bars (except to last bar)
        current_top = chart_bottom
        for i, val in enumerate(values[:-1]):
            bar_height = abs(val) * scale
            bar_top = current_top - bar_height
            current_top = bar_top

            # Draw horizontal connector to next bar
            x_start = left_start + bar_gap + i * (bar_width + bar_gap) + bar_width
            x_end = left_start + bar_gap + (i + 1) * (bar_width + bar_gap)

            connector = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                Inches(x_start), Inches(bar_top),
                Inches(x_end - x_start), Inches(0.015)
            )
            connector.fill.solid()
            connector.fill.fore_color.rgb = self.COLORS["text_light"]
            connector.line.fill.background()

    def _add_chart_placeholder(self, slide: Slide) -> None:
        """Add a placeholder for chart data."""
        pos = self.POSITIONS["chart"]
        shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(pos["left"]), Inches(pos["top"]),
            Inches(pos["width"]), Inches(pos["height"])
        )
        shape.fill.solid()
        shape.fill.fore_color.rgb = self.COLORS["light_gray"]
        shape.line.color.rgb = self.COLORS["text_light"]

        tf = shape.text_frame
        p = tf.paragraphs[0]
        p.text = "[Chart Placeholder]"
        p.alignment = PP_ALIGN.CENTER
        self._apply_font(p, "body")

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


def main():
    """Test the template renderer."""
    import argparse

    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

    parser = argparse.ArgumentParser(description="Template Renderer Test")
    parser.add_argument(
        "--template",
        default="C:/Users/wpcol/claudecode/pptx-design/pptx_templates/pptx_template_business_case/template_business_case.pptx",
        help="Path to template"
    )
    parser.add_argument(
        "--output",
        default="C:/Users/wpcol/claudecode/pptx-design/pptx_generator/output/template_render_test.pptx",
        help="Output file"
    )

    args = parser.parse_args()

    renderer = TemplateRenderer(args.template)
    prs = renderer.create_presentation()

    # Delete existing slides (keep first 5 as reference)
    while len(prs.slides) > 0:
        rId = prs.slides._sldIdLst[0].rId
        prs.part.drop_rel(rId)
        del prs.slides._sldIdLst[0]

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
