"""
Style Guide Configuration Module

Centralizes style guide specifications for use throughout the pptx_generator.
Loads specifications from style guide markdown files and provides easy access
to formatting values for charts, tables, typography, and colors.

Usage:
    from pptx_generator.modules.style_guide_config import StyleGuideConfig, get_style_config

    # Get current style configuration
    config = get_style_config()

    # Access chart settings
    gridline_color = config.chart.gridline_color
    gridline_width = config.chart.gridline_width_pt

    # Access table settings
    header_color = config.table.header_color
    cell_margin_lr = config.table.cell_margin_lr_inches
"""

import json
import logging
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt

logger = logging.getLogger(__name__)


def hex_to_rgb(hex_color: str) -> RGBColor:
    """Convert hex color string to RGBColor."""
    hex_color = hex_color.lstrip('#')
    r = int(hex_color[0:2], 16)
    g = int(hex_color[2:4], 16)
    b = int(hex_color[4:6], 16)
    return RGBColor(r, g, b)


@dataclass
class ChartStyleConfig:
    """Chart formatting specifications."""
    # Gridlines
    gridline_enabled: bool = True
    gridline_width_pt: float = 0.5
    gridline_color: str = "#D9D9D9"

    # Tick marks
    value_axis_tick_mark: str = "NONE"
    category_axis_tick_mark: str = "NONE"

    # Data labels
    data_label_font_size_pt: float = 10.0
    data_label_font_name: str = "Arial"

    # Axis labels
    axis_label_font_size_pt: float = 12.0
    axis_label_font_name: str = "Arial"
    axis_label_color: str = "#061F32"

    # Series colors
    primary_series_color: str = "#051C2C"
    secondary_series_color: str = "#4A90A4"
    tertiary_series_color: str = "#7FB3D5"

    # Line chart specific
    line_width_pt: float = 3.0

    @property
    def gridline_rgb(self) -> RGBColor:
        return hex_to_rgb(self.gridline_color)

    @property
    def primary_series_rgb(self) -> RGBColor:
        return hex_to_rgb(self.primary_series_color)

    @property
    def secondary_series_rgb(self) -> RGBColor:
        return hex_to_rgb(self.secondary_series_color)

    @property
    def axis_label_rgb(self) -> RGBColor:
        return hex_to_rgb(self.axis_label_color)

    @property
    def gridline_width(self):
        """Return gridline width as Pt for python-pptx."""
        return Pt(self.gridline_width_pt)

    @property
    def line_width(self):
        """Return line width as Pt for python-pptx."""
        return Pt(self.line_width_pt)


@dataclass
class TableStyleConfig:
    """Table formatting specifications."""
    # Header row
    header_color: str = "#051C2C"
    header_text_color: str = "#FFFFFF"
    header_font_bold: bool = True
    header_font_size_pt: float = 11.0
    header_font_name: str = "Arial"

    # Data rows (alternating)
    row_odd_color: str = "#FFFFFF"
    row_even_color: str = "#F5F5F5"
    row_text_color: str = "#061F32"
    row_font_size_pt: float = 10.0
    row_font_name: str = "Arial"

    # Cell margins
    cell_margin_lr_inches: float = 0.1
    cell_margin_tb_inches: float = 0.05

    # Borders
    border_width_pt: float = 0.0  # No visible borders
    border_color: str = "#FFFFFF"

    @property
    def header_rgb(self) -> RGBColor:
        return hex_to_rgb(self.header_color)

    @property
    def header_text_rgb(self) -> RGBColor:
        return hex_to_rgb(self.header_text_color)

    @property
    def row_odd_rgb(self) -> RGBColor:
        return hex_to_rgb(self.row_odd_color)

    @property
    def row_even_rgb(self) -> RGBColor:
        return hex_to_rgb(self.row_even_color)

    @property
    def row_text_rgb(self) -> RGBColor:
        return hex_to_rgb(self.row_text_color)

    @property
    def cell_margin_lr(self):
        """Return left/right margin as Inches for python-pptx."""
        return Inches(self.cell_margin_lr_inches)

    @property
    def cell_margin_tb(self):
        """Return top/bottom margin as Inches for python-pptx."""
        return Inches(self.cell_margin_tb_inches)


@dataclass
class MetricBoxStyleConfig:
    """Metric/KPI box formatting specifications."""
    background_color: str = "#051C2C"
    value_font_size_pt: float = 28.0
    value_font_bold: bool = True
    value_text_color: str = "#FFFFFF"
    label_font_size_pt: float = 14.0
    label_font_bold: bool = False
    label_text_color: str = "#FFFFFF"
    font_name: str = "Arial"
    corner_radius: float = 0.1  # inches

    @property
    def background_rgb(self) -> RGBColor:
        return hex_to_rgb(self.background_color)

    @property
    def value_text_rgb(self) -> RGBColor:
        return hex_to_rgb(self.value_text_color)

    @property
    def label_text_rgb(self) -> RGBColor:
        return hex_to_rgb(self.label_text_color)


@dataclass
class TypographyConfig:
    """Typography specifications."""
    # Font families
    primary_font: str = "Arial"
    secondary_font: str = "Arial"

    # Title
    title_size_pt: float = 32.0
    title_bold: bool = True
    title_color: str = "#061F32"

    # Subtitle/Thesis
    subtitle_size_pt: float = 18.0
    subtitle_color: str = "#061F32"

    # Body/Bullets
    body_size_pt: float = 14.0
    body_color: str = "#061F32"
    bullet_char: str = "•"

    # Section label
    section_label_size_pt: float = 9.0
    section_label_color: str = "#A6A6A6"

    # Source/Footer
    source_size_pt: float = 6.0
    source_color: str = "#A6A6A6"
    footer_size_pt: float = 8.0
    footer_color: str = "#A6A6A6"

    @property
    def title_rgb(self) -> RGBColor:
        return hex_to_rgb(self.title_color)

    @property
    def body_rgb(self) -> RGBColor:
        return hex_to_rgb(self.body_color)

    @property
    def section_label_rgb(self) -> RGBColor:
        return hex_to_rgb(self.section_label_color)

    @property
    def source_rgb(self) -> RGBColor:
        return hex_to_rgb(self.source_color)


@dataclass
class ColorPaletteConfig:
    """Color palette specifications."""
    primary_navy: str = "#051C2C"
    body_text_dark: str = "#061F32"
    white: str = "#FFFFFF"
    light_gray: str = "#F5F5F5"
    medium_gray: str = "#A6A6A6"
    gridline_gray: str = "#D9D9D9"

    @property
    def primary_navy_rgb(self) -> RGBColor:
        return hex_to_rgb(self.primary_navy)

    @property
    def body_text_dark_rgb(self) -> RGBColor:
        return hex_to_rgb(self.body_text_dark)

    @property
    def white_rgb(self) -> RGBColor:
        return hex_to_rgb(self.white)

    @property
    def light_gray_rgb(self) -> RGBColor:
        return hex_to_rgb(self.light_gray)

    @property
    def medium_gray_rgb(self) -> RGBColor:
        return hex_to_rgb(self.medium_gray)


@dataclass
class SlideLayoutConfig:
    """Slide layout specifications."""
    # Slide dimensions
    width_inches: float = 11.0
    height_inches: float = 8.5

    # Margins
    margin_left_inches: float = 0.4
    margin_right_inches: float = 0.4
    margin_top_inches: float = 0.4
    margin_bottom_inches: float = 0.5

    # Content area positions
    title_top_inches: float = 0.4
    content_top_inches: float = 2.8
    footer_top_inches: float = 7.5


@dataclass
class StyleGuideConfig:
    """Complete style guide configuration."""
    version: str = "2026.01.01"
    name: str = "PCCP CS Style Guide"

    # Component configs
    chart: ChartStyleConfig = field(default_factory=ChartStyleConfig)
    table: TableStyleConfig = field(default_factory=TableStyleConfig)
    metric_box: MetricBoxStyleConfig = field(default_factory=MetricBoxStyleConfig)
    typography: TypographyConfig = field(default_factory=TypographyConfig)
    colors: ColorPaletteConfig = field(default_factory=ColorPaletteConfig)
    layout: SlideLayoutConfig = field(default_factory=SlideLayoutConfig)

    @classmethod
    def from_style_guide_file(cls, style_guide_path: Path) -> "StyleGuideConfig":
        """
        Load configuration from a style guide markdown file.

        Currently uses default values; future versions could parse the markdown.
        """
        config = cls()

        # Extract version from filename
        if "style_guide" in style_guide_path.stem:
            parts = style_guide_path.stem.split("_")
            if len(parts) >= 4:
                # e.g., pccp_cs_style_guide_2026.01.01
                config.version = parts[-1]

        logger.info(f"Loaded style guide config version: {config.version}")
        return config

    @classmethod
    def from_json(cls, json_path: Path) -> "StyleGuideConfig":
        """Load configuration from a JSON file."""
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)

        config = cls()
        config.version = data.get("version", config.version)
        config.name = data.get("name", config.name)

        # Load chart config
        if "chart" in data:
            for key, value in data["chart"].items():
                if hasattr(config.chart, key):
                    setattr(config.chart, key, value)

        # Load table config
        if "table" in data:
            for key, value in data["table"].items():
                if hasattr(config.table, key):
                    setattr(config.table, key, value)

        # Load metric_box config
        if "metric_box" in data:
            for key, value in data["metric_box"].items():
                if hasattr(config.metric_box, key):
                    setattr(config.metric_box, key, value)

        # Load typography config
        if "typography" in data:
            for key, value in data["typography"].items():
                if hasattr(config.typography, key):
                    setattr(config.typography, key, value)

        # Load colors config
        if "colors" in data:
            for key, value in data["colors"].items():
                if hasattr(config.colors, key):
                    setattr(config.colors, key, value)

        # Load layout config
        if "layout" in data:
            for key, value in data["layout"].items():
                if hasattr(config.layout, key):
                    setattr(config.layout, key, value)

        return config

    def to_json(self, json_path: Path) -> None:
        """Save configuration to a JSON file."""
        data = {
            "version": self.version,
            "name": self.name,
            "chart": {
                "gridline_enabled": self.chart.gridline_enabled,
                "gridline_width_pt": self.chart.gridline_width_pt,
                "gridline_color": self.chart.gridline_color,
                "value_axis_tick_mark": self.chart.value_axis_tick_mark,
                "category_axis_tick_mark": self.chart.category_axis_tick_mark,
                "data_label_font_size_pt": self.chart.data_label_font_size_pt,
                "axis_label_font_size_pt": self.chart.axis_label_font_size_pt,
                "axis_label_color": self.chart.axis_label_color,
                "primary_series_color": self.chart.primary_series_color,
                "secondary_series_color": self.chart.secondary_series_color,
                "line_width_pt": self.chart.line_width_pt,
            },
            "table": {
                "header_color": self.table.header_color,
                "header_text_color": self.table.header_text_color,
                "header_font_bold": self.table.header_font_bold,
                "header_font_size_pt": self.table.header_font_size_pt,
                "row_odd_color": self.table.row_odd_color,
                "row_even_color": self.table.row_even_color,
                "row_text_color": self.table.row_text_color,
                "row_font_size_pt": self.table.row_font_size_pt,
                "cell_margin_lr_inches": self.table.cell_margin_lr_inches,
                "cell_margin_tb_inches": self.table.cell_margin_tb_inches,
            },
            "metric_box": {
                "background_color": self.metric_box.background_color,
                "value_font_size_pt": self.metric_box.value_font_size_pt,
                "label_font_size_pt": self.metric_box.label_font_size_pt,
            },
            "typography": {
                "primary_font": self.typography.primary_font,
                "title_size_pt": self.typography.title_size_pt,
                "body_size_pt": self.typography.body_size_pt,
                "source_size_pt": self.typography.source_size_pt,
            },
            "colors": {
                "primary_navy": self.colors.primary_navy,
                "body_text_dark": self.colors.body_text_dark,
                "white": self.colors.white,
                "light_gray": self.colors.light_gray,
                "medium_gray": self.colors.medium_gray,
                "gridline_gray": self.colors.gridline_gray,
            },
            "layout": {
                "width_inches": self.layout.width_inches,
                "height_inches": self.layout.height_inches,
                "margin_left_inches": self.layout.margin_left_inches,
                "margin_right_inches": self.layout.margin_right_inches,
            },
        }

        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2)


# Global style config instance
_style_config: Optional[StyleGuideConfig] = None


def get_style_config(version: Optional[str] = None) -> StyleGuideConfig:
    """
    Get the current style guide configuration.

    Args:
        version: Specific version to load (e.g., "2026.01.01").
                 If None, uses the latest available.

    Returns:
        StyleGuideConfig instance
    """
    global _style_config

    if _style_config is not None and (version is None or _style_config.version == version):
        return _style_config

    # Find style guides directory
    style_guides_dir = Path(__file__).parent.parent.parent / "style_guides"

    if not style_guides_dir.exists():
        logger.warning(f"Style guides directory not found: {style_guides_dir}")
        _style_config = StyleGuideConfig()
        return _style_config

    # Check for JSON config first (more precise)
    if version:
        json_path = style_guides_dir / f"pccp_cs_style_guide_{version}.json"
        md_path = style_guides_dir / f"pccp_cs_style_guide_{version}.md"
    else:
        # Find latest
        json_files = list(style_guides_dir.glob("pccp_cs_style_guide_*.json"))
        md_files = list(style_guides_dir.glob("pccp_cs_style_guide_*.md"))

        if json_files:
            json_files.sort(key=lambda p: p.stem, reverse=True)
            json_path = json_files[0]
        else:
            json_path = None

        if md_files:
            md_files.sort(key=lambda p: p.stem, reverse=True)
            md_path = md_files[0]
        else:
            md_path = None

    # Load from JSON if available, otherwise from markdown
    if json_path and json_path.exists():
        _style_config = StyleGuideConfig.from_json(json_path)
    elif md_path and md_path.exists():
        _style_config = StyleGuideConfig.from_style_guide_file(md_path)
    else:
        logger.warning("No style guide found, using defaults")
        _style_config = StyleGuideConfig()

    return _style_config


def reset_style_config() -> None:
    """Reset the global style config (useful for testing)."""
    global _style_config
    _style_config = None


def create_style_guide_json(output_path: Optional[Path] = None) -> Path:
    """
    Create a JSON version of the current style guide for easier editing.

    Args:
        output_path: Path to save the JSON file.
                     If None, saves to style_guides/pccp_cs_style_guide_current.json

    Returns:
        Path to the created JSON file
    """
    config = StyleGuideConfig()

    if output_path is None:
        style_guides_dir = Path(__file__).parent.parent.parent / "style_guides"
        style_guides_dir.mkdir(exist_ok=True)
        output_path = style_guides_dir / f"pccp_cs_style_guide_{config.version}.json"

    config.to_json(output_path)
    logger.info(f"Created style guide JSON: {output_path}")
    return output_path


if __name__ == "__main__":
    # Create JSON version of the style guide
    logging.basicConfig(level=logging.INFO)
    path = create_style_guide_json()
    print(f"Created: {path}")

    # Test loading
    config = get_style_config()
    print(f"\nLoaded style guide: {config.name} v{config.version}")
    print(f"  Chart gridline: {config.chart.gridline_width_pt}pt, {config.chart.gridline_color}")
    print(f"  Table header: {config.table.header_color}")
    print(f"  Table margins: {config.table.cell_margin_lr_inches}\" LR, {config.table.cell_margin_tb_inches}\" TB")
