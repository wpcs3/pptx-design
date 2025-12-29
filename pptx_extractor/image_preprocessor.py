"""
Image Preprocessing Module for Vision-Based Slide Analysis

Optimizes images for LLM vision models with:
- Measurement grid overlay for calibration
- Contrast/brightness normalization
- Optimal resolution scaling
- Edge detection for shape boundaries
- Color analysis preprocessing
"""

import io
import logging
from pathlib import Path
from typing import Optional, Tuple, Dict, Any
from dataclasses import dataclass

import numpy as np
from PIL import Image, ImageDraw, ImageFont, ImageEnhance, ImageFilter

logger = logging.getLogger(__name__)

# Standard PowerPoint dimensions
SLIDE_WIDTH_INCHES = 13.333
SLIDE_HEIGHT_INCHES = 7.5
OPTIMAL_WIDTH_PX = 1920  # Optimal for vision models
OPTIMAL_HEIGHT_PX = 1080


@dataclass
class PreprocessingConfig:
    """Configuration for image preprocessing."""
    add_grid: bool = True
    grid_spacing_inches: float = 1.0
    grid_color: str = "#FF000040"  # Semi-transparent red
    grid_label_color: str = "#FF0000"

    add_rulers: bool = True
    ruler_color: str = "#333333"
    ruler_text_color: str = "#000000"

    normalize_contrast: bool = True
    target_contrast: float = 1.2

    normalize_brightness: bool = True
    target_brightness: float = 1.0

    add_edge_overlay: bool = False
    edge_color: str = "#00FF0080"  # Semi-transparent green

    scale_to_optimal: bool = True
    optimal_width: int = OPTIMAL_WIDTH_PX
    optimal_height: int = OPTIMAL_HEIGHT_PX

    output_format: str = "PNG"
    output_quality: int = 95


class ImagePreprocessor:
    """
    Preprocesses slide images for optimal LLM vision analysis.
    """

    def __init__(self, config: Optional[PreprocessingConfig] = None):
        """Initialize preprocessor with configuration."""
        self.config = config or PreprocessingConfig()
        self._font = None

    def _get_font(self, size: int = 12) -> ImageFont.FreeTypeFont:
        """Get a font for drawing text."""
        try:
            return ImageFont.truetype("arial.ttf", size)
        except:
            try:
                return ImageFont.truetype("DejaVuSans.ttf", size)
            except:
                return ImageFont.load_default()

    def preprocess(
        self,
        image_path: Path,
        output_path: Optional[Path] = None,
        config: Optional[PreprocessingConfig] = None
    ) -> Tuple[Path, Dict[str, Any]]:
        """
        Preprocess an image for vision analysis.

        Args:
            image_path: Path to the source image
            output_path: Optional output path (defaults to temp file)
            config: Optional override configuration

        Returns:
            Tuple of (output_path, preprocessing_metadata)
        """
        cfg = config or self.config
        image_path = Path(image_path)

        if not image_path.exists():
            raise FileNotFoundError(f"Image not found: {image_path}")

        # Load image
        img = Image.open(image_path)
        original_size = img.size

        # Convert to RGB if necessary
        if img.mode != 'RGB':
            img = img.convert('RGB')

        metadata = {
            "original_size": original_size,
            "original_path": str(image_path),
            "preprocessing_applied": []
        }

        # Scale to optimal resolution
        if cfg.scale_to_optimal:
            img = self._scale_image(img, cfg.optimal_width, cfg.optimal_height)
            metadata["preprocessing_applied"].append("scaled_to_optimal")
            metadata["scaled_size"] = img.size

        # Normalize contrast
        if cfg.normalize_contrast:
            img = self._normalize_contrast(img, cfg.target_contrast)
            metadata["preprocessing_applied"].append("contrast_normalized")

        # Normalize brightness
        if cfg.normalize_brightness:
            img = self._normalize_brightness(img, cfg.target_brightness)
            metadata["preprocessing_applied"].append("brightness_normalized")

        # Add edge overlay
        if cfg.add_edge_overlay:
            img = self._add_edge_overlay(img, cfg.edge_color)
            metadata["preprocessing_applied"].append("edge_overlay")

        # Add measurement grid
        if cfg.add_grid:
            img = self._add_measurement_grid(
                img,
                cfg.grid_spacing_inches,
                cfg.grid_color,
                cfg.grid_label_color
            )
            metadata["preprocessing_applied"].append("measurement_grid")

        # Add rulers
        if cfg.add_rulers:
            img = self._add_rulers(img, cfg.ruler_color, cfg.ruler_text_color)
            metadata["preprocessing_applied"].append("rulers")

        # Determine output path
        if output_path is None:
            output_path = image_path.parent / f"{image_path.stem}_preprocessed.png"

        # Save
        img.save(output_path, cfg.output_format, quality=cfg.output_quality)
        metadata["output_path"] = str(output_path)
        metadata["output_size"] = img.size

        logger.info(f"Preprocessed image saved to: {output_path}")
        return output_path, metadata

    def preprocess_for_comparison(
        self,
        original_path: Path,
        generated_path: Path,
        output_dir: Optional[Path] = None
    ) -> Tuple[Path, Path, Dict[str, Any]]:
        """
        Preprocess two images for side-by-side comparison.

        Args:
            original_path: Path to original template image
            generated_path: Path to generated recreation image
            output_dir: Directory for output files

        Returns:
            Tuple of (original_preprocessed_path, generated_preprocessed_path, metadata)
        """
        if output_dir is None:
            output_dir = Path(original_path).parent

        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)

        # Use a lighter config for comparison (no grid on compared images)
        comparison_config = PreprocessingConfig(
            add_grid=False,
            add_rulers=False,
            normalize_contrast=True,
            normalize_brightness=True,
            add_edge_overlay=False,
            scale_to_optimal=True
        )

        orig_out = output_dir / f"original_preprocessed.png"
        gen_out = output_dir / f"generated_preprocessed.png"

        _, orig_meta = self.preprocess(original_path, orig_out, comparison_config)
        _, gen_meta = self.preprocess(generated_path, gen_out, comparison_config)

        # Create side-by-side comparison
        side_by_side_path = self._create_side_by_side(
            orig_out, gen_out, output_dir / "comparison_side_by_side.png"
        )

        metadata = {
            "original": orig_meta,
            "generated": gen_meta,
            "side_by_side": str(side_by_side_path)
        }

        return orig_out, gen_out, metadata

    def _scale_image(
        self,
        img: Image.Image,
        target_width: int,
        target_height: int
    ) -> Image.Image:
        """Scale image while maintaining aspect ratio."""
        # Calculate scaling factor
        width_ratio = target_width / img.width
        height_ratio = target_height / img.height
        ratio = min(width_ratio, height_ratio)

        new_width = int(img.width * ratio)
        new_height = int(img.height * ratio)

        return img.resize((new_width, new_height), Image.Resampling.LANCZOS)

    def _normalize_contrast(self, img: Image.Image, factor: float) -> Image.Image:
        """Normalize image contrast."""
        enhancer = ImageEnhance.Contrast(img)
        return enhancer.enhance(factor)

    def _normalize_brightness(self, img: Image.Image, factor: float) -> Image.Image:
        """Normalize image brightness."""
        enhancer = ImageEnhance.Brightness(img)
        return enhancer.enhance(factor)

    def _add_edge_overlay(self, img: Image.Image, edge_color: str) -> Image.Image:
        """Add edge detection overlay to help identify shape boundaries."""
        # Convert to grayscale for edge detection
        gray = img.convert('L')

        # Apply edge detection
        edges = gray.filter(ImageFilter.FIND_EDGES)

        # Create colored edge overlay
        edge_rgba = Image.new('RGBA', img.size, (0, 0, 0, 0))

        # Parse color
        r, g, b, a = self._parse_color(edge_color)

        # Apply edges as overlay
        edges_array = np.array(edges)
        edge_mask = edges_array > 30  # Threshold

        overlay_array = np.zeros((*img.size[::-1], 4), dtype=np.uint8)
        overlay_array[edge_mask] = [r, g, b, a]

        edge_overlay = Image.fromarray(overlay_array, 'RGBA')

        # Composite
        result = img.convert('RGBA')
        result = Image.alpha_composite(result, edge_overlay)
        return result.convert('RGB')

    def _add_measurement_grid(
        self,
        img: Image.Image,
        spacing_inches: float,
        grid_color: str,
        label_color: str
    ) -> Image.Image:
        """Add a measurement grid overlay."""
        # Calculate pixels per inch
        ppi_x = img.width / SLIDE_WIDTH_INCHES
        ppi_y = img.height / SLIDE_HEIGHT_INCHES

        # Create overlay
        overlay = Image.new('RGBA', img.size, (0, 0, 0, 0))
        draw = ImageDraw.Draw(overlay)

        # Parse colors
        grid_rgba = self._parse_color(grid_color)

        # Draw vertical lines
        x = 0
        inch = 0
        while x < img.width:
            x = int(inch * ppi_x)
            if x < img.width:
                draw.line([(x, 0), (x, img.height)], fill=grid_rgba, width=1)
            inch += spacing_inches

        # Draw horizontal lines
        y = 0
        inch = 0
        while y < img.height:
            y = int(inch * ppi_y)
            if y < img.height:
                draw.line([(0, y), (img.width, y)], fill=grid_rgba, width=1)
            inch += spacing_inches

        # Composite
        result = img.convert('RGBA')
        result = Image.alpha_composite(result, overlay)
        return result.convert('RGB')

    def _add_rulers(
        self,
        img: Image.Image,
        ruler_color: str,
        text_color: str
    ) -> Image.Image:
        """Add rulers to the edges of the image."""
        ruler_height = 25
        ruler_width = 25

        # Create new image with space for rulers
        new_width = img.width + ruler_width
        new_height = img.height + ruler_height

        result = Image.new('RGB', (new_width, new_height), color='white')
        result.paste(img, (ruler_width, ruler_height))

        draw = ImageDraw.Draw(result)
        font = self._get_font(10)

        # Calculate pixels per inch
        ppi_x = img.width / SLIDE_WIDTH_INCHES
        ppi_y = img.height / SLIDE_HEIGHT_INCHES

        # Draw horizontal ruler (top)
        for inch in range(int(SLIDE_WIDTH_INCHES) + 1):
            x = ruler_width + int(inch * ppi_x)
            # Major tick
            draw.line([(x, 0), (x, ruler_height - 5)], fill=ruler_color, width=1)
            # Label
            draw.text((x + 2, 2), f"{inch}\"", fill=text_color, font=font)

            # Minor ticks (0.5 inch)
            if inch < SLIDE_WIDTH_INCHES:
                x_half = ruler_width + int((inch + 0.5) * ppi_x)
                draw.line([(x_half, ruler_height - 10), (x_half, ruler_height - 5)],
                         fill=ruler_color, width=1)

        # Draw vertical ruler (left)
        for inch in range(int(SLIDE_HEIGHT_INCHES) + 1):
            y = ruler_height + int(inch * ppi_y)
            # Major tick
            draw.line([(0, y), (ruler_width - 5, y)], fill=ruler_color, width=1)
            # Label
            draw.text((2, y + 2), f"{inch}\"", fill=text_color, font=font)

            # Minor ticks
            if inch < SLIDE_HEIGHT_INCHES:
                y_half = ruler_height + int((inch + 0.5) * ppi_y)
                draw.line([(ruler_width - 10, y_half), (ruler_width - 5, y_half)],
                         fill=ruler_color, width=1)

        return result

    def _create_side_by_side(
        self,
        img1_path: Path,
        img2_path: Path,
        output_path: Path
    ) -> Path:
        """Create a side-by-side comparison image."""
        img1 = Image.open(img1_path)
        img2 = Image.open(img2_path)

        # Ensure same size
        if img1.size != img2.size:
            target_size = (max(img1.width, img2.width), max(img1.height, img2.height))
            img1 = img1.resize(target_size, Image.Resampling.LANCZOS)
            img2 = img2.resize(target_size, Image.Resampling.LANCZOS)

        # Create combined image with labels
        gap = 20
        label_height = 30
        combined_width = img1.width * 2 + gap
        combined_height = img1.height + label_height

        combined = Image.new('RGB', (combined_width, combined_height), 'white')

        # Add labels
        draw = ImageDraw.Draw(combined)
        font = self._get_font(16)
        draw.text((img1.width // 2 - 30, 5), "ORIGINAL", fill='green', font=font)
        draw.text((img1.width + gap + img2.width // 2 - 30, 5), "GENERATED", fill='blue', font=font)

        # Paste images
        combined.paste(img1, (0, label_height))
        combined.paste(img2, (img1.width + gap, label_height))

        combined.save(output_path)
        return output_path

    def _parse_color(self, color_str: str) -> Tuple[int, int, int, int]:
        """Parse color string to RGBA tuple."""
        color_str = color_str.lstrip('#')

        if len(color_str) == 6:
            r = int(color_str[0:2], 16)
            g = int(color_str[2:4], 16)
            b = int(color_str[4:6], 16)
            a = 255
        elif len(color_str) == 8:
            r = int(color_str[0:2], 16)
            g = int(color_str[2:4], 16)
            b = int(color_str[4:6], 16)
            a = int(color_str[6:8], 16)
        else:
            r, g, b, a = 128, 128, 128, 255

        return (r, g, b, a)

    def extract_colors(self, image_path: Path, num_colors: int = 10) -> list:
        """
        Extract dominant colors from an image.

        Args:
            image_path: Path to image
            num_colors: Number of colors to extract

        Returns:
            List of hex color strings sorted by frequency
        """
        img = Image.open(image_path)
        img = img.convert('RGB')

        # Resize for faster processing
        img.thumbnail((200, 200))

        # Get colors
        colors = img.getcolors(maxcolors=50000)
        if colors is None:
            return []

        # Sort by frequency
        colors.sort(key=lambda x: x[0], reverse=True)

        # Convert to hex
        hex_colors = []
        for count, (r, g, b) in colors[:num_colors]:
            hex_color = f"#{r:02X}{g:02X}{b:02X}"
            hex_colors.append({"color": hex_color, "frequency": count})

        return hex_colors

    def analyze_layout(self, image_path: Path) -> Dict[str, Any]:
        """
        Analyze the basic layout structure of a slide image.

        Returns regions of interest and layout characteristics.
        """
        img = Image.open(image_path)
        img = img.convert('RGB')

        width, height = img.size
        img_array = np.array(img)

        # Calculate regions (divide into 3x3 grid)
        region_width = width // 3
        region_height = height // 3

        regions = {}
        for row in range(3):
            for col in range(3):
                x1, y1 = col * region_width, row * region_height
                x2, y2 = (col + 1) * region_width, (row + 1) * region_height

                region = img_array[y1:y2, x1:x2]

                # Calculate region statistics
                mean_color = region.mean(axis=(0, 1))
                variance = region.var()

                region_name = f"region_{row}_{col}"
                regions[region_name] = {
                    "bounds": {"x1": x1, "y1": y1, "x2": x2, "y2": y2},
                    "mean_color": f"#{int(mean_color[0]):02X}{int(mean_color[1]):02X}{int(mean_color[2]):02X}",
                    "variance": float(variance),
                    "has_content": variance > 100  # High variance suggests content
                }

        # Detect if this is likely a title slide (content concentrated in center)
        center_variance = regions["region_1_1"]["variance"]
        edge_variance = (
            regions["region_0_0"]["variance"] +
            regions["region_0_2"]["variance"] +
            regions["region_2_0"]["variance"] +
            regions["region_2_2"]["variance"]
        ) / 4

        layout_analysis = {
            "regions": regions,
            "suggested_type": "title_slide" if center_variance > edge_variance * 2 else "content_slide",
            "dominant_colors": self.extract_colors(image_path, 5)
        }

        return layout_analysis


def create_calibration_image(
    output_path: Path,
    width: int = OPTIMAL_WIDTH_PX,
    height: int = OPTIMAL_HEIGHT_PX
) -> Path:
    """
    Create a calibration reference image with precise measurements.

    This can be used to verify the vision model's measurement accuracy.
    """
    img = Image.new('RGB', (width, height), 'white')
    draw = ImageDraw.Draw(img)

    # Calculate pixels per inch
    ppi_x = width / SLIDE_WIDTH_INCHES
    ppi_y = height / SLIDE_HEIGHT_INCHES

    try:
        font = ImageFont.truetype("arial.ttf", 14)
        font_large = ImageFont.truetype("arial.ttf", 24)
    except:
        font = ImageFont.load_default()
        font_large = font

    # Draw grid
    for inch in range(int(SLIDE_WIDTH_INCHES) + 1):
        x = int(inch * ppi_x)
        draw.line([(x, 0), (x, height)], fill='#CCCCCC', width=1)
        draw.text((x + 5, 5), f"{inch}\"", fill='black', font=font)

    for inch in range(int(SLIDE_HEIGHT_INCHES) + 1):
        y = int(inch * ppi_y)
        draw.line([(0, y), (width, y)], fill='#CCCCCC', width=1)
        draw.text((5, y + 5), f"{inch}\"", fill='black', font=font)

    # Draw reference rectangles
    # 1-inch square at (1, 1)
    x1, y1 = int(1 * ppi_x), int(1 * ppi_y)
    x2, y2 = int(2 * ppi_x), int(2 * ppi_y)
    draw.rectangle([x1, y1, x2, y2], outline='blue', width=2)
    draw.text((x1 + 10, y1 + 10), "1\" x 1\"", fill='blue', font=font)

    # 2x1 rectangle at (4, 2)
    x1, y1 = int(4 * ppi_x), int(2 * ppi_y)
    x2, y2 = int(6 * ppi_x), int(3 * ppi_y)
    draw.rectangle([x1, y1, x2, y2], outline='green', width=2)
    draw.text((x1 + 10, y1 + 10), "2\" x 1\"", fill='green', font=font)

    # Center info
    cx, cy = width // 2, height // 2
    draw.text(
        (cx - 100, cy - 50),
        f"Slide: {SLIDE_WIDTH_INCHES}\" x {SLIDE_HEIGHT_INCHES}\"\n"
        f"Image: {width} x {height} px\n"
        f"PPI: {ppi_x:.1f} x {ppi_y:.1f}",
        fill='black',
        font=font_large
    )

    output_path = Path(output_path)
    img.save(output_path)

    logger.info(f"Calibration image saved to: {output_path}")
    return output_path


# Convenience functions
def preprocess_for_analysis(
    image_path: Path,
    output_path: Optional[Path] = None,
    add_grid: bool = True,
    add_rulers: bool = True
) -> Tuple[Path, Dict[str, Any]]:
    """
    Quick preprocessing for slide analysis.

    Args:
        image_path: Source image path
        output_path: Optional output path
        add_grid: Whether to add measurement grid
        add_rulers: Whether to add rulers

    Returns:
        Tuple of (output_path, metadata)
    """
    config = PreprocessingConfig(
        add_grid=add_grid,
        add_rulers=add_rulers,
        normalize_contrast=True,
        normalize_brightness=True,
        scale_to_optimal=True
    )

    preprocessor = ImagePreprocessor(config)
    return preprocessor.preprocess(image_path, output_path)


if __name__ == "__main__":
    import sys

    logging.basicConfig(level=logging.INFO)

    if len(sys.argv) > 1:
        input_path = Path(sys.argv[1])
        output_path = Path(sys.argv[2]) if len(sys.argv) > 2 else None

        result_path, metadata = preprocess_for_analysis(input_path, output_path)
        print(f"Preprocessed image: {result_path}")
        print(f"Metadata: {metadata}")
    else:
        # Create calibration image
        cal_path = create_calibration_image(Path("calibration_reference.png"))
        print(f"Created calibration image: {cal_path}")
