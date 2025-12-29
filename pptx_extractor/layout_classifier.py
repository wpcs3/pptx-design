"""
ML-Based Slide Layout Classification

Uses LayoutParser and computer vision to classify slide layouts
from rendered images, detecting semantic regions like titles,
text blocks, figures, and tables.

Phase 4 Enhancement (2025-12-29):
- Region detection using LayoutParser with PubLayNet model
- Layout type inference from detected regions
- Confidence scoring for classifications
- Batch processing for multiple slides

Usage:
    from pptx_extractor.layout_classifier import SlideLayoutClassifier

    classifier = SlideLayoutClassifier()
    result = classifier.classify("slide_image.png")
    print(result.layout_type)  # "content_with_figure"
    print(result.regions)      # List of detected regions

Requirements:
    pip install layoutparser
    pip install "layoutparser[detectron2]"  # For PubLayNet model

    # Or use ONNX backend (lighter weight):
    pip install "layoutparser[ocr]"

Note:
    This module provides graceful fallbacks when ML models are unavailable.
    Basic classification works without deep learning dependencies.
"""

import logging
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

logger = logging.getLogger(__name__)

# Try to import layoutparser and related dependencies
LAYOUTPARSER_AVAILABLE = False
DETECTRON2_AVAILABLE = False

try:
    import layoutparser as lp
    LAYOUTPARSER_AVAILABLE = True
    logger.info("LayoutParser available")

    try:
        import detectron2
        DETECTRON2_AVAILABLE = True
        logger.info("Detectron2 backend available")
    except ImportError:
        logger.info("Detectron2 not available, using fallback methods")
except ImportError:
    logger.info("LayoutParser not installed, using basic classification")


@dataclass
class DetectedRegion:
    """A detected region in a slide image."""
    region_type: str  # title, text, figure, table, list
    bbox: Tuple[float, float, float, float]  # x1, y1, x2, y2
    confidence: float
    area_ratio: float = 0.0  # Ratio of slide area
    center: Tuple[float, float] = (0.0, 0.0)

    @property
    def width(self) -> float:
        return self.bbox[2] - self.bbox[0]

    @property
    def height(self) -> float:
        return self.bbox[3] - self.bbox[1]

    @property
    def is_top_half(self) -> bool:
        """Check if region is in top half of slide."""
        return self.center[1] < 0.5

    @property
    def is_left_half(self) -> bool:
        """Check if region is in left half of slide."""
        return self.center[0] < 0.5


@dataclass
class LayoutClassification:
    """Result of slide layout classification."""
    layout_type: str
    confidence: float
    regions: List[DetectedRegion]
    slide_width: int = 0
    slide_height: int = 0
    analysis: Dict[str, Any] = field(default_factory=dict)

    @property
    def region_summary(self) -> Dict[str, int]:
        """Count of each region type."""
        summary = {}
        for region in self.regions:
            rtype = region.region_type
            summary[rtype] = summary.get(rtype, 0) + 1
        return summary


class SlideLayoutClassifier:
    """
    ML-based slide layout classifier using LayoutParser.

    Detects semantic regions in slide images and infers layout types
    based on the arrangement and types of detected elements.

    Layout Types Detected:
    - title_slide: Large centered title, optional subtitle
    - content_slide: Title + body text/bullets
    - content_with_figure: Title + text + image/figure
    - content_with_table: Title + text + table
    - two_column: Content split into two columns
    - image_slide: Dominant figure/image
    - table_slide: Dominant table
    - chart_slide: Dominant chart/figure with data
    - section_divider: Simple title, minimal content
    - blank: No detected content
    """

    # PubLayNet label mapping
    PUBLAYNET_LABELS = {
        0: "text",
        1: "title",
        2: "list",
        3: "table",
        4: "figure"
    }

    # Layout type inference rules
    LAYOUT_RULES = [
        # (condition_func, layout_type, base_confidence)
        ("_is_title_slide", "title_slide", 0.9),
        ("_is_section_divider", "section_divider", 0.85),
        ("_is_table_slide", "table_slide", 0.85),
        ("_is_image_slide", "image_slide", 0.85),
        ("_is_two_column", "two_column", 0.8),
        ("_is_content_with_table", "content_with_table", 0.8),
        ("_is_content_with_figure", "content_with_figure", 0.8),
        ("_is_content_slide", "content_slide", 0.75),
    ]

    def __init__(
        self,
        model_config: str = "lp://PubLayNet/faster_rcnn_R_50_FPN_3x/config",
        confidence_threshold: float = 0.5,
        use_gpu: bool = False
    ):
        """
        Initialize the layout classifier.

        Args:
            model_config: LayoutParser model configuration path.
            confidence_threshold: Minimum confidence for region detection.
            use_gpu: Whether to use GPU for inference.
        """
        self.confidence_threshold = confidence_threshold
        self.use_gpu = use_gpu
        self.model = None
        self.model_config = model_config

        if LAYOUTPARSER_AVAILABLE and DETECTRON2_AVAILABLE:
            try:
                device = "cuda" if use_gpu else "cpu"
                self.model = lp.Detectron2LayoutModel(
                    model_config,
                    extra_config=["MODEL.DEVICE", device],
                    label_map=self.PUBLAYNET_LABELS
                )
                logger.info(f"Loaded LayoutParser model on {device}")
            except Exception as e:
                logger.warning(f"Failed to load LayoutParser model: {e}")
                self.model = None

    def classify(self, image_path: str) -> LayoutClassification:
        """
        Classify the layout of a slide image.

        Args:
            image_path: Path to the slide image (PNG, JPG).

        Returns:
            LayoutClassification with detected regions and inferred layout type.
        """
        image_path = Path(image_path)
        if not image_path.exists():
            logger.error(f"Image not found: {image_path}")
            return LayoutClassification(
                layout_type="unknown",
                confidence=0.0,
                regions=[]
            )

        # Load image
        try:
            from PIL import Image
            image = Image.open(image_path)
            width, height = image.size
        except ImportError:
            logger.error("PIL not installed")
            return LayoutClassification(layout_type="unknown", confidence=0.0, regions=[])
        except Exception as e:
            logger.error(f"Failed to load image: {e}")
            return LayoutClassification(layout_type="unknown", confidence=0.0, regions=[])

        # Detect regions
        if self.model:
            regions = self._detect_with_layoutparser(image, width, height)
        else:
            regions = self._detect_basic(image, width, height)

        # Infer layout type
        layout_type, confidence, analysis = self._infer_layout_type(regions, width, height)

        return LayoutClassification(
            layout_type=layout_type,
            confidence=confidence,
            regions=regions,
            slide_width=width,
            slide_height=height,
            analysis=analysis
        )

    def classify_batch(self, image_paths: List[str]) -> List[LayoutClassification]:
        """
        Classify multiple slide images.

        Args:
            image_paths: List of paths to slide images.

        Returns:
            List of LayoutClassification results.
        """
        results = []
        for path in image_paths:
            result = self.classify(path)
            results.append(result)
            logger.debug(f"Classified {path}: {result.layout_type}")
        return results

    def _detect_with_layoutparser(
        self,
        image,
        width: int,
        height: int
    ) -> List[DetectedRegion]:
        """Detect regions using LayoutParser model."""
        try:
            import numpy as np
            image_array = np.array(image)

            # Run detection
            layout = self.model.detect(image_array)

            regions = []
            slide_area = width * height

            for block in layout:
                if block.score < self.confidence_threshold:
                    continue

                # Get bounding box coordinates
                x1, y1, x2, y2 = block.coordinates

                # Calculate normalized values
                area = (x2 - x1) * (y2 - y1)
                area_ratio = area / slide_area
                center_x = ((x1 + x2) / 2) / width
                center_y = ((y1 + y2) / 2) / height

                region = DetectedRegion(
                    region_type=block.type.lower(),
                    bbox=(x1, y1, x2, y2),
                    confidence=block.score,
                    area_ratio=area_ratio,
                    center=(center_x, center_y)
                )
                regions.append(region)

            return regions

        except Exception as e:
            logger.error(f"LayoutParser detection failed: {e}")
            return self._detect_basic(image, width, height)

    def _detect_basic(
        self,
        image,
        width: int,
        height: int
    ) -> List[DetectedRegion]:
        """
        Basic region detection without ML model.

        Uses simple heuristics and edge detection.
        """
        regions = []

        try:
            import numpy as np
            from PIL import ImageFilter

            # Convert to grayscale
            gray = image.convert('L')
            gray_array = np.array(gray)

            # Simple edge detection
            edges = image.filter(ImageFilter.FIND_EDGES).convert('L')
            edges_array = np.array(edges)

            # Find content regions using simple thresholding
            # Divide into grid and check for content
            grid_rows, grid_cols = 4, 3
            cell_height = height // grid_rows
            cell_width = width // grid_cols

            content_cells = []
            for row in range(grid_rows):
                for col in range(grid_cols):
                    y1 = row * cell_height
                    y2 = (row + 1) * cell_height
                    x1 = col * cell_width
                    x2 = (col + 1) * cell_width

                    # Check if cell has content (non-white pixels)
                    cell = gray_array[y1:y2, x1:x2]
                    content_ratio = np.mean(cell < 250) # Non-white pixels

                    if content_ratio > 0.1:
                        content_cells.append((row, col, content_ratio))

            # Infer regions from content distribution
            if content_cells:
                # Check for title region (top area with content)
                top_cells = [c for c in content_cells if c[0] == 0]
                if top_cells:
                    regions.append(DetectedRegion(
                        region_type="title",
                        bbox=(0, 0, width, cell_height),
                        confidence=0.6,
                        area_ratio=1/grid_rows,
                        center=(0.5, 0.5/grid_rows)
                    ))

                # Check for body content
                body_cells = [c for c in content_cells if c[0] > 0]
                if body_cells:
                    # Determine if content is more figure-like or text-like
                    avg_density = np.mean([c[2] for c in body_cells])
                    region_type = "text" if avg_density < 0.3 else "figure"

                    regions.append(DetectedRegion(
                        region_type=region_type,
                        bbox=(0, cell_height, width, height),
                        confidence=0.5,
                        area_ratio=(grid_rows-1)/grid_rows,
                        center=(0.5, 0.5 + 0.5/grid_rows)
                    ))

        except Exception as e:
            logger.warning(f"Basic detection failed: {e}")

        return regions

    def _infer_layout_type(
        self,
        regions: List[DetectedRegion],
        width: int,
        height: int
    ) -> Tuple[str, float, Dict[str, Any]]:
        """Infer layout type from detected regions."""
        analysis = {
            "region_count": len(regions),
            "region_types": [r.region_type for r in regions],
            "has_title": any(r.region_type == "title" for r in regions),
            "has_text": any(r.region_type in ("text", "list") for r in regions),
            "has_figure": any(r.region_type == "figure" for r in regions),
            "has_table": any(r.region_type == "table" for r in regions),
        }

        # Apply layout rules
        for rule_method, layout_type, base_confidence in self.LAYOUT_RULES:
            method = getattr(self, rule_method)
            if method(regions, analysis, width, height):
                # Adjust confidence based on region confidence scores
                if regions:
                    avg_confidence = sum(r.confidence for r in regions) / len(regions)
                    confidence = base_confidence * avg_confidence
                else:
                    confidence = base_confidence * 0.5
                return layout_type, confidence, analysis

        # Default fallback
        return "content_slide", 0.5, analysis

    def _is_title_slide(
        self,
        regions: List[DetectedRegion],
        analysis: Dict,
        width: int,
        height: int
    ) -> bool:
        """Check if this is a title slide."""
        titles = [r for r in regions if r.region_type == "title"]
        if not titles:
            return False

        # Title slide: large centered title, minimal other content
        main_title = max(titles, key=lambda r: r.area_ratio)
        has_large_title = main_title.area_ratio > 0.05

        # Little body text
        text_regions = [r for r in regions if r.region_type in ("text", "list")]
        little_text = sum(r.area_ratio for r in text_regions) < 0.15

        # No figures/tables
        no_visuals = not analysis["has_figure"] and not analysis["has_table"]

        return has_large_title and little_text and no_visuals

    def _is_section_divider(
        self,
        regions: List[DetectedRegion],
        analysis: Dict,
        width: int,
        height: int
    ) -> bool:
        """Check if this is a section divider slide."""
        # Similar to title slide but centered vertically
        titles = [r for r in regions if r.region_type == "title"]
        if not titles:
            return False

        main_title = max(titles, key=lambda r: r.area_ratio)

        # Title in middle third of slide
        is_centered = 0.3 < main_title.center[1] < 0.7

        # Very little other content
        other_content = sum(r.area_ratio for r in regions if r.region_type != "title")
        minimal_content = other_content < 0.05

        return is_centered and minimal_content

    def _is_table_slide(
        self,
        regions: List[DetectedRegion],
        analysis: Dict,
        width: int,
        height: int
    ) -> bool:
        """Check if this is a table-dominant slide."""
        tables = [r for r in regions if r.region_type == "table"]
        if not tables:
            return False

        # Table takes significant portion of slide
        table_area = sum(r.area_ratio for r in tables)
        return table_area > 0.25

    def _is_image_slide(
        self,
        regions: List[DetectedRegion],
        analysis: Dict,
        width: int,
        height: int
    ) -> bool:
        """Check if this is an image-dominant slide."""
        figures = [r for r in regions if r.region_type == "figure"]
        if not figures:
            return False

        # Figure takes significant portion of slide
        figure_area = sum(r.area_ratio for r in figures)
        return figure_area > 0.35

    def _is_two_column(
        self,
        regions: List[DetectedRegion],
        analysis: Dict,
        width: int,
        height: int
    ) -> bool:
        """Check if this is a two-column layout."""
        content_regions = [r for r in regions if r.region_type in ("text", "list", "figure")]
        if len(content_regions) < 2:
            return False

        # Check for content on both sides
        left_content = [r for r in content_regions if r.is_left_half]
        right_content = [r for r in content_regions if not r.is_left_half]

        has_both_sides = len(left_content) > 0 and len(right_content) > 0

        # Both sides have similar vertical positions
        if has_both_sides:
            left_centers = [r.center[1] for r in left_content]
            right_centers = [r.center[1] for r in right_content]

            if left_centers and right_centers:
                left_avg = sum(left_centers) / len(left_centers)
                right_avg = sum(right_centers) / len(right_centers)
                similar_height = abs(left_avg - right_avg) < 0.2
                return similar_height

        return False

    def _is_content_with_table(
        self,
        regions: List[DetectedRegion],
        analysis: Dict,
        width: int,
        height: int
    ) -> bool:
        """Check if this is content slide with a table."""
        return (
            analysis["has_title"] and
            analysis["has_table"] and
            (analysis["has_text"] or not analysis["has_figure"])
        )

    def _is_content_with_figure(
        self,
        regions: List[DetectedRegion],
        analysis: Dict,
        width: int,
        height: int
    ) -> bool:
        """Check if this is content slide with a figure."""
        return (
            analysis["has_title"] and
            analysis["has_figure"]
        )

    def _is_content_slide(
        self,
        regions: List[DetectedRegion],
        analysis: Dict,
        width: int,
        height: int
    ) -> bool:
        """Check if this is a standard content slide."""
        return analysis["has_title"] and analysis["has_text"]


def classify_slide(image_path: str, use_gpu: bool = False) -> LayoutClassification:
    """
    Convenience function to classify a single slide image.

    Args:
        image_path: Path to slide image.
        use_gpu: Whether to use GPU for inference.

    Returns:
        LayoutClassification result.
    """
    classifier = SlideLayoutClassifier(use_gpu=use_gpu)
    return classifier.classify(image_path)


def classify_presentation_slides(
    slide_image_dir: str,
    pattern: str = "*.png"
) -> Dict[str, LayoutClassification]:
    """
    Classify all slide images in a directory.

    Args:
        slide_image_dir: Directory containing slide images.
        pattern: Glob pattern for image files.

    Returns:
        Dictionary mapping filename to classification.
    """
    from pathlib import Path

    slide_dir = Path(slide_image_dir)
    if not slide_dir.exists():
        logger.error(f"Directory not found: {slide_image_dir}")
        return {}

    image_files = sorted(slide_dir.glob(pattern))
    classifier = SlideLayoutClassifier()

    results = {}
    for image_file in image_files:
        classification = classifier.classify(str(image_file))
        results[image_file.name] = classification
        logger.info(f"{image_file.name}: {classification.layout_type} "
                   f"(confidence: {classification.confidence:.2f})")

    return results


# CLI interface
if __name__ == "__main__":
    import sys

    if len(sys.argv) > 1:
        image_path = sys.argv[1]

        print(f"Classifying: {image_path}")
        print(f"LayoutParser available: {LAYOUTPARSER_AVAILABLE}")
        print(f"Detectron2 available: {DETECTRON2_AVAILABLE}")
        print()

        result = classify_slide(image_path)

        print(f"Layout Type: {result.layout_type}")
        print(f"Confidence: {result.confidence:.2f}")
        print(f"Image Size: {result.slide_width}x{result.slide_height}")
        print()

        print("Detected Regions:")
        for i, region in enumerate(result.regions):
            print(f"  {i+1}. {region.region_type}")
            print(f"      Confidence: {region.confidence:.2f}")
            print(f"      Area Ratio: {region.area_ratio:.3f}")
            print(f"      Center: ({region.center[0]:.2f}, {region.center[1]:.2f})")
        print()

        print("Analysis:")
        for key, value in result.analysis.items():
            print(f"  {key}: {value}")
    else:
        print("Usage: python -m pptx_extractor.layout_classifier <image_path>")
        print()
        print("Classify slide layout from rendered image.")
        print()
        print("Dependencies:")
        print("  pip install layoutparser")
        print("  pip install 'layoutparser[detectron2]'  # For ML model")
