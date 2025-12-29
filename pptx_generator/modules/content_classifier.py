"""
Content Classifier Module

Provides automated classification for:
1. Images - categorize into logos, icons, photos, screenshots, backgrounds, charts
2. Icons - extract and tag icons for easy reuse
3. Diagrams - identify diagram patterns for templates
"""

import json
import logging
import hashlib
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Any
from dataclasses import dataclass, field, asdict
from enum import Enum

logger = logging.getLogger(__name__)


class ImageCategory(Enum):
    """Categories for image classification."""
    LOGO = "logo"
    ICON = "icon"
    PHOTO = "photo"
    SCREENSHOT = "screenshot"
    BACKGROUND = "background"
    CHART_IMAGE = "chart_image"
    DIAGRAM = "diagram"
    DECORATIVE = "decorative"
    UNKNOWN = "unknown"


class IconCategory(Enum):
    """Categories for icon classification."""
    SOFTWARE = "software"  # Microsoft, Google, etc.
    BUSINESS = "business"  # Charts, graphs, money
    PEOPLE = "people"  # Users, teams
    COMMUNICATION = "communication"  # Email, phone, chat
    DOCUMENT = "document"  # Files, folders
    LOCATION = "location"  # Maps, pins
    TIME = "time"  # Calendar, clock
    ARROW = "arrow"  # Directional
    CHECK = "check"  # Checkmarks, completion
    WARNING = "warning"  # Alerts, caution
    BUILDING = "building"  # Real estate, property
    ANALYTICS = "analytics"  # Data, metrics
    GENERIC = "generic"


class DiagramType(Enum):
    """Types of diagram patterns."""
    PROCESS_FLOW = "process_flow"
    TIMELINE = "timeline"
    HIERARCHY = "hierarchy"
    CYCLE = "cycle"
    COMPARISON = "comparison"
    PYRAMID = "pyramid"
    MATRIX = "matrix"
    FUNNEL = "funnel"
    RADIAL = "radial"
    STEP_PROCESS = "step_process"


@dataclass
class ClassifiedImage:
    """Classified image with metadata."""
    id: str
    filename: str
    category: str
    subcategory: str = ""
    confidence: float = 0.0
    width_inches: float = 0.0
    height_inches: float = 0.0
    aspect_ratio: float = 0.0
    size_bytes: int = 0
    tags: List[str] = field(default_factory=list)
    use_cases: List[str] = field(default_factory=list)


@dataclass
class ClassifiedIcon:
    """Classified icon with usage metadata."""
    id: str
    source_type: str  # "image" or "shape"
    category: str
    subcategory: str = ""
    size_px: Tuple[int, int] = (0, 0)
    colors: List[str] = field(default_factory=list)
    tags: List[str] = field(default_factory=list)
    keywords: List[str] = field(default_factory=list)


@dataclass
class DiagramTemplate:
    """Diagram template pattern."""
    id: str
    type: str
    name: str
    component_count: int = 0
    shape_types: List[str] = field(default_factory=list)
    connectors: int = 0
    layout: str = ""  # horizontal, vertical, radial
    tags: List[str] = field(default_factory=list)
    source_slide: Dict[str, Any] = field(default_factory=dict)


class ContentClassifier:
    """
    Classifies library content into categories for easy retrieval.
    """

    def __init__(self, library_path: Optional[Path] = None):
        """Initialize the classifier."""
        self.library_path = library_path or Path("pptx_component_library")
        self.index_path = self.library_path / "library_index.json"
        self.classified_path = self.library_path / "classified_index.json"

        self.images: Dict[str, ClassifiedImage] = {}
        self.icons: Dict[str, ClassifiedIcon] = {}
        self.diagrams: Dict[str, DiagramTemplate] = {}

        # Load existing classifications if available
        self._load_classifications()

    def _load_classifications(self):
        """Load existing classifications from file."""
        if self.classified_path.exists():
            try:
                with open(self.classified_path, 'r') as f:
                    data = json.load(f)

                for img_data in data.get('images', []):
                    img = ClassifiedImage(**img_data)
                    self.images[img.id] = img

                for icon_data in data.get('icons', []):
                    icon = ClassifiedIcon(**icon_data)
                    self.icons[icon.id] = icon

                for diag_data in data.get('diagrams', []):
                    diag = DiagramTemplate(**diag_data)
                    self.diagrams[diag.id] = diag

                logger.info(f"Loaded {len(self.images)} images, {len(self.icons)} icons, {len(self.diagrams)} diagrams")
            except Exception as e:
                logger.warning(f"Could not load classifications: {e}")

    def save_classifications(self):
        """Save classifications to file."""
        data = {
            'images': [asdict(img) for img in self.images.values()],
            'icons': [asdict(icon) for icon in self.icons.values()],
            'diagrams': [asdict(diag) for diag in self.diagrams.values()]
        }

        with open(self.classified_path, 'w') as f:
            json.dump(data, f, indent=2)

        logger.info(f"Saved {len(self.images)} images, {len(self.icons)} icons, {len(self.diagrams)} diagrams")

    # ==================== Image Classification ====================

    def classify_images(self, library_index: dict) -> Dict[str, ClassifiedImage]:
        """
        Classify all images in the library.

        Classification rules:
        - Icons: Small square images (< 1" and aspect ratio ~1:1)
        - Logos: Wide short images (aspect ratio > 3:1, height < 1")
        - Backgrounds: Large images (> 10" in any dimension)
        - Screenshots: Medium-large images with table/spreadsheet content
        - Photos: Medium images with typical photo aspect ratios
        - Chart images: Images that look like rendered charts
        """
        images = library_index.get('components', {}).get('images', [])

        for img in images:
            classified = self._classify_single_image(img)
            self.images[classified.id] = classified

        logger.info(f"Classified {len(self.images)} images")
        return self.images

    def _classify_single_image(self, img: dict) -> ClassifiedImage:
        """Classify a single image based on its properties."""
        width = img.get('width_inches', 0)
        height = img.get('height_inches', 0)
        size_bytes = img.get('size_bytes', 0)

        # Calculate aspect ratio
        aspect_ratio = width / height if height > 0 else 0

        # Determine category based on dimensions
        category = ImageCategory.UNKNOWN.value
        subcategory = ""
        confidence = 0.5
        tags = []
        use_cases = []

        # Icon detection: small square images
        if width < 1.0 and height < 1.0 and 0.7 < aspect_ratio < 1.4:
            category = ImageCategory.ICON.value
            confidence = 0.9
            tags = ["small", "square", "icon"]
            use_cases = ["bullet_icon", "accent", "list_marker"]

        # Logo detection: wide and short
        elif aspect_ratio > 3.0 and height < 1.5:
            category = ImageCategory.LOGO.value
            confidence = 0.85
            tags = ["wide", "logo", "branding"]
            use_cases = ["header", "footer", "title_slide"]

        # Background detection: large images
        elif width > 10 or height > 6:
            category = ImageCategory.BACKGROUND.value
            confidence = 0.8
            tags = ["large", "background", "cover"]
            use_cases = ["slide_background", "section_divider", "cover_slide"]

        # Screenshot detection: medium-large with high file size (complex content)
        elif width > 4 and height > 3 and size_bytes > 100000:
            category = ImageCategory.SCREENSHOT.value
            confidence = 0.7
            tags = ["screenshot", "document", "complex"]
            use_cases = ["example", "reference", "detail"]

        # Photo detection: medium images with typical photo ratios
        elif 1.2 < aspect_ratio < 2.0 and width > 1.5:
            category = ImageCategory.PHOTO.value
            confidence = 0.6
            tags = ["photo", "image"]
            use_cases = ["illustration", "case_study", "property_image"]

        # Chart image: square-ish medium images
        elif 0.8 < aspect_ratio < 1.5 and 1.5 < width < 5:
            category = ImageCategory.CHART_IMAGE.value
            confidence = 0.5
            tags = ["chart", "data"]
            use_cases = ["data_visualization", "comparison"]

        # Decorative: small non-square
        elif width < 2 and height < 2:
            category = ImageCategory.DECORATIVE.value
            confidence = 0.6
            tags = ["decorative", "accent"]
            use_cases = ["decoration", "separator"]

        return ClassifiedImage(
            id=img['id'],
            filename=img.get('filename', ''),
            category=category,
            subcategory=subcategory,
            confidence=confidence,
            width_inches=width,
            height_inches=height,
            aspect_ratio=round(aspect_ratio, 2),
            size_bytes=size_bytes,
            tags=tags,
            use_cases=use_cases
        )

    def get_images_by_category(self, category: str) -> List[ClassifiedImage]:
        """Get all images of a specific category."""
        return [img for img in self.images.values() if img.category == category]

    def get_image_stats(self) -> Dict[str, int]:
        """Get count of images by category."""
        stats = {}
        for img in self.images.values():
            stats[img.category] = stats.get(img.category, 0) + 1
        return stats

    # ==================== Icon Extraction ====================

    def extract_icons(self, library_index: dict) -> Dict[str, ClassifiedIcon]:
        """
        Extract and classify icons from both images and shapes.

        Icon sources:
        1. Small square images (already classified as icons)
        2. Small shapes with specific types (circles, rounded rects)
        """
        # Get icons from classified images
        for img_id, img in self.images.items():
            if img.category == ImageCategory.ICON.value:
                icon = self._create_icon_from_image(img)
                self.icons[icon.id] = icon

        # Get icons from shapes
        shapes = library_index.get('components', {}).get('shapes', [])
        for shape in shapes:
            if self._is_icon_shape(shape):
                icon = self._create_icon_from_shape(shape)
                self.icons[icon.id] = icon

        logger.info(f"Extracted {len(self.icons)} icons")
        return self.icons

    def _is_icon_shape(self, shape: dict) -> bool:
        """Determine if a shape qualifies as an icon."""
        width = shape.get('width_inches') or 0
        height = shape.get('height_inches') or 0
        shape_type = (shape.get('shape_type') or '').upper()

        # Icon criteria: small, roughly square
        if width > 1.5 or height > 1.5:
            return False

        if width < 0.1 or height < 0.1:
            return False  # Too small, probably a line

        aspect = width / height if height > 0 else 0
        if aspect < 0.5 or aspect > 2.0:
            return False  # Not square enough

        # Exclude certain shape types
        excluded_types = ['LINE', 'CONNECTOR', 'RECTANGLE']
        if shape_type in excluded_types:
            # Rectangles are ok if they're small and square
            if shape_type == 'RECTANGLE' and not (0.8 < aspect < 1.2):
                return False

        return True

    def _create_icon_from_image(self, img: ClassifiedImage) -> ClassifiedIcon:
        """Create an icon entry from a classified image."""
        # Determine icon category based on size and other heuristics
        category = self._guess_icon_category(img.tags)

        return ClassifiedIcon(
            id=f"img_{img.id}",
            source_type="image",
            category=category,
            size_px=(int(img.width_inches * 96), int(img.height_inches * 96)),
            tags=img.tags,
            keywords=img.use_cases
        )

    def _create_icon_from_shape(self, shape: dict) -> ClassifiedIcon:
        """Create an icon entry from a shape."""
        shape_type = (shape.get('shape_type') or '').upper()
        category = self._guess_icon_category_from_shape(shape_type)

        width = shape.get('width_inches', 0)
        height = shape.get('height_inches', 0)

        return ClassifiedIcon(
            id=f"shp_{shape.get('id', '')}",
            source_type="shape",
            category=category,
            subcategory=shape_type.lower(),
            size_px=(int(width * 96), int(height * 96)),
            colors=shape.get('colors', []),
            tags=[shape_type.lower()],
            keywords=[]
        )

    def _guess_icon_category(self, tags: List[str]) -> str:
        """Guess icon category from tags."""
        # Default to generic
        return IconCategory.GENERIC.value

    def _guess_icon_category_from_shape(self, shape_type: str) -> str:
        """Guess icon category from shape type."""
        shape_type = shape_type.upper()

        if 'ARROW' in shape_type or 'CHEVRON' in shape_type:
            return IconCategory.ARROW.value
        elif 'CHECK' in shape_type:
            return IconCategory.CHECK.value
        elif 'STAR' in shape_type or 'BURST' in shape_type:
            return IconCategory.DECORATIVE.value if hasattr(IconCategory, 'DECORATIVE') else IconCategory.GENERIC.value
        else:
            return IconCategory.GENERIC.value

    def get_icons_by_category(self, category: str) -> List[ClassifiedIcon]:
        """Get all icons of a specific category."""
        return [icon for icon in self.icons.values() if icon.category == category]

    def get_icon_stats(self) -> Dict[str, int]:
        """Get count of icons by category."""
        stats = {}
        for icon in self.icons.values():
            stats[icon.category] = stats.get(icon.category, 0) + 1
        return stats

    # ==================== Diagram Templates ====================

    def extract_diagram_templates(self, library_index: dict) -> Dict[str, DiagramTemplate]:
        """
        Extract diagram patterns from the library.

        Diagrams are identified by:
        1. Groups of connected shapes
        2. Diagrams in the diagrams component list
        3. Patterns of arrows + shapes
        """
        diagrams = library_index.get('components', {}).get('diagrams', [])

        for diag in diagrams:
            template = self._classify_diagram(diag)
            self.diagrams[template.id] = template

        logger.info(f"Extracted {len(self.diagrams)} diagram templates")
        return self.diagrams

    def _classify_diagram(self, diag: dict) -> DiagramTemplate:
        """Classify a diagram and create a template."""
        diag_id = diag.get('id', '')
        shape_count = diag.get('shape_count', 0)
        shapes = diag.get('shapes', [])

        # Analyze shape types
        shape_types = []
        arrow_count = 0
        connector_count = 0

        for shape in shapes:
            stype = (shape.get('shape_type') or '').upper()
            shape_types.append(stype)

            if 'ARROW' in stype or 'CHEVRON' in stype:
                arrow_count += 1
            if 'CONNECTOR' in stype or 'LINE' in stype:
                connector_count += 1

        # Determine diagram type based on shapes
        diagram_type = self._guess_diagram_type(shape_types, arrow_count, connector_count, shape_count)

        # Determine layout
        layout = self._guess_diagram_layout(shapes)

        # Generate name
        name = f"{diagram_type.replace('_', ' ').title()} ({shape_count} shapes)"

        return DiagramTemplate(
            id=diag_id,
            type=diagram_type,
            name=name,
            component_count=shape_count,
            shape_types=list(set(shape_types)),
            connectors=connector_count,
            layout=layout,
            tags=[diagram_type, layout],
            source_slide={
                'template': diag.get('references', [{}])[0].get('template', ''),
                'slide': diag.get('references', [{}])[0].get('slide', 0)
            }
        )

    def _guess_diagram_type(
        self,
        shape_types: List[str],
        arrow_count: int,
        connector_count: int,
        total_shapes: int
    ) -> str:
        """Guess diagram type from shape analysis."""
        # High arrow ratio suggests process flow
        arrow_ratio = arrow_count / total_shapes if total_shapes > 0 else 0

        if arrow_ratio > 0.3:
            return DiagramType.PROCESS_FLOW.value

        # Check for specific patterns
        has_chevrons = any('CHEVRON' in s for s in shape_types)
        has_circles = any('OVAL' in s or 'CIRCLE' in s for s in shape_types)
        has_rectangles = any('RECT' in s for s in shape_types)

        if has_chevrons and not has_circles:
            return DiagramType.STEP_PROCESS.value

        if has_circles and connector_count > 2:
            return DiagramType.CYCLE.value

        if total_shapes >= 6 and has_rectangles:
            # Could be a matrix
            return DiagramType.MATRIX.value

        # Default to process flow
        return DiagramType.PROCESS_FLOW.value

    def _guess_diagram_layout(self, shapes: List[dict]) -> str:
        """Guess diagram layout from shape positions."""
        if not shapes:
            return "unknown"

        # Get positions
        positions = []
        for shape in shapes:
            pos = shape.get('position', {})
            if pos:
                positions.append((
                    pos.get('left_inches', 0),
                    pos.get('top_inches', 0)
                ))

        if len(positions) < 2:
            return "single"

        # Calculate spread
        x_values = [p[0] for p in positions]
        y_values = [p[1] for p in positions]

        x_spread = max(x_values) - min(x_values)
        y_spread = max(y_values) - min(y_values)

        if x_spread > y_spread * 2:
            return "horizontal"
        elif y_spread > x_spread * 2:
            return "vertical"
        elif x_spread > 5 and y_spread > 3:
            return "grid"
        else:
            return "compact"

    def get_diagrams_by_type(self, diagram_type: str) -> List[DiagramTemplate]:
        """Get all diagrams of a specific type."""
        return [diag for diag in self.diagrams.values() if diag.type == diagram_type]

    def get_diagram_stats(self) -> Dict[str, int]:
        """Get count of diagrams by type."""
        stats = {}
        for diag in self.diagrams.values():
            stats[diag.type] = stats.get(diag.type, 0) + 1
        return stats

    # ==================== Full Classification ====================

    def classify_all(self, library_index: dict) -> Dict[str, Any]:
        """
        Run full classification on all library content.

        Returns summary statistics.
        """
        logger.info("Starting full content classification...")

        # Classify images first
        self.classify_images(library_index)

        # Extract icons (depends on image classification)
        self.extract_icons(library_index)

        # Extract diagram templates
        self.extract_diagram_templates(library_index)

        # Save results
        self.save_classifications()

        return {
            'images': self.get_image_stats(),
            'icons': self.get_icon_stats(),
            'diagrams': self.get_diagram_stats()
        }

    # ==================== Search Methods ====================

    def find_logo(self, prefer_wide: bool = True) -> Optional[ClassifiedImage]:
        """Find a suitable logo image."""
        logos = self.get_images_by_category(ImageCategory.LOGO.value)
        if not logos:
            return None

        if prefer_wide:
            logos.sort(key=lambda x: x.aspect_ratio, reverse=True)

        return logos[0]

    def find_background(self, min_width: float = 10.0) -> Optional[ClassifiedImage]:
        """Find a suitable background image."""
        backgrounds = self.get_images_by_category(ImageCategory.BACKGROUND.value)
        suitable = [b for b in backgrounds if b.width_inches >= min_width]

        if suitable:
            return suitable[0]
        return backgrounds[0] if backgrounds else None

    def find_icons_for_list(self, count: int = 4) -> List[ClassifiedIcon]:
        """Find icons suitable for bullet lists."""
        icons = list(self.icons.values())
        # Prefer smaller icons
        icons.sort(key=lambda x: x.size_px[0] if x.size_px else 999)
        return icons[:count]

    def find_diagram_for_process(self, step_count: int = 4) -> Optional[DiagramTemplate]:
        """Find a diagram template suitable for a process with N steps."""
        process_diagrams = self.get_diagrams_by_type(DiagramType.PROCESS_FLOW.value)
        process_diagrams.extend(self.get_diagrams_by_type(DiagramType.STEP_PROCESS.value))

        if not process_diagrams:
            return None

        # Find closest match to step count
        process_diagrams.sort(key=lambda x: abs(x.component_count - step_count))
        return process_diagrams[0]


def main():
    """Run classification from command line."""
    import argparse

    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

    parser = argparse.ArgumentParser(description="Content Classifier")
    parser.add_argument("--classify", action="store_true", help="Run full classification")
    parser.add_argument("--stats", action="store_true", help="Show classification stats")
    parser.add_argument("--images", action="store_true", help="Show image stats")
    parser.add_argument("--icons", action="store_true", help="Show icon stats")
    parser.add_argument("--diagrams", action="store_true", help="Show diagram stats")

    args = parser.parse_args()

    classifier = ContentClassifier()

    if args.classify:
        # Load library index
        index_path = Path("component_library/library_index.json")
        if not index_path.exists():
            print("Library index not found!")
            return

        with open(index_path, 'r') as f:
            library_index = json.load(f)

        stats = classifier.classify_all(library_index)
        print("\nClassification Complete!")
        print(f"\nImages: {stats['images']}")
        print(f"Icons: {stats['icons']}")
        print(f"Diagrams: {stats['diagrams']}")

    if args.stats or args.images:
        print("\nImage Classification:")
        for cat, count in sorted(classifier.get_image_stats().items(), key=lambda x: -x[1]):
            print(f"  {cat}: {count}")

    if args.stats or args.icons:
        print("\nIcon Classification:")
        for cat, count in sorted(classifier.get_icon_stats().items(), key=lambda x: -x[1]):
            print(f"  {cat}: {count}")

    if args.stats or args.diagrams:
        print("\nDiagram Classification:")
        for dtype, count in sorted(classifier.get_diagram_stats().items(), key=lambda x: -x[1]):
            print(f"  {dtype}: {count}")


if __name__ == "__main__":
    main()
