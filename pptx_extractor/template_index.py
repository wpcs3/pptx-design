"""
Template Index Builder

Creates a searchable index of extracted slide templates for the generator module.
Groups slides by type, identifies unique patterns, and builds a library that can
be used to match content to appropriate templates.

Usage:
    builder = TemplateIndexBuilder()
    builder.add_extraction(slide_path, extraction_result)
    builder.build_index()
    builder.save("template_library.json")
"""

import json
import logging
from pathlib import Path
from typing import Optional, Dict, Any, List, Tuple
from dataclasses import dataclass, field, asdict
from collections import defaultdict
import hashlib

logger = logging.getLogger(__name__)


@dataclass
class TemplateEntry:
    """A single template in the library."""
    id: str
    source_file: str
    slide_index: int
    slide_type: str
    template_category: str
    layout_category: str
    complexity_score: int
    reusability_score: int
    color_scheme: str
    visual_style: str
    primary_font: str
    element_count: int
    has_chart: bool = False
    has_table: bool = False
    has_image: bool = False
    content_zones: List[Dict] = field(default_factory=list)
    style_tokens: Dict[str, Any] = field(default_factory=dict)
    description_path: str = ""
    similarity_hash: str = ""


@dataclass
class TemplateLibrary:
    """Complete template library with index and statistics."""
    version: str = "1.0"
    total_templates: int = 0
    unique_patterns: int = 0
    templates: List[TemplateEntry] = field(default_factory=list)
    categories: Dict[str, List[str]] = field(default_factory=dict)
    style_clusters: Dict[str, List[str]] = field(default_factory=dict)
    statistics: Dict[str, Any] = field(default_factory=dict)


class TemplateIndexBuilder:
    """
    Builds a searchable template library from extracted slide descriptions.

    Features:
    - Groups templates by type and style
    - Identifies unique patterns (deduplication)
    - Creates searchable index for generator
    - Tracks statistics and coverage
    """

    def __init__(self, output_dir: Optional[Path] = None):
        """
        Initialize the template index builder.

        Args:
            output_dir: Directory for saving the library
        """
        self.output_dir = Path(output_dir) if output_dir else Path("descriptions")
        self.output_dir.mkdir(parents=True, exist_ok=True)

        self.templates: List[TemplateEntry] = []
        self.categories: Dict[str, List[str]] = defaultdict(list)
        self.similarity_hashes: Dict[str, str] = {}  # hash -> first template id

    def add_extraction(
        self,
        source_file: str,
        slide_index: int,
        extraction: Dict[str, Any],
        description_path: Optional[str] = None
    ) -> Optional[TemplateEntry]:
        """
        Add an extracted slide to the library.

        Args:
            source_file: Original PPTX filename
            slide_index: Slide number (0-indexed)
            extraction: The extraction result dictionary
            description_path: Path to saved description JSON

        Returns:
            TemplateEntry if added, None if duplicate
        """
        # Extract key information (support both old and new format)
        # New format has: metadata, generator_hints, typography_system
        # Old format has: slide_dimensions, background, elements, color_palette, design_notes
        metadata = extraction.get("metadata", {})
        generator_hints = extraction.get("generator_hints", {})
        color_palette = extraction.get("color_palette", {})
        typography = extraction.get("typography_system", {})
        elements = extraction.get("elements", [])

        # Handle old format - infer metadata from elements if missing
        if not metadata:
            metadata = {
                "slide_type": self._infer_slide_type(elements, extraction.get("design_notes", "")),
                "complexity_score": min(5, max(1, len(elements) // 3 + 1))
            }

        # Generate template ID
        template_id = f"{Path(source_file).stem}_slide_{slide_index:03d}"

        # Calculate similarity hash for deduplication
        similarity_hash = self._calculate_similarity_hash(extraction)

        # Check for duplicates
        if similarity_hash in self.similarity_hashes:
            logger.info(f"Skipping duplicate pattern: {template_id}")
            return None

        # Extract style tokens
        style_tokens = generator_hints.get("style_tokens", {})
        if not style_tokens:
            # Infer from extraction if not provided
            style_tokens = self._infer_style_tokens(extraction)

        # Create template entry
        entry = TemplateEntry(
            id=template_id,
            source_file=source_file,
            slide_index=slide_index,
            slide_type=metadata.get("slide_type", "unknown"),
            template_category=generator_hints.get("template_category", metadata.get("slide_type", "content")),
            layout_category=self._infer_layout_category(elements),
            complexity_score=metadata.get("complexity_score", 3),
            reusability_score=generator_hints.get("reusability_score", 3),
            color_scheme=style_tokens.get("color_scheme", "corporate"),
            visual_style=style_tokens.get("visual_style", "classic"),
            primary_font=typography.get("title_font", style_tokens.get("primary_font", "Arial")),
            element_count=len(elements),
            has_chart=any(e.get("type") == "chart" for e in elements),
            has_table=any(e.get("type") == "table" for e in elements),
            has_image=any(e.get("type") == "image" for e in elements),
            content_zones=generator_hints.get("layout_zones", []),
            style_tokens=style_tokens,
            description_path=description_path or "",
            similarity_hash=similarity_hash
        )

        self.templates.append(entry)
        self.similarity_hashes[similarity_hash] = template_id
        self.categories[entry.template_category].append(template_id)

        logger.info(f"Added template: {template_id} ({entry.template_category})")
        return entry

    def _calculate_similarity_hash(self, extraction: Dict[str, Any]) -> str:
        """
        Calculate a hash representing the visual pattern of a slide.

        Similar slides will have similar hashes, enabling deduplication.
        """
        # Key features for similarity
        features = []

        # Slide type
        features.append(extraction.get("metadata", {}).get("slide_type", ""))

        # Element types and rough positions
        elements = extraction.get("elements", [])
        for elem in sorted(elements, key=lambda e: e.get("z_order", 0)):
            elem_type = elem.get("type", "")
            pos = elem.get("position", {})
            # Round positions to reduce sensitivity
            left = round(pos.get("left_inches", 0) * 2) / 2  # Round to 0.5"
            top = round(pos.get("top_inches", 0) * 2) / 2
            features.append(f"{elem_type}@{left},{top}")

        # Color scheme (simplified) - handle both list and dict formats
        palette = extraction.get("color_palette", {})
        if isinstance(palette, list):
            # Old format: list of colors
            features.append(palette[0] if palette else "")
            features.append(palette[1] if len(palette) > 1 else "")
        else:
            # New format: dict with named colors
            features.append(palette.get("primary", ""))
            features.append(palette.get("background", ""))

        # Create hash
        feature_str = "|".join(str(f) for f in features)
        return hashlib.md5(feature_str.encode()).hexdigest()[:12]

    def _infer_layout_category(self, elements: List[Dict]) -> str:
        """Infer layout category from element positions."""
        if not elements:
            return "blank"

        # Get element positions
        positions = []
        for elem in elements:
            pos = elem.get("position", {})
            if pos:
                center_x = pos.get("left_inches", 0) + pos.get("width_inches", 0) / 2
                positions.append(center_x)

        if not positions:
            return "unknown"

        # Check for column patterns
        left_count = sum(1 for x in positions if x < 5)
        right_count = sum(1 for x in positions if x > 8)
        center_count = sum(1 for x in positions if 5 <= x <= 8)

        if left_count > 0 and right_count > 0 and left_count + right_count > center_count:
            return "two_column"
        elif center_count > left_count + right_count:
            return "centered"
        elif left_count > right_count * 2:
            return "single_column"
        elif len(elements) > 6:
            return "grid"
        else:
            return "asymmetric"

    def _infer_slide_type(self, elements: List[Dict], design_notes: str = "") -> str:
        """Infer slide type from elements and design notes."""
        notes_lower = design_notes.lower()

        # Check design notes for hints
        if "title" in notes_lower and "slide" in notes_lower:
            return "title"
        if "section" in notes_lower or "divider" in notes_lower:
            return "section_divider"
        if "comparison" in notes_lower:
            return "comparison"
        if "timeline" in notes_lower or "process" in notes_lower:
            return "process"

        # Check element types
        has_chart = any(e.get("type") == "chart" for e in elements)
        has_table = any(e.get("type") == "table" for e in elements)
        has_image = any(e.get("type") == "image" for e in elements)

        if has_chart:
            return "data_chart"
        if has_table:
            return "content"
        if has_image and len(elements) <= 3:
            return "image_focus"

        # Check element count and layout
        text_elements = [e for e in elements if e.get("type") in ["textbox", "placeholder"]]
        if len(text_elements) == 1 and len(elements) <= 2:
            return "title"
        if len(text_elements) == 2 and len(elements) <= 4:
            return "section_divider"

        return "content"

    def _infer_style_tokens(self, extraction: Dict[str, Any]) -> Dict[str, Any]:
        """Infer style tokens when not provided by generator_hints."""
        palette = extraction.get("color_palette", {})
        typography = extraction.get("typography_system", {})

        # Handle both list and dict formats for color_palette
        if isinstance(palette, list):
            # Old format: list of colors - assume first is primary/background
            bg_color = (palette[0] if palette else "#FFFFFF").upper()
            primary = (palette[1] if len(palette) > 1 else "#000000").upper()
        else:
            # New format: dict with named colors
            bg_color = palette.get("background", "#FFFFFF").upper()
            primary = palette.get("primary", "#000000").upper()

        if bg_color in ["#FFFFFF", "#FAFAFA", "#F5F5F5"]:
            if primary in ["#000000", "#333333", "#666666"]:
                color_scheme = "monochrome"
            else:
                color_scheme = "light"
        elif bg_color in ["#000000", "#1A1A1A", "#2D2D2D"]:
            color_scheme = "dark"
        else:
            color_scheme = "corporate"

        # Determine visual style
        elements = extraction.get("elements", [])
        has_decorative = any(
            "decorative" in e.get("id", "").lower() or
            "line" in e.get("type", "").lower()
            for e in elements
        )

        if len(elements) <= 3 and not has_decorative:
            visual_style = "minimal"
        elif has_decorative:
            visual_style = "classic"
        else:
            visual_style = "modern"

        return {
            "primary_font": typography.get("title_font", "Arial"),
            "heading_weight": "bold",
            "color_scheme": color_scheme,
            "visual_style": visual_style
        }

    def build_index(self) -> TemplateLibrary:
        """
        Build the complete template library index.

        Returns:
            TemplateLibrary with all templates and metadata
        """
        # Build style clusters
        style_clusters = defaultdict(list)
        for template in self.templates:
            cluster_key = f"{template.color_scheme}_{template.visual_style}"
            style_clusters[cluster_key].append(template.id)

        # Calculate statistics
        statistics = {
            "total_templates": len(self.templates),
            "unique_patterns": len(self.similarity_hashes),
            "by_category": {k: len(v) for k, v in self.categories.items()},
            "by_style": {k: len(v) for k, v in style_clusters.items()},
            "avg_complexity": sum(t.complexity_score for t in self.templates) / len(self.templates) if self.templates else 0,
            "avg_reusability": sum(t.reusability_score for t in self.templates) / len(self.templates) if self.templates else 0,
            "with_charts": sum(1 for t in self.templates if t.has_chart),
            "with_tables": sum(1 for t in self.templates if t.has_table),
            "with_images": sum(1 for t in self.templates if t.has_image)
        }

        library = TemplateLibrary(
            total_templates=len(self.templates),
            unique_patterns=len(self.similarity_hashes),
            templates=self.templates,
            categories=dict(self.categories),
            style_clusters=dict(style_clusters),
            statistics=statistics
        )

        logger.info(f"Built library with {len(self.templates)} templates in {len(self.categories)} categories")
        return library

    def save(self, filename: str = "template_library.json") -> Path:
        """
        Save the template library to disk.

        Args:
            filename: Output filename

        Returns:
            Path to saved file
        """
        library = self.build_index()
        output_path = self.output_dir / filename

        # Convert to serializable format
        data = {
            "version": library.version,
            "total_templates": library.total_templates,
            "unique_patterns": library.unique_patterns,
            "templates": [asdict(t) for t in library.templates],
            "categories": library.categories,
            "style_clusters": library.style_clusters,
            "statistics": library.statistics
        }

        with open(output_path, 'w') as f:
            json.dump(data, f, indent=2)

        logger.info(f"Saved template library to: {output_path}")
        return output_path

    def load(self, filename: str = "template_library.json") -> TemplateLibrary:
        """
        Load an existing template library.

        Args:
            filename: Input filename

        Returns:
            Loaded TemplateLibrary
        """
        input_path = self.output_dir / filename

        with open(input_path, 'r') as f:
            data = json.load(f)

        templates = [TemplateEntry(**t) for t in data.get("templates", [])]

        library = TemplateLibrary(
            version=data.get("version", "1.0"),
            total_templates=data.get("total_templates", 0),
            unique_patterns=data.get("unique_patterns", 0),
            templates=templates,
            categories=data.get("categories", {}),
            style_clusters=data.get("style_clusters", {}),
            statistics=data.get("statistics", {})
        )

        # Rebuild internal state
        self.templates = templates
        self.categories = defaultdict(list, library.categories)
        self.similarity_hashes = {t.similarity_hash: t.id for t in templates}

        logger.info(f"Loaded template library: {len(templates)} templates")
        return library


class TemplateSearcher:
    """
    Searches the template library to find matching templates for content.
    """

    def __init__(self, library: TemplateLibrary):
        """
        Initialize searcher with a template library.

        Args:
            library: The template library to search
        """
        self.library = library
        self._build_search_index()

    def _build_search_index(self):
        """Build internal search indices."""
        self.by_category = defaultdict(list)
        self.by_style = defaultdict(list)
        self.by_type = defaultdict(list)

        for template in self.library.templates:
            self.by_category[template.template_category].append(template)
            self.by_style[f"{template.color_scheme}_{template.visual_style}"].append(template)
            self.by_type[template.slide_type].append(template)

    def find_templates(
        self,
        slide_type: Optional[str] = None,
        template_category: Optional[str] = None,
        color_scheme: Optional[str] = None,
        visual_style: Optional[str] = None,
        min_reusability: int = 1,
        needs_chart: bool = False,
        needs_table: bool = False,
        limit: int = 10
    ) -> List[Tuple[TemplateEntry, float]]:
        """
        Find templates matching the given criteria.

        Args:
            slide_type: Filter by slide type
            template_category: Filter by template category
            color_scheme: Preferred color scheme
            visual_style: Preferred visual style
            min_reusability: Minimum reusability score
            needs_chart: Whether the slide needs a chart
            needs_table: Whether the slide needs a table
            limit: Maximum results

        Returns:
            List of (template, score) tuples sorted by relevance
        """
        results = []

        for template in self.library.templates:
            score = 0.0

            # Hard filters
            if template.reusability_score < min_reusability:
                continue
            if needs_chart and not template.has_chart:
                continue
            if needs_table and not template.has_table:
                continue

            # Scoring
            if slide_type and template.slide_type == slide_type:
                score += 3.0
            if template_category and template.template_category == template_category:
                score += 2.0
            if color_scheme and template.color_scheme == color_scheme:
                score += 1.0
            if visual_style and template.visual_style == visual_style:
                score += 1.0

            # Bonus for high reusability
            score += template.reusability_score * 0.2

            if score > 0:
                results.append((template, score))

        # Sort by score
        results.sort(key=lambda x: x[1], reverse=True)
        return results[:limit]

    def get_best_template(
        self,
        content_type: str,
        style_preferences: Optional[Dict[str, str]] = None
    ) -> Optional[TemplateEntry]:
        """
        Get the single best template for a content type.

        Args:
            content_type: Type of content (e.g., "section_header", "data_table")
            style_preferences: Optional style preferences

        Returns:
            Best matching template or None
        """
        prefs = style_preferences or {}

        results = self.find_templates(
            template_category=content_type,
            color_scheme=prefs.get("color_scheme"),
            visual_style=prefs.get("visual_style"),
            min_reusability=3,
            limit=1
        )

        return results[0][0] if results else None

    def get_style_summary(self) -> Dict[str, Any]:
        """Get a summary of available styles in the library."""
        return {
            "categories": list(self.by_category.keys()),
            "styles": list(self.by_style.keys()),
            "types": list(self.by_type.keys()),
            "total_templates": len(self.library.templates)
        }
