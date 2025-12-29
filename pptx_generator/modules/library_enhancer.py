"""
Library Enhancer Module

Provides tools for:
1. Adding domain-specific tags to library components
2. Semantic search across components
3. Auto-matching components to presentation content
4. Style/color palette application
"""

import json
import logging
import re
from pathlib import Path
from typing import Optional, List, Dict, Any, TYPE_CHECKING
from dataclasses import dataclass, field

if TYPE_CHECKING:
    from .content_classifier import ContentClassifier

logger = logging.getLogger(__name__)

# Real estate domain categories for tagging
REAL_ESTATE_DOMAINS = {
    "market_analysis": [
        "vacancy", "occupancy", "rent growth", "cap rate", "absorption",
        "supply", "demand", "inventory", "pipeline", "construction",
        "transaction volume", "pricing", "fundamentals"
    ],
    "financial_metrics": [
        "irr", "multiple", "yield", "noi", "cash flow", "revenue",
        "return", "equity", "debt", "leverage", "distribution",
        "appreciation", "income", "expense", "margin"
    ],
    "property_types": [
        "industrial", "warehouse", "logistics", "flex", "office",
        "multifamily", "retail", "hospitality", "mixed-use",
        "self-storage", "data center", "life science"
    ],
    "geographic": [
        "market", "submarket", "metro", "msa", "region", "sun belt",
        "gateway", "secondary", "tertiary", "coastal", "midwest"
    ],
    "investment": [
        "fund", "portfolio", "acquisition", "disposition", "hold",
        "value-add", "core", "opportunistic", "development", "stabilized"
    ],
    "comparison": [
        "vs", "comparison", "benchmark", "peer", "competitor",
        "historical", "trend", "forecast", "projection"
    ],
    "timeline": [
        "year", "quarter", "month", "ytd", "ttm", "forecast",
        "projection", "historical", "trend"
    ]
}

# Chart type mappings for better matching
CHART_PURPOSE_MAP = {
    "column": ["comparison", "time_series", "ranking"],
    "bar": ["comparison", "ranking", "distribution"],
    "line": ["time_series", "trend", "projection"],
    "pie": ["composition", "breakdown", "allocation"],
    "area": ["cumulative", "time_series", "stacked"],
    "scatter": ["correlation", "distribution", "clustering"],
    "waterfall": ["bridge", "contribution", "breakdown"]
}

# Shape purpose classifications
SHAPE_PURPOSES = {
    "icon": ["OVAL", "ROUNDED_RECTANGLE", "DIAMOND", "HEXAGON"],
    "separator": ["LINE", "RECTANGLE"],  # thin shapes
    "background": ["RECTANGLE"],  # full-width shapes
    "callout": ["ROUNDED_RECTANGLE", "CALLOUT"],
    "arrow": ["ARROW", "CHEVRON", "RIGHT_ARROW"],
    "container": ["RECTANGLE", "ROUNDED_RECTANGLE"]
}


@dataclass
class EnhancedComponent:
    """Enhanced component with domain tags and semantic metadata."""
    id: str
    type: str
    original_category: str
    domain_tags: List[str] = field(default_factory=list)
    purpose: str = ""
    semantic_score: float = 0.0
    content_keywords: List[str] = field(default_factory=list)
    style_profile: Dict[str, Any] = field(default_factory=dict)


class LibraryEnhancer:
    """
    Enhances the ComponentLibrary with domain-specific tagging,
    semantic search, and smart matching capabilities.
    """

    def __init__(self, library):
        """
        Initialize with a ComponentLibrary instance.

        Args:
            library: ComponentLibrary instance
        """
        self.library = library
        self.enhanced_index: Dict[str, EnhancedComponent] = {}
        self.domain_index: Dict[str, List[str]] = {}  # domain -> component_ids
        self._build_enhanced_index()

    def _build_enhanced_index(self):
        """Build enhanced index with domain tags."""
        # Process charts
        for chart in self.library.index.get('components', {}).get('charts', []):
            enhanced = self._enhance_chart(chart)
            self.enhanced_index[enhanced.id] = enhanced
            self._index_by_domain(enhanced)

        # Process tables
        for table in self.library.index.get('components', {}).get('tables', []):
            enhanced = self._enhance_table(table)
            self.enhanced_index[enhanced.id] = enhanced
            self._index_by_domain(enhanced)

        # Process shapes
        for shape in self.library.index.get('components', {}).get('shapes', []):
            enhanced = self._enhance_shape(shape)
            self.enhanced_index[enhanced.id] = enhanced
            self._index_by_domain(enhanced)

        logger.info(f"Enhanced {len(self.enhanced_index)} components")

    def _enhance_chart(self, chart: dict) -> EnhancedComponent:
        """Add domain tags to a chart based on its content."""
        enhanced = EnhancedComponent(
            id=chart['id'],
            type='chart',
            original_category=chart.get('category', 'uncategorized')
        )

        # Get chart data for content analysis
        chart_data = self.library.get_chart_data(chart['id'])
        if chart_data:
            # Extract keywords from categories and series names
            keywords = []
            for cat in chart_data.get('categories', []):
                keywords.extend(self._extract_keywords(str(cat)))
            for series in chart_data.get('series', []):
                keywords.extend(self._extract_keywords(series.get('name', '')))

            enhanced.content_keywords = list(set(keywords))

            # Match to domain tags
            enhanced.domain_tags = self._match_domains(keywords)

        # Determine chart purpose based on type
        chart_type = chart.get('chart_type', '').lower()
        for ctype, purposes in CHART_PURPOSE_MAP.items():
            if ctype in chart_type:
                enhanced.purpose = purposes[0]
                break

        # Add structure metadata
        enhanced.style_profile = {
            'series_count': chart.get('series_count', 1),
            'category_count': chart.get('category_count', 0),
            'chart_type': chart.get('chart_type', ''),
            'width': chart.get('width_inches', 0),
            'height': chart.get('height_inches', 0)
        }

        return enhanced

    def _enhance_table(self, table: dict) -> EnhancedComponent:
        """Add domain tags to a table."""
        enhanced = EnhancedComponent(
            id=table['id'],
            type='table',
            original_category=table.get('category', 'uncategorized')
        )

        # Get table data for content analysis
        table_data = self.library.get_table_data(table['id'])
        if table_data:
            keywords = []
            for row in table_data.get('data', []):
                for cell in row:
                    keywords.extend(self._extract_keywords(str(cell)))

            enhanced.content_keywords = list(set(keywords))
            enhanced.domain_tags = self._match_domains(keywords)

        # Set purpose based on category
        category = table.get('category', '')
        if 'comparison' in category:
            enhanced.purpose = 'comparison'
        elif 'data' in category:
            enhanced.purpose = 'data_display'
        else:
            enhanced.purpose = 'general'

        # Structure metadata
        enhanced.style_profile = {
            'rows': table.get('rows', 0),
            'cols': table.get('cols', 0),
            'width': table.get('width_inches', 0),
            'height': table.get('height_inches', 0)
        }

        return enhanced

    def _enhance_shape(self, shape: dict) -> EnhancedComponent:
        """Classify shape by purpose."""
        enhanced = EnhancedComponent(
            id=shape['id'],
            type='shape',
            original_category=shape.get('category', 'uncategorized')
        )

        shape_type = (shape.get('shape_type') or '').upper()
        width = shape.get('width_inches') or 0
        height = shape.get('height_inches') or 0

        # Classify by dimensions and type
        if width > 12 and height > 6:
            enhanced.purpose = 'background'
        elif width > 10 and height < 0.1:
            enhanced.purpose = 'separator'
        elif height > 10 and width < 0.1:
            enhanced.purpose = 'separator'
        elif width < 1 and height < 1:
            enhanced.purpose = 'icon'
        elif 'ARROW' in shape_type or 'CHEVRON' in shape_type:
            enhanced.purpose = 'arrow'
        elif 'CALLOUT' in shape_type:
            enhanced.purpose = 'callout'
        else:
            enhanced.purpose = 'container'

        enhanced.style_profile = {
            'shape_type': shape_type,
            'width': width,
            'height': height,
            'has_text': shape.get('has_text', False),
            'colors': shape.get('colors', [])
        }

        return enhanced

    def _extract_keywords(self, text: str) -> List[str]:
        """Extract meaningful keywords from text."""
        # Clean and split
        text = text.lower()
        text = re.sub(r'[^\w\s]', ' ', text)
        words = text.split()

        # Filter out common placeholders and short words
        stopwords = {'the', 'a', 'an', 'and', 'or', 'is', 'in', 'to', 'of', 'for',
                     'category', 'series', 'segment', 'answer', 'column', 'row',
                     'parameter', 'value', 'item', 'type', 'name'}

        return [w for w in words if len(w) > 2 and w not in stopwords]

    def _match_domains(self, keywords: List[str]) -> List[str]:
        """Match keywords to domain categories."""
        matched_domains = set()

        for domain, domain_keywords in REAL_ESTATE_DOMAINS.items():
            for keyword in keywords:
                for dk in domain_keywords:
                    if keyword in dk or dk in keyword:
                        matched_domains.add(domain)
                        break

        return list(matched_domains)

    def _index_by_domain(self, component: EnhancedComponent):
        """Add component to domain index."""
        for domain in component.domain_tags:
            if domain not in self.domain_index:
                self.domain_index[domain] = []
            self.domain_index[domain].append(component.id)

    # ==================== Search Methods ====================

    def search_by_domain(
        self,
        domains: List[str],
        component_type: Optional[str] = None,
        limit: int = 20
    ) -> List[EnhancedComponent]:
        """
        Search for components matching domain categories.

        Args:
            domains: List of domain categories to match
            component_type: Filter by component type
            limit: Maximum results

        Returns:
            List of matching EnhancedComponent objects
        """
        matching_ids = set()

        for domain in domains:
            matching_ids.update(self.domain_index.get(domain, []))

        results = []
        for comp_id in matching_ids:
            comp = self.enhanced_index.get(comp_id)
            if comp:
                if component_type and comp.type != component_type:
                    continue
                results.append(comp)

        # Sort by number of matching domains
        results.sort(key=lambda c: len(set(c.domain_tags) & set(domains)), reverse=True)

        return results[:limit]

    def search_by_purpose(
        self,
        purpose: str,
        component_type: Optional[str] = None,
        limit: int = 20
    ) -> List[EnhancedComponent]:
        """
        Search for components by purpose.

        Args:
            purpose: Purpose to match (e.g., 'comparison', 'icon', 'separator')
            component_type: Filter by type
            limit: Maximum results
        """
        results = []

        for comp in self.enhanced_index.values():
            if comp.purpose == purpose:
                if component_type and comp.type != component_type:
                    continue
                results.append(comp)

        return results[:limit]

    def find_best_chart(
        self,
        chart_type: str,
        series_count: int = 1,
        category_count: int = 4,
        domains: Optional[List[str]] = None
    ) -> Optional[str]:
        """
        Find the best matching chart from the library.

        Args:
            chart_type: Desired chart type (column, bar, line, pie)
            series_count: Number of data series needed
            category_count: Number of categories needed
            domains: Optional domain tags to prioritize

        Returns:
            Component ID of best match, or None
        """
        candidates = []

        for comp in self.enhanced_index.values():
            if comp.type != 'chart':
                continue

            profile = comp.style_profile
            comp_type = profile.get('chart_type', '').lower()

            # Must match chart type
            if chart_type.lower() not in comp_type:
                continue

            # Score based on structure match
            series_diff = abs(profile.get('series_count', 1) - series_count)
            cat_diff = abs(profile.get('category_count', 4) - category_count)
            structure_score = 10 - (series_diff + cat_diff)

            # Bonus for domain match
            domain_score = 0
            if domains:
                domain_score = len(set(comp.domain_tags) & set(domains)) * 2

            total_score = structure_score + domain_score
            candidates.append((comp.id, total_score))

        if not candidates:
            return None

        # Return best match
        candidates.sort(key=lambda x: x[1], reverse=True)
        return candidates[0][0]

    def find_best_table(
        self,
        rows: int,
        cols: int,
        purpose: str = 'data_display',
        domains: Optional[List[str]] = None
    ) -> Optional[str]:
        """
        Find the best matching table from the library.

        Args:
            rows: Number of rows needed
            cols: Number of columns needed
            purpose: Table purpose (data_display, comparison)
            domains: Optional domain tags

        Returns:
            Component ID of best match, or None
        """
        candidates = []

        for comp in self.enhanced_index.values():
            if comp.type != 'table':
                continue

            profile = comp.style_profile

            # Score based on dimension match (prefer larger tables that can accommodate)
            comp_rows = profile.get('rows', 0)
            comp_cols = profile.get('cols', 0)

            if comp_rows < rows or comp_cols < cols:
                continue  # Can't fit our data

            # Prefer closer matches
            size_diff = (comp_rows - rows) + (comp_cols - cols)
            size_score = 10 - min(size_diff, 10)

            # Purpose match
            purpose_score = 3 if comp.purpose == purpose else 0

            # Domain match
            domain_score = 0
            if domains:
                domain_score = len(set(comp.domain_tags) & set(domains)) * 2

            total_score = size_score + purpose_score + domain_score
            candidates.append((comp.id, total_score))

        if not candidates:
            return None

        candidates.sort(key=lambda x: x[1], reverse=True)
        return candidates[0][0]

    def find_icons(self, purpose: str = 'icon', limit: int = 10) -> List[str]:
        """Find icon shapes from the library."""
        return [
            comp.id for comp in self.search_by_purpose(purpose, 'shape', limit)
        ]

    def find_separators(self, limit: int = 10) -> List[str]:
        """Find separator shapes (lines, dividers)."""
        return [
            comp.id for comp in self.search_by_purpose('separator', 'shape', limit)
        ]

    # ==================== Style Retrieval ====================

    def get_color_palette_for_domain(self, domain: str) -> Optional[dict]:
        """Get a suitable color palette for a domain."""
        # For now, return first available palette
        # Could be enhanced to match domains to specific palettes
        palettes = self.library.get_color_palettes()
        return palettes[0] if palettes else None

    def get_typography_for_element(self, element_type: str) -> Optional[dict]:
        """Get typography preset for an element type (title, body, etc.)."""
        presets = self.library.get_typography_presets(preset_type=element_type)
        return presets[0] if presets else None

    # ==================== Content Analysis ====================

    def analyze_outline_content(self, outline: dict) -> Dict[str, List[str]]:
        """
        Analyze outline content to determine relevant domains.

        Args:
            outline: Presentation outline dictionary

        Returns:
            Dict mapping section names to relevant domain tags
        """
        section_domains = {}

        # Extract text from outline
        title = outline.get('title', '')
        context = outline.get('context', {})

        # Global domains from title and context
        global_keywords = self._extract_keywords(title)
        for key, value in context.items():
            if isinstance(value, str):
                global_keywords.extend(self._extract_keywords(value))

        global_domains = self._match_domains(global_keywords)

        # Analyze each section
        for section in outline.get('sections', []):
            section_name = section.get('name', '')
            keywords = self._extract_keywords(section_name)

            # Add keywords from slide content
            for slide in section.get('slides', []):
                content = slide.get('content', {})
                title = content.get('title', '')
                keywords.extend(self._extract_keywords(title))

                # Bullets
                for bullet in content.get('bullets', []):
                    keywords.extend(self._extract_keywords(bullet))

            domains = self._match_domains(keywords)
            section_domains[section_name] = list(set(domains + global_domains))

        return section_domains

    def suggest_components_for_outline(self, outline: dict) -> Dict[str, List[dict]]:
        """
        Suggest library components for each slide in an outline.

        Args:
            outline: Presentation outline

        Returns:
            Dict mapping slide keys to suggested components
        """
        suggestions = {}
        section_domains = self.analyze_outline_content(outline)

        for section in outline.get('sections', []):
            section_name = section.get('name', '')
            domains = section_domains.get(section_name, [])

            for i, slide in enumerate(section.get('slides', [])):
                slide_key = f"{section_name}_{i}"
                slide_type = slide.get('slide_type', '')
                content = slide.get('content', {})

                slide_suggestions = []

                # Suggest charts for data_chart slides
                if slide_type == 'data_chart':
                    chart_data = content.get('chart_data', {})
                    chart_type = chart_data.get('type', 'column')
                    series = chart_data.get('series', [])
                    categories = chart_data.get('categories', [])

                    best_chart = self.find_best_chart(
                        chart_type=chart_type,
                        series_count=len(series),
                        category_count=len(categories),
                        domains=domains
                    )

                    if best_chart:
                        slide_suggestions.append({
                            'type': 'chart',
                            'component_id': best_chart,
                            'reason': f'Matched {chart_type} chart with {len(series)} series'
                        })

                # Suggest tables for table_slide
                elif slide_type == 'table_slide':
                    headers = content.get('headers', [])
                    data = content.get('data', [])

                    best_table = self.find_best_table(
                        rows=len(data) + 1,  # +1 for header
                        cols=len(headers),
                        domains=domains
                    )

                    if best_table:
                        slide_suggestions.append({
                            'type': 'table',
                            'component_id': best_table,
                            'reason': f'Matched {len(data)+1}x{len(headers)} table'
                        })

                if slide_suggestions:
                    suggestions[slide_key] = slide_suggestions

        return suggestions

    # ==================== Statistics ====================

    def get_domain_stats(self) -> Dict[str, int]:
        """Get count of components per domain."""
        return {domain: len(ids) for domain, ids in self.domain_index.items()}

    def get_purpose_stats(self) -> Dict[str, Dict[str, int]]:
        """Get count of components per purpose, by type."""
        stats = {}

        for comp in self.enhanced_index.values():
            if comp.type not in stats:
                stats[comp.type] = {}

            purpose = comp.purpose or 'unknown'
            stats[comp.type][purpose] = stats[comp.type].get(purpose, 0) + 1

        return stats

    # ==================== Classified Content Access ====================

    def get_classified_content(self) -> Optional['ContentClassifier']:
        """
        Get access to classified content (images, icons, diagrams).

        Returns:
            ContentClassifier instance if available, None otherwise
        """
        try:
            from .content_classifier import ContentClassifier
            classifier = ContentClassifier(Path(self.library.library_path))

            # If no classifications exist, run classification
            if not classifier.images:
                classifier.classify_all(self.library.index)

            return classifier
        except Exception as e:
            logger.warning(f"Could not load ContentClassifier: {e}")
            return None

    def find_logo(self) -> Optional[str]:
        """
        Find a logo image from the library.

        Returns:
            Image ID if found, None otherwise
        """
        classifier = self.get_classified_content()
        if classifier:
            logo = classifier.find_logo()
            if logo:
                return logo.id
        return None

    def find_background_image(self, min_width: float = 10.0) -> Optional[str]:
        """
        Find a background image from the library.

        Args:
            min_width: Minimum width in inches

        Returns:
            Image ID if found, None otherwise
        """
        classifier = self.get_classified_content()
        if classifier:
            bg = classifier.find_background(min_width)
            if bg:
                return bg.id
        return None

    def find_icons(self, count: int = 4) -> List[str]:
        """
        Find icon images suitable for bullet lists.

        Args:
            count: Number of icons to find

        Returns:
            List of image/shape IDs
        """
        classifier = self.get_classified_content()
        if classifier:
            icons = classifier.find_icons_for_list(count)
            return [icon.id for icon in icons]
        return []

    def find_photo_images(self, limit: int = 10) -> List[str]:
        """
        Find photo images from the library.

        Args:
            limit: Maximum number to return

        Returns:
            List of image IDs
        """
        classifier = self.get_classified_content()
        if classifier:
            photos = classifier.get_images_by_category('photo')
            return [p.id for p in photos[:limit]]
        return []

    def find_diagram_template(self, step_count: int = 4) -> Optional[Dict[str, Any]]:
        """
        Find a diagram template suitable for a process.

        Args:
            step_count: Number of steps in the process

        Returns:
            Diagram template dict if found, None otherwise
        """
        classifier = self.get_classified_content()
        if classifier:
            template = classifier.find_diagram_for_process(step_count)
            if template:
                return {
                    'id': template.id,
                    'type': template.type,
                    'name': template.name,
                    'component_count': template.component_count,
                    'layout': template.layout
                }
        return None

    def get_image_categories(self) -> Dict[str, int]:
        """
        Get statistics on image categories.

        Returns:
            Dict mapping category to count
        """
        classifier = self.get_classified_content()
        if classifier:
            return classifier.get_image_stats()
        return {}


def main():
    """Test the library enhancer."""
    import argparse
    from .component_library import ComponentLibrary

    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

    parser = argparse.ArgumentParser(description="Library Enhancer")
    parser.add_argument("--stats", action="store_true", help="Show enhanced stats")
    parser.add_argument("--domains", action="store_true", help="Show domain distribution")
    parser.add_argument("--purposes", action="store_true", help="Show purpose distribution")
    parser.add_argument("--search-domain", type=str, help="Search by domain")
    parser.add_argument("--find-chart", type=str, help="Find chart type")

    args = parser.parse_args()

    library = ComponentLibrary()
    enhancer = LibraryEnhancer(library)

    if args.stats:
        print(f"\nEnhanced {len(enhancer.enhanced_index)} components")

    if args.domains:
        print("\nDomain Distribution:")
        for domain, count in sorted(enhancer.get_domain_stats().items(), key=lambda x: -x[1]):
            print(f"  {domain}: {count}")

    if args.purposes:
        print("\nPurpose Distribution:")
        for comp_type, purposes in enhancer.get_purpose_stats().items():
            print(f"\n  {comp_type}:")
            for purpose, count in sorted(purposes.items(), key=lambda x: -x[1]):
                print(f"    {purpose}: {count}")

    if args.search_domain:
        results = enhancer.search_by_domain([args.search_domain], limit=10)
        print(f"\nComponents matching domain '{args.search_domain}':")
        for comp in results:
            print(f"  [{comp.type}] {comp.id} - {comp.purpose}")

    if args.find_chart:
        chart_id = enhancer.find_best_chart(args.find_chart, series_count=2, category_count=5)
        if chart_id:
            print(f"\nBest {args.find_chart} chart: {chart_id}")
        else:
            print(f"\nNo matching {args.find_chart} chart found")


if __name__ == "__main__":
    main()
