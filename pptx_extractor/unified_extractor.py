"""
Unified Library Extractor - Master extractor that runs all component extractors.

This module provides a single entry point to extract all types of components:
- Basic components (images, charts, tables, shapes, diagrams)
- Styles (colors, typography, effects, gradients, shadows)
- Chart styles
- Layout blueprints
- Diagram templates
- Text patterns
- Slide sequences

The unified extractor maintains a master index that references all component types.
"""

import json
import logging
from pathlib import Path
from datetime import datetime
from typing import Optional, Dict, List, Any

from .library_extractor import LibraryExtractor
from .style_extractor import StyleExtractor
from .chart_style_extractor import ChartStyleExtractor
from .layout_blueprint_extractor import LayoutBlueprintExtractor
from .diagram_template_extractor import DiagramTemplateExtractor
from .text_template_extractor import TextTemplateExtractor
from .sequence_extractor import SequenceExtractor

logger = logging.getLogger(__name__)


class UnifiedLibraryExtractor:
    """
    Master extractor that runs all component extractors and maintains a unified index.
    """

    def __init__(self, output_dir: Path):
        """
        Initialize the unified extractor.

        Args:
            output_dir: Base directory for the component library
        """
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)

        # Initialize all sub-extractors
        self.extractors = {
            'components': LibraryExtractor(self.output_dir),
            'styles': StyleExtractor(self.output_dir),
            'chart_styles': ChartStyleExtractor(self.output_dir),
            'layouts': LayoutBlueprintExtractor(self.output_dir),
            'diagrams': DiagramTemplateExtractor(self.output_dir),
            'text_templates': TextTemplateExtractor(self.output_dir),
            'sequences': SequenceExtractor(self.output_dir),
        }

        # Master index path
        self.master_index_path = self.output_dir / 'master_index.json'
        self.master_index = self._load_master_index()

    def _load_master_index(self) -> dict:
        """Load or create master index."""
        if self.master_index_path.exists():
            with open(self.master_index_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        return {
            'metadata': {
                'created': datetime.now().isoformat(),
                'last_updated': datetime.now().isoformat(),
                'version': '2.0',
            },
            'templates': {},
            'summaries': {},
            'indexes': {
                'components': 'library_index.json',
                'styles': 'styles/style_index.json',
                'chart_styles': 'styles/chart_style_index.json',
                'layouts': 'layouts/layout_index.json',
                'diagrams': 'diagrams/diagram_template_index.json',
                'text_templates': 'text_templates/text_template_index.json',
                'sequences': 'sequences/sequence_index.json',
            },
        }

    def _save_master_index(self):
        """Save master index."""
        self.master_index['metadata']['last_updated'] = datetime.now().isoformat()
        with open(self.master_index_path, 'w', encoding='utf-8') as f:
            json.dump(self.master_index, f, indent=2, ensure_ascii=False)

    def extract_template(self, pptx_path: Path, template_name: Optional[str] = None) -> dict:
        """
        Extract all components from a PowerPoint template using all extractors.

        Args:
            pptx_path: Path to the PowerPoint file
            template_name: Optional name for the template

        Returns:
            Comprehensive extraction results
        """
        pptx_path = Path(pptx_path)
        template_name = template_name or pptx_path.stem

        logger.info(f"Starting unified extraction for: {template_name}")
        print(f"\n{'='*60}")
        print(f"Extracting: {template_name}")
        print(f"{'='*60}")

        results = {
            'template': template_name,
            'source_file': str(pptx_path),
            'extracted_date': datetime.now().isoformat(),
            'extraction_results': {},
        }

        # Run each extractor
        extractors_to_run = [
            ('components', 'Basic Components', 'extract_template'),
            ('styles', 'Styles & Effects', 'extract_template_styles'),
            ('chart_styles', 'Chart Styles', 'extract_template_chart_styles'),
            ('layouts', 'Layout Blueprints', 'extract_template_layouts'),
            ('diagrams', 'Diagram Templates', 'extract_template_diagrams'),
            ('text_templates', 'Text Patterns', 'extract_template_text_patterns'),
            ('sequences', 'Slide Sequences', 'extract_template_sequences'),
        ]

        for extractor_name, display_name, method_name in extractors_to_run:
            print(f"\n  Extracting {display_name}...", end=' ')
            try:
                extractor = self.extractors[extractor_name]
                method = getattr(extractor, method_name)
                result = method(pptx_path, template_name)
                results['extraction_results'][extractor_name] = self._summarize_result(result)
                print("Done")
            except Exception as e:
                logger.warning(f"Failed to extract {display_name}: {e}")
                print(f"Error: {e}")
                results['extraction_results'][extractor_name] = {'error': str(e)}

        # Update master index
        self.master_index['templates'][template_name] = {
            'source_file': str(pptx_path),
            'extracted_date': results['extracted_date'],
            'summary': self._create_template_summary(results),
        }
        self._save_master_index()

        return results

    def _summarize_result(self, result: dict) -> dict:
        """Create a summary of extraction result."""
        if isinstance(result, dict):
            summary = {}
            for key, value in result.items():
                if isinstance(value, list):
                    summary[key] = len(value)
                elif isinstance(value, dict) and 'components' in key.lower():
                    summary[key] = {k: len(v) if isinstance(v, list) else v
                                   for k, v in value.items()}
                elif key not in ('template', 'source_file', 'slide_dimensions'):
                    summary[key] = value
            return summary
        return result

    def _create_template_summary(self, results: dict) -> dict:
        """Create a summary for the template."""
        summary = {}
        for extractor_name, result in results.get('extraction_results', {}).items():
            if isinstance(result, dict) and 'error' not in result:
                summary[extractor_name] = result
        return summary

    def extract_all_templates(self, templates_dir: Path) -> dict:
        """
        Extract all components from all templates in a directory.

        Args:
            templates_dir: Directory containing PPTX templates

        Returns:
            Summary of all extractions
        """
        templates_dir = Path(templates_dir)

        # Find all PPTX files
        pptx_files = list(templates_dir.rglob('*.pptx'))

        print(f"\nFound {len(pptx_files)} PowerPoint files")
        print("="*60)

        all_results = {}

        for pptx_path in pptx_files:
            try:
                result = self.extract_template(pptx_path)
                all_results[pptx_path.stem] = result
            except Exception as e:
                print(f"Error processing {pptx_path.name}: {e}")
                all_results[pptx_path.stem] = {'error': str(e)}

        # Update summaries in master index
        self.master_index['summaries'] = self.get_full_summary()
        self._save_master_index()

        return all_results

    def get_full_summary(self) -> dict:
        """Get a full summary of all extracted content."""
        summary = {
            'total_templates': len(self.master_index['templates']),
            'components': self.extractors['components'].get_summary(),
            'styles': self.extractors['styles'].get_summary(),
            'chart_styles': self.extractors['chart_styles'].get_summary(),
            'layouts': self.extractors['layouts'].get_summary(),
            'diagrams': self.extractors['diagrams'].get_summary(),
            'text_templates': self.extractors['text_templates'].get_summary(),
            'sequences': self.extractors['sequences'].get_summary(),
        }
        return summary

    def search(self, query: str = None, component_type: str = None,
               category: str = None, template: str = None) -> dict:
        """
        Search across all component types.

        Args:
            query: Text query to search
            component_type: Specific component type to search
            category: Category filter
            template: Template name filter

        Returns:
            Search results organized by component type
        """
        results = {}

        # Search basic components
        if not component_type or component_type in ('images', 'charts', 'tables', 'shapes', 'diagrams'):
            results['components'] = self.extractors['components'].search(
                component_type=component_type,
                category=category,
                template=template
            )

        # Search styles
        if not component_type or component_type in ('colors', 'typography', 'effects'):
            if component_type == 'colors':
                results['colors'] = self.extractors['styles'].search_colors(
                    query=query, template=template
                )
            elif component_type == 'typography':
                results['typography'] = self.extractors['styles'].search_typography(
                    template=template
                )
            elif component_type == 'effects':
                results['effects'] = self.extractors['styles'].search_effects(
                    template=template
                )

        # Search chart styles
        if not component_type or component_type == 'chart_styles':
            results['chart_styles'] = self.extractors['chart_styles'].search(
                template=template
            )

        # Search layouts
        if not component_type or component_type in ('layouts', 'blueprints', 'grids'):
            results['layouts'] = self.extractors['layouts'].search_blueprints(
                category=category, template=template
            )
            results['grids'] = self.extractors['layouts'].search_grids()

        # Search diagram templates
        if not component_type or component_type == 'diagram_templates':
            results['diagram_templates'] = self.extractors['diagrams'].search(
                category=category, template=template
            )

        # Search text templates
        if not component_type or component_type in ('text', 'bullets', 'titles'):
            if component_type == 'bullets':
                results['bullet_patterns'] = self.extractors['text_templates'].search_bullet_patterns(
                    template=template
                )
            elif component_type == 'titles':
                results['title_patterns'] = self.extractors['text_templates'].search_title_patterns(
                    template=template
                )

        # Search sequences
        if not component_type or component_type == 'sequences':
            results['sequences'] = self.extractors['sequences'].search_sequences(
                template=template
            )

        return results

    def get_component(self, component_type: str, component_id: str) -> Optional[dict]:
        """
        Get a specific component by type and ID.

        Args:
            component_type: Type of component
            component_id: Component ID

        Returns:
            Component data or None
        """
        type_mapping = {
            'chart_style': ('chart_styles', 'get_style_by_id'),
            'diagram_template': ('diagrams', 'get_by_id'),
            'text_template': ('text_templates', 'get_by_id'),
            'layout': ('layouts', 'search_blueprints'),
        }

        if component_type in type_mapping:
            extractor_name, method_name = type_mapping[component_type]
            extractor = self.extractors[extractor_name]
            method = getattr(extractor, method_name)
            return method(component_id)

        return None

    def find_layout_for_content(self, content_requirements: dict) -> List[dict]:
        """
        Find suitable layouts for given content requirements.

        Args:
            content_requirements: Dict with content counts, e.g.:
                {'chart': 2, 'table': 1}

        Returns:
            List of matching layout blueprints
        """
        return self.extractors['layouts'].find_layout_for_content(content_requirements)

    def generate_deck_outline(self, structure_type: str, topic: str = 'Presentation') -> dict:
        """
        Generate a deck outline based on structure type.

        Args:
            structure_type: Type of structure (executive_presentation, data_heavy, etc.)
            topic: Topic/title for the presentation

        Returns:
            Deck outline with slide placeholders
        """
        return self.extractors['sequences'].generate_deck_outline(structure_type, topic)


def extract_all(templates_dir: Path, output_dir: Path) -> dict:
    """
    Convenience function to extract all components from all templates.

    Args:
        templates_dir: Directory containing PPTX templates
        output_dir: Output directory for the library

    Returns:
        Summary of extraction
    """
    extractor = UnifiedLibraryExtractor(output_dir)
    extractor.extract_all_templates(templates_dir)
    return extractor.get_full_summary()


if __name__ == '__main__':
    import sys

    if len(sys.argv) < 3:
        print("Usage: python unified_extractor.py <templates_dir> <output_dir>")
        sys.exit(1)

    templates_dir = Path(sys.argv[1])
    output_dir = Path(sys.argv[2])

    summary = extract_all(templates_dir, output_dir)

    print("\n" + "="*60)
    print("UNIFIED EXTRACTION COMPLETE")
    print("="*60)
    print(json.dumps(summary, indent=2))
