"""
Optimized Slide Extractor

Implements the two-phase extraction strategy for maximum cost efficiency:

Phase 1: Fast Categorization (Haiku + Batch API)
- Quick classification of all slides
- Identifies slide types and patterns
- Groups similar slides for deduplication
- Cost: ~$4.50 for 500 slides

Phase 2: Detailed Extraction (Sonnet + Prompt Caching)
- Full extraction of unique template patterns only
- Generator-optimized output with hints
- Cost: ~$5-10 for 20-50 unique templates

Total savings: ~70-80% compared to extracting all slides with Sonnet

Usage:
    extractor = OptimizedExtractor(output_dir="descriptions")
    library = extractor.process_template("template.pptx")
"""

import json
import logging
from pathlib import Path
from typing import Optional, List, Dict, Any, Tuple
from dataclasses import dataclass, field
from collections import defaultdict
from datetime import datetime
import hashlib

from rich.console import Console
from rich.progress import Progress, SpinnerColumn, TextColumn, BarColumn
from rich.table import Table

logger = logging.getLogger(__name__)
console = Console()


@dataclass
class ExtractionConfig:
    """Configuration for the extraction process."""
    # Phase 1: Categorization
    categorize_model: str = "claude-haiku-4-5"
    use_batch_for_categorize: bool = True

    # Phase 2: Detailed extraction
    extract_model: str = "claude-sonnet-4-5"
    use_prompt_cache: bool = True

    # Deduplication
    similarity_threshold: float = 0.85
    max_templates_per_category: int = 5

    # Output
    output_dir: Path = field(default_factory=lambda: Path("descriptions"))
    save_all_categorizations: bool = True
    save_detailed_only: bool = False


@dataclass
class ExtractionStats:
    """Statistics from the extraction process."""
    total_slides: int = 0
    categorized: int = 0
    unique_patterns: int = 0
    detailed_extracted: int = 0
    duplicates_skipped: int = 0
    phase1_cost: float = 0.0
    phase1_tokens: int = 0
    phase2_cost: float = 0.0
    phase2_tokens: int = 0
    total_time_seconds: float = 0.0

    @property
    def total_cost(self) -> float:
        return self.phase1_cost + self.phase2_cost

    @property
    def savings_vs_full_extraction(self) -> float:
        # Estimate what full Sonnet extraction would cost
        full_cost = self.total_slides * 0.054  # ~$54 per 1000 slides with Sonnet
        return max(0, full_cost - self.total_cost)


class OptimizedExtractor:
    """
    Two-phase slide extractor optimized for cost efficiency.

    Workflow:
    1. Render all slides to images
    2. Phase 1: Categorize all slides with Haiku (cheap + fast)
    3. Group slides by type and identify unique patterns
    4. Phase 2: Extract detailed specs for unique patterns with Sonnet
    5. Build template library for generator
    """

    def __init__(self, config: Optional[ExtractionConfig] = None):
        """
        Initialize the optimized extractor.

        Args:
            config: Extraction configuration
        """
        self.config = config or ExtractionConfig()
        self.config.output_dir.mkdir(parents=True, exist_ok=True)
        self.stats = ExtractionStats()

    def process_template(
        self,
        template_path: Path,
        render_slides: bool = True
    ) -> Dict[str, Any]:
        """
        Process a complete template with optimized extraction.

        Args:
            template_path: Path to the PPTX template
            render_slides: Whether to render slides (skip if already rendered)

        Returns:
            Dictionary with extraction results and library
        """
        template_path = Path(template_path)
        template_name = template_path.stem

        console.print(f"\n[bold blue]Processing template:[/] {template_name}")

        start_time = datetime.now()

        # Step 1: Render slides if needed
        if render_slides:
            image_paths = self._render_slides(template_path)
        else:
            image_paths = self._find_rendered_slides(template_path)

        self.stats.total_slides = len(image_paths)
        console.print(f"  Found [cyan]{len(image_paths)}[/] slides")

        # Step 2: Phase 1 - Categorization
        console.print("\n[bold]Phase 1:[/] Fast categorization...")
        categorizations = self._phase1_categorize(image_paths)

        # Step 3: Group and deduplicate
        console.print("\n[bold]Grouping:[/] Identifying unique patterns...")
        unique_patterns = self._identify_unique_patterns(categorizations, image_paths)

        # Step 4: Phase 2 - Detailed extraction
        console.print("\n[bold]Phase 2:[/] Detailed extraction of unique templates...")
        detailed_extractions = self._phase2_extract(unique_patterns)

        # Step 5: Build template library
        console.print("\n[bold]Building:[/] Template library...")
        library = self._build_library(template_name, detailed_extractions)

        # Calculate stats
        end_time = datetime.now()
        self.stats.total_time_seconds = (end_time - start_time).total_seconds()

        # Print summary
        self._print_summary()

        return {
            "template_name": template_name,
            "total_slides": len(image_paths),
            "unique_patterns": len(unique_patterns),
            "library_path": str(library),
            "stats": self.stats
        }

    def _render_slides(self, template_path: Path) -> List[Path]:
        """Render template slides to images."""
        from .renderer import render_template

        console.print("  Rendering slides to PNG...")

        output_dir = self.config.output_dir / template_path.stem
        output_dir.mkdir(parents=True, exist_ok=True)

        try:
            image_paths = render_template(template_path, output_dir)
            console.print(f"  Rendered [green]{len(image_paths)}[/] slides")
            return image_paths
        except Exception as e:
            logger.error(f"Failed to render slides: {e}")
            raise

    def _find_rendered_slides(self, template_path: Path) -> List[Path]:
        """Find already-rendered slide images."""
        output_dir = self.config.output_dir / template_path.stem
        patterns = ["slide_*.png", "*.png"]

        for pattern in patterns:
            images = sorted(output_dir.glob(pattern))
            if images:
                return images

        # Try outputs directory
        output_dir = Path("outputs") / template_path.stem
        for pattern in patterns:
            images = sorted(output_dir.glob(pattern))
            if images:
                return images

        raise FileNotFoundError(f"No rendered slides found for {template_path.stem}")

    def _phase1_categorize(self, image_paths: List[Path]) -> List[Dict[str, Any]]:
        """
        Phase 1: Fast categorization of all slides.

        Uses Haiku model with batch processing for minimum cost.
        """
        from .batch_processor import SyncBatchProcessor, BatchMode

        processor = SyncBatchProcessor(
            model=self.config.categorize_model,
            max_parallel=5
        )

        results = []
        with Progress(
            SpinnerColumn(),
            TextColumn("[progress.description]{task.description}"),
            BarColumn(),
            TextColumn("[progress.percentage]{task.percentage:>3.0f}%"),
            console=console
        ) as progress:
            task = progress.add_task("Categorizing...", total=len(image_paths))

            def on_progress(completed, total):
                progress.update(task, completed=completed)

            results = processor.process_images(
                image_paths,
                mode=BatchMode.CATEGORIZE,
                progress_callback=on_progress
            )

        # Calculate costs (Haiku pricing)
        total_tokens = sum(r.tokens_used for r in results)
        self.stats.phase1_tokens = total_tokens
        self.stats.phase1_cost = (total_tokens / 1_000_000) * 3  # ~$3/MTok average
        self.stats.categorized = len([r for r in results if r.success])

        # Convert to list of dicts
        categorizations = []
        for i, result in enumerate(results):
            cat = {
                "slide_index": i,
                "image_path": str(image_paths[i]),
                "success": result.success,
                "categorization": result.result if result.success else None,
                "error": result.error
            }
            categorizations.append(cat)

        console.print(f"  Categorized [green]{self.stats.categorized}[/] slides")
        console.print(f"  Phase 1 cost: [yellow]${self.stats.phase1_cost:.2f}[/]")

        # Save categorizations
        if self.config.save_all_categorizations:
            cat_path = self.config.output_dir / "categorizations.json"
            with open(cat_path, 'w') as f:
                json.dump(categorizations, f, indent=2)

        return categorizations

    def _identify_unique_patterns(
        self,
        categorizations: List[Dict],
        image_paths: List[Path]
    ) -> List[Tuple[int, Path, Dict]]:
        """
        Identify unique slide patterns from categorizations.

        Groups similar slides and selects representative examples.
        """
        # Group by slide type
        by_type = defaultdict(list)
        for cat in categorizations:
            if cat["success"] and cat["categorization"]:
                slide_type = cat["categorization"].get("slide_type", "unknown")
                by_type[slide_type].append(cat)

        # Find unique patterns within each type
        unique_patterns = []
        pattern_hashes = set()

        for slide_type, slides in by_type.items():
            type_patterns = []

            for slide in slides:
                cat = slide["categorization"]
                idx = slide["slide_index"]

                # Create pattern hash
                pattern_features = [
                    cat.get("layout_category", ""),
                    str(cat.get("element_count", 0)),
                    str(cat.get("has_chart", False)),
                    str(cat.get("has_table", False)),
                    cat.get("complexity_score", 3)
                ]
                pattern_hash = hashlib.md5(
                    "|".join(str(f) for f in pattern_features).encode()
                ).hexdigest()[:8]

                if pattern_hash not in pattern_hashes:
                    pattern_hashes.add(pattern_hash)
                    type_patterns.append((idx, image_paths[idx], cat))

            # Limit patterns per category
            selected = type_patterns[:self.config.max_templates_per_category]
            unique_patterns.extend(selected)

            console.print(f"    {slide_type}: {len(slides)} slides → {len(selected)} unique patterns")

        self.stats.unique_patterns = len(unique_patterns)
        self.stats.duplicates_skipped = self.stats.categorized - len(unique_patterns)

        console.print(f"  Found [cyan]{len(unique_patterns)}[/] unique patterns")
        console.print(f"  Skipping [dim]{self.stats.duplicates_skipped}[/] duplicate patterns")

        return unique_patterns

    def _phase2_extract(
        self,
        unique_patterns: List[Tuple[int, Path, Dict]]
    ) -> List[Dict[str, Any]]:
        """
        Phase 2: Detailed extraction of unique patterns.

        Uses Sonnet model with prompt caching for quality + efficiency.
        """
        from .cached_vision import CachedVisionClient

        client = CachedVisionClient(model=self.config.extract_model)

        extractions = []
        with Progress(
            SpinnerColumn(),
            TextColumn("[progress.description]{task.description}"),
            BarColumn(),
            TextColumn("[progress.percentage]{task.percentage:>3.0f}%"),
            console=console
        ) as progress:
            task = progress.add_task("Extracting...", total=len(unique_patterns))

            for idx, image_path, categorization in unique_patterns:
                try:
                    # Add categorization as context
                    context = f"This is a {categorization.get('slide_type', 'content')} slide with {categorization.get('layout_category', 'unknown')} layout."

                    result = client.analyze_image(
                        image_path,
                        mode="detailed",
                        additional_context=context
                    )

                    extraction = {
                        "slide_index": idx,
                        "image_path": str(image_path),
                        "categorization": categorization,
                        "extraction": result,
                        "success": "_error" not in result
                    }
                    extractions.append(extraction)

                    # Save individual extraction
                    if not self.config.save_detailed_only:
                        save_path = self.config.output_dir / f"slide_{idx:03d}_detailed.json"
                        with open(save_path, 'w') as f:
                            json.dump(result, f, indent=2)

                except Exception as e:
                    logger.error(f"Failed to extract slide {idx}: {e}")
                    extractions.append({
                        "slide_index": idx,
                        "image_path": str(image_path),
                        "error": str(e),
                        "success": False
                    })

                progress.advance(task)

        # Get cache stats
        cache_stats = client.get_stats()
        self.stats.phase2_tokens = cache_stats["total_input_tokens"] + cache_stats["total_output_tokens"]
        self.stats.phase2_cost = (
            (cache_stats["total_input_tokens"] / 1_000_000) * 3 +  # Sonnet input
            (cache_stats["total_output_tokens"] / 1_000_000) * 15 -  # Sonnet output
            cache_stats["estimated_savings_usd"]  # Cache savings
        )
        self.stats.detailed_extracted = len([e for e in extractions if e.get("success")])

        console.print(f"  Extracted [green]{self.stats.detailed_extracted}[/] detailed specs")
        console.print(f"  Cache hit rate: [cyan]{cache_stats['cache_hit_rate']:.1%}[/]")
        console.print(f"  Phase 2 cost: [yellow]${self.stats.phase2_cost:.2f}[/]")

        return extractions

    def _build_library(
        self,
        template_name: str,
        extractions: List[Dict[str, Any]]
    ) -> Path:
        """Build the template library from extractions."""
        from .template_index import TemplateIndexBuilder

        builder = TemplateIndexBuilder(output_dir=self.config.output_dir)

        for extraction in extractions:
            if extraction.get("success") and "extraction" in extraction:
                builder.add_extraction(
                    source_file=f"{template_name}.pptx",
                    slide_index=extraction["slide_index"],
                    extraction=extraction["extraction"],
                    description_path=str(
                        self.config.output_dir / f"slide_{extraction['slide_index']:03d}_detailed.json"
                    )
                )

        library_path = builder.save(f"{template_name}_library.json")
        console.print(f"  Saved library to [green]{library_path}[/]")

        return library_path

    def _print_summary(self):
        """Print extraction summary."""
        console.print("\n" + "=" * 60)
        console.print("[bold]Extraction Summary[/]")
        console.print("=" * 60)

        table = Table(show_header=False, box=None)
        table.add_column("Metric", style="cyan")
        table.add_column("Value", style="white")

        table.add_row("Total slides", str(self.stats.total_slides))
        table.add_row("Unique patterns", str(self.stats.unique_patterns))
        table.add_row("Duplicates skipped", str(self.stats.duplicates_skipped))
        table.add_row("Detailed extractions", str(self.stats.detailed_extracted))
        table.add_row("", "")
        table.add_row("Phase 1 cost (Haiku)", f"${self.stats.phase1_cost:.2f}")
        table.add_row("Phase 2 cost (Sonnet)", f"${self.stats.phase2_cost:.2f}")
        table.add_row("[bold]Total cost[/]", f"[bold]${self.stats.total_cost:.2f}[/]")
        table.add_row("", "")
        table.add_row("Estimated savings", f"[green]${self.stats.savings_vs_full_extraction:.2f}[/]")
        table.add_row("Processing time", f"{self.stats.total_time_seconds:.1f}s")

        console.print(table)
        console.print("=" * 60)


def quick_estimate(num_slides: int) -> Dict[str, Any]:
    """
    Quick cost estimate for processing a template.

    Args:
        num_slides: Number of slides in the template

    Returns:
        Cost estimate dictionary
    """
    from .cached_vision import estimate_batch_cost

    # Estimate unique patterns (typically 10-20% of slides)
    unique_ratio = 0.15
    estimated_unique = max(5, int(num_slides * unique_ratio))

    # Phase 1: Categorize all with Haiku
    phase1 = estimate_batch_cost(
        num_slides=num_slides,
        mode="categorize",
        model="claude-haiku-4.5",
        use_batch_api=True,
        use_prompt_cache=True
    )

    # Phase 2: Extract unique with Sonnet
    phase2 = estimate_batch_cost(
        num_slides=estimated_unique,
        mode="detailed",
        model="claude-sonnet-4.5",
        use_batch_api=False,  # Sync for caching
        use_prompt_cache=True
    )

    # Full extraction comparison
    full_extraction = estimate_batch_cost(
        num_slides=num_slides,
        mode="detailed",
        model="claude-sonnet-4.5",
        use_batch_api=False,
        use_prompt_cache=False
    )

    return {
        "num_slides": num_slides,
        "estimated_unique_patterns": estimated_unique,
        "phase1_cost": phase1["final_cost_usd"],
        "phase2_cost": phase2["final_cost_usd"],
        "optimized_total": phase1["final_cost_usd"] + phase2["final_cost_usd"],
        "full_extraction_cost": full_extraction["final_cost_usd"],
        "savings_usd": full_extraction["final_cost_usd"] - (phase1["final_cost_usd"] + phase2["final_cost_usd"]),
        "savings_percentage": round(
            (1 - (phase1["final_cost_usd"] + phase2["final_cost_usd"]) / full_extraction["final_cost_usd"]) * 100, 1
        )
    }
