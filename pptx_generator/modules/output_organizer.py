"""
Output Organizer Module

Manages output file organization into topic-specific subfolders.
Provides automatic topic detection and file organization utilities.
"""

import logging
import re
import shutil
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

logger = logging.getLogger(__name__)


@dataclass
class TopicConfig:
    """Configuration for a presentation topic."""
    name: str  # Display name
    folder_name: str  # Subfolder name (lowercase, underscores)
    patterns: List[str] = field(default_factory=list)  # Filename patterns to match
    keywords: List[str] = field(default_factory=list)  # Title keywords to match


# Default topic configurations
DEFAULT_TOPICS: Dict[str, TopicConfig] = {
    "light_industrial": TopicConfig(
        name="Light Industrial",
        folder_name="light_industrial",
        patterns=[
            r"light[_\s]?industrial",
            r"industrial[_\s]?thesis",
            r"industrial[_\s]?logistics",
        ],
        keywords=[
            "light industrial",
            "industrial thesis",
            "industrial logistics",
            "warehouse",
            "logistics fund",
        ]
    ),
    "btr": TopicConfig(
        name="Build-to-Rent",
        folder_name="btr",
        patterns=[
            r"btr[_\s]",
            r"build[_\s]?to[_\s]?rent",
            r"single[_\s]?family[_\s]?rental",
            r"sfr[_\s]",
        ],
        keywords=[
            "build-to-rent",
            "build to rent",
            "btr",
            "single family rental",
            "sfr",
            "rental housing",
        ]
    ),
    "multifamily": TopicConfig(
        name="Multifamily",
        folder_name="multifamily",
        patterns=[
            r"multifamily",
            r"multi[_\s]?family",
            r"apartment",
        ],
        keywords=[
            "multifamily",
            "multi-family",
            "apartment",
            "residential rental",
        ]
    ),
    "office": TopicConfig(
        name="Office",
        folder_name="office",
        patterns=[
            r"office[_\s]",
            r"commercial[_\s]?office",
        ],
        keywords=[
            "office",
            "commercial office",
            "workspace",
        ]
    ),
    "retail": TopicConfig(
        name="Retail",
        folder_name="retail",
        patterns=[
            r"retail[_\s]",
            r"shopping[_\s]?center",
        ],
        keywords=[
            "retail",
            "shopping center",
            "retail center",
        ]
    ),
    "mixed_use": TopicConfig(
        name="Mixed Use",
        folder_name="mixed_use",
        patterns=[
            r"mixed[_\s]?use",
        ],
        keywords=[
            "mixed use",
            "mixed-use",
        ]
    ),
    "hospitality": TopicConfig(
        name="Hospitality",
        folder_name="hospitality",
        patterns=[
            r"hospitality",
            r"hotel",
        ],
        keywords=[
            "hospitality",
            "hotel",
            "resort",
        ]
    ),
    "fund_overview": TopicConfig(
        name="Fund Overview",
        folder_name="fund_overview",
        patterns=[
            r"fund[_\s]?overview",
            r"investor[_\s]?deck",
            r"pitch[_\s]?deck",
        ],
        keywords=[
            "fund overview",
            "investor deck",
            "pitch deck",
            "fund presentation",
        ]
    ),
}


class OutputOrganizer:
    """Manages output file organization into topic subfolders."""

    def __init__(
        self,
        output_dir: Path,
        topics: Optional[Dict[str, TopicConfig]] = None,
        create_folders: bool = True
    ):
        """
        Initialize the output organizer.

        Args:
            output_dir: Base output directory
            topics: Custom topic configurations (merges with defaults)
            create_folders: Whether to create folders automatically
        """
        self.output_dir = Path(output_dir)
        self.topics = {**DEFAULT_TOPICS}
        if topics:
            self.topics.update(topics)
        self.create_folders = create_folders

        # Ensure base output directory exists
        self.output_dir.mkdir(parents=True, exist_ok=True)

    def detect_topic(
        self,
        filename: Optional[str] = None,
        title: Optional[str] = None
    ) -> Optional[str]:
        """
        Detect the topic from filename or title.

        Args:
            filename: Filename to check
            title: Presentation title to check

        Returns:
            Topic key if detected, None otherwise
        """
        # Check filename patterns first (more specific)
        if filename:
            filename_lower = filename.lower()
            for topic_key, config in self.topics.items():
                for pattern in config.patterns:
                    if re.search(pattern, filename_lower, re.IGNORECASE):
                        logger.debug(f"Detected topic '{topic_key}' from filename pattern '{pattern}'")
                        return topic_key

        # Check title keywords
        if title:
            title_lower = title.lower()
            for topic_key, config in self.topics.items():
                for keyword in config.keywords:
                    if keyword.lower() in title_lower:
                        logger.debug(f"Detected topic '{topic_key}' from title keyword '{keyword}'")
                        return topic_key

        return None

    def get_topic_folder(self, topic_key: str) -> Path:
        """
        Get the folder path for a topic.

        Args:
            topic_key: Topic identifier

        Returns:
            Path to the topic folder
        """
        if topic_key not in self.topics:
            # Use the key as folder name if not in config
            folder_name = topic_key.lower().replace(" ", "_").replace("-", "_")
        else:
            folder_name = self.topics[topic_key].folder_name

        folder_path = self.output_dir / folder_name

        if self.create_folders:
            folder_path.mkdir(parents=True, exist_ok=True)

        return folder_path

    def get_output_path(
        self,
        filename: str,
        topic: Optional[str] = None,
        title: Optional[str] = None,
        auto_detect: bool = True
    ) -> Path:
        """
        Get the full output path for a file.

        Args:
            filename: The filename to save
            topic: Explicit topic (overrides auto-detection)
            title: Presentation title (for auto-detection)
            auto_detect: Whether to auto-detect topic from filename/title

        Returns:
            Full path including subfolder
        """
        # Determine topic
        if topic:
            detected_topic = topic
        elif auto_detect:
            detected_topic = self.detect_topic(filename, title)
        else:
            detected_topic = None

        # Get the appropriate folder
        if detected_topic:
            folder = self.get_topic_folder(detected_topic)
            logger.info(f"Using topic folder: {folder.name}")
        else:
            folder = self.output_dir
            logger.debug(f"Using base output folder (no topic detected)")

        return folder / filename

    def organize_existing_files(
        self,
        dry_run: bool = False,
        include_patterns: Optional[List[str]] = None
    ) -> Dict[str, List[Tuple[Path, Path]]]:
        """
        Organize existing files in the output directory into topic subfolders.

        Args:
            dry_run: If True, don't actually move files
            include_patterns: File patterns to include (e.g., ["*.pptx", "*.pdf"])

        Returns:
            Dictionary mapping topics to list of (source, destination) tuples
        """
        if include_patterns is None:
            include_patterns = ["*.pptx", "*.pdf"]

        moves: Dict[str, List[Tuple[Path, Path]]] = {
            topic_key: [] for topic_key in self.topics
        }
        moves["_unorganized"] = []

        # Find all matching files in base directory (not subfolders)
        for pattern in include_patterns:
            for file_path in self.output_dir.glob(pattern):
                # Skip files already in subfolders
                if file_path.parent != self.output_dir:
                    continue

                # Skip temporary files
                if file_path.name.startswith("~$") or file_path.name.startswith("_temp"):
                    continue

                # Detect topic
                topic = self.detect_topic(file_path.name)

                if topic:
                    dest_folder = self.get_topic_folder(topic)
                    dest_path = dest_folder / file_path.name
                    moves[topic].append((file_path, dest_path))

                    if not dry_run:
                        try:
                            # Move the file
                            shutil.move(str(file_path), str(dest_path))
                            logger.info(f"Moved: {file_path.name} -> {dest_folder.name}/")
                        except PermissionError:
                            logger.warning(f"Skipped (file locked): {file_path.name}")
                            moves["_skipped"] = moves.get("_skipped", [])
                            moves["_skipped"].append((file_path, dest_path))
                else:
                    moves["_unorganized"].append((file_path, file_path))

        return moves

    def list_topics(self) -> List[Dict[str, str]]:
        """List all configured topics."""
        return [
            {
                "key": key,
                "name": config.name,
                "folder": config.folder_name,
            }
            for key, config in self.topics.items()
        ]

    def get_topic_stats(self) -> Dict[str, Dict[str, int]]:
        """
        Get file statistics for each topic folder.

        Returns:
            Dictionary with file counts per topic
        """
        stats = {}

        for topic_key, config in self.topics.items():
            folder = self.output_dir / config.folder_name
            if folder.exists():
                pptx_count = len(list(folder.glob("*.pptx")))
                pdf_count = len(list(folder.glob("*.pdf")))
                total = len(list(folder.iterdir()))
                stats[topic_key] = {
                    "pptx": pptx_count,
                    "pdf": pdf_count,
                    "total": total,
                }
            else:
                stats[topic_key] = {"pptx": 0, "pdf": 0, "total": 0}

        # Count files in base directory
        base_pptx = len([f for f in self.output_dir.glob("*.pptx") if f.parent == self.output_dir])
        base_pdf = len([f for f in self.output_dir.glob("*.pdf") if f.parent == self.output_dir])
        stats["_base_directory"] = {
            "pptx": base_pptx,
            "pdf": base_pdf,
            "total": base_pptx + base_pdf,
        }

        return stats

    def add_topic(
        self,
        key: str,
        name: str,
        folder_name: Optional[str] = None,
        patterns: Optional[List[str]] = None,
        keywords: Optional[List[str]] = None
    ) -> None:
        """
        Add a new topic configuration.

        Args:
            key: Topic identifier
            name: Display name
            folder_name: Subfolder name (defaults to key)
            patterns: Filename regex patterns
            keywords: Title keywords
        """
        self.topics[key] = TopicConfig(
            name=name,
            folder_name=folder_name or key.lower().replace(" ", "_"),
            patterns=patterns or [],
            keywords=keywords or [],
        )


def organize_output_directory(
    output_dir: str = "C:/Users/wpcol/claudecode/pptx-design/pptx_generator/output",
    dry_run: bool = False
) -> Dict[str, List[Tuple[Path, Path]]]:
    """
    Convenience function to organize the output directory.

    Args:
        output_dir: Path to output directory
        dry_run: If True, don't actually move files

    Returns:
        Dictionary of moves performed
    """
    organizer = OutputOrganizer(Path(output_dir))
    return organizer.organize_existing_files(dry_run=dry_run)


def main():
    """CLI for organizing output files."""
    import argparse

    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

    parser = argparse.ArgumentParser(description="Organize presentation output files")
    parser.add_argument(
        "--output-dir",
        default="C:/Users/wpcol/claudecode/pptx-design/pptx_generator/output",
        help="Output directory to organize"
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Show what would be moved without actually moving"
    )
    parser.add_argument(
        "--stats",
        action="store_true",
        help="Show statistics only"
    )
    parser.add_argument(
        "--list-topics",
        action="store_true",
        help="List configured topics"
    )

    args = parser.parse_args()

    organizer = OutputOrganizer(Path(args.output_dir))

    if args.list_topics:
        print("\nConfigured Topics:")
        print("-" * 50)
        for topic in organizer.list_topics():
            print(f"  {topic['key']:20} -> {topic['folder']}/")
        return

    if args.stats:
        print("\nOutput Directory Statistics:")
        print("-" * 50)
        stats = organizer.get_topic_stats()
        for topic, counts in stats.items():
            if counts["total"] > 0:
                print(f"  {topic:20}: {counts['pptx']} pptx, {counts['pdf']} pdf")
        return

    # Organize files
    action = "Would move" if args.dry_run else "Moved"
    print(f"\n{'DRY RUN: ' if args.dry_run else ''}Organizing output files...")
    print("-" * 50)

    moves = organizer.organize_existing_files(dry_run=args.dry_run)

    total_moved = 0
    for topic, file_moves in moves.items():
        if topic == "_unorganized":
            continue
        if file_moves:
            print(f"\n{topic}:")
            for src, dest in file_moves:
                print(f"  {action}: {src.name}")
                total_moved += 1

    unorganized = moves.get("_unorganized", [])
    if unorganized:
        print(f"\nUnorganized ({len(unorganized)} files):")
        for src, _ in unorganized[:5]:
            print(f"  {src.name}")
        if len(unorganized) > 5:
            print(f"  ... and {len(unorganized) - 5} more")

    print(f"\n{'Would move' if args.dry_run else 'Moved'}: {total_moved} files")


if __name__ == "__main__":
    main()
