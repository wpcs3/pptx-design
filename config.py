"""
Central configuration for the PPTX design system.

This module provides all configuration settings including paths,
iteration limits, and tool paths for LibreOffice and Poppler.
"""
import os
from pathlib import Path

# Base paths
PROJECT_ROOT = Path(r"C:\Users\wpcol\claudecode\pptx-design")
TEMPLATE_DIR = PROJECT_ROOT / "pptx_templates"
OUTPUT_DIR = PROJECT_ROOT / "outputs"
DESCRIPTION_DIR = PROJECT_ROOT / "descriptions"
DIFF_DIR = PROJECT_ROOT / "diffs"
SRC_DIR = PROJECT_ROOT / "pptx_extractor"

# Ensure directories exist
for dir_path in [OUTPUT_DIR, DESCRIPTION_DIR, DIFF_DIR]:
    dir_path.mkdir(parents=True, exist_ok=True)

# LibreOffice path (adjust if needed for Windows)
# Common locations:
LIBREOFFICE_PATH = r"C:\Program Files\LibreOffice\program\soffice.exe"
# Alternative: r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"


def get_poppler_path() -> str | None:
    """
    Detect Poppler installation path.

    Priority:
    1. Active conda environment (CONDA_PREFIX)
    2. Explicit fallback path for the pptx-design environment
    3. None (relies on system PATH)

    Returns:
        Path to Poppler bin directory, or None if not found.
    """
    # Check if we're in an activated conda environment
    conda_prefix = os.environ.get('CONDA_PREFIX')
    if conda_prefix:
        poppler_bin = Path(conda_prefix) / 'Library' / 'bin'
        if (poppler_bin / 'pdftoppm.exe').exists():
            return str(poppler_bin)

    # Fallback: check the expected pptx-design environment location
    fallback_paths = [
        Path(os.path.expanduser('~')) / 'miniforge3' / 'envs' / 'pptx-design' / 'Library' / 'bin',
        Path(os.path.expanduser('~')) / 'miniconda3' / 'envs' / 'pptx-design' / 'Library' / 'bin',
        Path(os.path.expanduser('~')) / 'anaconda3' / 'envs' / 'pptx-design' / 'Library' / 'bin',
    ]

    for fallback in fallback_paths:
        if fallback.exists() and (fallback / 'pdftoppm.exe').exists():
            return str(fallback)

    # Last resort: assume it's in system PATH
    return None


def check_libreoffice() -> bool:
    """Check if LibreOffice is installed at the configured path."""
    return Path(LIBREOFFICE_PATH).exists()


def check_poppler() -> bool:
    """Check if Poppler is available."""
    poppler_path = get_poppler_path()
    if poppler_path:
        return (Path(poppler_path) / 'pdftoppm.exe').exists()
    # Check system PATH
    import shutil
    return shutil.which('pdftoppm') is not None


POPPLER_PATH = get_poppler_path()

# Iteration settings
MAX_ITERATIONS = 10
SIMILARITY_THRESHOLD = 0.95  # 0-1 scale for programmatic comparison

# Rendering settings
RENDER_DPI = 150  # Higher = more detail but slower

# Slide dimensions (standard widescreen 16:9)
DEFAULT_SLIDE_WIDTH_INCHES = 13.333
DEFAULT_SLIDE_HEIGHT_INCHES = 7.5


def print_config_status():
    """Print configuration status for debugging."""
    print("=" * 60)
    print("PPTX Design System Configuration")
    print("=" * 60)
    print(f"Project Root: {PROJECT_ROOT}")
    print(f"Template Dir: {TEMPLATE_DIR} (exists: {TEMPLATE_DIR.exists()})")
    print(f"Output Dir:   {OUTPUT_DIR} (exists: {OUTPUT_DIR.exists()})")
    print(f"Description Dir: {DESCRIPTION_DIR} (exists: {DESCRIPTION_DIR.exists()})")
    print(f"Diff Dir:     {DIFF_DIR} (exists: {DIFF_DIR.exists()})")
    print("-" * 60)
    print(f"LibreOffice Path: {LIBREOFFICE_PATH}")
    print(f"LibreOffice Installed: {check_libreoffice()}")
    print("-" * 60)
    print(f"Poppler Path: {POPPLER_PATH}")
    print(f"Poppler Available: {check_poppler()}")
    print("-" * 60)
    print(f"Max Iterations: {MAX_ITERATIONS}")
    print(f"Similarity Threshold: {SIMILARITY_THRESHOLD}")
    print(f"Render DPI: {RENDER_DPI}")
    print("=" * 60)


if __name__ == "__main__":
    print_config_status()
