"""
PPTX to PNG Renderer

Converts PowerPoint slides to PNG images using LibreOffice (headless mode)
for PPTX to PDF conversion, and Poppler (via pdf2image) for PDF to PNG conversion.
"""
import subprocess
import tempfile
import shutil
import logging
from pathlib import Path
from typing import Optional

# Add parent directory to path for config import
import sys
sys.path.insert(0, str(Path(__file__).parent.parent))

from config import (
    LIBREOFFICE_PATH,
    POPPLER_PATH,
    RENDER_DPI,
    OUTPUT_DIR,
    check_libreoffice,
    check_poppler
)

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class RenderError(Exception):
    """Exception raised when rendering fails."""
    pass


def verify_dependencies() -> tuple[bool, list[str]]:
    """
    Verify that all required dependencies are available.

    Returns:
        Tuple of (all_ok, list_of_missing_dependencies)
    """
    missing = []

    if not check_libreoffice():
        missing.append(f"LibreOffice not found at: {LIBREOFFICE_PATH}")

    if not check_poppler():
        missing.append("Poppler not found (required for PDF to PNG conversion)")

    return (len(missing) == 0, missing)


def pptx_to_pdf(pptx_path: Path, output_dir: Path) -> Path:
    """
    Convert a PPTX file to PDF using LibreOffice headless mode.

    Args:
        pptx_path: Path to the PPTX file
        output_dir: Directory to save the PDF

    Returns:
        Path to the generated PDF file

    Raises:
        RenderError: If conversion fails
    """
    if not pptx_path.exists():
        raise RenderError(f"PPTX file not found: {pptx_path}")

    if not check_libreoffice():
        raise RenderError(f"LibreOffice not found at: {LIBREOFFICE_PATH}")

    output_dir.mkdir(parents=True, exist_ok=True)

    # LibreOffice command for headless PDF conversion
    cmd = [
        LIBREOFFICE_PATH,
        "--headless",
        "--convert-to", "pdf",
        "--outdir", str(output_dir),
        str(pptx_path)
    ]

    logger.info(f"Converting PPTX to PDF: {pptx_path.name}")
    logger.debug(f"Command: {' '.join(cmd)}")

    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=120  # 2 minute timeout
        )

        if result.returncode != 0:
            logger.error(f"LibreOffice stderr: {result.stderr}")
            raise RenderError(f"LibreOffice conversion failed: {result.stderr}")

    except subprocess.TimeoutExpired:
        raise RenderError("LibreOffice conversion timed out after 120 seconds")
    except FileNotFoundError:
        raise RenderError(f"LibreOffice executable not found: {LIBREOFFICE_PATH}")

    # LibreOffice outputs the PDF with the same name as input
    pdf_path = output_dir / f"{pptx_path.stem}.pdf"

    if not pdf_path.exists():
        raise RenderError(f"PDF was not created at expected path: {pdf_path}")

    logger.info(f"PDF created: {pdf_path}")
    return pdf_path


def pdf_to_images(pdf_path: Path, output_dir: Path, dpi: int = None) -> list[Path]:
    """
    Convert PDF pages to PNG images using Poppler via pdf2image.

    Args:
        pdf_path: Path to the PDF file
        output_dir: Directory to save PNG images
        dpi: DPI for rendering (defaults to RENDER_DPI from config)

    Returns:
        List of paths to generated PNG files

    Raises:
        RenderError: If conversion fails
    """
    if dpi is None:
        dpi = RENDER_DPI

    if not pdf_path.exists():
        raise RenderError(f"PDF file not found: {pdf_path}")

    try:
        from pdf2image import convert_from_path
    except ImportError:
        raise RenderError("pdf2image not installed. Run: pip install pdf2image")

    output_dir.mkdir(parents=True, exist_ok=True)

    logger.info(f"Converting PDF to images: {pdf_path.name} at {dpi} DPI")

    try:
        images = convert_from_path(
            pdf_path,
            dpi=dpi,
            poppler_path=POPPLER_PATH,
            fmt='png'
        )
    except Exception as e:
        raise RenderError(f"pdf2image conversion failed: {e}")

    output_paths = []
    base_name = pdf_path.stem

    for i, image in enumerate(images):
        output_path = output_dir / f"{base_name}_slide_{i + 1:03d}.png"
        image.save(output_path, 'PNG')
        output_paths.append(output_path)
        logger.debug(f"Saved slide {i + 1}: {output_path}")

    logger.info(f"Generated {len(output_paths)} PNG images")
    return output_paths


def render_slide(
    pptx_path: Path,
    slide_index: int,
    output_path: Optional[Path] = None,
    dpi: int = None
) -> Path:
    """
    Render a single slide from a PPTX file to PNG.

    Args:
        pptx_path: Path to the PPTX file
        slide_index: Zero-based index of the slide to render
        output_path: Optional specific path for the output PNG
        dpi: DPI for rendering (defaults to RENDER_DPI from config)

    Returns:
        Path to the generated PNG file

    Raises:
        RenderError: If rendering fails
        IndexError: If slide_index is out of range
    """
    pptx_path = Path(pptx_path)

    # Create a temporary directory for intermediate files
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_path = Path(temp_dir)

        # Convert PPTX to PDF
        pdf_path = pptx_to_pdf(pptx_path, temp_path)

        # Convert PDF to images
        images = pdf_to_images(pdf_path, temp_path, dpi)

        if slide_index >= len(images):
            raise IndexError(
                f"Slide index {slide_index} out of range. "
                f"PPTX has {len(images)} slides (0-{len(images) - 1})"
            )

        # Determine output path
        if output_path is None:
            output_path = OUTPUT_DIR / f"{pptx_path.stem}_slide_{slide_index + 1:03d}.png"
        else:
            output_path = Path(output_path)

        output_path.parent.mkdir(parents=True, exist_ok=True)

        # Copy the specific slide image to the output path
        shutil.copy(images[slide_index], output_path)

        logger.info(f"Rendered slide {slide_index + 1} to: {output_path}")
        return output_path


def render_all_slides(
    pptx_path: Path,
    output_dir: Optional[Path] = None,
    dpi: int = None
) -> list[Path]:
    """
    Render all slides from a PPTX file to PNG images.

    Args:
        pptx_path: Path to the PPTX file
        output_dir: Directory to save PNG images (defaults to OUTPUT_DIR)
        dpi: DPI for rendering (defaults to RENDER_DPI from config)

    Returns:
        List of paths to generated PNG files

    Raises:
        RenderError: If rendering fails
    """
    pptx_path = Path(pptx_path)

    if output_dir is None:
        output_dir = OUTPUT_DIR / pptx_path.stem
    else:
        output_dir = Path(output_dir)

    output_dir.mkdir(parents=True, exist_ok=True)

    # Create a temporary directory for intermediate files
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_path = Path(temp_dir)

        # Convert PPTX to PDF
        pdf_path = pptx_to_pdf(pptx_path, temp_path)

        # Convert PDF to images in temp directory
        temp_images = pdf_to_images(pdf_path, temp_path, dpi)

        # Copy images to final output directory with proper naming
        output_paths = []
        for i, temp_image in enumerate(temp_images):
            output_path = output_dir / f"slide_{i + 1:03d}.png"
            shutil.copy(temp_image, output_path)
            output_paths.append(output_path)

        logger.info(f"Rendered {len(output_paths)} slides to: {output_dir}")
        return output_paths


def get_slide_count(pptx_path: Path) -> int:
    """
    Get the number of slides in a PPTX file.

    Args:
        pptx_path: Path to the PPTX file

    Returns:
        Number of slides in the presentation
    """
    try:
        from pptx import Presentation
    except ImportError:
        raise RenderError("python-pptx not installed. Run: pip install python-pptx")

    prs = Presentation(pptx_path)
    return len(prs.slides)


if __name__ == "__main__":
    # Test the renderer
    import sys

    print("PPTX Renderer - Dependency Check")
    print("=" * 50)

    ok, missing = verify_dependencies()

    if ok:
        print("All dependencies are available!")
    else:
        print("Missing dependencies:")
        for m in missing:
            print(f"  - {m}")
        sys.exit(1)

    # If a PPTX path is provided, try to render it
    if len(sys.argv) > 1:
        pptx_path = Path(sys.argv[1])
        print(f"\nRendering: {pptx_path}")

        try:
            slide_count = get_slide_count(pptx_path)
            print(f"Slide count: {slide_count}")

            output_paths = render_all_slides(pptx_path)
            print(f"\nGenerated {len(output_paths)} images:")
            for p in output_paths:
                print(f"  - {p}")

        except Exception as e:
            print(f"Error: {e}")
            sys.exit(1)
