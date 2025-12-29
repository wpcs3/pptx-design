"""
Visual Testing Module

Automated visual comparison for presentation generation:
1. Generate slides from descriptions
2. Render to PNG
3. Compare with reference images using SSIM
4. Report differences

Usage:
    from pptx_design.testing import VisualTester

    tester = VisualTester()
    result = tester.test_slide("my_slide.json", "reference.png")
    print(f"SSIM: {result.ssim_score}")
"""

import json
import logging
import subprocess
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional, Tuple

logger = logging.getLogger(__name__)


@dataclass
class TestResult:
    """Result of a visual test."""
    name: str
    passed: bool
    ssim_score: float
    threshold: float
    generated_path: Optional[Path] = None
    reference_path: Optional[Path] = None
    diff_path: Optional[Path] = None
    error: Optional[str] = None

    def __str__(self) -> str:
        status = "PASS" if self.passed else "FAIL"
        return f"[{status}] {self.name}: SSIM={self.ssim_score:.4f} (threshold={self.threshold})"


@dataclass
class TestSuite:
    """Collection of test results."""
    name: str
    results: List[TestResult]

    @property
    def passed(self) -> int:
        return sum(1 for r in self.results if r.passed)

    @property
    def failed(self) -> int:
        return sum(1 for r in self.results if not r.passed)

    @property
    def total(self) -> int:
        return len(self.results)

    @property
    def pass_rate(self) -> float:
        return self.passed / self.total if self.total > 0 else 0.0

    def summary(self) -> str:
        lines = [
            f"Test Suite: {self.name}",
            f"  Passed: {self.passed}/{self.total} ({self.pass_rate:.1%})",
            f"  Failed: {self.failed}",
            "",
        ]
        for result in self.results:
            lines.append(f"  {result}")
        return "\n".join(lines)


class VisualTester:
    """
    Visual testing for presentation generation.

    Compares generated slides against reference images using SSIM.
    """

    # Default SSIM threshold for passing
    DEFAULT_THRESHOLD = 0.90

    # LibreOffice path for PPTX to PDF conversion
    LIBREOFFICE_PATH = r"C:\Program Files\LibreOffice\program\soffice.exe"

    def __init__(
        self,
        output_dir: Optional[Path] = None,
        reference_dir: Optional[Path] = None,
        threshold: float = DEFAULT_THRESHOLD
    ):
        """
        Initialize tester.

        Args:
            output_dir: Directory for test outputs
            reference_dir: Directory containing reference images
            threshold: SSIM threshold for pass/fail (0.0-1.0)
        """
        self.output_dir = Path(output_dir) if output_dir else Path("test_outputs")
        self.reference_dir = Path(reference_dir) if reference_dir else Path("test_references")
        self.threshold = threshold

        self.output_dir.mkdir(parents=True, exist_ok=True)

    def compare_images(
        self,
        image1_path: Path,
        image2_path: Path
    ) -> Tuple[float, Optional[Path]]:
        """
        Compare two images using SSIM.

        Args:
            image1_path: First image
            image2_path: Second image

        Returns:
            Tuple of (ssim_score, diff_image_path)
        """
        try:
            from skimage.metrics import structural_similarity as ssim
            from PIL import Image
            import numpy as np

            # Load images
            img1 = Image.open(image1_path).convert("RGB")
            img2 = Image.open(image2_path).convert("RGB")

            # Resize to match if needed
            if img1.size != img2.size:
                img2 = img2.resize(img1.size, Image.Resampling.LANCZOS)

            # Convert to numpy arrays
            arr1 = np.array(img1)
            arr2 = np.array(img2)

            # Calculate SSIM
            score, diff = ssim(arr1, arr2, full=True, channel_axis=2)

            # Create diff image
            diff_normalized = (diff * 255).astype(np.uint8)
            diff_img = Image.fromarray(diff_normalized)
            diff_path = self.output_dir / f"diff_{image1_path.stem}.png"
            diff_img.save(diff_path)

            return float(score), diff_path

        except ImportError as e:
            logger.warning(f"Could not compare images: {e}")
            return 0.0, None
        except Exception as e:
            logger.error(f"Image comparison failed: {e}")
            return 0.0, None

    def pptx_to_png(
        self,
        pptx_path: Path,
        output_dir: Optional[Path] = None,
        slide_index: int = 0
    ) -> Optional[Path]:
        """
        Convert PPTX to PNG via PDF.

        Args:
            pptx_path: Path to PPTX file
            output_dir: Output directory for PNG
            slide_index: Which slide to render (0-indexed)

        Returns:
            Path to PNG file or None on failure
        """
        output_dir = output_dir or self.output_dir

        try:
            # Step 1: PPTX to PDF using LibreOffice
            pdf_path = output_dir / f"{pptx_path.stem}.pdf"
            subprocess.run([
                self.LIBREOFFICE_PATH,
                "--headless",
                "--convert-to", "pdf",
                "--outdir", str(output_dir),
                str(pptx_path)
            ], check=True, capture_output=True, timeout=60)

            if not pdf_path.exists():
                logger.error(f"PDF conversion failed: {pdf_path}")
                return None

            # Step 2: PDF to PNG using pdf2image
            try:
                from pdf2image import convert_from_path

                images = convert_from_path(str(pdf_path), dpi=150)
                if slide_index < len(images):
                    png_path = output_dir / f"{pptx_path.stem}_slide_{slide_index:03d}.png"
                    images[slide_index].save(png_path, "PNG")
                    return png_path
                else:
                    logger.error(f"Slide index {slide_index} out of range")
                    return None
            except ImportError:
                logger.warning("pdf2image not available, returning PDF path")
                return pdf_path

        except subprocess.TimeoutExpired:
            logger.error("PPTX conversion timed out")
            return None
        except subprocess.CalledProcessError as e:
            logger.error(f"PPTX conversion failed: {e}")
            return None
        except Exception as e:
            logger.error(f"Conversion error: {e}")
            return None

    def test_slide(
        self,
        description_path: Path,
        reference_path: Path,
        template: str = "consulting_toolkit",
        threshold: Optional[float] = None
    ) -> TestResult:
        """
        Test a single slide against a reference image.

        Args:
            description_path: Path to slide description JSON
            reference_path: Path to reference PNG
            template: Template name
            threshold: SSIM threshold (uses default if not specified)

        Returns:
            TestResult
        """
        threshold = threshold or self.threshold
        test_name = description_path.stem

        try:
            # Load description
            with open(description_path, "r", encoding="utf-8") as f:
                description = json.load(f)

            # Generate PPTX
            from .builder import PresentationBuilder

            builder = PresentationBuilder(template)
            slide_type = description.get("type", "content")
            content = description.get("content", {})

            # Add slide based on type
            if slide_type == "title":
                builder.add_title_slide(
                    content.get("title", ""),
                    content.get("subtitle", "")
                )
            elif slide_type == "agenda":
                builder.add_agenda(
                    content.get("bullets", content.get("body", [])),
                    content.get("title", "Agenda")
                )
            else:
                builder.add_content_slide(
                    content.get("title", ""),
                    body=content.get("body", ""),
                    bullets=content.get("bullets", [])
                )

            # Save PPTX
            pptx_path = self.output_dir / f"{test_name}_test.pptx"
            builder.save(pptx_path)

            # Convert to PNG
            generated_path = self.pptx_to_png(pptx_path)
            if not generated_path or not generated_path.exists():
                return TestResult(
                    name=test_name,
                    passed=False,
                    ssim_score=0.0,
                    threshold=threshold,
                    error="Failed to generate PNG"
                )

            # Compare with reference
            ssim_score, diff_path = self.compare_images(
                generated_path,
                Path(reference_path)
            )

            return TestResult(
                name=test_name,
                passed=ssim_score >= threshold,
                ssim_score=ssim_score,
                threshold=threshold,
                generated_path=generated_path,
                reference_path=Path(reference_path),
                diff_path=diff_path
            )

        except Exception as e:
            return TestResult(
                name=test_name,
                passed=False,
                ssim_score=0.0,
                threshold=threshold,
                error=str(e)
            )

    def run_suite(
        self,
        test_cases: List[Tuple[Path, Path]],
        suite_name: str = "Visual Tests"
    ) -> TestSuite:
        """
        Run a suite of visual tests.

        Args:
            test_cases: List of (description_path, reference_path) tuples
            suite_name: Name for the test suite

        Returns:
            TestSuite with results
        """
        results = []
        for desc_path, ref_path in test_cases:
            result = self.test_slide(desc_path, ref_path)
            results.append(result)
            logger.info(str(result))

        return TestSuite(name=suite_name, results=results)

    def create_reference(
        self,
        pptx_path: Path,
        output_path: Path,
        slide_index: int = 0
    ) -> Optional[Path]:
        """
        Create a reference image from a PPTX file.

        Args:
            pptx_path: Source PPTX
            output_path: Where to save reference PNG
            slide_index: Which slide to capture

        Returns:
            Path to reference image
        """
        with tempfile.TemporaryDirectory() as tmp_dir:
            png_path = self.pptx_to_png(pptx_path, Path(tmp_dir), slide_index)
            if png_path and png_path.exists():
                output_path.parent.mkdir(parents=True, exist_ok=True)
                import shutil
                shutil.copy(png_path, output_path)
                return output_path
        return None


def quick_compare(image1: str, image2: str) -> float:
    """
    Quick SSIM comparison of two images.

    Args:
        image1: Path to first image
        image2: Path to second image

    Returns:
        SSIM score (0.0-1.0)
    """
    tester = VisualTester()
    score, _ = tester.compare_images(Path(image1), Path(image2))
    return score
