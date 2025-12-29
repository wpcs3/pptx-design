"""
Visual Comparison Module

Compares two images (original template render vs. generated output) using
structural similarity (SSIM) and generates visual diff images.
"""
import base64
import logging
from pathlib import Path
from typing import Optional

import numpy as np
from PIL import Image

# Add parent directory to path for config import
import sys
sys.path.insert(0, str(Path(__file__).parent.parent))

from config import DIFF_DIR

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class ComparisonError(Exception):
    """Exception raised when comparison fails."""
    pass


def load_image(image_path: Path) -> np.ndarray:
    """
    Load an image and convert to numpy array.

    Args:
        image_path: Path to the image file

    Returns:
        Numpy array of the image in RGB format

    Raises:
        ComparisonError: If image cannot be loaded
    """
    image_path = Path(image_path)

    if not image_path.exists():
        raise ComparisonError(f"Image not found: {image_path}")

    try:
        with Image.open(image_path) as img:
            # Convert to RGB if necessary (handles RGBA, grayscale, etc.)
            if img.mode != 'RGB':
                img = img.convert('RGB')
            return np.array(img)
    except Exception as e:
        raise ComparisonError(f"Failed to load image {image_path}: {e}")


def resize_to_match(img1: np.ndarray, img2: np.ndarray) -> tuple[np.ndarray, np.ndarray]:
    """
    Resize images to match dimensions (uses the larger dimensions).

    Args:
        img1: First image array
        img2: Second image array

    Returns:
        Tuple of (resized_img1, resized_img2)
    """
    h1, w1 = img1.shape[:2]
    h2, w2 = img2.shape[:2]

    if (h1, w1) == (h2, w2):
        return img1, img2

    # Use the larger dimensions
    target_h = max(h1, h2)
    target_w = max(w1, w2)

    logger.debug(f"Resizing images from ({w1}x{h1}), ({w2}x{h2}) to ({target_w}x{target_h})")

    def resize(img: np.ndarray, target_w: int, target_h: int) -> np.ndarray:
        pil_img = Image.fromarray(img)
        resized = pil_img.resize((target_w, target_h), Image.Resampling.LANCZOS)
        return np.array(resized)

    img1_resized = resize(img1, target_w, target_h) if (h1, w1) != (target_h, target_w) else img1
    img2_resized = resize(img2, target_w, target_h) if (h2, w2) != (target_h, target_w) else img2

    return img1_resized, img2_resized


def compute_similarity(
    image1_path: Path,
    image2_path: Path,
    resize: bool = True
) -> float:
    """
    Compute structural similarity between two images.

    Uses SSIM (Structural Similarity Index) from scikit-image.

    Args:
        image1_path: Path to the first image
        image2_path: Path to the second image
        resize: If True, resize images to match dimensions

    Returns:
        Similarity score from 0.0 (completely different) to 1.0 (identical)

    Raises:
        ComparisonError: If comparison fails
    """
    try:
        from skimage.metrics import structural_similarity as ssim
    except ImportError:
        raise ComparisonError(
            "scikit-image not installed. Run: pip install scikit-image"
        )

    img1 = load_image(image1_path)
    img2 = load_image(image2_path)

    if resize:
        img1, img2 = resize_to_match(img1, img2)

    # Check dimensions match
    if img1.shape != img2.shape:
        raise ComparisonError(
            f"Image dimensions do not match: {img1.shape} vs {img2.shape}. "
            "Set resize=True to automatically resize."
        )

    # Compute SSIM
    # Use channel_axis=2 for color images (RGB has channels in axis 2)
    try:
        score = ssim(img1, img2, channel_axis=2, data_range=255)
    except Exception as e:
        raise ComparisonError(f"SSIM computation failed: {e}")

    logger.info(f"Similarity score: {score:.4f}")
    return float(score)


def generate_diff_image(
    image1_path: Path,
    image2_path: Path,
    output_path: Optional[Path] = None,
    mode: str = "highlight"
) -> Path:
    """
    Generate a visual diff image highlighting differences.

    Args:
        image1_path: Path to the original/reference image
        image2_path: Path to the generated/comparison image
        output_path: Path for the output diff image (auto-generated if None)
        mode: Diff visualization mode:
            - "highlight": Highlight differences in red overlay
            - "sidebyside": Side-by-side comparison
            - "difference": Raw pixel difference

    Returns:
        Path to the generated diff image

    Raises:
        ComparisonError: If diff generation fails
    """
    img1 = load_image(image1_path)
    img2 = load_image(image2_path)
    img1, img2 = resize_to_match(img1, img2)

    if output_path is None:
        name1 = Path(image1_path).stem
        name2 = Path(image2_path).stem
        output_path = DIFF_DIR / f"diff_{name1}_vs_{name2}.png"
    else:
        output_path = Path(output_path)

    output_path.parent.mkdir(parents=True, exist_ok=True)

    if mode == "highlight":
        diff_img = _create_highlight_diff(img1, img2)
    elif mode == "sidebyside":
        diff_img = _create_sidebyside_diff(img1, img2)
    elif mode == "difference":
        diff_img = _create_difference_diff(img1, img2)
    else:
        raise ComparisonError(f"Unknown diff mode: {mode}")

    # Save the diff image
    Image.fromarray(diff_img).save(output_path)
    logger.info(f"Diff image saved: {output_path}")

    return output_path


def _create_highlight_diff(img1: np.ndarray, img2: np.ndarray) -> np.ndarray:
    """Create a diff that highlights differences in red on the original."""
    # Compute absolute difference
    diff = np.abs(img1.astype(np.int16) - img2.astype(np.int16))

    # Create mask where differences exceed threshold
    threshold = 30  # Sensitivity threshold
    mask = np.any(diff > threshold, axis=2)

    # Create output image (copy of original)
    output = img1.copy()

    # Highlight differences in red
    output[mask] = [255, 0, 0]

    return output


def _create_sidebyside_diff(img1: np.ndarray, img2: np.ndarray) -> np.ndarray:
    """Create a side-by-side comparison image."""
    h, w = img1.shape[:2]

    # Create separator line
    separator_width = 4
    separator = np.full((h, separator_width, 3), [128, 128, 128], dtype=np.uint8)

    # Concatenate horizontally
    output = np.hstack([img1, separator, img2])

    return output


def _create_difference_diff(img1: np.ndarray, img2: np.ndarray) -> np.ndarray:
    """Create a raw pixel difference image (amplified for visibility)."""
    # Compute absolute difference
    diff = np.abs(img1.astype(np.int16) - img2.astype(np.int16))

    # Amplify differences for visibility (multiply by 3, clamp to 255)
    amplified = np.clip(diff * 3, 0, 255).astype(np.uint8)

    return amplified


def image_to_base64(image_path: Path) -> str:
    """
    Convert an image file to base64 string for API calls.

    Args:
        image_path: Path to the image file

    Returns:
        Base64-encoded string of the image

    Raises:
        ComparisonError: If encoding fails
    """
    image_path = Path(image_path)

    if not image_path.exists():
        raise ComparisonError(f"Image not found: {image_path}")

    try:
        with open(image_path, 'rb') as f:
            return base64.standard_b64encode(f.read()).decode('utf-8')
    except Exception as e:
        raise ComparisonError(f"Failed to encode image to base64: {e}")


def prepare_comparison_prompt(
    image1_path: Path,
    image2_path: Path
) -> dict:
    """
    Prepare data structure for sending to vision model for comparison.

    Args:
        image1_path: Path to the original/reference image
        image2_path: Path to the generated/comparison image

    Returns:
        Dictionary with base64-encoded images and metadata, ready for API call
    """
    img1_path = Path(image1_path)
    img2_path = Path(image2_path)

    # Get image info
    with Image.open(img1_path) as img1:
        img1_size = img1.size
    with Image.open(img2_path) as img2:
        img2_size = img2.size

    return {
        "original": {
            "filename": img1_path.name,
            "base64": image_to_base64(img1_path),
            "size": img1_size,
            "media_type": "image/png"
        },
        "generated": {
            "filename": img2_path.name,
            "base64": image_to_base64(img2_path),
            "size": img2_size,
            "media_type": "image/png"
        },
        "similarity": compute_similarity(img1_path, img2_path)
    }


def compare_slides(
    original_path: Path,
    generated_path: Path,
    generate_diff: bool = True,
    diff_mode: str = "sidebyside"
) -> dict:
    """
    Full comparison of two slide images.

    Args:
        original_path: Path to the original slide image
        generated_path: Path to the generated slide image
        generate_diff: If True, generate a visual diff image
        diff_mode: Mode for diff visualization

    Returns:
        Dictionary with comparison results:
        {
            "similarity": float,
            "matches_threshold": bool,
            "diff_path": Path or None,
            "original_path": Path,
            "generated_path": Path
        }
    """
    from config import SIMILARITY_THRESHOLD

    similarity = compute_similarity(original_path, generated_path)

    result = {
        "similarity": similarity,
        "matches_threshold": similarity >= SIMILARITY_THRESHOLD,
        "diff_path": None,
        "original_path": Path(original_path),
        "generated_path": Path(generated_path)
    }

    if generate_diff:
        result["diff_path"] = generate_diff_image(
            original_path,
            generated_path,
            mode=diff_mode
        )

    return result


if __name__ == "__main__":
    # Test the comparator
    import sys

    if len(sys.argv) != 3:
        print("Usage: python comparator.py <image1> <image2>")
        print("\nThis will compute similarity and generate diff images.")
        sys.exit(1)

    img1_path = Path(sys.argv[1])
    img2_path = Path(sys.argv[2])

    print(f"Comparing: {img1_path.name} vs {img2_path.name}")
    print("=" * 50)

    try:
        result = compare_slides(img1_path, img2_path)

        print(f"Similarity: {result['similarity']:.4f}")
        print(f"Matches threshold: {result['matches_threshold']}")
        if result['diff_path']:
            print(f"Diff image: {result['diff_path']}")

    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)
