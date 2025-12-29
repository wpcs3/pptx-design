"""
Image Search Integration

Provides automatic image sourcing from Pexels and Unsplash APIs
for presentation slides.

Phase 3 Enhancement (2025-12-29):
- Pexels API integration
- Unsplash API integration
- Keyword extraction from slide content
- Probabilistic image selection
- Image caching for performance

Usage:
    from pptx_generator.modules.image_search import ImageSearch

    search = ImageSearch()
    images = search.search("business meeting", count=3)
    image_path = search.download(images[0])

Environment Variables:
    PEXELS_API_KEY: API key for Pexels (https://www.pexels.com/api/)
    UNSPLASH_ACCESS_KEY: Access key for Unsplash (https://unsplash.com/developers)
"""

import hashlib
import logging
import os
import random
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional
from urllib.parse import urljoin

import requests

logger = logging.getLogger(__name__)


# =============================================================================
# Data Classes
# =============================================================================

@dataclass
class ImageResult:
    """Represents an image search result."""
    id: str
    url: str  # Full-size URL
    thumbnail_url: str  # Small preview URL
    width: int
    height: int
    photographer: str
    photographer_url: str
    source: str  # 'pexels' or 'unsplash'
    alt_text: str = ""
    download_url: Optional[str] = None  # Direct download URL

    @property
    def aspect_ratio(self) -> float:
        """Calculate aspect ratio (width/height)."""
        return self.width / self.height if self.height > 0 else 1.0

    @property
    def is_landscape(self) -> bool:
        """Check if image is landscape orientation."""
        return self.aspect_ratio > 1.0

    @property
    def is_portrait(self) -> bool:
        """Check if image is portrait orientation."""
        return self.aspect_ratio < 1.0


# =============================================================================
# API Providers
# =============================================================================

class PexelsProvider:
    """Pexels API provider."""

    BASE_URL = "https://api.pexels.com/v1/"

    def __init__(self, api_key: str):
        self.api_key = api_key
        self.headers = {"Authorization": api_key}

    def search(
        self,
        query: str,
        count: int = 10,
        orientation: str = None,
        size: str = "medium"
    ) -> List[ImageResult]:
        """
        Search for images on Pexels.

        Args:
            query: Search keywords
            count: Number of results (max 80)
            orientation: 'landscape', 'portrait', or 'square'
            size: 'large', 'medium', or 'small'

        Returns:
            List of ImageResult objects
        """
        params = {
            "query": query,
            "per_page": min(count, 80),
            "size": size
        }

        if orientation:
            params["orientation"] = orientation

        try:
            response = requests.get(
                urljoin(self.BASE_URL, "search"),
                headers=self.headers,
                params=params,
                timeout=10
            )
            response.raise_for_status()
            data = response.json()

            results = []
            for photo in data.get("photos", []):
                results.append(ImageResult(
                    id=str(photo["id"]),
                    url=photo["src"]["original"],
                    thumbnail_url=photo["src"]["medium"],
                    width=photo["width"],
                    height=photo["height"],
                    photographer=photo["photographer"],
                    photographer_url=photo["photographer_url"],
                    source="pexels",
                    alt_text=photo.get("alt", ""),
                    download_url=photo["src"]["original"]
                ))

            return results

        except requests.RequestException as e:
            logger.error(f"Pexels API error: {e}")
            return []

    def get_curated(self, count: int = 10) -> List[ImageResult]:
        """Get curated photos from Pexels."""
        try:
            response = requests.get(
                urljoin(self.BASE_URL, "curated"),
                headers=self.headers,
                params={"per_page": min(count, 80)},
                timeout=10
            )
            response.raise_for_status()
            data = response.json()

            results = []
            for photo in data.get("photos", []):
                results.append(ImageResult(
                    id=str(photo["id"]),
                    url=photo["src"]["original"],
                    thumbnail_url=photo["src"]["medium"],
                    width=photo["width"],
                    height=photo["height"],
                    photographer=photo["photographer"],
                    photographer_url=photo["photographer_url"],
                    source="pexels",
                    alt_text=photo.get("alt", ""),
                    download_url=photo["src"]["original"]
                ))

            return results

        except requests.RequestException as e:
            logger.error(f"Pexels API error: {e}")
            return []


class UnsplashProvider:
    """Unsplash API provider."""

    BASE_URL = "https://api.unsplash.com/"

    def __init__(self, access_key: str):
        self.access_key = access_key
        self.headers = {"Authorization": f"Client-ID {access_key}"}

    def search(
        self,
        query: str,
        count: int = 10,
        orientation: str = None
    ) -> List[ImageResult]:
        """
        Search for images on Unsplash.

        Args:
            query: Search keywords
            count: Number of results (max 30)
            orientation: 'landscape', 'portrait', or 'squarish'

        Returns:
            List of ImageResult objects
        """
        params = {
            "query": query,
            "per_page": min(count, 30)
        }

        if orientation:
            params["orientation"] = orientation

        try:
            response = requests.get(
                urljoin(self.BASE_URL, "search/photos"),
                headers=self.headers,
                params=params,
                timeout=10
            )
            response.raise_for_status()
            data = response.json()

            results = []
            for photo in data.get("results", []):
                results.append(ImageResult(
                    id=photo["id"],
                    url=photo["urls"]["full"],
                    thumbnail_url=photo["urls"]["small"],
                    width=photo["width"],
                    height=photo["height"],
                    photographer=photo["user"]["name"],
                    photographer_url=photo["user"]["links"]["html"],
                    source="unsplash",
                    alt_text=photo.get("alt_description", "") or "",
                    download_url=photo["links"]["download"]
                ))

            return results

        except requests.RequestException as e:
            logger.error(f"Unsplash API error: {e}")
            return []

    def get_random(
        self,
        query: str = None,
        count: int = 1,
        orientation: str = None
    ) -> List[ImageResult]:
        """Get random photos from Unsplash."""
        params = {"count": min(count, 30)}

        if query:
            params["query"] = query
        if orientation:
            params["orientation"] = orientation

        try:
            response = requests.get(
                urljoin(self.BASE_URL, "photos/random"),
                headers=self.headers,
                params=params,
                timeout=10
            )
            response.raise_for_status()
            data = response.json()

            # Handle single vs multiple results
            if isinstance(data, dict):
                data = [data]

            results = []
            for photo in data:
                results.append(ImageResult(
                    id=photo["id"],
                    url=photo["urls"]["full"],
                    thumbnail_url=photo["urls"]["small"],
                    width=photo["width"],
                    height=photo["height"],
                    photographer=photo["user"]["name"],
                    photographer_url=photo["user"]["links"]["html"],
                    source="unsplash",
                    alt_text=photo.get("alt_description", "") or "",
                    download_url=photo["links"]["download"]
                ))

            return results

        except requests.RequestException as e:
            logger.error(f"Unsplash API error: {e}")
            return []


# =============================================================================
# Unified Image Search
# =============================================================================

class ImageSearch:
    """
    Unified image search across multiple providers.

    Usage:
        search = ImageSearch()
        images = search.search("business meeting", count=3)
        path = search.download(images[0], output_dir="./images")
    """

    # Keywords to extract from slide content
    BUSINESS_KEYWORDS = [
        "business", "meeting", "office", "team", "collaboration",
        "strategy", "growth", "success", "innovation", "technology",
        "data", "analytics", "finance", "investment", "market",
        "professional", "corporate", "leadership", "presentation"
    ]

    def __init__(
        self,
        pexels_key: str = None,
        unsplash_key: str = None,
        cache_dir: str = None
    ):
        """
        Initialize image search with API keys.

        Args:
            pexels_key: Pexels API key (or PEXELS_API_KEY env var)
            unsplash_key: Unsplash access key (or UNSPLASH_ACCESS_KEY env var)
            cache_dir: Directory for caching downloaded images
        """
        self.pexels_key = pexels_key or os.getenv("PEXELS_API_KEY")
        self.unsplash_key = unsplash_key or os.getenv("UNSPLASH_ACCESS_KEY")

        self.providers = []
        if self.pexels_key:
            self.providers.append(PexelsProvider(self.pexels_key))
            logger.info("Pexels provider initialized")
        if self.unsplash_key:
            self.providers.append(UnsplashProvider(self.unsplash_key))
            logger.info("Unsplash provider initialized")

        if not self.providers:
            logger.warning(
                "No image search providers configured. "
                "Set PEXELS_API_KEY or UNSPLASH_ACCESS_KEY environment variables."
            )

        self.cache_dir = Path(cache_dir) if cache_dir else Path.cwd() / ".image_cache"
        self.cache_dir.mkdir(parents=True, exist_ok=True)

    @property
    def is_available(self) -> bool:
        """Check if any provider is available."""
        return len(self.providers) > 0

    def search(
        self,
        query: str,
        count: int = 5,
        orientation: str = None,
        provider: str = None
    ) -> List[ImageResult]:
        """
        Search for images across all providers.

        Args:
            query: Search keywords
            count: Number of results
            orientation: 'landscape', 'portrait', or 'square'
            provider: Specific provider ('pexels' or 'unsplash')

        Returns:
            List of ImageResult objects
        """
        if not self.providers:
            logger.warning("No image providers available")
            return []

        results = []

        for p in self.providers:
            if provider and not isinstance(p, self._get_provider_class(provider)):
                continue

            try:
                provider_results = p.search(
                    query,
                    count=count,
                    orientation=orientation
                )
                results.extend(provider_results)
            except Exception as e:
                logger.error(f"Provider search error: {e}")

        # Shuffle to mix providers
        random.shuffle(results)

        return results[:count]

    def search_for_slide(
        self,
        slide_content: Dict[str, Any],
        count: int = 3,
        orientation: str = "landscape"
    ) -> List[ImageResult]:
        """
        Search for images based on slide content.

        Automatically extracts keywords from slide title and body.

        Args:
            slide_content: Dictionary with 'title', 'body', 'bullets' etc.
            count: Number of results
            orientation: Preferred orientation

        Returns:
            List of ImageResult objects
        """
        # Extract keywords from slide content
        keywords = self.extract_keywords(slide_content)

        if not keywords:
            # Fallback to generic business keywords
            keywords = random.sample(self.BUSINESS_KEYWORDS, min(3, len(self.BUSINESS_KEYWORDS)))

        query = " ".join(keywords[:5])  # Use top 5 keywords
        logger.info(f"Searching for slide images: '{query}'")

        return self.search(query, count=count, orientation=orientation)

    def extract_keywords(self, content: Dict[str, Any]) -> List[str]:
        """
        Extract relevant keywords from slide content.

        Args:
            content: Slide content dictionary

        Returns:
            List of keywords
        """
        text_parts = []

        # Collect all text from content
        if "title" in content:
            text_parts.append(content["title"])
        if "subtitle" in content:
            text_parts.append(content["subtitle"])
        if "body" in content:
            if isinstance(content["body"], list):
                text_parts.extend(content["body"])
            else:
                text_parts.append(content["body"])
        if "bullets" in content:
            text_parts.extend(content["bullets"])

        # Combine and clean text
        full_text = " ".join(str(t) for t in text_parts).lower()

        # Remove common stop words
        stop_words = {
            "the", "a", "an", "and", "or", "but", "in", "on", "at", "to",
            "for", "of", "with", "by", "from", "is", "are", "was", "were",
            "be", "been", "being", "have", "has", "had", "do", "does", "did",
            "will", "would", "could", "should", "may", "might", "must", "shall"
        }

        # Extract words (alphanumeric only)
        words = re.findall(r'\b[a-z]{3,}\b', full_text)
        keywords = [w for w in words if w not in stop_words]

        # Prioritize business keywords
        prioritized = []
        for kw in self.BUSINESS_KEYWORDS:
            if kw in keywords:
                prioritized.append(kw)
                keywords.remove(kw)

        # Add remaining keywords
        prioritized.extend(keywords[:10])

        return prioritized

    def download(
        self,
        image: ImageResult,
        output_dir: str = None,
        filename: str = None
    ) -> Optional[Path]:
        """
        Download an image to local storage.

        Args:
            image: ImageResult to download
            output_dir: Directory to save to (uses cache_dir if not specified)
            filename: Custom filename (auto-generated if not specified)

        Returns:
            Path to downloaded image, or None if failed
        """
        output_dir = Path(output_dir) if output_dir else self.cache_dir
        output_dir.mkdir(parents=True, exist_ok=True)

        # Generate filename from URL hash if not provided
        if not filename:
            url_hash = hashlib.md5(image.url.encode()).hexdigest()[:10]
            filename = f"{image.source}_{image.id}_{url_hash}.jpg"

        output_path = output_dir / filename

        # Check cache
        if output_path.exists():
            logger.debug(f"Using cached image: {output_path}")
            return output_path

        # Download
        try:
            download_url = image.download_url or image.url
            response = requests.get(download_url, timeout=30)
            response.raise_for_status()

            with open(output_path, 'wb') as f:
                f.write(response.content)

            logger.info(f"Downloaded image: {output_path}")
            return output_path

        except Exception as e:
            logger.error(f"Error downloading image: {e}")
            return None

    def _get_provider_class(self, provider_name: str):
        """Get provider class by name."""
        providers = {
            "pexels": PexelsProvider,
            "unsplash": UnsplashProvider
        }
        return providers.get(provider_name.lower())

    def get_attribution(self, image: ImageResult) -> str:
        """
        Get attribution text for an image.

        Args:
            image: ImageResult

        Returns:
            Attribution string (required by Pexels/Unsplash TOS)
        """
        if image.source == "pexels":
            return f"Photo by {image.photographer} on Pexels"
        elif image.source == "unsplash":
            return f"Photo by {image.photographer} on Unsplash"
        else:
            return f"Photo by {image.photographer}"


# =============================================================================
# Convenience Functions
# =============================================================================

def search_images(query: str, count: int = 5) -> List[ImageResult]:
    """
    Quick image search using default settings.

    Args:
        query: Search keywords
        count: Number of results

    Returns:
        List of ImageResult objects
    """
    search = ImageSearch()
    return search.search(query, count=count)


def get_slide_image(slide_content: Dict[str, Any]) -> Optional[ImageResult]:
    """
    Get a single relevant image for a slide.

    Args:
        slide_content: Slide content dictionary

    Returns:
        Single ImageResult or None
    """
    search = ImageSearch()
    results = search.search_for_slide(slide_content, count=1)
    return results[0] if results else None


# =============================================================================
# CLI
# =============================================================================

def main():
    """CLI for testing image search."""
    import argparse

    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

    parser = argparse.ArgumentParser(description="Image Search Test")
    parser.add_argument("query", help="Search query")
    parser.add_argument("-c", "--count", type=int, default=3, help="Number of results")
    parser.add_argument("-o", "--orientation", choices=["landscape", "portrait", "square"],
                        help="Image orientation")
    parser.add_argument("-d", "--download", action="store_true", help="Download first result")

    args = parser.parse_args()

    search = ImageSearch()

    if not search.is_available:
        print("No API keys configured. Set PEXELS_API_KEY or UNSPLASH_ACCESS_KEY.")
        return

    print(f"Searching for: {args.query}")
    results = search.search(args.query, count=args.count, orientation=args.orientation)

    print(f"\nFound {len(results)} results:")
    for i, img in enumerate(results, 1):
        print(f"\n{i}. {img.source.upper()}")
        print(f"   Size: {img.width}x{img.height} ({img.aspect_ratio:.2f})")
        print(f"   Photographer: {img.photographer}")
        print(f"   URL: {img.thumbnail_url}")

    if args.download and results:
        print("\nDownloading first result...")
        path = search.download(results[0])
        if path:
            print(f"Saved to: {path}")


if __name__ == "__main__":
    main()
