"""
Gemini Image Generator for Presentation Section Headers

Uses Google's Gemini API to generate professional images for PowerPoint
presentation section headers and title slides.

Usage:
    from pptx_generator.modules.image_generator import GeminiImageGenerator

    generator = GeminiImageGenerator()
    generator.generate(prompt, output_path)

    # Or generate all section images at once
    from pptx_generator.modules.image_generator import generate_all_section_images
    images = generate_all_section_images("cache/images/")
"""

import logging
import os
from pathlib import Path
from typing import Dict, Optional

logger = logging.getLogger(__name__)

# Section header image prompts - REALISTIC PHOTOGRAPHY ONLY
# Focus: SMALL BAY LIGHT INDUSTRIAL buildings (under 50,000 SF, 20-28 ft clear heights)
# NOT large bulk logistics warehouses
SECTION_IMAGE_PROMPTS = {
    "title_slide": """
Professional aerial photograph of a SMALL BAY LIGHT INDUSTRIAL business park during
golden hour. Multiple single-story flex/industrial buildings under 50,000 SF each with
multiple roll-up doors and suite entrances. Clean tilt-up concrete construction with
modest clear heights (20-28 feet). Parking areas with pickup trucks, vans, and small
box trucks typical of local businesses. Sunbelt suburban setting. High-end commercial
real estate marketing photography.

IMPORTANT: Show SMALL BAY light industrial (NOT large bulk logistics warehouses).
This must be a realistic photograph, NOT an illustration, diagram, or cartoon.
""",

    "market_fundamentals": """
Professional photograph of a modern SMALL BAY LIGHT INDUSTRIAL flex park at dusk.
Multiple smaller industrial buildings (10,000-40,000 SF each) with multiple tenant
suites visible. Roll-up doors, glass storefronts for showrooms, parking in front.
The buildings should look like multi-tenant flex/light industrial, NOT massive
distribution centers. Suburban business park setting.

IMPORTANT: Show SMALL BAY light industrial (NOT large bulk logistics warehouses).
This must be a realistic photograph, NOT an illustration, diagram, or cartoon.
""",

    "target_markets": """
Aerial photograph of a SMALL BAY LIGHT INDUSTRIAL business park in a Sunbelt city.
Multiple low-rise flex/industrial buildings arranged in a campus setting. Each building
has multiple tenant suites with individual entrances and loading areas. Interstate
visible nearby. Nashville, Tampa, or Phoenix style suburban industrial area. Clear day.

IMPORTANT: Show SMALL BAY light industrial parks (NOT large bulk logistics warehouses).
This must be a realistic photograph, NOT an illustration, diagram, or cartoon.
""",

    "demand_drivers": """
Interior photograph of a SMALL BAY LIGHT INDUSTRIAL space. A flex warehouse unit
around 5,000-15,000 SF with 24-foot clear height, showing a small business operation.
Could be light manufacturing, assembly, or local distribution. Modest racking, work
benches, small equipment. A few workers. The feel of a small/medium business tenant.

IMPORTANT: Show SMALL BAY interior (NOT massive e-commerce fulfillment centers).
This must be a realistic photograph, NOT an illustration, diagram, or cartoon.
""",

    "investment_strategy": """
Professional architectural photograph of a SMALL BAY LIGHT INDUSTRIAL building exterior.
A single-story flex/industrial building around 20,000-40,000 SF with 4-6 roll-up dock
doors and several glass entry doors for office/showroom suites. Clean tilt-up concrete
construction. Professional landscaping. Multiple tenant signage visible. Clear blue sky.
Institutional quality but clearly small-bay format.

IMPORTANT: Show SMALL BAY light industrial (NOT large bulk logistics warehouses).
This must be a realistic photograph, NOT an illustration, diagram, or cartoon.
""",

    "competitive_positioning": """
Photograph of a MULTI-TENANT SMALL BAY LIGHT INDUSTRIAL building with 6-10 suite
entrances and individual roll-up doors for each tenant. Various small business tenant
activity - service vans, pickup trucks, local delivery vehicles. Signage showing
diverse tenant mix: contractors, distributors, light manufacturers. Suburban setting.

IMPORTANT: Show SMALL BAY multi-tenant (NOT large single-tenant warehouses).
This must be a realistic photograph, NOT an illustration, diagram, or cartoon.
""",

    "risk_management": """
Photograph of a well-maintained SMALL BAY LIGHT INDUSTRIAL property showing quality
construction and professional management. Clean facades, fresh paint, maintained
landscaping, clear tenant signage, organized parking. The property conveys stability
and institutional quality for a multi-tenant small bay asset. Daytime with good lighting.

IMPORTANT: Show SMALL BAY light industrial (NOT large bulk logistics warehouses).
This must be a realistic photograph, NOT an illustration, diagram, or cartoon.
""",

    "esg_strategy": """
Aerial photograph of a SMALL BAY LIGHT INDUSTRIAL building with rooftop solar panels.
A flex/industrial building around 30,000-50,000 SF with solar array on roof. EV charging
stations in parking lot. Native drought-tolerant landscaping. Modern clean building
with multiple tenant suites. Bright sunny day. Sustainability through real infrastructure.

IMPORTANT: Show SMALL BAY light industrial (NOT massive distribution centers).
This must be a realistic photograph, NOT an illustration, diagram, or cartoon.
""",

    "jv_structure": """
Professional photograph of a SMALL BAY LIGHT INDUSTRIAL business park campus. Multiple
flex/industrial buildings (each under 50,000 SF) within a master-planned development.
Coordinated architecture, shared landscaping, visible tenant activity. The scene conveys
institutional-quality small bay development suitable for pension fund investment.

IMPORTANT: Show SMALL BAY light industrial parks (NOT bulk logistics).
This must be a realistic photograph, NOT an illustration, diagram, or cartoon.
""",

    "conclusion": """
Panoramic sunrise photograph of a thriving SMALL BAY LIGHT INDUSTRIAL business park.
Multiple flex/industrial buildings with tenant activity beginning for the day. Local
delivery trucks, service vehicles, workers arriving. American flag visible. Golden
morning light. The image conveys the essential nature of small bay industrial for
local economies. Optimistic, prosperous mood.

IMPORTANT: Show SMALL BAY light industrial (NOT large bulk logistics warehouses).
This must be a realistic photograph, NOT an illustration, diagram, or cartoon.
""",

    "end_slide": """
Professional twilight photograph of a premium SMALL BAY LIGHT INDUSTRIAL property.
A well-lit flex/industrial building at dusk with warm interior lights glowing through
windows and glass doors. Clean modern architecture. The mood should be professional,
successful, and inviting - suitable for an ending slide. Deep blue sky gradient.

IMPORTANT: Show SMALL BAY light industrial (NOT large bulk logistics warehouses).
This must be a realistic photograph, NOT an illustration, diagram, or cartoon.
""",
}


class GeminiImageGenerator:
    """Generate presentation images using Google Gemini API."""

    def __init__(self, api_key: Optional[str] = None):
        """
        Initialize with API key from param or environment.

        Args:
            api_key: Optional Gemini API key. If not provided, reads from
                     GOOGLE_API_KEY or GEMINI_API_KEY environment variables.
        """
        from google import genai

        # Try to load from .env file
        env_path = Path(__file__).parent.parent.parent / ".env"
        if env_path.exists():
            self._load_env(env_path)

        key = api_key or os.environ.get("GOOGLE_API_KEY") or os.environ.get("GEMINI_API_KEY")
        if not key:
            raise ValueError(
                "Gemini API key required. Set GOOGLE_API_KEY environment variable "
                "or pass api_key parameter."
            )

        self.client = genai.Client(api_key=key)
        logger.info("Gemini API client initialized successfully")

    def _load_env(self, env_path: Path):
        """Load environment variables from .env file."""
        try:
            with open(env_path, 'r') as f:
                for line in f:
                    line = line.strip()
                    if line and not line.startswith('#') and '=' in line:
                        key, value = line.split('=', 1)
                        os.environ[key.strip()] = value.strip()
        except Exception as e:
            logger.warning(f"Could not load .env file: {e}")

    def generate(
        self,
        prompt: str,
        output_path: Path,
        width: int = 1920,
        height: int = 1080
    ) -> Path:
        """
        Generate image from prompt and save to path using Imagen 3.

        Args:
            prompt: Natural language image description
            output_path: Where to save the PNG
            width: Output width in pixels (default 1920 for HD)
            height: Output height in pixels (default 1080 for HD)

        Returns:
            Path to saved image
        """
        from PIL import Image
        from google.genai import types

        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)

        # Append slide formatting guidance
        full_prompt = (
            f"{prompt}\n\n"
            f"Generate this as a high-quality image suitable for a "
            f"PowerPoint presentation slide. Do not include any text, "
            f"labels, titles, or watermarks in the image."
        )

        try:
            # Use Gemini 3 Pro Image Preview for image generation
            response = self.client.models.generate_content(
                model="gemini-3-pro-image-preview",
                contents=full_prompt,
                config=types.GenerateContentConfig(
                    response_modalities=["IMAGE", "TEXT"],
                )
            )

            # Extract image from response
            import io
            for part in response.candidates[0].content.parts:
                if hasattr(part, 'inline_data') and part.inline_data is not None:
                    image_bytes = part.inline_data.data
                    pil_image = Image.open(io.BytesIO(image_bytes))

                    # Resize to exact dimensions if needed
                    if pil_image.size != (width, height):
                        pil_image = pil_image.resize((width, height), Image.LANCZOS)

                    pil_image.save(output_path, "PNG", quality=95)
                    logger.info(f"Generated image saved to: {output_path}")
                    return output_path

            raise RuntimeError("Gemini did not return any images")

        except Exception as e:
            logger.error(f"Image generation failed: {e}")
            raise

    def generate_with_gemini_flash(
        self,
        prompt: str,
        output_path: Path,
        width: int = 1920,
        height: int = 1080
    ) -> Path:
        """
        Alternative method using Gemini 3 Flash Preview with native image output.

        This uses the multimodal generation capability of Gemini models
        that can output images directly.
        """
        from PIL import Image
        from google.genai import types
        import io

        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)

        full_prompt = (
            f"{prompt}\n\n"
            f"Generate this as a high-quality image suitable for a PowerPoint presentation slide. "
            f"Do not include any text, labels, titles, or watermarks in the image."
        )

        try:
            # Use Gemini 3 Flash Preview with image generation
            response = self.client.models.generate_content(
                model="gemini-3-flash-preview",
                contents=full_prompt,
                config=types.GenerateContentConfig(
                    response_modalities=["IMAGE", "TEXT"],
                )
            )

            # Extract image from response
            for part in response.candidates[0].content.parts:
                if hasattr(part, 'inline_data') and part.inline_data is not None:
                    image_bytes = part.inline_data.data
                    pil_image = Image.open(io.BytesIO(image_bytes))

                    # Resize to target dimensions
                    if pil_image.size != (width, height):
                        pil_image = pil_image.resize((width, height), Image.LANCZOS)

                    pil_image.save(output_path, "PNG", quality=95)
                    logger.info(f"Generated image saved to: {output_path}")
                    return output_path

            raise RuntimeError("No image in Gemini response")

        except Exception as e:
            logger.error(f"Gemini Flash image generation failed: {e}")
            raise


def generate_all_section_images(
    output_dir: str = "cache/images/",
    use_flash: bool = False
) -> Dict[str, Path]:
    """
    Generate all section header images for the presentation.

    Args:
        output_dir: Directory to save generated images
        use_flash: If True, use Gemini Flash instead of Imagen 3

    Returns:
        Dictionary mapping section names to image file paths
    """
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)

    generator = GeminiImageGenerator()
    generated_images = {}

    for section_name, prompt in SECTION_IMAGE_PROMPTS.items():
        image_path = output_path / f"{section_name}.png"

        try:
            if use_flash:
                generated_images[section_name] = generator.generate_with_gemini_flash(
                    prompt, image_path
                )
            else:
                generated_images[section_name] = generator.generate(prompt, image_path)
            print(f"  Generated: {section_name}")
        except Exception as e:
            print(f"  Failed {section_name}: {e}")
            logger.error(f"Failed to generate {section_name}: {e}")

    return generated_images


def get_section_image(
    section_name: str,
    cache_dir: str = "cache/images/"
) -> Optional[Path]:
    """
    Get section image from cache, generating if needed.

    Args:
        section_name: Name of the section (e.g., 'title_slide', 'market_fundamentals')
        cache_dir: Directory for cached images

    Returns:
        Path to image file, or None if generation fails
    """
    cached_path = Path(cache_dir) / f"{section_name}.png"

    if cached_path.exists():
        return cached_path

    try:
        generator = GeminiImageGenerator()
        prompt = SECTION_IMAGE_PROMPTS.get(section_name)
        if prompt:
            return generator.generate(prompt, cached_path)
    except Exception as e:
        logger.warning(f"Image generation failed, returning None: {e}")

    return None


# CLI interface
if __name__ == "__main__":
    import sys

    print("Gemini Image Generator for PPTX Presentations")
    print("=" * 50)

    if len(sys.argv) > 1:
        # Generate specific section
        section = sys.argv[1]
        if section in SECTION_IMAGE_PROMPTS:
            print(f"\nGenerating: {section}")
            try:
                generator = GeminiImageGenerator()
                output = Path(f"cache/images/{section}.png")
                generator.generate(SECTION_IMAGE_PROMPTS[section], output)
                print(f"Saved to: {output}")
            except Exception as e:
                print(f"Error: {e}")
        elif section == "--flash":
            # Use Gemini Flash for all
            print("\nGenerating all section images with Gemini Flash...")
            images = generate_all_section_images(use_flash=True)
            print(f"\nGenerated {len(images)} images")
        else:
            print(f"Unknown section: {section}")
            print(f"Available sections: {', '.join(SECTION_IMAGE_PROMPTS.keys())}")
    else:
        # Generate all sections
        print("\nGenerating all section images with Imagen 3...")
        print(f"Sections: {', '.join(SECTION_IMAGE_PROMPTS.keys())}")
        print()

        images = generate_all_section_images()

        print()
        print(f"Generated {len(images)} images")
        for name, path in images.items():
            print(f"  {name}: {path}")
