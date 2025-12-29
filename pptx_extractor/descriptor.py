"""
Slide Description Generator Module

Generates natural language descriptions of slide designs using vision models.
This is the core intelligence of the system.

Supports multiple LLM providers:
- Anthropic (Claude): claude-opus-4.5, claude-sonnet-4.5, claude-haiku-4.5
- OpenAI (GPT): gpt-5.2, gpt-5.1, gpt-5, gpt-5-mini, gpt-5-nano
- Google (Gemini): gemini-3-pro, gemini-3-flash
"""
import json
import logging
import re
from pathlib import Path
from typing import Optional

# Add parent directory to path for config import
import sys
sys.path.insert(0, str(Path(__file__).parent.parent))

from config import DESCRIPTION_DIR, SRC_DIR

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Default model for vision tasks
DEFAULT_VISION_MODEL = "gemini-3-flash"


class DescriptorError(Exception):
    """Exception raised when description generation fails."""
    pass


def load_prompt(prompt_name: str) -> str:
    """
    Load a prompt template from the prompts directory.

    Args:
        prompt_name: Name of the prompt file (without .txt extension)

    Returns:
        Prompt template string
    """
    prompt_path = SRC_DIR / "prompts" / f"{prompt_name}.txt"

    if not prompt_path.exists():
        raise DescriptorError(f"Prompt template not found: {prompt_path}")

    with open(prompt_path, 'r', encoding='utf-8') as f:
        return f.read()


def describe_slide_design(
    image_path: Path,
    use_anthropic: bool = False,
    anthropic_client=None,
    model: Optional[str] = None
) -> dict:
    """
    Analyze a slide image and generate a structured description.

    This function can work in multiple modes:
    1. Using configurable LLM provider (model parameter specified)
    2. Standalone mode with Anthropic API (use_anthropic=True, legacy)
    3. Returns prompt data for external processing (default)

    Args:
        image_path: Path to the slide image
        use_anthropic: If True, use Anthropic API directly (legacy)
        anthropic_client: Pre-configured Anthropic client (optional, legacy)
        model: Model name to use (e.g., "claude-sonnet-4.5", "gpt-5", "gemini-3-pro")
               If specified, uses the pptx_generator LLM provider

    Returns:
        Structured description dict or prompt data for external processing
    """
    from pptx_extractor.comparator import image_to_base64

    image_path = Path(image_path)
    if not image_path.exists():
        raise DescriptorError(f"Image not found: {image_path}")

    prompt = load_prompt("description_prompt")
    image_b64 = image_to_base64(image_path)

    # Use configurable LLM provider if model is specified
    if model:
        return _call_llm_vision(prompt, image_b64, model)
    elif use_anthropic:
        return _call_anthropic_vision(prompt, image_b64, anthropic_client)
    else:
        # Return data for external processing (e.g., by Claude Code)
        return {
            "mode": "external",
            "prompt": prompt,
            "image_base64": image_b64,
            "image_path": str(image_path),
            "media_type": "image/png"
        }


def _call_llm_vision(prompt: str, image_b64: str, model: str) -> dict:
    """
    Call LLM with vision capabilities using the configurable provider.

    Args:
        prompt: The prompt to send
        image_b64: Base64-encoded image
        model: Model name from the LLM provider

    Returns:
        Parsed JSON description
    """
    try:
        from pptx_generator.modules.llm_provider import LLMManager, ModelProvider, AVAILABLE_MODELS
    except ImportError:
        raise DescriptorError("pptx_generator module not found. Ensure it's in the Python path.")

    model_config = AVAILABLE_MODELS.get(model)
    if not model_config:
        raise DescriptorError(f"Unknown model: {model}. Available: {list(AVAILABLE_MODELS.keys())}")

    provider = model_config.provider

    if provider == ModelProvider.ANTHROPIC:
        return _call_anthropic_vision_with_model(prompt, image_b64, model_config.model_id)
    elif provider == ModelProvider.OPENAI:
        return _call_openai_vision(prompt, image_b64, model_config.model_id)
    elif provider == ModelProvider.GOOGLE:
        return _call_google_vision(prompt, image_b64, model_config.model_id)
    else:
        raise DescriptorError(f"Unsupported provider for vision: {provider}")


def _call_anthropic_vision_with_model(prompt: str, image_b64: str, model_id: str) -> dict:
    """Call Anthropic API with a specific model."""
    try:
        import anthropic
    except ImportError:
        raise DescriptorError("anthropic package not installed. Run: pip install anthropic")

    from pptx_generator.modules.llm_provider import load_env_file
    load_env_file()

    import os
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        raise DescriptorError("ANTHROPIC_API_KEY not found in environment")

    client = anthropic.Anthropic(api_key=api_key)

    message = client.messages.create(
        model=model_id,
        max_tokens=4096,
        messages=[
            {
                "role": "user",
                "content": [
                    {
                        "type": "image",
                        "source": {
                            "type": "base64",
                            "media_type": "image/png",
                            "data": image_b64
                        }
                    },
                    {
                        "type": "text",
                        "text": prompt
                    }
                ]
            }
        ]
    )

    response_text = message.content[0].text
    return _parse_json_response(response_text)


def _call_openai_vision(prompt: str, image_b64: str, model_id: str) -> dict:
    """Call OpenAI API with vision."""
    try:
        from openai import OpenAI
    except ImportError:
        raise DescriptorError("openai package not installed. Run: pip install openai")

    from pptx_generator.modules.llm_provider import load_env_file
    load_env_file()

    import os
    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        raise DescriptorError("OPENAI_API_KEY not found in environment")

    client = OpenAI(api_key=api_key)

    response = client.chat.completions.create(
        model=model_id,
        messages=[
            {
                "role": "user",
                "content": [
                    {
                        "type": "image_url",
                        "image_url": {
                            "url": f"data:image/png;base64,{image_b64}"
                        }
                    },
                    {
                        "type": "text",
                        "text": prompt
                    }
                ]
            }
        ],
        max_tokens=4096
    )

    response_text = response.choices[0].message.content
    return _parse_json_response(response_text)


def _call_google_vision(prompt: str, image_b64: str, model_id: str) -> dict:
    """Call Google Gemini API with vision."""
    try:
        import google.generativeai as genai
    except ImportError:
        raise DescriptorError("google-generativeai package not installed. Run: pip install google-generativeai")

    from pptx_generator.modules.llm_provider import load_env_file
    load_env_file()

    import os
    import base64
    api_key = os.environ.get("GOOGLE_API_KEY")
    if not api_key:
        raise DescriptorError("GOOGLE_API_KEY not found in environment")

    genai.configure(api_key=api_key)
    model = genai.GenerativeModel(model_id)

    # Decode base64 image for Gemini
    image_bytes = base64.b64decode(image_b64)

    response = model.generate_content([
        {
            "mime_type": "image/png",
            "data": image_bytes
        },
        prompt
    ])

    response_text = response.text
    return _parse_json_response(response_text)


def _call_anthropic_vision(prompt: str, image_b64: str, client=None) -> dict:
    """
    Call Anthropic API with vision to analyze the image.

    Args:
        prompt: The prompt to send
        image_b64: Base64-encoded image
        client: Pre-configured Anthropic client

    Returns:
        Parsed JSON description
    """
    try:
        import anthropic
    except ImportError:
        raise DescriptorError("anthropic package not installed. Run: pip install anthropic")

    if client is None:
        client = anthropic.Anthropic()

    message = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=4096,
        messages=[
            {
                "role": "user",
                "content": [
                    {
                        "type": "image",
                        "source": {
                            "type": "base64",
                            "media_type": "image/png",
                            "data": image_b64
                        }
                    },
                    {
                        "type": "text",
                        "text": prompt
                    }
                ]
            }
        ]
    )

    response_text = message.content[0].text
    return _parse_json_response(response_text)


def _parse_json_response(response_text: str) -> dict:
    """
    Parse JSON from the model response.

    Handles cases where JSON is wrapped in markdown code blocks.

    Args:
        response_text: Raw response text from the model

    Returns:
        Parsed JSON as dict

    Raises:
        DescriptorError: If JSON parsing fails
    """
    # Try to extract JSON from markdown code blocks
    json_match = re.search(r'```(?:json)?\s*([\s\S]*?)\s*```', response_text)
    if json_match:
        json_str = json_match.group(1)
    else:
        # Try parsing the whole response as JSON
        json_str = response_text.strip()

    try:
        return json.loads(json_str)
    except json.JSONDecodeError as e:
        logger.error(f"Failed to parse JSON response: {e}")
        logger.error(f"Response text: {response_text[:500]}...")
        raise DescriptorError(f"Failed to parse description JSON: {e}")


def compare_slides_for_differences(
    original_image_path: Path,
    generated_image_path: Path,
    use_anthropic: bool = False,
    anthropic_client=None
) -> str:
    """
    Compare two slide images and describe the differences.

    Args:
        original_image_path: Path to the original slide image
        generated_image_path: Path to the generated slide image
        use_anthropic: If True, use Anthropic API directly
        anthropic_client: Pre-configured Anthropic client

    Returns:
        String describing the differences, or prompt data for external processing
    """
    from pptx_extractor.comparator import image_to_base64

    prompt = load_prompt("comparison_prompt")

    original_b64 = image_to_base64(original_image_path)
    generated_b64 = image_to_base64(generated_image_path)

    if use_anthropic:
        return _call_anthropic_comparison(
            prompt, original_b64, generated_b64, anthropic_client
        )
    else:
        return {
            "mode": "external",
            "prompt": prompt,
            "original_image_base64": original_b64,
            "generated_image_base64": generated_b64,
            "original_path": str(original_image_path),
            "generated_path": str(generated_image_path),
            "media_type": "image/png"
        }


def _call_anthropic_comparison(
    prompt: str,
    original_b64: str,
    generated_b64: str,
    client=None
) -> str:
    """
    Call Anthropic API with two images for comparison.

    Args:
        prompt: The comparison prompt
        original_b64: Base64-encoded original image
        generated_b64: Base64-encoded generated image
        client: Pre-configured Anthropic client

    Returns:
        String describing differences
    """
    try:
        import anthropic
    except ImportError:
        raise DescriptorError("anthropic package not installed")

    if client is None:
        client = anthropic.Anthropic()

    message = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=4096,
        messages=[
            {
                "role": "user",
                "content": [
                    {
                        "type": "text",
                        "text": "IMAGE 1 (Original Template):"
                    },
                    {
                        "type": "image",
                        "source": {
                            "type": "base64",
                            "media_type": "image/png",
                            "data": original_b64
                        }
                    },
                    {
                        "type": "text",
                        "text": "IMAGE 2 (Generated Recreation):"
                    },
                    {
                        "type": "image",
                        "source": {
                            "type": "base64",
                            "media_type": "image/png",
                            "data": generated_b64
                        }
                    },
                    {
                        "type": "text",
                        "text": prompt
                    }
                ]
            }
        ]
    )

    return message.content[0].text


def refine_description(
    current_description: dict,
    diff_feedback: str,
    use_anthropic: bool = False,
    anthropic_client=None
) -> dict:
    """
    Refine a description based on feedback about differences.

    Args:
        current_description: Current structured description
        diff_feedback: Feedback about what differs
        use_anthropic: If True, use Anthropic API directly
        anthropic_client: Pre-configured Anthropic client

    Returns:
        Updated description dict, or prompt data for external processing
    """
    prompt_template = load_prompt("refinement_prompt")
    prompt = prompt_template.format(
        current_description=json.dumps(current_description, indent=2),
        diff_feedback=diff_feedback
    )

    if use_anthropic:
        return _call_anthropic_refine(prompt, anthropic_client)
    else:
        return {
            "mode": "external",
            "prompt": prompt,
            "current_description": current_description,
            "diff_feedback": diff_feedback
        }


def _call_anthropic_refine(prompt: str, client=None) -> dict:
    """
    Call Anthropic API to refine the description.

    Args:
        prompt: The refinement prompt with current description and feedback
        client: Pre-configured Anthropic client

    Returns:
        Updated description dict
    """
    try:
        import anthropic
    except ImportError:
        raise DescriptorError("anthropic package not installed")

    if client is None:
        client = anthropic.Anthropic()

    message = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=4096,
        messages=[
            {
                "role": "user",
                "content": prompt
            }
        ]
    )

    return _parse_json_response(message.content[0].text)


def save_description(
    description: dict,
    template_name: str,
    output_dir: Optional[Path] = None
) -> tuple[Path, Path]:
    """
    Save a description to JSON and human-readable markdown.

    Args:
        description: Structured description dict
        template_name: Name for the output files (without extension)
        output_dir: Directory to save files (defaults to DESCRIPTION_DIR)

    Returns:
        Tuple of (json_path, markdown_path)
    """
    if output_dir is None:
        output_dir = DESCRIPTION_DIR
    else:
        output_dir = Path(output_dir)

    output_dir.mkdir(parents=True, exist_ok=True)

    # Clean template name for filename
    safe_name = re.sub(r'[^\w\-]', '_', template_name)

    # Save JSON
    json_path = output_dir / f"{safe_name}.json"
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(description, f, indent=2, ensure_ascii=False)

    # Save markdown
    markdown_path = output_dir / f"{safe_name}.md"
    markdown_content = _description_to_markdown(description, template_name)
    with open(markdown_path, 'w', encoding='utf-8') as f:
        f.write(markdown_content)

    logger.info(f"Description saved: {json_path}")
    logger.info(f"Markdown saved: {markdown_path}")

    return json_path, markdown_path


def _description_to_markdown(description: dict, template_name: str) -> str:
    """
    Convert a structured description to human-readable markdown.

    Args:
        description: Structured description dict
        template_name: Name of the template

    Returns:
        Markdown formatted string
    """
    md = [f"# Template: {template_name}\n"]

    # Slide dimensions
    dims = description.get("slide_dimensions", {})
    md.append("## Slide Dimensions\n")
    md.append(f"- Width: {dims.get('width_inches', 'N/A')} inches")
    md.append(f"- Height: {dims.get('height_inches', 'N/A')} inches")
    md.append(f"- Aspect Ratio: {dims.get('aspect_ratio', 'N/A')}\n")

    # Background
    bg = description.get("background", {})
    md.append("## Background\n")
    md.append(f"- Type: {bg.get('type', 'N/A')}")
    if bg.get("color"):
        md.append(f"- Color: {bg['color']}")
    if bg.get("gradient_start"):
        md.append(f"- Gradient: {bg['gradient_start']} to {bg['gradient_end']}")
    md.append("")

    # Elements
    elements = description.get("elements", [])
    md.append(f"## Elements ({len(elements)} total)\n")

    for i, elem in enumerate(elements, 1):
        md.append(f"### Element {i}: {elem.get('id', 'unnamed')}\n")
        md.append(f"- Type: {elem.get('type', 'N/A')}")

        pos = elem.get("position", {})
        md.append(f"- Position: ({pos.get('left_inches', 0)}\", {pos.get('top_inches', 0)}\")")
        md.append(f"- Size: {pos.get('width_inches', 0)}\" x {pos.get('height_inches', 0)}\"")

        text_props = elem.get("text_properties", {})
        if text_props:
            md.append(f"- Font: {text_props.get('font_family', 'N/A')} {text_props.get('font_size_pt', 'N/A')}pt")
            md.append(f"- Text Color: {text_props.get('font_color', 'N/A')}")
            md.append(f"- Alignment: {text_props.get('alignment', 'N/A')}")
            if text_props.get("placeholder_text"):
                md.append(f"- Placeholder: \"{text_props['placeholder_text']}\"")

        shape_props = elem.get("shape_properties", {})
        if shape_props:
            if shape_props.get("fill_color"):
                md.append(f"- Fill: {shape_props['fill_color']}")
            if shape_props.get("shape_type"):
                md.append(f"- Shape: {shape_props['shape_type']}")

        md.append("")

    # Color palette
    palette = description.get("color_palette", [])
    if palette:
        md.append("## Color Palette\n")
        for color in palette:
            md.append(f"- {color}")
        md.append("")

    # Design notes
    notes = description.get("design_notes", "")
    if notes:
        md.append("## Design Notes\n")
        md.append(notes)
        md.append("")

    return "\n".join(md)


def load_description(template_name: str, description_dir: Optional[Path] = None) -> dict:
    """
    Load a saved description from JSON.

    Args:
        template_name: Name of the template (matches the JSON filename)
        description_dir: Directory to load from (defaults to DESCRIPTION_DIR)

    Returns:
        Description dict

    Raises:
        DescriptorError: If file not found or invalid JSON
    """
    if description_dir is None:
        description_dir = DESCRIPTION_DIR
    else:
        description_dir = Path(description_dir)

    # Try with and without extension
    json_path = description_dir / f"{template_name}.json"
    if not json_path.exists():
        json_path = description_dir / template_name
        if not json_path.exists():
            raise DescriptorError(f"Description file not found: {template_name}")

    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except json.JSONDecodeError as e:
        raise DescriptorError(f"Invalid JSON in {json_path}: {e}")


if __name__ == "__main__":
    # Test loading prompts
    print("Testing prompt loading...")
    try:
        desc_prompt = load_prompt("description_prompt")
        print(f"Description prompt loaded ({len(desc_prompt)} chars)")

        comp_prompt = load_prompt("comparison_prompt")
        print(f"Comparison prompt loaded ({len(comp_prompt)} chars)")

        refine_prompt = load_prompt("refinement_prompt")
        print(f"Refinement prompt loaded ({len(refine_prompt)} chars)")

        print("\nAll prompts loaded successfully!")
    except Exception as e:
        print(f"Error: {e}")
