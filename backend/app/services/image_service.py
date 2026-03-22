from __future__ import annotations

import base64
from pathlib import Path
from typing import Optional

from openai import OpenAI

from app.core.config import settings


def _safe_filename(name: str) -> str:
    cleaned = "".join(c if c.isalnum() or c in ("-", "_") else "_" for c in name)
    cleaned = cleaned.strip("_")
    return cleaned[:80] or "slide"


def _build_prompt(title: str, visual_idea: str) -> str:
    return f"""
Create a minimalist presentation illustration.

Slide topic: {title}
Visual direction: {visual_idea}

Requirements:
- modern presentation style
- soft pastel palette
- minimalist
- organic shapes
- clean composition
- no text
- no letters
- no logos
- no watermark
- no UI mockups
- suitable for a business or academic slide
""".strip()


def generate_slide_image(title: str, visual_idea: str, slide_number: int) -> Optional[str]:
    api_key = getattr(settings, "openai_api_key", None) or getattr(settings, "OPENAI_API_KEY", None)
    if not api_key:
        print("[image generation] OPENAI_API_KEY not set in settings")
        return None

    out_dir = Path(settings.generated_dir) / "images"
    out_dir.mkdir(parents=True, exist_ok=True)

    filename = f"slide_{slide_number:02d}_{_safe_filename(title)}.png"
    out_path = out_dir / filename

    try:
        client = OpenAI(api_key=api_key)

        result = client.images.generate(
            model=getattr(settings, "openai_image_model", "gpt-image-1"),
            prompt=_build_prompt(title, visual_idea),
            size=getattr(settings, "openai_image_size", "1024x1024"),
        )

        if not result or not getattr(result, "data", None):
            print(f"[image generation] empty response for slide {slide_number}")
            return None

        first = result.data[0]
        image_b64 = getattr(first, "b64_json", None)

        if not image_b64:
            print(f"[image generation] no b64_json returned for slide {slide_number}")
            print(f"[image generation] raw first item: {first}")
            return None

        out_path.write_bytes(base64.b64decode(image_b64))
        print(f"[image generation] saved: {out_path}")
        return str(out_path)

    except Exception as e:
        print(f"[image generation] failed for slide {slide_number}: {e}")
        return None