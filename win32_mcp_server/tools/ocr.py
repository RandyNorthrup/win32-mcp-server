"""
OCR (Optical Character Recognition) tools.

Tools:
  - ocr_screen              Plain text from full screen
  - ocr_region              Plain text from screen region
  - ocr_window              Plain text from a specific window
  - ocr_screen_structured   Text with bounding boxes from full screen
  - ocr_region_structured   Text with bounding boxes from region
"""

import asyncio
import json
import logging

from PIL import Image
import pytesseract
from mcp.types import TextContent

from ..registry import registry
from ..config import config, PreprocessMode, VALID_PREPROCESS_MODES
from ..utils.errors import ToolError
from ..utils.imaging import (
    preprocess_for_ocr,
    check_tesseract,
)
from .capture import capture_screen_impl, capture_region_impl, capture_window_impl

logger = logging.getLogger("win32-mcp")


# ===================================================================
# Tesseract Guard
# ===================================================================

def _ensure_tesseract():
    """Raise ToolError if Tesseract is not installed."""
    ok, msg = check_tesseract()
    if not ok:
        raise ToolError(msg, suggestion="Install Tesseract from https://github.com/UB-Mannheim/tesseract/wiki")


def _validate_preprocess(value: str | None) -> PreprocessMode | None:
    """Validate and cast a preprocess string from user input."""
    if value is None:
        return None
    if value not in VALID_PREPROCESS_MODES:
        raise ToolError(
            f"Invalid preprocess mode: '{value}'",
            suggestion=f"Valid modes: {', '.join(sorted(VALID_PREPROCESS_MODES))}",
        )
    return value  # type: ignore[return-value]


# ===================================================================
# Implementation functions
# ===================================================================

async def ocr_plain_impl(
    img: Image.Image,
    lang: str | None = None,
    preprocess: PreprocessMode | None = None,
) -> str:
    """Run OCR on a PIL Image and return plain text."""
    _ensure_tesseract()

    lang = lang or config.ocr.lang
    effective_preprocess: PreprocessMode = preprocess if preprocess is not None else config.ocr.preprocess_mode

    processed, _ = preprocess_for_ocr(
        img,
        mode=effective_preprocess,
        scale_small=config.ocr.scale_small_images,
        min_dimension=config.ocr.min_dimension_for_scaling,
        upscale_factor=config.ocr.upscale_factor,
    )

    def _ocr():
        return pytesseract.image_to_string(processed, lang=lang).strip()

    return await asyncio.to_thread(_ocr)


async def ocr_structured_impl(
    img: Image.Image,
    lang: str | None = None,
    preprocess: PreprocessMode | None = None,
    confidence_threshold: int | None = None,
) -> list[dict]:
    """Run OCR and return structured results with bounding boxes.

    Each result dict has:
      text, x, y, width, height, confidence, line_num, block_num

    Coordinates are in the original (pre-preprocessing) image space.
    """
    _ensure_tesseract()

    lang = lang or config.ocr.lang
    effective_preprocess: PreprocessMode = preprocess if preprocess is not None else config.ocr.preprocess_mode
    conf_threshold = confidence_threshold if confidence_threshold is not None else config.ocr.confidence_threshold

    processed, scale_factor = preprocess_for_ocr(
        img,
        mode=effective_preprocess,
        scale_small=config.ocr.scale_small_images,
        min_dimension=config.ocr.min_dimension_for_scaling,
        upscale_factor=config.ocr.upscale_factor,
    )

    def _ocr():
        return pytesseract.image_to_data(
            processed, lang=lang, output_type=pytesseract.Output.DICT,
        )

    data = await asyncio.to_thread(_ocr)

    results: list[dict] = []
    n = len(data["text"])
    for i in range(n):
        text = data["text"][i].strip()
        if not text:
            continue

        raw_conf = data["conf"][i]
        try:
            conf = int(float(raw_conf))
        except (ValueError, TypeError):
            conf = 0

        if conf < conf_threshold:
            continue

        # Map coordinates back to original image space
        x = int(data["left"][i] / scale_factor)
        y = int(data["top"][i] / scale_factor)
        w = int(data["width"][i] / scale_factor)
        h = int(data["height"][i] / scale_factor)

        results.append({
            "text": text,
            "x": x,
            "y": y,
            "width": w,
            "height": h,
            "confidence": conf,
            "line_num": data["line_num"][i],
            "block_num": data["block_num"][i],
            "word_num": data["word_num"][i],
        })

    return results


# ===================================================================
# MCP Tool Handlers
# ===================================================================

@registry.register("ocr_screen", "Extract text from the full screen using OCR", {
    "type": "object",
    "properties": {
        "lang": {
            "type": "string",
            "description": "Tesseract language code (default: eng). Use 'eng+fra' for multiple.",
        },
        "preprocess": {
            "type": "string",
            "enum": ["auto", "light_bg", "dark_bg", "high_contrast", "none"],
            "description": "Image preprocessing mode (default: auto)",
        },
    },
})
async def handle_ocr_screen(arguments: dict):
    img = await capture_screen_impl()
    text = await ocr_plain_impl(
        img,
        lang=arguments.get("lang"),
        preprocess=_validate_preprocess(arguments.get("preprocess")),
    )
    return [TextContent(type="text", text=f"OCR Result ({img.width}x{img.height}):\n{text}")]


@registry.register("ocr_region", "Extract text from a screen region using OCR", {
    "type": "object",
    "properties": {
        "x": {"type": "number", "description": "Left edge of region"},
        "y": {"type": "number", "description": "Top edge of region"},
        "width": {"type": "number", "description": "Region width"},
        "height": {"type": "number", "description": "Region height"},
        "lang": {"type": "string"},
        "preprocess": {
            "type": "string",
            "enum": ["auto", "light_bg", "dark_bg", "high_contrast", "none"],
        },
    },
    "required": ["x", "y", "width", "height"],
})
async def handle_ocr_region(arguments: dict):
    x, y = int(arguments["x"]), int(arguments["y"])
    w, h = int(arguments["width"]), int(arguments["height"])

    img = await capture_region_impl(x, y, w, h)
    text = await ocr_plain_impl(
        img,
        lang=arguments.get("lang"),
        preprocess=_validate_preprocess(arguments.get("preprocess")),
    )
    return [TextContent(type="text", text=f"OCR Region ({x},{y} {w}x{h}):\n{text}")]


@registry.register("ocr_window", "Extract text from a specific window using OCR", {
    "type": "object",
    "properties": {
        "window_title": {"type": "string", "description": "Full or partial window title"},
        "lang": {"type": "string"},
        "preprocess": {
            "type": "string",
            "enum": ["auto", "light_bg", "dark_bg", "high_contrast", "none"],
        },
    },
    "required": ["window_title"],
})
async def handle_ocr_window(arguments: dict):
    img, win_info = await capture_window_impl(arguments["window_title"])
    text = await ocr_plain_impl(
        img,
        lang=arguments.get("lang"),
        preprocess=_validate_preprocess(arguments.get("preprocess")),
    )
    return [TextContent(
        type="text",
        text=f"OCR Window '{win_info['title']}' ({win_info['width']}x{win_info['height']}):\n{text}",
    )]


@registry.register("ocr_screen_structured", "Extract text with bounding boxes from full screen", {
    "type": "object",
    "properties": {
        "lang": {"type": "string"},
        "preprocess": {
            "type": "string",
            "enum": ["auto", "light_bg", "dark_bg", "high_contrast", "none"],
        },
        "confidence_threshold": {
            "type": "number",
            "description": "Minimum confidence 0-100 to include (default: 60)",
        },
    },
})
async def handle_ocr_screen_structured(arguments: dict):
    img = await capture_screen_impl()
    results = await ocr_structured_impl(
        img,
        lang=arguments.get("lang"),
        preprocess=_validate_preprocess(arguments.get("preprocess")),
        confidence_threshold=arguments.get("confidence_threshold"),
    )
    return [TextContent(type="text", text=json.dumps({
        "screen_size": f"{img.width}x{img.height}",
        "element_count": len(results),
        "elements": results,
    }, indent=2))]


@registry.register("ocr_region_structured", "Extract text with bounding boxes from a region", {
    "type": "object",
    "properties": {
        "x": {"type": "number"},
        "y": {"type": "number"},
        "width": {"type": "number"},
        "height": {"type": "number"},
        "lang": {"type": "string"},
        "preprocess": {
            "type": "string",
            "enum": ["auto", "light_bg", "dark_bg", "high_contrast", "none"],
        },
        "confidence_threshold": {"type": "number"},
    },
    "required": ["x", "y", "width", "height"],
})
async def handle_ocr_region_structured(arguments: dict):
    x, y = int(arguments["x"]), int(arguments["y"])
    w, h = int(arguments["width"]), int(arguments["height"])

    img = await capture_region_impl(x, y, w, h)
    results = await ocr_structured_impl(
        img,
        lang=arguments.get("lang"),
        preprocess=_validate_preprocess(arguments.get("preprocess")),
        confidence_threshold=arguments.get("confidence_threshold"),
    )

    # Offset coordinates to screen space
    for r in results:
        r["screen_x"] = r["x"] + x
        r["screen_y"] = r["y"] + y

    return [TextContent(type="text", text=json.dumps({
        "region": {"x": x, "y": y, "width": w, "height": h},
        "element_count": len(results),
        "elements": results,
    }, indent=2))]
