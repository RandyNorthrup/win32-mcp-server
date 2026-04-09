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
import hashlib
import json
import logging
import time
from typing import Any

import pytesseract
from mcp.types import TextContent
from PIL import Image

from ..config import VALID_PREPROCESS_MODES, PreprocessMode, config
from ..registry import registry
from ..utils.coordinates import validate_region
from ..utils.errors import ToolError
from ..utils.imaging import (
    check_tesseract,
    preprocess_for_ocr,
)
from .capture import capture_region_impl, capture_screen_impl, capture_window_impl

logger = logging.getLogger("win32-mcp")


# ===================================================================
# OCR Result Cache
# ===================================================================

# Cache keyed by (image_hash, lang, preprocess, result_type).
# Entries expire after `_OCR_CACHE_TTL` seconds.
_OCR_CACHE_TTL: float = 2.0
_ocr_cache: dict[tuple[str, str, str, str], tuple[float, Any]] = {}


def _image_hash(img: Image.Image) -> str:
    """Fast perceptual hash: downsample to 32x32 grayscale, then SHA-256."""
    small = img.resize((32, 32), Image.Resampling.NEAREST).convert("L")
    return hashlib.sha256(small.tobytes()).hexdigest()


def _cache_get(key: tuple[str, str, str, str]) -> Any | None:
    """Return cached value if present and not expired, else None."""
    entry = _ocr_cache.get(key)
    if entry is None:
        return None
    ts, value = entry
    if time.monotonic() - ts > _OCR_CACHE_TTL:
        del _ocr_cache[key]
        return None
    logger.debug("OCR cache hit for %s", key[:3])
    return value


def _cache_put(key: tuple[str, str, str, str], value: Any) -> None:
    """Store a value in the cache and evict expired entries."""
    now = time.monotonic()
    # Evict stale entries (keep cache bounded)
    stale = [k for k, (ts, _) in _ocr_cache.items() if now - ts > _OCR_CACHE_TTL]
    for k in stale:
        del _ocr_cache[k]
    _ocr_cache[key] = (now, value)


# ===================================================================
# Tesseract Guard
# ===================================================================


def _ensure_tesseract() -> None:
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


def _is_mixed_brightness(img: Image.Image) -> bool:
    """Detect if an image has mixed bright/dark regions (e.g. dark-themed
    apps with highlights, toolbars, or dialogs). Global inversion fails on
    these images, so a dual-pass OCR is needed."""
    import numpy as np

    gray = np.array(img.convert("L"), dtype=np.float64)
    mean_val = float(np.mean(gray))
    std_val = float(np.std(gray))
    # Dark overall (mean < 130) but high variance (std > 40) = mixed content
    # e.g. dark editor with bright toolbar/status bar/highlights
    return mean_val < 130 and std_val > 40


async def ocr_plain_impl(
    img: Image.Image,
    lang: str | None = None,
    preprocess: PreprocessMode | None = None,
) -> str:
    """Run OCR on a PIL Image and return plain text.

    When preprocess is 'auto' and the image has mixed brightness (e.g.
    dark-themed windows), runs OCR twice (light_bg + dark_bg) and returns
    the longer result for better coverage.

    Results are cached for _OCR_CACHE_TTL seconds keyed by a perceptual
    hash of the image, so repeated calls on an unchanged screen are free.
    """
    _ensure_tesseract()

    lang = lang or config.ocr.lang
    effective_preprocess: PreprocessMode = preprocess if preprocess is not None else config.ocr.preprocess_mode

    # Check cache
    img_h = _image_hash(img)
    cache_key = (img_h, lang, effective_preprocess, "plain")
    cached = _cache_get(cache_key)
    if cached is not None:
        return str(cached)

    # Dual-pass for auto mode on mixed-brightness images
    if effective_preprocess == "auto" and _is_mixed_brightness(img):
        results = []
        for mode in ("light_bg", "dark_bg"):
            processed, _ = preprocess_for_ocr(
                img,
                mode=mode,
                scale_small=config.ocr.scale_small_images,
                min_dimension=config.ocr.min_dimension_for_scaling,
                upscale_factor=config.ocr.upscale_factor,
            )
            pass_text = await asyncio.to_thread(
                pytesseract.image_to_string,
                processed,
                lang=lang,
            )
            results.append(pass_text.strip())
        # Return the longer result -- it captured more text
        result = str(max(results, key=len))
        _cache_put(cache_key, result)
        return result

    processed, _ = preprocess_for_ocr(
        img,
        mode=effective_preprocess,
        scale_small=config.ocr.scale_small_images,
        min_dimension=config.ocr.min_dimension_for_scaling,
        upscale_factor=config.ocr.upscale_factor,
    )

    def _ocr() -> str:
        result: str = pytesseract.image_to_string(processed, lang=lang).strip()
        return result

    text: str = await asyncio.to_thread(_ocr)
    _cache_put(cache_key, text)
    return text


async def ocr_structured_impl(
    img: Image.Image,
    lang: str | None = None,
    preprocess: PreprocessMode | None = None,
    confidence_threshold: int | None = None,
) -> list[dict[str, Any]]:
    """Run OCR and return structured results with bounding boxes.

    Each result dict has:
      text, x, y, width, height, confidence, line_num, block_num

    Coordinates are in the original (pre-preprocessing) image space.

    When preprocess is 'auto' and the image has mixed brightness,
    runs OCR twice (light_bg + dark_bg) and merges results, keeping
    the higher-confidence word at each approximate position.
    """
    _ensure_tesseract()

    lang = lang or config.ocr.lang
    effective_preprocess: PreprocessMode = preprocess if preprocess is not None else config.ocr.preprocess_mode
    conf_threshold = confidence_threshold if confidence_threshold is not None else config.ocr.confidence_threshold

    # Check cache
    img_h = _image_hash(img)
    cache_key = (img_h, lang, effective_preprocess, f"structured:{conf_threshold}")
    cached = _cache_get(cache_key)
    if cached is not None:
        return list(cached)

    # Dual-pass for auto mode on mixed-brightness images
    if effective_preprocess == "auto" and _is_mixed_brightness(img):
        all_results = []
        for mode in ("light_bg", "dark_bg"):
            results = await _ocr_structured_single_pass(
                img,
                lang=lang,
                preprocess=mode,
                confidence_threshold=conf_threshold,
            )
            all_results.append(results)
        merged = _merge_ocr_results(all_results[0], all_results[1])
        _cache_put(cache_key, merged)
        return merged

    result = await _ocr_structured_single_pass(
        img,
        lang=lang,
        preprocess=effective_preprocess,
        confidence_threshold=conf_threshold,
    )
    _cache_put(cache_key, result)
    return result


async def _ocr_structured_single_pass(
    img: Image.Image,
    lang: str,
    preprocess: PreprocessMode,
    confidence_threshold: int,
) -> list[dict[str, Any]]:
    """Single-pass structured OCR with a specific preprocessing mode."""
    processed, scale_factor = preprocess_for_ocr(
        img,
        mode=preprocess,
        scale_small=config.ocr.scale_small_images,
        min_dimension=config.ocr.min_dimension_for_scaling,
        upscale_factor=config.ocr.upscale_factor,
    )

    def _ocr() -> dict[str, Any]:
        result: dict[str, Any] = pytesseract.image_to_data(
            processed,
            lang=lang,
            output_type=pytesseract.Output.DICT,
        )
        return result

    data = await asyncio.to_thread(_ocr)

    results: list[dict[str, Any]] = []
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

        if conf < confidence_threshold:
            continue

        # Map coordinates back to original image space
        x = int(data["left"][i] / scale_factor)
        y = int(data["top"][i] / scale_factor)
        w = int(data["width"][i] / scale_factor)
        h = int(data["height"][i] / scale_factor)

        results.append(
            {
                "text": text,
                "x": x,
                "y": y,
                "width": w,
                "height": h,
                "confidence": conf,
                "line_num": data["line_num"][i],
                "block_num": data["block_num"][i],
                "word_num": data["word_num"][i],
            }
        )

    return results


def _merge_ocr_results(
    results_a: list[dict[str, Any]],
    results_b: list[dict[str, Any]],
    overlap_threshold: int = 20,
) -> list[dict[str, Any]]:
    """Merge two OCR result sets, deduplicating overlapping words.

    For words at approximately the same position (within overlap_threshold
    pixels), keep the one with higher confidence. Words unique to either
    set are included as-is.
    """
    merged = list(results_a)
    used = [False] * len(results_a)

    for rb in results_b:
        best_overlap_idx = -1
        best_overlap_score = -1

        rb_cx = rb["x"] + rb["width"] // 2
        rb_cy = rb["y"] + rb["height"] // 2

        for i, ra in enumerate(results_a):
            ra_cx = ra["x"] + ra["width"] // 2
            ra_cy = ra["y"] + ra["height"] // 2

            dx = abs(ra_cx - rb_cx)
            dy = abs(ra_cy - rb_cy)

            if dx < overlap_threshold and dy < overlap_threshold and ra["confidence"] > best_overlap_score:
                best_overlap_idx = i
                best_overlap_score = ra["confidence"]

        if best_overlap_idx >= 0:
            # Overlapping word found — keep the higher confidence version
            used[best_overlap_idx] = True
            if rb["confidence"] > merged[best_overlap_idx]["confidence"]:
                merged[best_overlap_idx] = rb
        else:
            # Unique word from results_b — add it
            merged.append(rb)

    return merged


# ===================================================================
# MCP Tool Handlers
# ===================================================================


@registry.register(
    "ocr_screen",
    "Extract text from the full screen using OCR",
    {
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
    },
)
async def handle_ocr_screen(arguments: dict[str, Any]) -> list[TextContent]:
    img = await capture_screen_impl()
    text = await ocr_plain_impl(
        img,
        lang=arguments.get("lang"),
        preprocess=_validate_preprocess(arguments.get("preprocess")),
    )
    return [TextContent(type="text", text=f"OCR Result ({img.width}x{img.height}):\n{text}")]


@registry.register(
    "ocr_region",
    "Extract text from a screen region using OCR",
    {
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
    },
)
async def handle_ocr_region(arguments: dict[str, Any]) -> list[TextContent]:
    x, y = int(arguments["x"]), int(arguments["y"])
    w, h = int(arguments["width"]), int(arguments["height"])

    if config.validate_coordinates:
        x, y, w, h = validate_region(x, y, w, h)

    img = await capture_region_impl(x, y, w, h)
    text = await ocr_plain_impl(
        img,
        lang=arguments.get("lang"),
        preprocess=_validate_preprocess(arguments.get("preprocess")),
    )
    return [TextContent(type="text", text=f"OCR Region ({x},{y} {w}x{h}):\n{text}")]


@registry.register(
    "ocr_window",
    "Extract text from a specific window using OCR",
    {
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
    },
)
async def handle_ocr_window(arguments: dict[str, Any]) -> list[TextContent]:
    img, win_info = await capture_window_impl(arguments["window_title"])
    text = await ocr_plain_impl(
        img,
        lang=arguments.get("lang"),
        preprocess=_validate_preprocess(arguments.get("preprocess")),
    )
    return [
        TextContent(
            type="text",
            text=f"OCR Window '{win_info['title']}' ({win_info['width']}x{win_info['height']}):\n{text}",
        )
    ]


@registry.register(
    "ocr_screen_structured",
    "Extract text with bounding boxes from full screen",
    {
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
    },
)
async def handle_ocr_screen_structured(arguments: dict[str, Any]) -> list[TextContent]:
    img = await capture_screen_impl()
    results = await ocr_structured_impl(
        img,
        lang=arguments.get("lang"),
        preprocess=_validate_preprocess(arguments.get("preprocess")),
        confidence_threshold=arguments.get("confidence_threshold"),
    )
    return [
        TextContent(
            type="text",
            text=json.dumps(
                {
                    "screen_size": f"{img.width}x{img.height}",
                    "element_count": len(results),
                    "elements": results,
                },
                indent=2,
            ),
        )
    ]


@registry.register(
    "ocr_region_structured",
    "Extract text with bounding boxes from a region",
    {
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
    },
)
async def handle_ocr_region_structured(arguments: dict[str, Any]) -> list[TextContent]:
    x, y = int(arguments["x"]), int(arguments["y"])
    w, h = int(arguments["width"]), int(arguments["height"])

    if config.validate_coordinates:
        x, y, w, h = validate_region(x, y, w, h)

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

    return [
        TextContent(
            type="text",
            text=json.dumps(
                {
                    "region": {"x": x, "y": y, "width": w, "height": h},
                    "element_count": len(results),
                    "elements": results,
                },
                indent=2,
            ),
        )
    ]
