"""
Image processing utilities for screenshots and OCR.

Handles:
- mss screenshot → PIL conversion
- OCR image preprocessing (grayscale, threshold, invert, scale)
- Image compression and base64 encoding
- Tesseract availability checking
"""

import base64
import io
import logging
from typing import Any

import numpy as np
from PIL import Image, ImageFilter, ImageOps

from ..config import PreprocessMode

logger = logging.getLogger("win32-mcp")


# ---------------------------------------------------------------------------
# Screenshot Conversion
# ---------------------------------------------------------------------------


def mss_to_pil(screenshot: Any) -> Image.Image:
    """Convert an mss ScreenShot object to a PIL RGB Image."""
    return Image.frombytes("RGB", screenshot.size, screenshot.bgra, "raw", "BGRX")


# ---------------------------------------------------------------------------
# OCR Preprocessing
# ---------------------------------------------------------------------------


def preprocess_for_ocr(
    img: Image.Image,
    mode: PreprocessMode = "auto",
    scale_small: bool = True,
    min_dimension: int = 500,
    upscale_factor: int = 2,
) -> tuple[Image.Image, float]:
    """Preprocess an image for better OCR accuracy.

    Returns:
        (processed_image, scale_factor) — scale_factor is needed to map
        Tesseract coordinates back to original image space.
    """
    if mode == "none":
        return img, 1.0

    scale_factor = 1.0

    # Detect background brightness for auto mode
    if mode == "auto":
        gray_arr = np.array(img.convert("L"), dtype=np.float64)
        mean_brightness = float(np.mean(gray_arr))
        std_brightness = float(np.std(gray_arr))
        if mean_brightness < 100:
            mode = "dark_bg"
        elif mean_brightness > 200 and std_brightness < 30:
            # Very bright *and* low contrast — text is washed out
            mode = "high_contrast"
        else:
            mode = "light_bg"

    # Convert to grayscale
    gray = img.convert("L")

    # Invert for dark backgrounds (light text on dark bg)
    if mode == "dark_bg":
        gray = ImageOps.invert(gray)

    # Scale up small images for better character recognition
    if scale_small:
        w, h = gray.size
        if w < min_dimension or h < min_dimension:
            factor = max(upscale_factor, min(4, min_dimension // min(w, h) + 1))
            gray = gray.resize((w * factor, h * factor), Image.Resampling.LANCZOS)
            scale_factor = float(factor)

    # Sharpen to improve edge clarity
    gray = gray.filter(ImageFilter.SHARPEN)

    # Adaptive thresholding via numpy
    arr = np.array(gray, dtype=np.float64)

    if mode == "high_contrast":
        # Use Otsu-style threshold
        threshold = float(np.mean(arr))
        binary = ((arr > threshold) * 255).astype(np.uint8)
    else:
        # Gentle thresholding — preserve grey tones for Tesseract's own binarization
        # Stretch contrast to 0-255 range
        lo, hi = float(np.percentile(arr, 2)), float(np.percentile(arr, 98))
        if hi - lo < 10:
            hi = lo + 10
        stretched = np.clip((arr - lo) / (hi - lo) * 255, 0, 255).astype(np.uint8)
        binary = stretched

    result = Image.fromarray(binary, mode="L")
    return result, scale_factor


# ---------------------------------------------------------------------------
# Image Encoding
# ---------------------------------------------------------------------------


def image_to_base64(
    img: Image.Image,
    fmt: str = "png",
    quality: int = 85,
    scale: float = 1.0,
) -> tuple[str, str, int]:
    """Encode a PIL Image to base64.

    Args:
        img: Source image.
        fmt: Output format — "png", "jpeg", or "webp".
        quality: Compression quality (1-100) for jpeg/webp.
        scale: Resize factor (0.1-1.0). 0.5 = half size.

    Returns:
        (base64_data, mime_type, size_bytes)
    """
    if 0 < scale < 1.0:
        new_w = max(1, int(img.width * scale))
        new_h = max(1, int(img.height * scale))
        img = img.resize((new_w, new_h), Image.Resampling.LANCZOS)

    buf = io.BytesIO()
    fmt_lower = fmt.lower().strip()

    if fmt_lower in ("jpg", "jpeg"):
        img = img.convert("RGB")  # drop alpha
        img.save(buf, format="JPEG", quality=quality, optimize=True)
        mime = "image/jpeg"
    elif fmt_lower == "webp":
        img.save(buf, format="WEBP", quality=quality)
        mime = "image/webp"
    else:
        img.save(buf, format="PNG", optimize=True)
        mime = "image/png"

    data = buf.getvalue()
    return base64.b64encode(data).decode("ascii"), mime, len(data)


# ---------------------------------------------------------------------------
# Tesseract Detection
# ---------------------------------------------------------------------------


def check_tesseract() -> tuple[bool, str]:
    """Check if Tesseract OCR is installed and reachable.

    Returns:
        (is_available, version_string_or_error_message)
    """
    try:
        import pytesseract

        version = pytesseract.get_tesseract_version()
        return True, str(version)
    except Exception as exc:
        name = type(exc).__name__
        if "NotFound" in name or "not installed" in str(exc).lower():
            return False, (
                "Tesseract OCR is not installed. "
                "Download from https://github.com/UB-Mannheim/tesseract/wiki "
                "and ensure it is on PATH."
            )
        return False, f"Tesseract error: {exc}"


# ---------------------------------------------------------------------------
# Screenshot Diffing
# ---------------------------------------------------------------------------


def compute_image_diff(
    img_a: Image.Image,
    img_b: Image.Image,
    pixel_threshold: int = 30,
) -> dict[str, Any]:
    """Compare two images and return similarity metrics.

    Args:
        img_a: First image.
        img_b: Second image.
        pixel_threshold: Per-channel difference (0-255) below which pixels
            are considered identical.

    Returns:
        dict with similarity, change_percent, dimensions.
    """
    # Resize to match if needed
    if img_a.size != img_b.size:
        img_b = img_b.resize(img_a.size, Image.Resampling.LANCZOS)

    arr_a = np.array(img_a.convert("RGB"), dtype=np.float64)
    arr_b = np.array(img_b.convert("RGB"), dtype=np.float64)

    # Mean-squared error across all channels
    mse = float(np.mean((arr_a - arr_b) ** 2))
    similarity = 1.0 - mse / (255.0**2)

    # Percentage of meaningfully changed pixels
    diff = np.abs(arr_a - arr_b).mean(axis=2)
    changed_count = int(np.sum(diff > pixel_threshold))
    total_pixels = diff.size
    change_pct = (changed_count / total_pixels) * 100.0

    return {
        "similarity": round(similarity, 6),
        "change_percent": round(change_pct, 4),
        "mse": round(mse, 2),
        "changed_pixels": changed_count,
        "total_pixels": total_pixels,
        "size_a": list(img_a.size),
        "size_b": list(img_b.size),
    }
