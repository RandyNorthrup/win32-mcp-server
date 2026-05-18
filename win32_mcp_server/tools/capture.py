"""
Screenshot and display tools.

Tools:
  - capture_screen      Full screen screenshot with compression options
  - capture_window      Window-specific capture with proper waiting
  - capture_monitor     Capture a specific monitor by index
  - list_monitors       List all connected monitors
  - get_pixel_color     Get color of pixel at coordinates
  - compare_screenshots Compare current screen with reference image
"""

import asyncio
import base64
import io
import json
import logging
import threading
from typing import Any

from mcp.types import ImageContent, TextContent
from mss import mss
from PIL import Image

from ..config import VALID_CAPTURE_FORMATS, config
from ..registry import registry
from ..utils.args import get_bool, get_enum, get_float, get_int, get_str
from ..utils.coordinates import clamp_rect_to_virtual_screen, get_all_monitors, validate_coordinates, validate_region
from ..utils.errors import ToolError
from ..utils.imaging import compute_image_diff, image_to_base64, mss_to_pil
from ..utils.window_match import find_window_strict

logger = logging.getLogger("win32-mcp")


# ===================================================================
# mss per-thread instances — GDI device contexts are thread-affine
# ===================================================================

_mss_local = threading.local()


def _get_mss() -> Any:
    """Return a thread-local mss instance (GDI DCs are thread-affine on Windows)."""
    inst = getattr(_mss_local, "instance", None)
    if inst is None:
        inst = mss()
        _mss_local.instance = inst
    return inst


def _capture_options(arguments: dict[str, Any]) -> tuple[str, int, float]:
    fmt = get_enum(arguments, "format", VALID_CAPTURE_FORMATS, default=config.capture.default_format)
    quality = get_int(arguments, "quality", default=config.capture.default_quality, min_value=1, max_value=100)
    scale = get_float(arguments, "scale", default=config.capture.default_scale, min_value=0.1, max_value=1.0)
    return fmt, quality, scale


def _optional_region(arguments: dict[str, Any], key: str = "region") -> tuple[int, int, int, int] | None:
    region = arguments.get(key)
    if region is None:
        return None
    if not isinstance(region, dict):
        raise ToolError(f"Argument '{key}' must be an object")
    return (
        get_int(region, "x", required=True),
        get_int(region, "y", required=True),
        get_int(region, "width", required=True, min_value=1),
        get_int(region, "height", required=True, min_value=1),
    )


# ===================================================================
# Implementation functions (called by smart tools and handlers)
# ===================================================================


async def capture_screen_impl(
    monitor_index: int = 1,
) -> Image.Image:
    """Capture a monitor and return a PIL Image.

    Args:
        monitor_index: 1-based monitor index (1 = primary). Use 0 for
            the virtual screen spanning all monitors.
    """

    def _grab() -> Image.Image:
        sct = _get_mss()
        if monitor_index < 0 or monitor_index >= len(sct.monitors):
            raise ToolError(
                f"Monitor index {monitor_index} out of range (0\u2013{len(sct.monitors) - 1})",
                suggestion="Use list_monitors to see available monitors",
            )
        screenshot = sct.grab(sct.monitors[monitor_index])
        return mss_to_pil(screenshot)

    return await asyncio.to_thread(_grab)


async def capture_region_impl(x: int, y: int, width: int, height: int) -> Image.Image:
    """Capture a screen region and return a PIL Image."""

    def _grab() -> Image.Image:
        sct = _get_mss()
        region = {"top": y, "left": x, "width": width, "height": height}
        screenshot = sct.grab(region)
        return mss_to_pil(screenshot)

    return await asyncio.to_thread(_grab)


async def capture_window_impl(window_title: str, activate: bool = True) -> tuple[Image.Image, dict[str, Any]]:
    """Capture a window's contents.

    Returns:
        (pil_image, window_info_dict)
    """
    import ctypes

    win = find_window_strict(window_title)

    # Restore if minimized
    if win.isMinimized:
        if activate:
            try:
                win.restore()
            except Exception as exc:
                logger.debug("Could not restore window '%s': %s", win.title, exc)
            await asyncio.sleep(0.5)
        else:
            raise ToolError(
                f"Window '{win.title}' is minimized",
                suggestion="Restore it first or capture with activate=true",
            )

    # Bring to foreground
    if activate:
        try:
            win.activate()
        except Exception as exc:
            logger.debug("Could not activate window '%s': %s", win.title, exc)
        await asyncio.sleep(0.2)

        # Poll until window is foreground (up to 1.5s)
        hwnd = getattr(win, "_hWnd", None)
        if hwnd:
            for _ in range(15):
                try:
                    fg = ctypes.windll.user32.GetForegroundWindow()
                    if fg == hwnd:
                        break
                except Exception as exc:
                    logger.debug("Foreground poll interrupted: %s", exc)
                    break
                await asyncio.sleep(0.1)

    # Clamp to virtual screen — handles multi-monitor with negative coordinates
    left, top, width, height = clamp_rect_to_virtual_screen(
        win.left,
        win.top,
        win.width,
        win.height,
    )
    if width <= 0 or height <= 0:
        raise ToolError(
            f"Window '{win.title}' is entirely off-screen "
            f"(position: {win.left},{win.top} size: {win.width}x{win.height})",
            suggestion="Move the window on-screen first using move_window",
        )

    img = await capture_region_impl(left, top, width, height)

    info = {
        "title": win.title,
        "x": win.left,
        "y": win.top,
        "width": win.width,
        "height": win.height,
        "capture_x": left,
        "capture_y": top,
        "capture_width": width,
        "capture_height": height,
    }
    return img, info


# ===================================================================
# MCP Tool Handlers
# ===================================================================


@registry.register(
    "capture_screen",
    "Capture full screen screenshot with optional compression",
    {
        "type": "object",
        "properties": {
            "format": {
                "type": "string",
                "enum": ["png", "jpeg", "webp"],
                "description": "Image format (default: png). jpeg/webp are smaller.",
            },
            "quality": {
                "type": "number",
                "minimum": 1,
                "maximum": 100,
                "description": "Compression quality 1-100 for jpeg/webp (default: 85)",
            },
            "scale": {
                "type": "number",
                "minimum": 0.1,
                "maximum": 1.0,
                "description": "Resize factor 0.1-1.0 (default: 1.0). 0.5 = half size, 4x smaller file.",
            },
        },
    },
)
async def handle_capture_screen(arguments: dict[str, Any]) -> list[TextContent | ImageContent]:
    fmt, quality, scale = _capture_options(arguments)

    img = await capture_screen_impl()
    data, mime, size = image_to_base64(img, fmt=fmt, quality=quality, scale=scale)

    return [
        TextContent(
            type="text",
            text=json.dumps(
                {
                    "screen_size": f"{img.width}x{img.height}",
                    "format": fmt,
                    "file_size_kb": round(size / 1024, 1),
                }
            ),
        ),
        ImageContent(type="image", data=data, mimeType=mime),
    ]


@registry.register(
    "capture_window",
    "Capture a specific window by title (fuzzy match supported)",
    {
        "type": "object",
        "properties": {
            "window_title": {"type": "string", "description": "Full or partial window title"},
            "format": {"type": "string", "enum": ["png", "jpeg", "webp"]},
            "quality": {"type": "number", "minimum": 1, "maximum": 100},
            "scale": {"type": "number", "minimum": 0.1, "maximum": 1.0},
            "activate": {
                "type": "boolean",
                "description": "Bring window to front before capture (default: true)",
            },
        },
        "required": ["window_title"],
    },
)
async def handle_capture_window(arguments: dict[str, Any]) -> list[TextContent | ImageContent]:
    fmt, quality, scale = _capture_options(arguments)
    activate = get_bool(arguments, "activate", default=True)

    img, win_info = await capture_window_impl(
        get_str(arguments, "window_title", required=True, min_length=1, max_length=512),
        activate=activate,
    )
    data, mime, size = image_to_base64(img, fmt=fmt, quality=quality, scale=scale)

    return [
        TextContent(
            type="text",
            text=json.dumps(
                {
                    **win_info,
                    "format": fmt,
                    "file_size_kb": round(size / 1024, 1),
                }
            ),
        ),
        ImageContent(type="image", data=data, mimeType=mime),
    ]


@registry.register(
    "capture_monitor",
    "Capture a specific monitor by index",
    {
        "type": "object",
        "properties": {
            "monitor_index": {
                "type": "number",
                "minimum": 0,
                "description": "Monitor index (1=primary, 0=virtual screen). Use list_monitors to see indices.",
            },
            "format": {"type": "string", "enum": ["png", "jpeg", "webp"]},
            "quality": {"type": "number", "minimum": 1, "maximum": 100},
            "scale": {"type": "number", "minimum": 0.1, "maximum": 1.0},
        },
        "required": ["monitor_index"],
    },
)
async def handle_capture_monitor(arguments: dict[str, Any]) -> list[TextContent | ImageContent]:
    idx = get_int(arguments, "monitor_index", required=True, min_value=0)
    fmt, quality, scale = _capture_options(arguments)

    img = await capture_screen_impl(monitor_index=idx)
    data, mime, size = image_to_base64(img, fmt=fmt, quality=quality, scale=scale)

    return [
        TextContent(
            type="text",
            text=json.dumps(
                {
                    "monitor": idx,
                    "size": f"{img.width}x{img.height}",
                    "format": fmt,
                    "file_size_kb": round(size / 1024, 1),
                }
            ),
        ),
        ImageContent(type="image", data=data, mimeType=mime),
    ]


@registry.register(
    "list_monitors",
    "List all connected monitors with resolution and position",
    {
        "type": "object",
        "properties": {},
    },
)
async def handle_list_monitors(arguments: dict[str, Any]) -> dict[str, Any]:
    monitors = get_all_monitors()
    from ..utils.coordinates import get_scaling_factor, get_system_dpi

    return {
        "monitor_count": len(monitors),
        "dpi": get_system_dpi(),
        "scaling": f"{get_scaling_factor() * 100:.0f}%",
        "monitors": monitors,
    }


@registry.register(
    "get_pixel_color",
    "Get the color of a pixel at given coordinates",
    {
        "type": "object",
        "properties": {
            "x": {"type": "number", "description": "X coordinate"},
            "y": {"type": "number", "description": "Y coordinate"},
        },
        "required": ["x", "y"],
    },
)
async def handle_get_pixel_color(arguments: dict[str, Any]) -> dict[str, int | str]:
    x = get_int(arguments, "x", required=True)
    y = get_int(arguments, "y", required=True)
    if config.validate_coordinates:
        validate_coordinates(x, y, "get_pixel_color")

    def _grab() -> tuple[int, int, int]:
        sct = _get_mss()
        region = {"top": y, "left": x, "width": 1, "height": 1}
        screenshot = sct.grab(region)
        pixel: tuple[int, int, int] = screenshot.pixel(0, 0)
        return pixel

    r, g, b = await asyncio.to_thread(_grab)
    return {
        "x": x,
        "y": y,
        "r": r,
        "g": g,
        "b": b,
        "hex": f"#{r:02x}{g:02x}{b:02x}",
    }


@registry.register(
    "compare_screenshots",
    "Compare current screen with a reference image",
    {
        "type": "object",
        "properties": {
            "reference_image": {
                "type": "string",
                "description": "Base64-encoded reference image to compare against",
            },
            "region": {
                "type": "object",
                "description": "Optional region {x, y, width, height} to compare",
                "properties": {
                    "x": {"type": "number"},
                    "y": {"type": "number"},
                    "width": {"type": "number", "minimum": 1},
                    "height": {"type": "number", "minimum": 1},
                },
                "required": ["x", "y", "width", "height"],
            },
            "threshold": {
                "type": "number",
                "minimum": 0,
                "maximum": 1,
                "description": "Similarity threshold 0-1 (default: 0.95). Above = match.",
            },
        },
        "required": ["reference_image"],
    },
)
async def handle_compare_screenshots(arguments: dict[str, Any]) -> dict[str, Any]:
    threshold = get_float(arguments, "threshold", default=0.95, min_value=0.0, max_value=1.0)
    region = _optional_region(arguments)

    # Capture current
    if region:
        rx, ry, rw, rh = region
        if config.validate_coordinates:
            rx, ry, rw, rh = validate_region(rx, ry, rw, rh)
        current = await capture_region_impl(
            rx,
            ry,
            rw,
            rh,
        )
    else:
        current = await capture_screen_impl()

    # Decode reference — limit to ~50 MB decoded to prevent OOM
    max_b64_len = int(config.limits.max_reference_image_bytes * 4 / 3) + 4
    ref_b64 = get_str(arguments, "reference_image", required=True, min_length=1, max_length=max_b64_len)
    try:
        ref_bytes = base64.b64decode(ref_b64, validate=True)
    except Exception as exc:
        raise ToolError(f"Invalid reference image: {exc}") from exc
    if len(ref_bytes) > config.limits.max_reference_image_bytes:
        raise ToolError("Reference image too large (max ~50 MB)")
    try:
        reference = Image.open(io.BytesIO(ref_bytes)).convert("RGB")
    except Exception as exc:
        raise ToolError(f"Invalid reference image: {exc}") from exc

    diff = compute_image_diff(current, reference)
    diff["match"] = diff["similarity"] >= threshold
    diff["threshold"] = threshold
    return diff
