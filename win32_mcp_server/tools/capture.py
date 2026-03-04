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

from mss import mss
from PIL import Image
from mcp.types import TextContent, ImageContent

from ..registry import registry
from ..config import config
from ..utils.errors import ToolError
from ..utils.imaging import mss_to_pil, image_to_base64, compute_image_diff
from ..utils.coordinates import validate_coordinates, get_all_monitors
from ..utils.window_match import find_window_strict

logger = logging.getLogger("win32-mcp")


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
    def _grab():
        with mss() as sct:
            if monitor_index < 0 or monitor_index >= len(sct.monitors):
                raise ToolError(
                    f"Monitor index {monitor_index} out of range (0–{len(sct.monitors)-1})",
                    suggestion="Use list_monitors to see available monitors",
                )
            screenshot = sct.grab(sct.monitors[monitor_index])
            return mss_to_pil(screenshot)

    return await asyncio.to_thread(_grab)


async def capture_region_impl(x: int, y: int, width: int, height: int) -> Image.Image:
    """Capture a screen region and return a PIL Image."""
    def _grab():
        with mss() as sct:
            region = {"top": y, "left": x, "width": width, "height": height}
            screenshot = sct.grab(region)
            return mss_to_pil(screenshot)

    return await asyncio.to_thread(_grab)


async def capture_window_impl(window_title: str, activate: bool = True) -> tuple[Image.Image, dict]:
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
            except Exception:
                pass
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
        except Exception:
            pass  # already focused or not activatable
        await asyncio.sleep(0.2)

        # Poll until window is foreground (up to 1.5s)
        hwnd = getattr(win, "_hWnd", None)
        if hwnd:
            for _ in range(15):
                try:
                    fg = ctypes.windll.user32.GetForegroundWindow()
                    if fg == hwnd:
                        break
                except Exception:
                    break
                await asyncio.sleep(0.1)

    # Clamp to screen — handle windows extending off-screen
    left = max(0, win.left)
    top = max(0, win.top)
    # Adjust width/height for the portion clipped at the left/top edges
    width = win.width - (left - win.left)
    height = win.height - (top - win.top)
    if width <= 0 or height <= 0:
        raise ToolError(
            f"Window '{win.title}' is entirely off-screen (position: {win.left},{win.top} size: {win.width}x{win.height})",
            suggestion="Move the window on-screen first using move_window",
        )
    width = max(1, width)
    height = max(1, height)

    img = await capture_region_impl(left, top, width, height)

    info = {
        "title": win.title,
        "x": win.left,
        "y": win.top,
        "width": win.width,
        "height": win.height,
    }
    return img, info


# ===================================================================
# MCP Tool Handlers
# ===================================================================

@registry.register("capture_screen", "Capture full screen screenshot with optional compression", {
    "type": "object",
    "properties": {
        "format": {
            "type": "string",
            "enum": ["png", "jpeg", "webp"],
            "description": "Image format (default: png). jpeg/webp are smaller.",
        },
        "quality": {
            "type": "number",
            "description": "Compression quality 1-100 for jpeg/webp (default: 85)",
        },
        "scale": {
            "type": "number",
            "description": "Resize factor 0.1-1.0 (default: 1.0). 0.5 = half size, 4x smaller file.",
        },
    },
})
async def handle_capture_screen(arguments: dict):
    fmt = arguments.get("format", config.capture.default_format)
    quality = arguments.get("quality", config.capture.default_quality)
    scale = arguments.get("scale", config.capture.default_scale)

    img = await capture_screen_impl()
    data, mime, size = image_to_base64(img, fmt=fmt, quality=quality, scale=scale)

    return [
        TextContent(type="text", text=json.dumps({
            "screen_size": f"{img.width}x{img.height}",
            "format": fmt,
            "file_size_kb": round(size / 1024, 1),
        })),
        ImageContent(type="image", data=data, mimeType=mime),
    ]


@registry.register("capture_window", "Capture a specific window by title (fuzzy match supported)", {
    "type": "object",
    "properties": {
        "window_title": {"type": "string", "description": "Full or partial window title"},
        "format": {"type": "string", "enum": ["png", "jpeg", "webp"]},
        "quality": {"type": "number"},
        "scale": {"type": "number"},
        "activate": {
            "type": "boolean",
            "description": "Bring window to front before capture (default: true)",
        },
    },
    "required": ["window_title"],
})
async def handle_capture_window(arguments: dict):
    fmt = arguments.get("format", config.capture.default_format)
    quality = arguments.get("quality", config.capture.default_quality)
    scale = arguments.get("scale", config.capture.default_scale)
    activate = arguments.get("activate", True)

    img, win_info = await capture_window_impl(
        arguments["window_title"], activate=activate,
    )
    data, mime, size = image_to_base64(img, fmt=fmt, quality=quality, scale=scale)

    return [
        TextContent(type="text", text=json.dumps({
            **win_info,
            "format": fmt,
            "file_size_kb": round(size / 1024, 1),
        })),
        ImageContent(type="image", data=data, mimeType=mime),
    ]


@registry.register("capture_monitor", "Capture a specific monitor by index", {
    "type": "object",
    "properties": {
        "monitor_index": {
            "type": "number",
            "description": "1-based monitor index (1=primary). Use list_monitors to see indices.",
        },
        "format": {"type": "string", "enum": ["png", "jpeg", "webp"]},
        "quality": {"type": "number"},
        "scale": {"type": "number"},
    },
    "required": ["monitor_index"],
})
async def handle_capture_monitor(arguments: dict):
    idx = int(arguments["monitor_index"])
    fmt = arguments.get("format", config.capture.default_format)
    quality = arguments.get("quality", config.capture.default_quality)
    scale = arguments.get("scale", config.capture.default_scale)

    img = await capture_screen_impl(monitor_index=idx)
    data, mime, size = image_to_base64(img, fmt=fmt, quality=quality, scale=scale)

    return [
        TextContent(type="text", text=json.dumps({
            "monitor": idx,
            "size": f"{img.width}x{img.height}",
            "format": fmt,
            "file_size_kb": round(size / 1024, 1),
        })),
        ImageContent(type="image", data=data, mimeType=mime),
    ]


@registry.register("list_monitors", "List all connected monitors with resolution and position", {
    "type": "object",
    "properties": {},
})
async def handle_list_monitors(arguments: dict):
    monitors = get_all_monitors()
    from ..utils.coordinates import get_system_dpi, get_scaling_factor
    return {
        "monitor_count": len(monitors),
        "dpi": get_system_dpi(),
        "scaling": f"{get_scaling_factor() * 100:.0f}%",
        "monitors": monitors,
    }


@registry.register("get_pixel_color", "Get the color of a pixel at given coordinates", {
    "type": "object",
    "properties": {
        "x": {"type": "number", "description": "X coordinate"},
        "y": {"type": "number", "description": "Y coordinate"},
    },
    "required": ["x", "y"],
})
async def handle_get_pixel_color(arguments: dict):
    x, y = int(arguments["x"]), int(arguments["y"])
    if config.validate_coordinates:
        validate_coordinates(x, y, "get_pixel_color")

    def _grab():
        with mss() as sct:
            region = {"top": y, "left": x, "width": 1, "height": 1}
            screenshot = sct.grab(region)
            return screenshot.pixel(0, 0)  # returns (R, G, B)

    r, g, b = await asyncio.to_thread(_grab)
    return {
        "x": x, "y": y,
        "r": r, "g": g, "b": b,
        "hex": f"#{r:02x}{g:02x}{b:02x}",
    }


@registry.register("compare_screenshots", "Compare current screen with a reference image", {
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
                "width": {"type": "number"},
                "height": {"type": "number"},
            },
        },
        "threshold": {
            "type": "number",
            "description": "Similarity threshold 0-1 (default: 0.95). Above = match.",
        },
    },
    "required": ["reference_image"],
})
async def handle_compare_screenshots(arguments: dict):
    threshold = arguments.get("threshold", 0.95)
    region = arguments.get("region")

    # Capture current
    if region:
        current = await capture_region_impl(
            int(region["x"]), int(region["y"]),
            int(region["width"]), int(region["height"]),
        )
    else:
        current = await capture_screen_impl()

    # Decode reference
    try:
        ref_bytes = base64.b64decode(arguments["reference_image"])
        reference = Image.open(io.BytesIO(ref_bytes)).convert("RGB")
    except Exception as exc:
        raise ToolError(f"Invalid reference image: {exc}")

    diff = compute_image_diff(current, reference)
    diff["match"] = diff["similarity"] >= threshold
    diff["threshold"] = threshold
    return diff
