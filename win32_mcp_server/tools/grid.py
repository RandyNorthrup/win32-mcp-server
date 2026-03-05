"""
On-screen coordinate grid overlay tool.

Captures the screen and draws a labelled coordinate grid on top so the
AI (or user) can precisely identify pixel positions.  The grid
auto-adapts to the current resolution and DPI scaling, and uses
dual-colour rendering (dark outline + bright fill) so lines and labels
remain visible on *any* background.

Tools:
  - capture_grid       Capture screen with coordinate grid overlay
"""

import json
import logging

from PIL import Image, ImageDraw, ImageFont
from mcp.types import TextContent, ImageContent

from ..registry import registry
from ..config import config
from ..utils.coordinates import (
    get_all_monitors,
    get_system_dpi,
    get_scaling_factor,
)
from ..utils.imaging import image_to_base64
from .capture import capture_screen_impl, capture_region_impl

logger = logging.getLogger("win32-mcp")


# ───────────────────────────────────────────────────────────────────
# Constants
# ───────────────────────────────────────────────────────────────────

# Dual-colour scheme: outline colour + fill colour.
# Magenta/cyan are rarely dominant in UIs so they contrast well.
OUTLINE_COLOR = (0, 0, 0, 180)        # semi-transparent black outline
MAJOR_LINE_COLOR = (255, 0, 255, 160)  # magenta for major lines
MINOR_LINE_COLOR = (0, 255, 255, 100)  # cyan for minor lines
LABEL_BG_COLOR = (0, 0, 0, 200)       # dark badge behind text
LABEL_FG_COLOR = (255, 255, 0, 255)   # yellow text (high contrast on dark badge)

OUTLINE_WIDTH = 3   # background/outline stroke width
MAJOR_WIDTH = 1     # foreground major line width
MINOR_WIDTH = 1     # foreground minor line width


def _pick_grid_spacing(width: int, height: int, density: str) -> tuple[int, int]:
    """Return (major_spacing, minor_spacing) in pixels.

    *density* is one of "auto", "coarse", "normal", "fine", "ultra".
    "auto" picks based on resolution.
    """
    presets = {
        "coarse": (200, 0),
        "normal": (100, 50),
        "fine":   (50, 25),
        "ultra":  (25, 0),
    }

    if density != "auto":
        return presets.get(density, presets["normal"])

    total = max(width, height)
    if total <= 1280:
        return presets["normal"]
    elif total <= 1920:
        return presets["normal"]
    elif total <= 2560:
        return (100, 50)
    else:  # 4K+
        return (200, 100)


def _get_font(size: int) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    """Best-effort: use a TrueType font if available, else built-in."""
    # Try common Windows fonts first, then fall back
    for name in ("consola.ttf", "cour.ttf", "arial.ttf", "lucon.ttf"):
        try:
            return ImageFont.truetype(name, size)
        except (OSError, IOError):
            continue
    try:
        return ImageFont.truetype("C:/Windows/Fonts/consola.ttf", size)
    except Exception:
        pass
    return ImageFont.load_default()


def draw_grid_overlay(
    img: Image.Image,
    density: str = "auto",
    show_labels: bool = True,
    region_offset: tuple[int, int] = (0, 0),
) -> Image.Image:
    """Draw a coordinate grid on *img* (non-destructive – returns a copy).

    Parameters
    ----------
    img : PIL Image
        The screenshot to annotate.
    density : str
        Grid density preset: auto | coarse | normal | fine | ultra.
    show_labels : bool
        Whether to render coordinate labels at major intersections.
    region_offset : (ox, oy)
        If the image is a sub-region, offset labels so coordinates
        reflect the absolute screen position.

    Returns
    -------
    PIL Image (RGBA) with grid overlay.
    """
    w, h = img.size
    ox, oy = region_offset
    major_sp, minor_sp = _pick_grid_spacing(w, h, density)

    # Work in RGBA so we can use semi-transparency
    overlay = img.convert("RGBA")
    grid_layer = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    draw = ImageDraw.Draw(grid_layer)

    # Choose font size relative to grid spacing
    font_size = max(10, min(14, major_sp // 8))
    font = _get_font(font_size)

    # ── Minor grid lines ──────────────────────────────────────────
    if minor_sp and minor_sp < major_sp:
        for x in range(0, w, minor_sp):
            draw.line([(x, 0), (x, h)], fill=MINOR_LINE_COLOR, width=MINOR_WIDTH)
        for y in range(0, h, minor_sp):
            draw.line([(0, y), (w, y)], fill=MINOR_LINE_COLOR, width=MINOR_WIDTH)

    # ── Major grid lines (outline + fill for contrast) ────────────
    for x in range(0, w, major_sp):
        # Outline
        draw.line([(x, 0), (x, h)], fill=OUTLINE_COLOR, width=OUTLINE_WIDTH)
        # Fill
        draw.line([(x, 0), (x, h)], fill=MAJOR_LINE_COLOR, width=MAJOR_WIDTH)

    for y in range(0, h, major_sp):
        draw.line([(0, y), (w, y)], fill=OUTLINE_COLOR, width=OUTLINE_WIDTH)
        draw.line([(0, y), (w, y)], fill=MAJOR_LINE_COLOR, width=MAJOR_WIDTH)

    # ── Coordinate labels at major intersections ──────────────────
    if show_labels:
        # Top edge: x labels
        for x in range(0, w, major_sp):
            label = str(x + ox)
            bbox = draw.textbbox((0, 0), label, font=font)
            tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
            lx = x + 2
            ly = 2
            # Badge background
            draw.rectangle([lx - 1, ly - 1, lx + tw + 3, ly + th + 3], fill=LABEL_BG_COLOR)
            draw.text((lx, ly), label, fill=LABEL_FG_COLOR, font=font)

        # Left edge: y labels
        for y in range(major_sp, h, major_sp):  # skip 0,0 (already has x label there)
            label = str(y + oy)
            bbox = draw.textbbox((0, 0), label, font=font)
            tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
            lx = 2
            ly = y + 2
            draw.rectangle([lx - 1, ly - 1, lx + tw + 3, ly + th + 3], fill=LABEL_BG_COLOR)
            draw.text((lx, ly), label, fill=LABEL_FG_COLOR, font=font)

        # Interior intersection labels (every N major cells to avoid clutter)
        # Label every intersection for coarse/normal; every 2nd for fine/ultra
        label_every = 1 if major_sp >= 100 else 2
        count_x = 0
        for x in range(major_sp, w, major_sp):
            count_x += 1
            if count_x % label_every != 0:
                continue
            count_y = 0
            for y in range(major_sp, h, major_sp):
                count_y += 1
                if count_y % label_every != 0:
                    continue
                label = f"{x + ox},{y + oy}"
                bbox = draw.textbbox((0, 0), label, font=font)
                tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
                lx = x + 3
                ly = y + 3
                # Don't draw if it would go off-image
                if lx + tw + 4 > w or ly + th + 4 > h:
                    continue
                draw.rectangle([lx - 1, ly - 1, lx + tw + 3, ly + th + 3], fill=LABEL_BG_COLOR)
                draw.text((lx, ly), label, fill=LABEL_FG_COLOR, font=font)

    # ── Composite ────────────────────────────────────────────────
    overlay = Image.alpha_composite(overlay, grid_layer)
    return overlay


# ===================================================================
# MCP Tool Handlers
# ===================================================================

@registry.register("capture_grid", (
    "Capture screen with a coordinate grid overlay. "
    "Returns an annotated screenshot with labelled grid lines so you "
    "can precisely identify pixel coordinates for clicking, dragging, etc. "
    "Adapts to current resolution and DPI scaling."
), {
    "type": "object",
    "properties": {
        "monitor_index": {
            "type": "number",
            "description": (
                "1-based monitor index (1 = primary). "
                "Use 0 for the virtual screen spanning all monitors. "
                "Default: 1."
            ),
        },
        "density": {
            "type": "string",
            "enum": ["auto", "coarse", "normal", "fine", "ultra"],
            "description": (
                "Grid density. 'auto' picks based on resolution. "
                "coarse=200px, normal=100px, fine=50px, ultra=25px. "
                "Default: auto."
            ),
        },
        "show_labels": {
            "type": "boolean",
            "description": "Show coordinate labels at grid intersections (default: true).",
        },
        "region": {
            "type": "object",
            "description": (
                "Optional sub-region to capture {x, y, width, height}. "
                "Coordinates on the overlay will reflect absolute screen position."
            ),
            "properties": {
                "x": {"type": "number"},
                "y": {"type": "number"},
                "width": {"type": "number"},
                "height": {"type": "number"},
            },
            "required": ["x", "y", "width", "height"],
        },
        "format": {
            "type": "string",
            "enum": ["png", "jpeg", "webp"],
            "description": "Image format (default: png).",
        },
        "quality": {
            "type": "number",
            "description": "Compression quality 1-100 for jpeg/webp (default: 85).",
        },
        "scale": {
            "type": "number",
            "description": (
                "Resize factor 0.1-1.0 (default: 1.0). "
                "Scaling happens *after* the grid is drawn, so coordinates "
                "in the image stay correct but the image is smaller."
            ),
        },
    },
})
async def handle_capture_grid(arguments: dict):
    monitor_index = int(arguments.get("monitor_index", 1))
    density = arguments.get("density", "auto")
    show_labels = arguments.get("show_labels", True)
    region = arguments.get("region")
    fmt = arguments.get("format", config.capture.default_format)
    quality = arguments.get("quality", config.capture.default_quality)
    scale = arguments.get("scale", config.capture.default_scale)

    # ── Capture ───────────────────────────────────────────────────
    rx = ry = rw = rh = 0
    if region:
        rx, ry = int(region["x"]), int(region["y"])
        rw, rh = int(region["width"]), int(region["height"])
        img = await capture_region_impl(rx, ry, rw, rh)
        region_offset = (rx, ry)
    else:
        img = await capture_screen_impl(monitor_index=monitor_index)
        # If we captured a specific monitor, labels should reflect its
        # absolute position on the virtual desktop.
        if monitor_index >= 1:
            monitors = get_all_monitors()
            if monitor_index <= len(monitors):
                m = monitors[monitor_index - 1]
                region_offset = (m["x"], m["y"])
            else:
                region_offset = (0, 0)
        else:
            region_offset = (0, 0)

    # ── Draw grid ─────────────────────────────────────────────────
    annotated = draw_grid_overlay(
        img,
        density=density,
        show_labels=show_labels,
        region_offset=region_offset,
    )

    # Convert RGBA → RGB for JPEG compat, keep RGBA for PNG/WebP
    if fmt.lower() in ("jpg", "jpeg"):
        annotated = annotated.convert("RGB")

    data, mime, size = image_to_base64(annotated, fmt=fmt, quality=quality, scale=scale)

    # ── Metadata ──────────────────────────────────────────────────
    dpi = get_system_dpi()
    scaling = get_scaling_factor()
    major_sp, minor_sp = _pick_grid_spacing(img.width, img.height, density)

    meta = {
        "screen_size": f"{img.width}x{img.height}",
        "grid_spacing_major": major_sp,
        "grid_spacing_minor": minor_sp or "none",
        "density": density,
        "dpi": dpi,
        "scaling": f"{scaling * 100:.0f}%",
        "monitor_index": monitor_index,
        "format": fmt,
        "file_size_kb": round(size / 1024, 1),
        "hint": (
            "Use the grid coordinates shown on the image to identify "
            "exact pixel positions for click, drag, and other mouse "
            "operations. Labels show absolute screen coordinates."
        ),
    }

    if region:
        meta["region"] = {"x": rx, "y": ry, "width": rw, "height": rh}

    return [
        TextContent(type="text", text=json.dumps(meta)),
        ImageContent(type="image", data=data, mimeType=mime),
    ]
