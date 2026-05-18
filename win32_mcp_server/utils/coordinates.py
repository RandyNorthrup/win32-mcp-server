"""
Screen coordinate utilities, DPI awareness, and monitor info.

Initializes DPI awareness on import so all subsequent coordinate
operations use physical pixels.
"""

import ctypes
import ctypes.wintypes
import logging
from typing import Any

from mss import mss

from .errors import ToolError

logger = logging.getLogger("win32-mcp")


# ---------------------------------------------------------------------------
# DPI Awareness
# ---------------------------------------------------------------------------


def setup_dpi_awareness() -> bool:
    """Enable per-monitor DPI awareness. Call once at startup.

    Returns True if awareness was successfully set.
    """
    try:
        # PROCESS_PER_MONITOR_DPI_AWARE = 2
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
        logger.info("Per-monitor DPI awareness enabled")
        return True
    except Exception:
        logger.debug("SetProcessDpiAwareness failed, trying fallback")
    try:
        ctypes.windll.user32.SetProcessDPIAware()
        logger.info("System-level DPI awareness enabled")
        return True
    except Exception:
        logger.warning("Could not set DPI awareness — coordinates may be inaccurate on scaled displays")
        return False


def get_system_dpi() -> int:
    """Return the system DPI (96 = 100% scaling)."""
    try:
        result: int = ctypes.windll.user32.GetDpiForSystem()
        return result
    except Exception:
        # Fallback: query DC
        try:
            hdc = ctypes.windll.user32.GetDC(0)
            dpi: int = ctypes.windll.gdi32.GetDeviceCaps(hdc, 88)  # LOGPIXELSX
            ctypes.windll.user32.ReleaseDC(0, hdc)
            return dpi
        except Exception:
            return 96


def get_scaling_factor() -> float:
    """Return the display scaling factor (e.g. 1.25 for 125%)."""
    return get_system_dpi() / 96.0


# ---------------------------------------------------------------------------
# Screen Geometry
# ---------------------------------------------------------------------------


def get_primary_screen_size() -> tuple[int, int]:
    """Return (width, height) of the primary monitor."""
    with mss() as sct:
        m = sct.monitors[1]
        return m["width"], m["height"]


def get_virtual_screen_size() -> tuple[int, int]:
    """Return (width, height) of the virtual screen (all monitors)."""
    with mss() as sct:
        m = sct.monitors[0]
        return m["width"], m["height"]


def get_all_monitors() -> list[dict[str, Any]]:
    """Return a list of monitor dicts with position, size, and index."""
    with mss() as sct:
        monitors = []
        for i, m in enumerate(sct.monitors[1:], start=1):
            monitors.append(
                {
                    "index": i,
                    "x": m["left"],
                    "y": m["top"],
                    "width": m["width"],
                    "height": m["height"],
                    "is_primary": i == 1,
                }
            )
        return monitors


def point_in_monitor(x: int, y: int, monitor: dict[str, Any]) -> bool:
    """Return True if a point is inside a monitor rect."""
    left = int(monitor["x"])
    top = int(monitor["y"])
    right = left + int(monitor["width"])
    bottom = top + int(monitor["height"])
    return left <= x < right and top <= y < bottom


def point_in_any_monitor(x: int, y: int, monitors: list[dict[str, Any]] | None = None) -> bool:
    """Return True if a point is inside the actual union of monitor rects."""
    monitor_list = monitors if monitors is not None else get_all_monitors()
    return any(point_in_monitor(x, y, monitor) for monitor in monitor_list)


# ---------------------------------------------------------------------------
# Coordinate Validation
# ---------------------------------------------------------------------------


def validate_coordinates(x: int, y: int, context: str = "operation") -> None:
    """Raise ToolError if (x, y) is outside any monitor bounds.

    Checks actual monitor rectangles, not just virtual-screen bounding box.
    """
    monitors = get_all_monitors()
    if not point_in_any_monitor(x, y, monitors):
        bounds = [f"({m['x']},{m['y']})-({m['x'] + m['width']},{m['y'] + m['height']})" for m in monitors]
        raise ToolError(
            f"Coordinates ({x}, {y}) are outside monitor bounds for {context}: {', '.join(bounds)}",
            suggestion="Use capture_screen or list_monitors to check screen dimensions",
        )


def validate_region(x: int, y: int, width: int, height: int) -> tuple[int, int, int, int]:
    """Validate a rectangular region is within screen bounds and positive.

    Returns:
        (x, y, width, height) — clamped to screen bounds.
    """
    if width <= 0 or height <= 0:
        raise ToolError(
            f"Invalid region size: {width}x{height} — must be positive",
        )
    validate_coordinates(x, y, "region top-left")
    # Clamp bottom-right to virtual screen
    with mss() as sct:
        virtual = sct.monitors[0]
        max_x = virtual["left"] + virtual["width"]
        max_y = virtual["top"] + virtual["height"]
    if x + width > max_x:
        width = max_x - x
    if y + height > max_y:
        height = max_y - y
    return x, y, max(1, width), max(1, height)


# ---------------------------------------------------------------------------
# Rectangle Clamping (multi-monitor safe)
# ---------------------------------------------------------------------------


def clamp_rect_to_virtual_screen(
    x: int,
    y: int,
    width: int,
    height: int,
) -> tuple[int, int, int, int]:
    """Clamp a rectangle to the virtual screen bounds.

    Handles multi-monitor setups where coordinates can be negative
    (e.g., monitors to the left of or above the primary).

    Returns:
        (clamped_x, clamped_y, clamped_width, clamped_height)
        Width/height will be 0 if entirely off-screen.
    """
    with mss() as sct:
        virt = sct.monitors[0]
        min_x = virt["left"]
        min_y = virt["top"]
        max_x = min_x + virt["width"]
        max_y = min_y + virt["height"]

    clamped_x = max(min_x, x)
    clamped_y = max(min_y, y)
    clamped_w = min(max_x, x + width) - clamped_x
    clamped_h = min(max_y, y + height) - clamped_y

    return clamped_x, clamped_y, max(0, clamped_w), max(0, clamped_h)


# ---------------------------------------------------------------------------
# Auto-initialize DPI on import
# ---------------------------------------------------------------------------
_dpi_initialized = setup_dpi_awareness()
