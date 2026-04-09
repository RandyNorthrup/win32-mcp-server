"""
Window management tools.

Tools:
  - list_windows      List all open windows (deduplicated, with details)
  - get_window_info   Detailed information about a specific window
  - focus_window      Bring window to foreground (with retry)
  - close_window      Close window by title
  - minimize_window   Minimize window
  - maximize_window   Maximize window
  - restore_window    Restore window from minimized/maximized
  - resize_window     Resize window to specified dimensions
  - move_window       Move window to specified position
  - wait_for_window   Wait for a window to appear (polling)
"""

import asyncio
import json
import logging
import time
from typing import Any

from mcp.types import TextContent

from ..config import config
from ..registry import registry
from ..utils.errors import ToolError
from ..utils.window_match import (
    find_window,
    find_window_strict,
    get_all_windows_deduped,
    get_window_details,
)

logger = logging.getLogger("win32-mcp")


# ===================================================================
# Retry helper for window operations
# ===================================================================


async def _retry_window_op(op_name: str, func: Any, *args: Any, **kwargs: Any) -> Any:
    """Retry a window operation with exponential backoff."""
    last_exc = None
    for attempt in range(config.window_retry_attempts):
        try:
            return func(*args, **kwargs)
        except Exception as exc:
            last_exc = exc
            if attempt < config.window_retry_attempts - 1:
                delay = config.window_retry_delay * (attempt + 1)
                logger.debug("Retrying %s (attempt %d, delay %.1fs): %s", op_name, attempt + 1, delay, exc)
                await asyncio.sleep(delay)
    raise ToolError(
        f"{op_name} failed after {config.window_retry_attempts} attempts: {last_exc}",
        suggestion="The window may be unresponsive. Try again or check if it's still open.",
    )


# ===================================================================
# MCP Tool Handlers
# ===================================================================


@registry.register(
    "list_windows",
    "List all open windows with position, size, and state",
    {
        "type": "object",
        "properties": {
            "filter": {
                "type": "string",
                "description": "Optional text filter — only windows whose title contains this text",
            },
        },
    },
)
async def handle_list_windows(arguments: dict[str, Any]) -> list[TextContent]:
    filter_text = arguments.get("filter", "").lower().strip()

    windows = get_all_windows_deduped()
    results = []

    # Also strip punctuation from filter so "SAK" can match "S.A.K."
    from win32_mcp_server.utils.window_match import _strip_punct

    filter_stripped = _strip_punct(filter_text) if filter_text else ""

    for win in windows:
        if filter_text:
            title_lower = win.title.lower()
            title_stripped = _strip_punct(title_lower)
            if filter_text not in title_lower and filter_stripped not in title_stripped:
                continue
        try:
            results.append(get_window_details(win))
        except Exception as exc:
            logger.debug("Could not get details for window '%s': %s", win.title, exc)

    return [
        TextContent(
            type="text",
            text=json.dumps(
                {
                    "count": len(results),
                    "filter": filter_text or None,
                    "windows": results,
                },
                indent=2,
            ),
        )
    ]


@registry.register(
    "get_window_info",
    "Get detailed information about a specific window",
    {
        "type": "object",
        "properties": {
            "window_title": {"type": "string", "description": "Full or partial window title"},
        },
        "required": ["window_title"],
    },
)
async def handle_get_window_info(arguments: dict[str, Any]) -> dict[str, Any]:
    win = find_window_strict(arguments["window_title"])
    details = get_window_details(win)

    # Add process info if PID available
    if details.get("pid"):
        try:
            import psutil

            proc = psutil.Process(details["pid"])
            details["process_name"] = proc.name()
            details["process_exe"] = proc.exe()
            details["memory_mb"] = round(proc.memory_info().rss / 1024 / 1024, 1)
            details["cpu_percent"] = proc.cpu_percent(interval=0.1)
        except Exception as exc:
            logger.debug("Could not get process info for PID %s: %s", details.get("pid"), exc)

    return details


@registry.register(
    "focus_window",
    "Bring a window to the foreground and activate it",
    {
        "type": "object",
        "properties": {
            "window_title": {"type": "string", "description": "Full or partial window title"},
        },
        "required": ["window_title"],
    },
)
async def handle_focus_window(arguments: dict[str, Any]) -> list[TextContent]:
    win = find_window_strict(arguments["window_title"])

    if win.isMinimized:
        await _retry_window_op("restore", win.restore)
        await asyncio.sleep(0.3)

    await _retry_window_op("activate", win.activate)
    await asyncio.sleep(0.2)

    # Verify focus if requested
    if arguments.get("verify", False):
        import pygetwindow as gw

        fg = gw.getActiveWindow()
        if fg is None or fg.title != win.title:
            logger.warning(
                "Focus verify: foreground is '%s', expected '%s'",
                fg.title if fg else "<none>",
                win.title,
            )

    return [TextContent(type="text", text=f"Focused: {win.title}")]


@registry.register(
    "close_window",
    "Close a window by title",
    {
        "type": "object",
        "properties": {
            "window_title": {"type": "string", "description": "Full or partial window title"},
        },
        "required": ["window_title"],
    },
)
async def handle_close_window(arguments: dict[str, Any]) -> list[TextContent]:
    win = find_window_strict(arguments["window_title"])
    title = win.title
    await _retry_window_op("close", win.close)
    return [TextContent(type="text", text=f"Closed: {title}")]


@registry.register(
    "minimize_window",
    "Minimize a window",
    {
        "type": "object",
        "properties": {
            "window_title": {"type": "string", "description": "Full or partial window title"},
        },
        "required": ["window_title"],
    },
)
async def handle_minimize_window(arguments: dict[str, Any]) -> list[TextContent]:
    win = find_window_strict(arguments["window_title"])
    await _retry_window_op("minimize", win.minimize)
    return [TextContent(type="text", text=f"Minimized: {win.title}")]


@registry.register(
    "maximize_window",
    "Maximize a window",
    {
        "type": "object",
        "properties": {
            "window_title": {"type": "string", "description": "Full or partial window title"},
        },
        "required": ["window_title"],
    },
)
async def handle_maximize_window(arguments: dict[str, Any]) -> list[TextContent]:
    win = find_window_strict(arguments["window_title"])
    await _retry_window_op("maximize", win.maximize)
    return [TextContent(type="text", text=f"Maximized: {win.title}")]


@registry.register(
    "restore_window",
    "Restore a window from minimized or maximized state",
    {
        "type": "object",
        "properties": {
            "window_title": {"type": "string", "description": "Full or partial window title"},
        },
        "required": ["window_title"],
    },
)
async def handle_restore_window(arguments: dict[str, Any]) -> list[TextContent]:
    win = find_window_strict(arguments["window_title"])
    await _retry_window_op("restore", win.restore)
    return [TextContent(type="text", text=f"Restored: {win.title}")]


@registry.register(
    "resize_window",
    "Resize a window to specified dimensions",
    {
        "type": "object",
        "properties": {
            "window_title": {"type": "string", "description": "Full or partial window title"},
            "width": {"type": "number", "description": "New width in pixels"},
            "height": {"type": "number", "description": "New height in pixels"},
        },
        "required": ["window_title", "width", "height"],
    },
)
async def handle_resize_window(arguments: dict[str, Any]) -> list[TextContent]:
    win = find_window_strict(arguments["window_title"])
    w, h = int(arguments["width"]), int(arguments["height"])

    if w < 100 or h < 50:
        raise ToolError(
            f"Window size {w}x{h} is too small (min 100x50)",
            suggestion="Specify larger dimensions",
        )

    await _retry_window_op("resize", win.resizeTo, w, h)
    return [TextContent(type="text", text=f"Resized '{win.title}' to {w}x{h}")]


@registry.register(
    "move_window",
    "Move a window to specified screen position",
    {
        "type": "object",
        "properties": {
            "window_title": {"type": "string", "description": "Full or partial window title"},
            "x": {"type": "number", "description": "New X position"},
            "y": {"type": "number", "description": "New Y position"},
        },
        "required": ["window_title", "x", "y"],
    },
)
async def handle_move_window(arguments: dict[str, Any]) -> list[TextContent]:
    win = find_window_strict(arguments["window_title"])
    x, y = int(arguments["x"]), int(arguments["y"])

    await _retry_window_op("move", win.moveTo, x, y)
    return [TextContent(type="text", text=f"Moved '{win.title}' to ({x}, {y})")]


@registry.register(
    "wait_for_window",
    "Wait for a window with a matching title to appear",
    {
        "type": "object",
        "properties": {
            "window_title": {
                "type": "string",
                "description": "Full or partial window title to wait for",
            },
            "timeout_seconds": {
                "type": "number",
                "description": "Maximum time to wait in seconds (default: 10)",
            },
            "poll_interval": {
                "type": "number",
                "description": "Seconds between checks (default: 0.5)",
            },
        },
        "required": ["window_title"],
    },
)
async def handle_wait_for_window(arguments: dict[str, Any]) -> dict[str, Any]:
    title = arguments["window_title"]
    timeout = arguments.get("timeout_seconds", config.default_timeout)
    interval = arguments.get("poll_interval", 0.5)

    start = time.monotonic()
    while True:
        win, _ = find_window(title)
        if win is not None:
            elapsed = round(time.monotonic() - start, 2)
            details = get_window_details(win)
            details["found_after_seconds"] = elapsed
            return details

        elapsed = time.monotonic() - start
        if elapsed >= timeout:
            raise ToolError(
                f"Window '{title}' did not appear within {timeout}s",
                suggestion="Check if the application started. Use list_windows to see current windows.",
            )

        await asyncio.sleep(interval)
