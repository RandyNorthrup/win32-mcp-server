"""
Mouse interaction tools.

Tools:
  - click               Click at coordinates (left/right/middle)
  - double_click        Double-click at coordinates
  - triple_click        Triple-click to select line/paragraph
  - drag                Drag from start to end coordinates
  - mouse_position      Get current mouse position
  - mouse_move          Move mouse to position
  - scroll              Scroll vertically (with optional position)
  - scroll_horizontal   Scroll horizontally
"""

import asyncio
import logging
from typing import Any

import pyautogui
from mcp.types import TextContent

from ..config import config
from ..registry import registry
from ..utils.args import get_bool, get_enum, get_float, get_int
from ..utils.coordinates import validate_coordinates
from ..utils.errors import ToolError

logger = logging.getLogger("win32-mcp")
_MOUSE_BUTTONS = {"left", "right", "middle"}
_MAX_SCROLL_AMOUNT = 10_000

# Configurable failsafe. Default preserves previous uninterrupted automation behavior.
pyautogui.FAILSAFE = config.automation.pyautogui_failsafe


def _validate_if_enabled(x: int, y: int, tool: str) -> None:
    if config.validate_coordinates:
        validate_coordinates(x, y, tool)


# ===================================================================
# MCP Tool Handlers
# ===================================================================


@registry.register(
    "click",
    "Click at screen coordinates",
    {
        "type": "object",
        "properties": {
            "x": {"type": "number", "description": "X coordinate"},
            "y": {"type": "number", "description": "Y coordinate"},
            "button": {
                "type": "string",
                "enum": ["left", "right", "middle"],
                "description": "Mouse button (default: left)",
            },
            "clicks": {
                "type": "number",
                "minimum": 1,
                "maximum": 10,
                "description": "Number of clicks (default: 1)",
            },
            "verify": {
                "type": "boolean",
                "description": "Verify cursor position after clicking (default: false)",
            },
        },
        "required": ["x", "y"],
    },
)
async def handle_click(arguments: dict[str, Any]) -> list[TextContent]:
    x = get_int(arguments, "x", required=True)
    y = get_int(arguments, "y", required=True)
    button = get_enum(arguments, "button", _MOUSE_BUTTONS, default="left")
    clicks = get_int(arguments, "clicks", default=1, min_value=1, max_value=10)

    _validate_if_enabled(x, y, "click")

    await asyncio.to_thread(pyautogui.click, x, y, button=button, clicks=clicks)
    await asyncio.sleep(config.automation.click_delay)

    # Optional post-click verification
    if get_bool(arguments, "verify", default=False):
        pos = pyautogui.position()
        dx, dy = abs(pos.x - x), abs(pos.y - y)
        if dx > 5 or dy > 5:
            logger.warning("Click verify: mouse at (%d,%d), expected (%d,%d)", pos.x, pos.y, x, y)

    return [
        TextContent(
            type="text",
            text=f"Clicked {button} x{clicks} at ({x}, {y})",
        )
    ]


@registry.register(
    "double_click",
    "Double-click at screen coordinates",
    {
        "type": "object",
        "properties": {
            "x": {"type": "number", "description": "X coordinate"},
            "y": {"type": "number", "description": "Y coordinate"},
            "button": {
                "type": "string",
                "enum": ["left", "right", "middle"],
                "description": "Mouse button (default: left)",
            },
        },
        "required": ["x", "y"],
    },
)
async def handle_double_click(arguments: dict[str, Any]) -> list[TextContent]:
    x = get_int(arguments, "x", required=True)
    y = get_int(arguments, "y", required=True)
    button = get_enum(arguments, "button", _MOUSE_BUTTONS, default="left")

    _validate_if_enabled(x, y, "double_click")

    await asyncio.to_thread(pyautogui.doubleClick, x, y, button=button)
    await asyncio.sleep(config.automation.click_delay)

    return [TextContent(type="text", text=f"Double-clicked {button} at ({x}, {y})")]


@registry.register(
    "triple_click",
    "Triple-click to select entire line or paragraph",
    {
        "type": "object",
        "properties": {
            "x": {"type": "number", "description": "X coordinate"},
            "y": {"type": "number", "description": "Y coordinate"},
        },
        "required": ["x", "y"],
    },
)
async def handle_triple_click(arguments: dict[str, Any]) -> list[TextContent]:
    x = get_int(arguments, "x", required=True)
    y = get_int(arguments, "y", required=True)
    _validate_if_enabled(x, y, "triple_click")

    await asyncio.to_thread(pyautogui.click, x, y, clicks=3)
    await asyncio.sleep(config.automation.click_delay)

    return [TextContent(type="text", text=f"Triple-clicked at ({x}, {y})")]


@registry.register(
    "drag",
    "Drag from start to end coordinates",
    {
        "type": "object",
        "properties": {
            "start_x": {"type": "number", "description": "Start X coordinate"},
            "start_y": {"type": "number", "description": "Start Y coordinate"},
            "end_x": {"type": "number", "description": "End X coordinate"},
            "end_y": {"type": "number", "description": "End Y coordinate"},
            "duration": {
                "type": "number",
                "minimum": 0,
                "maximum": 30,
                "description": "Drag duration in seconds (default: 0.5)",
            },
            "button": {
                "type": "string",
                "enum": ["left", "right", "middle"],
                "description": "Mouse button to hold (default: left)",
            },
        },
        "required": ["start_x", "start_y", "end_x", "end_y"],
    },
)
async def handle_drag(arguments: dict[str, Any]) -> list[TextContent]:
    sx = get_int(arguments, "start_x", required=True)
    sy = get_int(arguments, "start_y", required=True)
    ex = get_int(arguments, "end_x", required=True)
    ey = get_int(arguments, "end_y", required=True)
    duration = get_float(arguments, "duration", default=config.automation.drag_duration, min_value=0, max_value=30)
    button = get_enum(arguments, "button", _MOUSE_BUTTONS, default="left")

    _validate_if_enabled(sx, sy, "drag start")
    _validate_if_enabled(ex, ey, "drag end")

    def _drag() -> None:
        pyautogui.moveTo(sx, sy)
        pyautogui.drag(ex - sx, ey - sy, duration=duration, button=button)

    await asyncio.to_thread(_drag)

    return [
        TextContent(
            type="text",
            text=f"Dragged {button} from ({sx}, {sy}) to ({ex}, {ey}) in {duration}s",
        )
    ]


@registry.register(
    "mouse_position",
    "Get current mouse cursor position",
    {
        "type": "object",
        "properties": {},
    },
)
async def handle_mouse_position(arguments: dict[str, Any]) -> dict[str, int]:
    pos = pyautogui.position()
    return {"x": pos.x, "y": pos.y}


@registry.register(
    "mouse_move",
    "Move mouse cursor to position",
    {
        "type": "object",
        "properties": {
            "x": {"type": "number", "description": "Target X coordinate"},
            "y": {"type": "number", "description": "Target Y coordinate"},
            "duration": {
                "type": "number",
                "minimum": 0,
                "maximum": 30,
                "description": "Movement duration in seconds (default: 0.25)",
            },
        },
        "required": ["x", "y"],
    },
)
async def handle_mouse_move(arguments: dict[str, Any]) -> list[TextContent]:
    x = get_int(arguments, "x", required=True)
    y = get_int(arguments, "y", required=True)
    duration = get_float(arguments, "duration", default=config.automation.move_duration, min_value=0, max_value=30)

    _validate_if_enabled(x, y, "mouse_move")

    await asyncio.to_thread(pyautogui.moveTo, x, y, duration=duration)

    return [TextContent(type="text", text=f"Moved mouse to ({x}, {y})")]


@registry.register(
    "scroll",
    "Scroll vertically at current or specified position",
    {
        "type": "object",
        "properties": {
            "amount": {
                "type": "number",
                "minimum": -_MAX_SCROLL_AMOUNT,
                "maximum": _MAX_SCROLL_AMOUNT,
                "description": "Scroll amount. Positive=up, negative=down.",
            },
            "x": {"type": "number", "description": "Optional X coordinate to scroll at"},
            "y": {"type": "number", "description": "Optional Y coordinate to scroll at"},
        },
        "required": ["amount"],
    },
)
async def handle_scroll(arguments: dict[str, Any]) -> list[TextContent]:
    amount = get_int(arguments, "amount", required=True, min_value=-_MAX_SCROLL_AMOUNT, max_value=_MAX_SCROLL_AMOUNT)
    x = arguments.get("x")
    y = arguments.get("y")

    if (x is None) != (y is None):
        raise ToolError(
            "Both x and y must be provided for positional scrolling, or neither",
            suggestion="Provide both x and y coordinates, or omit both to scroll at current position",
        )

    if x is not None and y is not None:
        ix = get_int(arguments, "x", required=True)
        iy = get_int(arguments, "y", required=True)
        _validate_if_enabled(ix, iy, "scroll")
        await asyncio.to_thread(pyautogui.scroll, amount, ix, iy)
        pos_text = f" at ({ix}, {iy})"
    else:
        await asyncio.to_thread(pyautogui.scroll, amount)
        pos_text = ""

    direction = "up" if amount > 0 else "down"
    return [TextContent(type="text", text=f"Scrolled {direction} {abs(amount)}{pos_text}")]


@registry.register(
    "scroll_horizontal",
    "Scroll horizontally at current or specified position",
    {
        "type": "object",
        "properties": {
            "amount": {
                "type": "number",
                "minimum": -_MAX_SCROLL_AMOUNT,
                "maximum": _MAX_SCROLL_AMOUNT,
                "description": "Scroll amount. Positive=right, negative=left.",
            },
            "x": {"type": "number", "description": "Optional X coordinate to scroll at"},
            "y": {"type": "number", "description": "Optional Y coordinate to scroll at"},
        },
        "required": ["amount"],
    },
)
async def handle_scroll_horizontal(arguments: dict[str, Any]) -> list[TextContent]:
    amount = get_int(arguments, "amount", required=True, min_value=-_MAX_SCROLL_AMOUNT, max_value=_MAX_SCROLL_AMOUNT)
    x = arguments.get("x")
    y = arguments.get("y")

    if (x is None) != (y is None):
        raise ToolError(
            "Both x and y must be provided for positional scrolling, or neither",
            suggestion="Provide both x and y coordinates, or omit both to scroll at current position",
        )

    if x is not None and y is not None:
        ix = get_int(arguments, "x", required=True)
        iy = get_int(arguments, "y", required=True)
        _validate_if_enabled(ix, iy, "scroll_horizontal")
        await asyncio.to_thread(pyautogui.hscroll, amount, ix, iy)
        pos_text = f" at ({ix}, {iy})"
    else:
        await asyncio.to_thread(pyautogui.hscroll, amount)
        pos_text = ""

    direction = "right" if amount > 0 else "left"
    return [TextContent(type="text", text=f"Scrolled {direction} {abs(amount)}{pos_text}")]
