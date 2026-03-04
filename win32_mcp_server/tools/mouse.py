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

import pyautogui
from mcp.types import TextContent

from ..registry import registry
from ..config import config
from ..utils.coordinates import validate_coordinates

logger = logging.getLogger("win32-mcp")

# Disable PyAutoGUI's failsafe (moving mouse to corner aborts)
pyautogui.FAILSAFE = False


def _validate_if_enabled(x: int, y: int, tool: str):
    if config.validate_coordinates:
        validate_coordinates(x, y, tool)


# ===================================================================
# MCP Tool Handlers
# ===================================================================

@registry.register("click", "Click at screen coordinates", {
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
            "description": "Number of clicks (default: 1)",
        },
    },
    "required": ["x", "y"],
})
async def handle_click(arguments: dict):
    x, y = int(arguments["x"]), int(arguments["y"])
    button = arguments.get("button", "left")
    clicks = int(arguments.get("clicks", 1))

    _validate_if_enabled(x, y, "click")

    await asyncio.to_thread(pyautogui.click, x, y, button=button, clicks=clicks)
    await asyncio.sleep(config.automation.click_delay)

    return [TextContent(
        type="text",
        text=f"Clicked {button} x{clicks} at ({x}, {y})",
    )]


@registry.register("double_click", "Double-click at screen coordinates", {
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
})
async def handle_double_click(arguments: dict):
    x, y = int(arguments["x"]), int(arguments["y"])
    button = arguments.get("button", "left")

    _validate_if_enabled(x, y, "double_click")

    await asyncio.to_thread(pyautogui.doubleClick, x, y, button=button)
    await asyncio.sleep(config.automation.click_delay)

    return [TextContent(type="text", text=f"Double-clicked {button} at ({x}, {y})")]


@registry.register("triple_click", "Triple-click to select entire line or paragraph", {
    "type": "object",
    "properties": {
        "x": {"type": "number", "description": "X coordinate"},
        "y": {"type": "number", "description": "Y coordinate"},
    },
    "required": ["x", "y"],
})
async def handle_triple_click(arguments: dict):
    x, y = int(arguments["x"]), int(arguments["y"])
    _validate_if_enabled(x, y, "triple_click")

    await asyncio.to_thread(pyautogui.click, x, y, clicks=3)
    await asyncio.sleep(config.automation.click_delay)

    return [TextContent(type="text", text=f"Triple-clicked at ({x}, {y})")]


@registry.register("drag", "Drag from start to end coordinates", {
    "type": "object",
    "properties": {
        "start_x": {"type": "number", "description": "Start X coordinate"},
        "start_y": {"type": "number", "description": "Start Y coordinate"},
        "end_x": {"type": "number", "description": "End X coordinate"},
        "end_y": {"type": "number", "description": "End Y coordinate"},
        "duration": {
            "type": "number",
            "description": "Drag duration in seconds (default: 0.5)",
        },
        "button": {
            "type": "string",
            "enum": ["left", "right", "middle"],
            "description": "Mouse button to hold (default: left)",
        },
    },
    "required": ["start_x", "start_y", "end_x", "end_y"],
})
async def handle_drag(arguments: dict):
    sx = int(arguments["start_x"])
    sy = int(arguments["start_y"])
    ex = int(arguments["end_x"])
    ey = int(arguments["end_y"])
    duration = arguments.get("duration", config.automation.drag_duration)
    button = arguments.get("button", "left")

    _validate_if_enabled(sx, sy, "drag start")
    _validate_if_enabled(ex, ey, "drag end")

    def _drag():
        pyautogui.moveTo(sx, sy)
        pyautogui.drag(ex - sx, ey - sy, duration=duration, button=button)

    await asyncio.to_thread(_drag)

    return [TextContent(
        type="text",
        text=f"Dragged {button} from ({sx}, {sy}) to ({ex}, {ey}) in {duration}s",
    )]


@registry.register("mouse_position", "Get current mouse cursor position", {
    "type": "object",
    "properties": {},
})
async def handle_mouse_position(arguments: dict):
    pos = pyautogui.position()
    return {"x": pos.x, "y": pos.y}


@registry.register("mouse_move", "Move mouse cursor to position", {
    "type": "object",
    "properties": {
        "x": {"type": "number", "description": "Target X coordinate"},
        "y": {"type": "number", "description": "Target Y coordinate"},
        "duration": {
            "type": "number",
            "description": "Movement duration in seconds (default: 0.25)",
        },
    },
    "required": ["x", "y"],
})
async def handle_mouse_move(arguments: dict):
    x, y = int(arguments["x"]), int(arguments["y"])
    duration = arguments.get("duration", config.automation.move_duration)

    _validate_if_enabled(x, y, "mouse_move")

    await asyncio.to_thread(pyautogui.moveTo, x, y, duration=duration)

    return [TextContent(type="text", text=f"Moved mouse to ({x}, {y})")]


@registry.register("scroll", "Scroll vertically at current or specified position", {
    "type": "object",
    "properties": {
        "amount": {
            "type": "number",
            "description": "Scroll amount. Positive=up, negative=down.",
        },
        "x": {"type": "number", "description": "Optional X coordinate to scroll at"},
        "y": {"type": "number", "description": "Optional Y coordinate to scroll at"},
    },
    "required": ["amount"],
})
async def handle_scroll(arguments: dict):
    amount = int(arguments["amount"])
    x = arguments.get("x")
    y = arguments.get("y")

    if x is not None and y is not None:
        ix, iy = int(x), int(y)
        _validate_if_enabled(ix, iy, "scroll")
        await asyncio.to_thread(pyautogui.scroll, amount, ix, iy)
        pos_text = f" at ({ix}, {iy})"
    else:
        await asyncio.to_thread(pyautogui.scroll, amount)
        pos_text = ""

    direction = "up" if amount > 0 else "down"
    return [TextContent(type="text", text=f"Scrolled {direction} {abs(amount)}{pos_text}")]


@registry.register("scroll_horizontal", "Scroll horizontally at current or specified position", {
    "type": "object",
    "properties": {
        "amount": {
            "type": "number",
            "description": "Scroll amount. Positive=right, negative=left.",
        },
        "x": {"type": "number", "description": "Optional X coordinate to scroll at"},
        "y": {"type": "number", "description": "Optional Y coordinate to scroll at"},
    },
    "required": ["amount"],
})
async def handle_scroll_horizontal(arguments: dict):
    amount = int(arguments["amount"])
    x = arguments.get("x")
    y = arguments.get("y")

    if x is not None and y is not None:
        ix, iy = int(x), int(y)
        _validate_if_enabled(ix, iy, "scroll_horizontal")
        await asyncio.to_thread(pyautogui.hscroll, amount, ix, iy)
        pos_text = f" at ({ix}, {iy})"
    else:
        await asyncio.to_thread(pyautogui.hscroll, amount)
        pos_text = ""

    direction = "right" if amount > 0 else "left"
    return [TextContent(type="text", text=f"Scrolled {direction} {abs(amount)}{pos_text}")]
