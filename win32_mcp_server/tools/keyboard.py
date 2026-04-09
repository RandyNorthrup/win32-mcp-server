"""
Keyboard interaction tools.

Tools:
  - type_text   Type text with Unicode support (auto-fallback to clipboard paste)
  - press_key   Press key or key combination (e.g. 'ctrl+c', 'enter')
  - hotkey      Execute hotkey combination from array of keys
"""

import asyncio
import logging
from typing import Any

import pyautogui
import pyperclip
from mcp.types import TextContent

from ..config import config
from ..registry import registry
from ..utils.errors import ToolError

logger = logging.getLogger("win32-mcp")


@registry.register(
    "type_text",
    "Type text at current cursor position (supports Unicode)",
    {
        "type": "object",
        "properties": {
            "text": {"type": "string", "description": "Text to type"},
            "interval": {
                "type": "number",
                "description": "Seconds between keystrokes (default: 0.01)",
            },
            "method": {
                "type": "string",
                "enum": ["auto", "type", "paste"],
                "description": (
                    "Typing method. 'type'=keystrokes (ASCII only), "
                    "'paste'=clipboard paste (any text), "
                    "'auto'=type if ASCII, paste if Unicode (default: auto)"
                ),
            },
        },
        "required": ["text"],
    },
)
async def handle_type_text(arguments: dict[str, Any]) -> list[TextContent]:
    text = arguments["text"]
    interval = arguments.get("interval", config.automation.type_interval)
    method = arguments.get("method", "auto")

    if not text:
        raise ToolError("Empty text — nothing to type")

    # Decide method
    use_paste = method == "paste" or (method == "auto" and not text.isascii())

    if use_paste:
        # Save current clipboard, paste text, restore
        try:
            old_clipboard = pyperclip.paste()
        except Exception as exc:
            logger.debug("Could not read clipboard: %s", exc)
            old_clipboard = ""

        pyperclip.copy(text)
        await asyncio.sleep(0.05)
        await asyncio.to_thread(pyautogui.hotkey, "ctrl", "v")
        await asyncio.sleep(0.1)

        # Restore previous clipboard
        try:
            pyperclip.copy(old_clipboard)
        except Exception as exc:
            logger.debug("Could not restore clipboard: %s", exc)

        return [TextContent(type="text", text=f"Typed (pasted): {text[:200]}")]
    await asyncio.to_thread(pyautogui.write, text, interval=interval)
    return [TextContent(type="text", text=f"Typed: {text[:200]}")]


@registry.register(
    "press_key",
    "Press a keyboard key or key combination",
    {
        "type": "object",
        "properties": {
            "keys": {
                "type": "string",
                "description": (
                    "Key name or combo separated by '+'. Examples: 'enter', 'tab', 'ctrl+c', 'alt+f4', 'ctrl+shift+s'"
                ),
            },
        },
        "required": ["keys"],
    },
)
async def handle_press_key(arguments: dict[str, Any]) -> list[TextContent]:
    keys_raw = arguments["keys"].strip()
    if not keys_raw:
        raise ToolError("Empty key string")

    keys_lower = keys_raw.lower()

    if "+" in keys_lower:
        parts = [k.strip() for k in keys_lower.split("+") if k.strip()]
        if not parts:
            raise ToolError(f"Invalid key combination: '{keys_raw}'")
        await asyncio.to_thread(pyautogui.hotkey, *parts)
    else:
        await asyncio.to_thread(pyautogui.press, keys_lower)

    return [TextContent(type="text", text=f"Pressed: {keys_raw}")]


@registry.register(
    "hotkey",
    "Execute a hotkey combination from an array of key names",
    {
        "type": "object",
        "properties": {
            "keys": {
                "type": "array",
                "items": {"type": "string"},
                "description": "Array of key names, e.g. ['ctrl', 'shift', 's']",
            },
        },
        "required": ["keys"],
    },
)
async def handle_hotkey(arguments: dict[str, Any]) -> list[TextContent]:
    keys = arguments["keys"]
    if not keys:
        raise ToolError("Empty key list")

    cleaned = [k.strip().lower() for k in keys if k.strip()]
    await asyncio.to_thread(pyautogui.hotkey, *cleaned)

    return [TextContent(type="text", text=f"Hotkey: {'+'.join(cleaned)}")]
