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
from ..utils.args import get_enum, get_float, get_str, get_text
from ..utils.errors import ToolError
from ..utils.security import redact_text

logger = logging.getLogger("win32-mcp")
_TYPE_METHODS = {"auto", "type", "paste"}
_MAX_KEY_COMBO_CHARS = 128
_MAX_HOTKEY_KEYS = 8
_VALID_KEYS = {str(key).lower() for key in pyautogui.KEYBOARD_KEYS}
_KEY_ALIASES = {
    "control": "ctrl",
    "delete": "del",
    "escape": "esc",
    "page_down": "pagedown",
    "page_up": "pageup",
    "pgdn": "pagedown",
    "pgup": "pageup",
    "return": "enter",
    "windows": "win",
}


def _normalize_key(key: str) -> str:
    normalized = _KEY_ALIASES.get(key.strip().lower(), key.strip().lower())
    if not normalized:
        raise ToolError("Empty key name")
    if normalized not in _VALID_KEYS:
        raise ToolError(
            f"Invalid key name: '{key}'",
            suggestion="Use a PyAutoGUI key name such as enter, tab, esc, ctrl, shift, alt, or f1.",
        )
    return normalized


def _parse_key_combo(keys_raw: str) -> list[str]:
    parts = [part.strip() for part in keys_raw.split("+") if part.strip()]
    if not parts:
        raise ToolError(f"Invalid key combination: '{keys_raw}'")
    if len(parts) > _MAX_HOTKEY_KEYS:
        raise ToolError(f"Too many hotkey keys (max {_MAX_HOTKEY_KEYS})")
    return [_normalize_key(part) for part in parts]


@registry.register(
    "type_text",
    "Type text at current cursor position (supports Unicode)",
    {
        "type": "object",
        "properties": {
            "text": {"type": "string", "minLength": 1, "maxLength": 20000, "description": "Text to type"},
            "interval": {
                "type": "number",
                "minimum": 0,
                "maximum": 1,
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
    text = get_text(arguments, "text")
    interval = get_float(arguments, "interval", default=config.automation.type_interval, min_value=0.0, max_value=1.0)
    method = get_enum(arguments, "method", _TYPE_METHODS, default="auto")

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

        return [TextContent(type="text", text=f"Typed (pasted): {redact_text(text)}")]
    await asyncio.to_thread(pyautogui.write, text, interval=interval)
    return [TextContent(type="text", text=f"Typed: {redact_text(text)}")]


@registry.register(
    "press_key",
    "Press a keyboard key or key combination",
    {
        "type": "object",
        "properties": {
            "keys": {
                "type": "string",
                "minLength": 1,
                "maxLength": _MAX_KEY_COMBO_CHARS,
                "description": (
                    "Key name or combo separated by '+'. Examples: 'enter', 'tab', 'ctrl+c', 'alt+f4', 'ctrl+shift+s'"
                ),
            },
        },
        "required": ["keys"],
    },
)
async def handle_press_key(arguments: dict[str, Any]) -> list[TextContent]:
    keys_raw = get_str(arguments, "keys", required=True, min_length=1, max_length=_MAX_KEY_COMBO_CHARS).strip()
    if not keys_raw:
        raise ToolError("Empty key string")

    parts = _parse_key_combo(keys_raw)
    if len(parts) > 1:
        await asyncio.to_thread(pyautogui.hotkey, *parts)
    else:
        await asyncio.to_thread(pyautogui.press, parts[0])

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
                "minItems": 1,
                "maxItems": _MAX_HOTKEY_KEYS,
                "description": "Array of key names, e.g. ['ctrl', 'shift', 's']",
            },
        },
        "required": ["keys"],
    },
)
async def handle_hotkey(arguments: dict[str, Any]) -> list[TextContent]:
    keys = arguments["keys"]
    if not isinstance(keys, list) or not keys:
        raise ToolError("Empty key list")
    if len(keys) > _MAX_HOTKEY_KEYS:
        raise ToolError(f"Too many hotkey keys (max {_MAX_HOTKEY_KEYS})")

    cleaned: list[str] = []
    for idx, key in enumerate(keys):
        if not isinstance(key, str):
            raise ToolError(f"Hotkey item {idx} must be a string")
        if key.strip():
            cleaned.append(_normalize_key(key))
    if not cleaned:
        raise ToolError("Empty key list")
    await asyncio.to_thread(pyautogui.hotkey, *cleaned)

    return [TextContent(type="text", text=f"Hotkey: {'+'.join(cleaned)}")]
