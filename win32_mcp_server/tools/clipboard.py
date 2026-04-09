"""
Clipboard tools.

Tools:
  - clipboard_copy   Copy text to the system clipboard
  - clipboard_paste  Read the current clipboard contents
"""

import logging
from typing import Any

import pyperclip
from mcp.types import TextContent

from ..registry import registry
from ..utils.errors import ToolError

logger = logging.getLogger("win32-mcp")


@registry.register(
    "clipboard_copy",
    "Copy text to the system clipboard",
    {
        "type": "object",
        "properties": {
            "text": {"type": "string", "description": "Text to copy to clipboard"},
        },
        "required": ["text"],
    },
)
async def handle_clipboard_copy(arguments: dict[str, Any]) -> list[TextContent]:
    text = arguments["text"]
    pyperclip.copy(text)
    return [TextContent(type="text", text=f"Copied {len(text)} characters to clipboard")]


@registry.register(
    "clipboard_paste",
    "Read the current clipboard text contents",
    {
        "type": "object",
        "properties": {},
    },
)
async def handle_clipboard_paste(arguments: dict[str, Any]) -> list[TextContent]:
    try:
        text = pyperclip.paste()
    except Exception as exc:
        raise ToolError(f"Failed to read clipboard: {exc}") from exc

    if not text:
        return [TextContent(type="text", text="Clipboard is empty")]

    return [TextContent(type="text", text=f"Clipboard ({len(text)} chars):\n{text}")]
