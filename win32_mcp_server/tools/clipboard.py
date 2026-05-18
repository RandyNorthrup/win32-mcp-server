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

from ..config import config
from ..registry import registry
from ..utils.args import get_bool, get_int, get_str
from ..utils.errors import ToolError
from ..utils.security import redact_text, truncate_text

logger = logging.getLogger("win32-mcp")


@registry.register(
    "clipboard_copy",
    "Copy text to the system clipboard",
    {
        "type": "object",
        "properties": {
            "text": {"type": "string", "maxLength": 20000, "description": "Text to copy to clipboard"},
        },
        "required": ["text"],
    },
)
async def handle_clipboard_copy(arguments: dict[str, Any]) -> list[TextContent]:
    text = get_str(arguments, "text", required=True, max_length=config.limits.max_clipboard_chars)
    pyperclip.copy(text)
    return [TextContent(type="text", text=f"Copied {len(text)} characters to clipboard")]


@registry.register(
    "clipboard_paste",
    "Read the current clipboard text contents",
    {
        "type": "object",
        "properties": {
            "include_content": {
                "type": "boolean",
                "description": "Return clipboard contents in response (default: false when redaction is enabled)",
            },
            "max_chars": {
                "type": "number",
                "minimum": 1,
                "maximum": 20000,
                "description": "Maximum clipboard characters to return",
            },
        },
    },
)
async def handle_clipboard_paste(arguments: dict[str, Any]) -> list[TextContent]:
    try:
        text = pyperclip.paste()
    except Exception as exc:
        raise ToolError(f"Failed to read clipboard: {exc}") from exc

    if not text:
        return [TextContent(type="text", text="Clipboard is empty")]

    include_content = get_bool(arguments, "include_content", default=False)
    max_chars = get_int(arguments, "max_chars", default=2000, min_value=1, max_value=config.limits.max_clipboard_chars)
    visible = truncate_text(text, max_chars=max_chars) if include_content else redact_text(text)

    return [TextContent(type="text", text=f"Clipboard ({len(text)} chars):\n{visible}")]
