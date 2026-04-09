"""
Tool registration system with decorator-based dispatch.

Provides a central registry for all MCP tools, handling:
- Tool definition (name, description, schema)
- Handler registration via @registry.register decorator
- Dispatch with structured error handling
- Automatic JSON serialization of dict/list results
"""

import asyncio
import json
import logging
import time
import traceback
from collections.abc import Callable
from typing import Any

from mcp.types import ImageContent, TextContent, Tool

from .config import config
from .utils.errors import ToolError

logger = logging.getLogger("win32-mcp")


class ToolRegistry:
    """Registry for MCP tool handlers."""

    def __init__(self) -> None:
        self._tools: dict[str, dict[str, Any]] = {}
        self._last_dispatch_time: float = 0.0

    def register(self, name: str, description: str, schema: dict[str, Any]) -> Callable[..., Any]:
        """Decorator to register a tool handler.

        Usage:
            @registry.register("my_tool", "Description", {"type": "object", ...})
            async def handle_my_tool(arguments: dict[str, Any]):
                ...
        """

        def decorator(func: Callable[..., Any]) -> Callable[..., Any]:
            self._tools[name] = {
                "handler": func,
                "tool": Tool(
                    name=name,
                    description=description,
                    inputSchema=schema,
                ),
            }
            return func

        return decorator

    def get_tools(self) -> list[Tool]:
        """Return all registered Tool definitions for MCP list_tools."""
        return [entry["tool"] for entry in self._tools.values()]

    def get_handler(self, name: str) -> Callable[..., Any] | None:
        """Return the handler function for a tool name, or None."""
        entry = self._tools.get(name)
        return entry["handler"] if entry else None

    @property
    def tool_names(self) -> list[str]:
        return sorted(self._tools.keys())

    async def dispatch(self, name: str, arguments: dict[str, Any]) -> list[TextContent | ImageContent]:
        """Dispatch a tool call with full error handling.

        - Enforces min_operation_interval between consecutive calls.
        - Known errors (ToolError) return structured JSON with suggestion.
        - Unknown errors return structured JSON with traceback in debug mode.
        - If handler returns a list, it's passed through as-is.
        - If handler returns a dict/str, it's wrapped in TextContent.
        """
        # Rate limiting — enforce minimum interval between operations
        interval = config.min_operation_interval
        if interval > 0:
            now = time.monotonic()
            elapsed = now - self._last_dispatch_time
            if elapsed < interval:
                await asyncio.sleep(interval - elapsed)
            self._last_dispatch_time = time.monotonic()

        handler = self.get_handler(name)
        if handler is None:
            return [
                TextContent(
                    type="text",
                    text=json.dumps(
                        {
                            "error": True,
                            "tool": name,
                            "message": f"Unknown tool: {name}",
                            "available_tools": self.tool_names,
                        },
                        indent=2,
                    ),
                )
            ]

        try:
            result = await handler(arguments)

            # Handler returned MCP content list — pass through
            if isinstance(result, list):
                return result

            # Handler returned a single MCP content item
            if isinstance(result, TextContent | ImageContent):
                return [result]

            # Handler returned a dict or other serializable — wrap in text
            if isinstance(result, dict):
                return [
                    TextContent(
                        type="text",
                        text=json.dumps(result, indent=2, default=str),
                    )
                ]

            # Fallback: stringify
            return [TextContent(type="text", text=str(result))]

        except ToolError as exc:
            logger.warning("Tool %s error: %s", name, exc)
            payload: dict[str, Any] = {
                "error": True,
                "tool": name,
                "message": str(exc),
            }
            if exc.suggestion:
                payload["suggestion"] = exc.suggestion
            return [
                TextContent(
                    type="text",
                    text=json.dumps(payload, indent=2),
                )
            ]

        except Exception as exc:
            logger.exception("Unexpected error in tool %s", name)
            payload = {
                "error": True,
                "tool": name,
                "message": f"{type(exc).__name__}: {exc}",
            }
            if config.debug:
                payload["traceback"] = traceback.format_exc()
            return [
                TextContent(
                    type="text",
                    text=json.dumps(payload, indent=2),
                )
            ]


# Module-level singleton
registry = ToolRegistry()
