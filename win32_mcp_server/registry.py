"""
Tool registration system with decorator-based dispatch.

Provides a central registry for all MCP tools, handling:
- Tool definition (name, description, schema)
- Handler registration via @registry.register decorator
- Dispatch with structured error handling
- Automatic JSON serialization of dict/list results
"""

import asyncio
import copy
import json
import logging
import time
import traceback
from collections.abc import Callable
from typing import Any

from mcp.types import ImageContent, TextContent, Tool

from .config import config
from .utils.errors import ToolError
from .utils.execution import estimate_timeout_seconds, run_controlled, wrap_timeout_error
from .utils.security import HIGH_RISK_TOOLS, MUTATING_TOOLS, enforce_tool_allowed

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
            input_schema = _with_common_security_schema(name, schema)
            self._tools[name] = {
                "handler": func,
                "tool": Tool(
                    name=name,
                    description=description,
                    inputSchema=input_schema,
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

    def get_schema(self, name: str) -> dict[str, Any] | None:
        """Return a tool input schema."""
        entry = self._tools.get(name)
        return entry["tool"].inputSchema if entry else None

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
            enforce_tool_allowed(name, arguments)
            schema = self.get_schema(name)
            if schema is not None:
                _validate_schema(arguments, schema, name)

            async def _call() -> Any:
                return await handler(arguments)

            result = await run_controlled(name, arguments, _call)

            # Handler returned MCP content list — pass through
            if isinstance(result, list):
                if all(isinstance(item, TextContent | ImageContent) for item in result):
                    return result
                return [
                    TextContent(
                        type="text",
                        text=json.dumps(_wrap_result(name, {"items": result}), indent=2, default=str),
                    )
                ]

            # Handler returned a single MCP content item
            if isinstance(result, TextContent | ImageContent):
                return [result]

            # Handler returned a dict or other serializable — wrap in text
            if isinstance(result, dict):
                payload = _wrap_result(name, result)
                return [
                    TextContent(
                        type="text",
                        text=json.dumps(payload, indent=2, default=str),
                    )
                ]

            # Fallback: stringify
            return [TextContent(type="text", text=str(result))]

        except ToolError as exc:
            logger.warning("Tool %s error: %s", name, exc)
            error_payload: dict[str, Any] = {
                "error": True,
                "tool": name,
                "message": str(exc),
            }
            if exc.suggestion:
                error_payload["suggestion"] = exc.suggestion
            return [
                TextContent(
                    type="text",
                    text=json.dumps(error_payload, indent=2),
                )
            ]

        except asyncio.TimeoutError:
            timeout = estimate_timeout_seconds(name, arguments)
            logger.warning("Tool %s timed out after %.1fs", name, timeout)
            return wrap_timeout_error(name, timeout)

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


def _validate_schema(value: Any, schema: dict[str, Any], tool_name: str) -> None:
    """Small JSON-schema subset validator for MCP tool inputs."""
    try:
        _validate_value(value, schema, "$")
    except ToolError as exc:
        raise ToolError(
            f"Invalid arguments for {tool_name}: {exc}",
            suggestion="Check the tool schema and retry with valid JSON values.",
        ) from exc


def _wrap_result(tool_name: str, result: dict[str, Any]) -> dict[str, Any]:
    if not config.result_envelope:
        return result
    return {
        "success": True,
        "tool": tool_name,
        "data": result,
    }


def _with_common_security_schema(tool_name: str, schema: dict[str, Any]) -> dict[str, Any]:
    """Advertise cross-cutting safety args supported by registry dispatch."""
    result = copy.deepcopy(schema)
    properties = result.setdefault("properties", {})
    if not isinstance(properties, dict):
        return result
    if tool_name in MUTATING_TOOLS:
        properties.setdefault(
            "dry_run",
            {
                "type": "boolean",
                "description": "Simulate the action and report what would run without changing OS state.",
            },
        )
    if tool_name in HIGH_RISK_TOOLS:
        properties.setdefault(
            "confirmation_token",
            {
                "type": "string",
                "description": "Required when WIN32_MCP_CONFIRMATION_TOKEN is configured.",
            },
        )
    return result


def _validate_value(value: Any, schema: dict[str, Any], path: str) -> None:
    if "enum" in schema and value not in schema["enum"]:
        raise ToolError(f"{path} must be one of {schema['enum']!r}")

    expected = schema.get("type")
    if expected is None:
        return

    if expected == "object":
        if not isinstance(value, dict):
            raise ToolError(f"{path} must be an object")
        required = schema.get("required", [])
        missing = [key for key in required if key not in value]
        if missing:
            raise ToolError(f"{path} missing required field(s): {', '.join(missing)}")
        properties = schema.get("properties", {})
        for key, item in value.items():
            if key in properties:
                _validate_value(item, properties[key], f"{path}.{key}")
        return

    if expected == "array":
        if not isinstance(value, list):
            raise ToolError(f"{path} must be an array")
        _validate_len(value, schema, path)
        item_schema = schema.get("items")
        if isinstance(item_schema, dict):
            for idx, item in enumerate(value):
                _validate_value(item, item_schema, f"{path}[{idx}]")
        return

    if expected == "string":
        if not isinstance(value, str):
            raise ToolError(f"{path} must be a string")
        _validate_len(value, schema, path)
        return

    if expected == "boolean":
        if not isinstance(value, bool):
            raise ToolError(f"{path} must be a boolean")
        return

    if expected in {"number", "integer"}:
        if isinstance(value, bool) or not isinstance(value, int | float):
            raise ToolError(f"{path} must be a {expected}")
        if expected == "integer" and not isinstance(value, int):
            raise ToolError(f"{path} must be an integer")
        minimum = schema.get("minimum")
        maximum = schema.get("maximum")
        if minimum is not None and value < minimum:
            raise ToolError(f"{path} must be >= {minimum}")
        if maximum is not None and value > maximum:
            raise ToolError(f"{path} must be <= {maximum}")


def _validate_len(value: str | list[Any], schema: dict[str, Any], path: str) -> None:
    min_len = schema.get("minLength", schema.get("minItems"))
    max_len = schema.get("maxLength", schema.get("maxItems"))
    if min_len is not None and len(value) < min_len:
        raise ToolError(f"{path} length must be >= {min_len}")
    if max_len is not None and len(value) > max_len:
        raise ToolError(f"{path} length must be <= {max_len}")


# Module-level singleton
registry = ToolRegistry()
