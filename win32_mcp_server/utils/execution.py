"""Central execution controls: serialization, dry-run, and timeouts."""

import asyncio
from collections.abc import Awaitable, Callable
from typing import Any

from mcp.types import ImageContent, TextContent

from ..config import config
from .security import HIGH_RISK_TOOLS, MUTATING_TOOLS, redact_arguments, should_dry_run

_interaction_lock = asyncio.Lock()
_process_lock = asyncio.Lock()

_PROCESS_TOOLS = {"start_process", "kill_process", "wait_for_idle", "list_processes"}
_LOCK_FREE_TOOLS = {"execute_sequence"}


async def run_controlled(
    tool_name: str,
    arguments: dict[str, Any],
    call: Callable[[], Awaitable[Any]],
) -> Any:
    """Run a tool under central dry-run, serialization, and timeout rules."""
    if should_dry_run(tool_name, arguments):
        return {
            "dry_run": True,
            "tool": tool_name,
            "risk": "high" if tool_name in HIGH_RISK_TOOLS else "mutating",
            "would_call": redact_arguments(tool_name, arguments),
        }

    timeout = estimate_timeout_seconds(tool_name, arguments)
    lock = _select_lock(tool_name)

    if lock is None:
        return await asyncio.wait_for(call(), timeout=timeout)

    async with lock:
        return await asyncio.wait_for(call(), timeout=timeout)


def estimate_timeout_seconds(tool_name: str, arguments: dict[str, Any]) -> float:
    """Estimate hard timeout for a tool call."""
    max_runtime = config.limits.max_tool_runtime_seconds
    if tool_name == "execute_sequence":
        steps = arguments.get("steps", [])
        delay_total = 0.0
        if isinstance(steps, list):
            for step in steps:
                if isinstance(step, dict):
                    delay_total += min(float(step.get("delay_ms", 0) or 0), 30_000) / 1000.0
        return min(max_runtime, max(config.default_timeout, delay_total + (len(steps) * config.default_timeout) + 5.0))

    requested = arguments.get("timeout_seconds", config.default_timeout)
    try:
        requested_float = float(requested)
    except (TypeError, ValueError):
        requested_float = config.default_timeout

    if tool_name.startswith("ocr_") or tool_name in {"find_text_on_screen", "click_text", "get_window_snapshot"}:
        requested_float = max(requested_float, config.default_timeout * 2)

    return min(max_runtime, max(1.0, requested_float + 5.0))


def _select_lock(tool_name: str) -> asyncio.Lock | None:
    if tool_name in _LOCK_FREE_TOOLS:
        return None
    if tool_name in _PROCESS_TOOLS:
        return _process_lock
    if tool_name in MUTATING_TOOLS:
        return _interaction_lock
    return None


def wrap_timeout_error(tool_name: str, timeout: float) -> list[TextContent | ImageContent]:
    """Return standard timeout payload."""
    return [
        TextContent(
            type="text",
            text=(
                "{\n"
                '  "error": true,\n'
                f'  "tool": "{tool_name}",\n'
                f'  "message": "Tool timed out after {timeout:.1f}s"\n'
                "}"
            ),
        )
    ]
