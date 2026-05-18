"""Safety policy, redaction, and output bounding helpers."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from ..config import config
from .errors import ToolError

READ_ONLY_TOOLS = {
    "assert_text_visible",
    "capture_monitor",
    "capture_screen",
    "capture_window",
    "clipboard_paste",
    "compare_screenshots",
    "find_text_on_screen",
    "get_pixel_color",
    "get_window_info",
    "get_window_snapshot",
    "health_check",
    "list_monitors",
    "list_processes",
    "list_windows",
    "mouse_position",
    "ocr_region",
    "ocr_region_structured",
    "ocr_screen",
    "ocr_screen_structured",
    "ocr_window",
    "uia_find_control",
    "uia_get_control_value",
    "uia_get_focused",
    "uia_inspect_window",
    "wait_for_idle",
    "wait_for_text",
    "wait_for_window",
}

HIGH_RISK_TOOLS = {
    "close_window",
    "kill_process",
    "start_process",
}

MUTATING_TOOLS = {
    "click",
    "click_text",
    "clipboard_copy",
    "double_click",
    "drag",
    "fill_field",
    "focus_window",
    "hotkey",
    "maximize_window",
    "minimize_window",
    "mouse_move",
    "move_window",
    "press_key",
    "resize_window",
    "restore_window",
    "right_click_menu",
    "scroll",
    "scroll_horizontal",
    "triple_click",
    "type_text",
    "uia_click_control",
    "uia_set_control_value",
} | HIGH_RISK_TOOLS

SENSITIVE_KEYS = {
    "password",
    "passphrase",
    "secret",
    "token",
    "api_key",
    "apikey",
    "authorization",
    "cookie",
    "value",
    "text",
    "reference_image",
}

SENSITIVE_TOOLS = {
    "clipboard_copy",
    "clipboard_paste",
    "fill_field",
    "type_text",
    "uia_set_control_value",
}


def enforce_tool_allowed(tool_name: str, arguments: dict[str, Any] | None = None) -> None:
    """Raise if active policy blocks this tool."""
    security = config.security
    args = arguments or {}

    if security.allowed_tools and tool_name not in security.allowed_tools:
        raise ToolError(
            f"Tool '{tool_name}' is not in WIN32_MCP_ALLOWED_TOOLS",
            suggestion="Add it to WIN32_MCP_ALLOWED_TOOLS or clear the allowlist.",
        )

    if tool_name in security.blocked_tools:
        raise ToolError(
            f"Tool '{tool_name}' is blocked by WIN32_MCP_BLOCKED_TOOLS",
            suggestion="Remove it from WIN32_MCP_BLOCKED_TOOLS if this action is expected.",
        )

    if security.profile == "read_only" and tool_name not in READ_ONLY_TOOLS and tool_name != "execute_sequence":
        raise ToolError(
            f"Tool '{tool_name}' is blocked by read_only security profile",
            suggestion="Set WIN32_MCP_SECURITY_PROFILE=interactive or unrestricted to allow OS mutation.",
        )

    if security.profile == "interactive" and tool_name in HIGH_RISK_TOOLS:
        raise ToolError(
            f"Tool '{tool_name}' is blocked by interactive security profile",
            suggestion="Set WIN32_MCP_SECURITY_PROFILE=unrestricted only in trusted environments.",
        )

    if tool_name in HIGH_RISK_TOOLS and security.confirmation_token:
        provided = str(args.get("confirmation_token", ""))
        if provided != security.confirmation_token:
            raise ToolError(
                f"Tool '{tool_name}' requires confirmation_token",
                suggestion="Provide the configured WIN32_MCP_CONFIRMATION_TOKEN value for high-risk actions.",
            )


def is_mutating_tool(tool_name: str) -> bool:
    """Return True if a tool can mutate local OS state."""
    return tool_name in MUTATING_TOOLS


def should_dry_run(tool_name: str, arguments: dict[str, Any]) -> bool:
    """Return True if this mutating call should be simulated."""
    if not is_mutating_tool(tool_name):
        return False
    if config.security.dry_run:
        return True
    value = arguments.get("dry_run", False)
    if not isinstance(value, bool):
        raise ToolError("Argument 'dry_run' must be a boolean")
    return value


def enforce_command_allowed(command: str) -> None:
    """Raise if process command violates command allow/block lists."""
    security = config.security
    normalized = _normalize_command(command)

    blocked = {_normalize_command(item) for item in security.blocked_commands}
    if normalized in blocked:
        raise ToolError(
            f"Command '{command}' is blocked by WIN32_MCP_BLOCKED_COMMANDS",
            suggestion="Remove it from the blocklist if this launch is expected.",
        )

    if security.allowed_commands:
        allowed = {_normalize_command(item) for item in security.allowed_commands}
        if normalized not in allowed:
            raise ToolError(
                f"Command '{command}' is not in WIN32_MCP_ALLOWED_COMMANDS",
                suggestion="Add the executable name/path to WIN32_MCP_ALLOWED_COMMANDS.",
            )


def _normalize_command(command: str) -> str:
    raw = command.strip().strip('"').strip("'")
    name = Path(raw).name if ("\\" in raw or "/" in raw) else raw
    return name.lower()


def redact_text(text: str, *, preserve_length: bool = True) -> str:
    """Return safe placeholder for sensitive text."""
    if not config.security.redact_sensitive_output:
        return text
    if preserve_length:
        return f"[redacted {len(text)} chars]"
    return "[redacted]"


def truncate_text(text: str, max_chars: int | None = None) -> str:
    """Bound text output and mark truncation."""
    limit = max_chars if max_chars is not None else config.limits.max_text_chars
    if len(text) <= limit:
        return text
    omitted = len(text) - limit
    return f"{text[:limit]}\n...[truncated {omitted} chars]"


def redact_arguments(tool_name: str, arguments: dict[str, Any]) -> dict[str, Any]:
    """Return redacted copy for audit logs."""
    if not arguments:
        return {}
    redacted = _redact_value(arguments)
    if tool_name in SENSITIVE_TOOLS and isinstance(redacted, dict):
        return redacted
    return redacted if isinstance(redacted, dict) else {}


def _redact_value(value: Any, key: str | None = None) -> Any:
    if key is not None and key.lower() in SENSITIVE_KEYS:
        if isinstance(value, str):
            return redact_text(value)
        return "[redacted]"
    if isinstance(value, dict):
        return {str(k): _redact_value(v, str(k)) for k, v in value.items()}
    if isinstance(value, list):
        return [_redact_value(item) for item in value]
    return value


def safe_json_dumps(value: Any, max_chars: int = 1_000) -> str:
    """Serialize for logs with hard size bound."""
    try:
        text = json.dumps(value, default=str)
    except TypeError:
        text = str(value)
    return truncate_text(text, max_chars=max_chars)
