"""
MCP Server for Windows UI Inspection and Control - v2.6

Enterprise-grade automation server with 53 tools for screen capture, OCR,
mouse/keyboard control, window management, process control, UI Automation,
and intelligent high-level automation (click_text, wait_for_text, fill_field, etc.).

Author: Randy Northrup
GitHub: https://github.com/RandyNorthrup/win32-mcp-server
"""

import argparse
import asyncio
import json
import logging
import platform
import sys
from collections.abc import Sequence
from typing import Any

from mcp.server import Server
from mcp.types import ImageContent, TextContent, Tool

from . import __version__
from .config import config
from .registry import registry
from .utils.security import redact_arguments, safe_json_dumps

# ---------------------------------------------------------------------------
# Logging setup
# ---------------------------------------------------------------------------

logging.basicConfig(
    level=logging.DEBUG if config.debug else logging.INFO,
    format="%(asctime)s [%(name)s] %(levelname)s: %(message)s",
    handlers=[logging.StreamHandler(sys.stderr)],
)
logger = logging.getLogger("win32-mcp")

# ---------------------------------------------------------------------------
# Import all tool modules to trigger registration
# ---------------------------------------------------------------------------

from . import tools  # noqa: F401  — registers all tools

# ---------------------------------------------------------------------------
# Server-level tools (health_check is registered here)
# ---------------------------------------------------------------------------


@registry.register(
    "health_check",
    "Verify all dependencies and report server status",
    {
        "type": "object",
        "properties": {},
    },
)
async def handle_health_check(arguments: dict[str, Any]) -> dict[str, Any]:
    from .utils.coordinates import get_all_monitors, get_scaling_factor, get_system_dpi
    from .utils.imaging import check_tesseract

    status: dict[str, Any] = {
        "server_version": __version__,
        "python_version": sys.version.split()[0],
        "platform": platform.platform(),
        "architecture": platform.machine(),
        "security_profile": config.security.profile,
        "result_envelope": config.result_envelope,
        "coordinate_validation": config.validate_coordinates,
        "pyautogui_failsafe": config.automation.pyautogui_failsafe,
        "dry_run": config.security.dry_run,
        "tool_policy": {
            "allowed_tools": len(config.security.allowed_tools),
            "blocked_tools": len(config.security.blocked_tools),
            "allowed_commands": len(config.security.allowed_commands),
            "blocked_commands": len(config.security.blocked_commands),
            "confirmation_token_required": bool(config.security.confirmation_token),
        },
        "capture_defaults": {
            "format": config.capture.default_format,
            "quality": config.capture.default_quality,
            "scale": config.capture.default_scale,
        },
        "ocr_defaults": {
            "lang": config.ocr.lang,
            "preprocess": config.ocr.preprocess_mode,
            "tesseract_path_configured": bool(config.ocr.tesseract_path),
        },
    }

    # DPI / Scaling
    try:
        dpi = get_system_dpi()
        status["dpi"] = dpi
        status["display_scaling"] = f"{get_scaling_factor() * 100:.0f}%"
    except Exception as exc:
        status["dpi_error"] = str(exc)

    # Monitors
    try:
        monitors = get_all_monitors()
        status["monitor_count"] = len(monitors)
        if monitors:
            primary = monitors[0]
            status["primary_resolution"] = f"{primary['width']}x{primary['height']}"
    except Exception as exc:
        status["monitor_error"] = str(exc)

    # Tesseract OCR
    ok, msg = check_tesseract()
    status["tesseract"] = {"installed": ok, "info": msg}

    # Dependencies
    deps = {}
    for mod_name in [
        "mcp",
        "mss",
        "PIL",
        "pyautogui",
        "pygetwindow",
        "pyperclip",
        "pytesseract",
        "psutil",
        "numpy",
        "uiautomation",
    ]:
        try:
            mod = __import__(mod_name)
            ver = getattr(mod, "__version__", getattr(mod, "VERSION", "installed"))
            deps[mod_name] = str(ver)
        except ImportError:
            deps[mod_name] = "NOT INSTALLED"
    status["dependencies"] = deps

    # Rapidfuzz (optional)
    try:
        import rapidfuzz

        deps["rapidfuzz"] = rapidfuzz.__version__
    except ImportError:
        deps["rapidfuzz"] = "not installed (using difflib fallback)"

    # Tool count
    status["registered_tools"] = len(registry.tool_names)
    status["tools"] = registry.tool_names

    return status


# ---------------------------------------------------------------------------
# MCP Server Instance
# ---------------------------------------------------------------------------

app = Server("win32-inspector")


@app.list_tools()  # type: ignore[untyped-decorator, no-untyped-call]
async def list_tools() -> list[Tool]:
    """Return all registered tool definitions."""
    return registry.get_tools()


@app.call_tool()  # type: ignore[untyped-decorator]
async def call_tool(name: str, arguments: Any) -> list[TextContent | ImageContent]:
    """Dispatch a tool call through the registry."""
    args = arguments if isinstance(arguments, dict) else {}
    logger.debug("Tool call: %s(%s)", name, safe_json_dumps(redact_arguments(name, args), max_chars=500))
    return await registry.dispatch(name, args)


# ---------------------------------------------------------------------------
# Entry Points
# ---------------------------------------------------------------------------


async def async_main() -> None:
    """Async entry point — runs the MCP server over stdio."""
    from mcp.server.stdio import stdio_server

    logger.info("win32-mcp-server v%s starting (%d tools registered)", __version__, len(registry.tool_names))

    async with stdio_server() as (read_stream, write_stream):
        await app.run(
            read_stream,
            write_stream,
            app.create_initialization_options(),
        )


def _build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Windows Automation Inspector MCP server")
    parser.add_argument("--version", action="store_true", help="Print server version and exit")
    parser.add_argument("--list-tools", action="store_true", help="Print registered tool names as JSON and exit")
    parser.add_argument("--health-check", action="store_true", help="Run health_check once as JSON and exit")
    return parser


def main(argv: Sequence[str] | None = None) -> None:
    """Synchronous entry point for console_scripts."""
    args = _build_arg_parser().parse_args(argv)
    if args.version:
        sys.stdout.write(f"{__version__}\n")
        return
    if args.list_tools:
        sys.stdout.write(f"{json.dumps(registry.tool_names, indent=2)}\n")
        return
    if args.health_check:
        sys.stdout.write(f"{json.dumps(asyncio.run(handle_health_check({})), indent=2, default=str)}\n")
        return

    try:
        asyncio.run(async_main())
    except KeyboardInterrupt:
        logger.info("Server stopped by user")
    except Exception as exc:
        logger.critical("Server crashed: %s", exc, exc_info=True)
        sys.exit(1)
