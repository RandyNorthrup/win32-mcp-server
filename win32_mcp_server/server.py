"""
MCP Server for Windows UI Inspection and Control - v2.5

Enterprise-grade automation server with 53 tools for screen capture, OCR,
mouse/keyboard control, window management, process control, UI Automation,
and intelligent high-level automation (click_text, wait_for_text, fill_field, etc.).

Author: Randy Northrup
GitHub: https://github.com/RandyNorthrup/win32-mcp-server
"""

import asyncio
import json
import logging
import platform
import sys
from typing import Any

from mcp.server import Server
from mcp.types import ImageContent, TextContent, Tool

from . import __version__
from .config import config
from .registry import registry

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


@app.list_tools()  # type: ignore[no-untyped-call]
async def list_tools() -> list[Tool]:
    """Return all registered tool definitions."""
    return registry.get_tools()


@app.call_tool()
async def call_tool(name: str, arguments: Any) -> list[TextContent | ImageContent]:
    """Dispatch a tool call through the registry."""
    args = arguments if isinstance(arguments, dict) else {}
    logger.debug("Tool call: %s(%s)", name, json.dumps(args, default=str)[:500])
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


def main() -> None:
    """Synchronous entry point for console_scripts."""
    try:
        asyncio.run(async_main())
    except KeyboardInterrupt:
        logger.info("Server stopped by user")
    except Exception as exc:
        logger.critical("Server crashed: %s", exc, exc_info=True)
        sys.exit(1)
