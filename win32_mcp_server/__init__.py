"""
Win32 MCP Server — Comprehensive Windows automation for AI agents.

Enterprise-grade MCP server with 53 tools for screen capture, OCR,
mouse/keyboard control, window management, process control, UI Automation,
and intelligent high-level automation.

Author: Randy Northrup
GitHub: https://github.com/RandyNorthrup/win32-mcp-server
"""

import sys
from collections.abc import Sequence

__version__ = "2.6.0"


def main(argv: Sequence[str] | None = None) -> None:
    """Lazy console entry point; avoids Win32 side effects on package import."""
    args = list(sys.argv[1:] if argv is None else argv)
    if args == ["--version"]:
        sys.stdout.write(f"{__version__}\n")
        return

    from .server import main as _server_main

    _server_main(args)


__all__ = ["__version__", "main"]
