"""
Win32 MCP Server — Comprehensive Windows automation for AI agents.

Enterprise-grade MCP server with 40+ tools for screen capture, OCR,
mouse/keyboard control, window management, process control, and
intelligent high-level automation.

Author: Randy Northrup
GitHub: https://github.com/RandyNorthrup/win32-mcp-server
"""

__version__ = "2.0.0"

from .server import main

__all__ = ["main", "__version__"]
