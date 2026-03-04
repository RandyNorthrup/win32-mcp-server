"""
Backward-compatibility wrapper for win32-mcp-server.

All implementation has moved to the win32_mcp_server package.
This file is kept so that `python server.py` still works.
"""

from win32_mcp_server import main

if __name__ == "__main__":
    main()
