# Windows MCP Server

**Comprehensive Windows automation MCP server for AI agents**

Full control over Windows desktop applications with 25+ tools: screenshots, OCR, mouse/keyboard control, window management, process control, clipboard operations, and more.

## Features

### Screen Capture
- Full screen screenshots
- Window-specific capture
- Region capture

### OCR (Optical Character Recognition)
- Full screen text extraction
- Region-based OCR
- Powered by Tesseract

### Mouse Control
- Click (left/right/middle)
- Double-click
- Drag and drop
- Mouse movement with duration
- Scroll (up/down)
- Get mouse position

### Keyboard Control
- Type text with configurable speed
- Press individual keys
- Execute hotkey combinations (Ctrl+C, Alt+F4, etc.)
- Full keyboard shortcuts support

### Clipboard
- Copy text to clipboard
- Paste/read clipboard content
- Seamless clipboard integration

### Window Management
- List all open windows
- Focus/activate windows
- Close windows
- Minimize/maximize/restore
- Resize windows
- Move windows
- Get window details (position, size, state)

### Process Management
- List running processes with PIDs
- Filter processes by name
- Kill processes by PID
- Memory usage monitoring

## Installation

### Prerequisites

1. **Python 3.10+** installed
2. **Tesseract OCR** for text recognition:
   - Download: https://github.com/UB-Mannheim/tesseract/wiki
   - Install to default location or add to PATH
   - Verify: `tesseract --version`

### Install Package

**Option 1: Install from PyPI (Recommended)**

```bash
pip install win32-mcp-server
```

**Option 2: Install from GitHub**

```bash
pip install git+https://github.com/RandyNorthrup/win32-mcp-server.git
```

**Option 3: Install from source**

```bash
# Clone repository
git clone https://github.com/RandyNorthrup/win32-mcp-server.git
cd win32-mcp-server

# Install with dependencies
pip install -e .
```

## Configuration

### VS Code with GitHub Copilot

After installing via pip, add to your MCP configuration (`%APPDATA%\Code\User\mcp.json`):

```json
{
  "servers": {
    "win32-inspector": {
      "type": "stdio",
      "command": "win32-mcp-server"
    }
  }
}
```

**Or install from VS Code MCP Extensions:**
1. Open VS Code
2. Press `Ctrl+Shift+P`
3. Type "MCP: Install Server"
4. Search for "Windows Automation Inspector"
5. Click Install

### Claude Desktop

After installing via pip, add to `%APPDATA%\Claude\claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "win32-inspector": {
      "command": "win32-mcp-server"
    }
  }
}
```

### Other MCP Clients

The server uses **STDIO transport** and works with any MCP-compatible client that supports stdio.

## Usage Examples

### Capture Screenshot
```
"Capture screenshot of the window titled 'Compliance Guard'"
```

### OCR Text Extraction
```
"Extract text from the screen using OCR"
"OCR the region at x=100, y=100, width=500, height=300"
```

### Automate UI Interactions
```
"Click at coordinates (500, 300)"
"Double-click the button at (450, 250)"
"Drag from (100, 100) to (500, 500)"
```

### Keyboard Automation
```
"Type 'Hello World' at the current cursor position"
"Press Ctrl+C to copy"
"Execute Alt+F4 to close the window"
```

### Window Management
```
"List all open windows"
"Focus the window titled 'Visual Studio Code'"
"Maximize the Chrome window"
"Resize Notepad to 800x600"
```

### Process Control
```
"List all running processes"
"Show processes containing 'chrome'"
"Kill process with PID 1234"
```

## Available Tools

| Tool | Description |
|------|-------------|
| `capture_screen` | Capture full screen screenshot |
| `capture_window` | Capture specific window by title |
| `list_windows` | List all open windows with details |
| `ocr_screen` | Extract text from full screen |
| `ocr_region` | Extract text from specified region |
| `click` | Click at coordinates (left/right/middle) |
| `double_click` | Double-click at coordinates |
| `drag` | Drag from start to end coordinates |
| `type_text` | Type text at current position |
| `press_key` | Press keyboard key or shortcut |
| `hotkey` | Execute hotkey combination |
| `clipboard_copy` | Copy text to clipboard |
| `clipboard_paste` | Get clipboard content |
| `mouse_position` | Get current mouse position |
| `mouse_move` | Move mouse to position |
| `scroll` | Scroll up/down |
| `list_processes` | List running processes with PIDs |
| `kill_process` | Terminate process by PID |
| `focus_window` | Activate window |
| `close_window` | Close window by title |
| `minimize_window` | Minimize window |
| `maximize_window` | Maximize window |
| `restore_window` | Restore window |
| `resize_window` | Resize window |
| `move_window` | Move window position |

## Security Considerations

**WARNING**: This server has powerful system control capabilities including:
- Mouse and keyboard control
- Process termination
- Clipboard access
- Screen capture

**Only use in trusted environments** where you control the MCP client.

### Recommended Security Practices

1. **Restrict Usage**: Only enable when actively needed
2. **Review Logs**: Monitor all automated actions
3. **Sandbox Testing**: Test in isolated environments first
4. **Access Control**: Limit who can access the MCP client
5. **Disable PyAutoGUI Failsafe**: Server disables failsafe for automation - be cautious

## Troubleshooting

### Tesseract Not Found
```
TesseractNotFoundError: tesseract is not installed
```
**Solution**: Install Tesseract OCR from https://github.com/UB-Mannheim/tesseract/wiki

### Permission Errors
```
PermissionError: [WinError 5] Access is denied
```
**Solution**: Run VS Code or MCP client as Administrator for process control features

### Module Not Found
```
ModuleNotFoundError: No module named 'mcp'
```
**Solution**: Reinstall dependencies: `pip install -e .`

### Window Not Found
```
Window not found: [title]
```
**Solution**: Use partial window title matching. Check exact title with `list_windows` first.

## Development

### Project Structure
```
win32-mcp-server/
├── server.py          # Main MCP server implementation
├── pyproject.toml     # Package configuration
├── README.md          # This file
└── LICENSE            # MIT License
```

### Dependencies
- **mcp**: Model Context Protocol SDK
- **mss**: Cross-platform screen capture
- **Pillow**: Image processing
- **pyautogui**: Mouse and keyboard automation
- **pygetwindow**: Window management
- **pyperclip**: Clipboard operations
- **pytesseract**: OCR text extraction
- **psutil**: Process management

## License

MIT License - see LICENSE file

## Contributing

Contributions welcome! Please:
1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Submit a pull request

## Links

- **Repository**: https://github.com/RandyNorthrup/win32-mcp-server
- **Issues**: https://github.com/RandyNorthrup/win32-mcp-server/issues
- **MCP Documentation**: https://modelcontextprotocol.io/

## Support

For bugs and feature requests, please use [GitHub Issues](https://github.com/RandyNorthrup/win32-mcp-server/issues).

---

**Made for Windows automation and AI agents**
