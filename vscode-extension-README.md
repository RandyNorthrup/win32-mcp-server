# Windows Automation Inspector (MCP)

Comprehensive Windows automation MCP server for VS Code with 25+ powerful tools.

## Features

- 📸 **Screen Capture** - Full screen and window-specific screenshots
- 🔍 **OCR** - Extract text from screen regions using Tesseract
- 🖱️ **Mouse Control** - Click, drag, move, scroll with precision
- ⌨️ **Keyboard Control** - Type text, press keys, execute hotkeys
- 📋 **Clipboard** - Copy and paste operations
- 🪟 **Window Management** - Focus, close, minimize, maximize, resize, move windows
- 🔧 **Process Management** - List and kill processes with PID filtering

## Installation

This extension automatically installs the required Python package `win32-mcp-server`.

### Prerequisites

1. **Python 3.10+** must be installed and in PATH
2. **Tesseract OCR** for text recognition:
   - Download: https://github.com/UB-Mannheim/tesseract/wiki
   - Add to PATH or install to default location

### Quick Start

1. Install this extension from VS Code Marketplace
2. Reload VS Code when prompted
3. The extension will automatically install Python dependencies
4. The MCP server will be available to GitHub Copilot

### Manual Installation

If automatic installation fails:

```bash
pip install win32-mcp-server
```

## Configuration

The extension adds the following settings:

- `win32-mcp.enabled` - Enable/disable the MCP server (default: true)
- `win32-mcp.autoInstall` - Automatically install Python package (default: true)

## Usage

Once installed, the MCP server is automatically available to GitHub Copilot and other MCP clients in VS Code.

### Example Commands

- "Capture screenshot of the window titled 'Chrome'"
- "Extract text from the screen using OCR"
- "Click at coordinates (500, 300)"
- "Type 'Hello World' at the cursor"
- "List all open windows"
- "Maximize the Notepad window"
- "List all running processes"

## Available Tools (25)

| Category | Tools |
|----------|-------|
| Screen Capture | `capture_screen`, `capture_window` |
| OCR | `ocr_screen`, `ocr_region` |
| Mouse | `click`, `double_click`, `drag`, `mouse_position`, `mouse_move`, `scroll` |
| Keyboard | `type_text`, `press_key`, `hotkey` |
| Clipboard | `clipboard_copy`, `clipboard_paste` |
| Windows | `list_windows`, `focus_window`, `close_window`, `minimize_window`, `maximize_window`, `restore_window`, `resize_window`, `move_window` |
| Processes | `list_processes`, `kill_process` |

## Troubleshooting

### Extension Not Working

1. Check Python is installed: `python --version`
2. Verify package is installed: `pip show win32-mcp-server`
3. Check Tesseract is installed: `tesseract --version`
4. Reload VS Code: `Ctrl+Shift+P` → "Developer: Reload Window"

### Manual MCP Configuration

If the extension doesn't auto-configure, add to `%APPDATA%\Code\User\mcp.json`:

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

## Security

This extension provides powerful system automation capabilities. Only use in trusted environments.

## Support

- **Issues**: https://github.com/RandyNorthrup/win32-mcp-server/issues
- **Documentation**: https://github.com/RandyNorthrup/win32-mcp-server#readme

## License

MIT License - see [LICENSE](https://github.com/RandyNorthrup/win32-mcp-server/blob/main/LICENSE)
