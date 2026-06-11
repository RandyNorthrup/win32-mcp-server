# Win32 MCP Server

**Enterprise-grade Windows automation for AI agents — 53 tools over MCP**

The most comprehensive Windows desktop automation server for the [Model Context Protocol](https://modelcontextprotocol.io/). Give any MCP-compatible AI agent full control over Windows applications: intelligent text finding and clicking, structured OCR, screenshot capture, mouse/keyboard input, window management, process control, and multi-step batch operations — all through a single MCP server.

[![Version](https://img.shields.io/badge/version-2.6.1-blue)](https://github.com/RandyNorthrup/win32-mcp-server/releases)
[![PyPI](https://img.shields.io/pypi/v/win32-mcp-server)](https://pypi.org/project/win32-mcp-server/)
[![VS Code Marketplace](https://img.shields.io/badge/VS%20Code-Marketplace-007ACC)](https://marketplace.visualstudio.com/items?itemName=RandyNorthrup.win32-mcp-inspector)
[![Python](https://img.shields.io/badge/python-3.10%2B-green)](https://python.org)
[![License](https://img.shields.io/badge/license-MIT-yellow)](LICENSE)
[![MCP](https://img.shields.io/badge/MCP-compatible-purple)](https://modelcontextprotocol.io/)

---

## What's New in v2.6

- **Security profiles** — `read_only`, `interactive`, allow/block lists, dry-run, and confirmation tokens
- **Central execution controls** — serialized OS mutation, hard timeouts, bounded subprocess output
- **Typed argument validation** — consistent bounds for mouse, keyboard, clipboard, window, capture, OCR, process, UIA, and smart tools
- **UIA-first smart actions** — `click_text` and `fill_field` try control-based automation before OCR fallback
- **Stable response option** — `WIN32_MCP_RESULT_ENVELOPE` for `{success, tool, data}` JSON dict responses
- **Advertised safety args** — mutating tool schemas expose `dry_run`; high-risk tool schemas expose `confirmation_token`
- **CI and release hygiene** — Windows GitHub Actions, contract tests, `SECURITY.md`, `CONTRIBUTING.md`, wheel package-data checks
- **VS Code extension hardening** — shell-free commands, pinned package version checks, custom Tesseract path support

---

## Features

### Smart Automation (the most powerful tools)
| Tool | Description |
|------|-------------|
| `click_text` | Find text on screen and click it — no coordinates needed |
| `find_text_on_screen` | Locate all occurrences of text with screen coordinates |
| `wait_for_text` | Poll until text appears on screen (with timeout) |
| `assert_text_visible` | Verify text is or is not visible (for UI testing) |
| `fill_field` | Click a labeled input field and type a value |
| `get_window_snapshot` | Screenshot + structured OCR in one call |
| `right_click_menu` | Right-click and OCR the context menu items |
| `execute_sequence` | Run up to 50 tools in sequence without round-trips |

### Screen Capture (6 tools)
- Full screen, per-window, and per-monitor capture
- PNG, JPEG, and WebP output with quality/scale controls
- Pixel color sampling at any coordinate
- Screenshot comparison with similarity metrics

### OCR — Optical Character Recognition (5 tools)
- Full screen and region-based text extraction
- Per-window OCR with automatic focus and capture
- **Structured mode** — every word with bounding box, confidence, line/block/word numbers
- Intelligent preprocessing: auto-detects light/dark backgrounds
- Coordinates map back to original screen space for accurate clicking

### Mouse Control (8 tools)
- Click, double-click, triple-click (left/right/middle buttons)
- Drag-and-drop with configurable duration and button
- Mouse move with smooth animation
- Vertical and horizontal scrolling at any position
- Current position reporting

### Keyboard Control (3 tools)
- Type text with Unicode support (auto-fallback to clipboard paste)
- Press individual keys or key combinations (`ctrl+c`, `alt+f4`)
- Execute hotkey combos from arrays (`["ctrl", "shift", "s"]`)

### Clipboard (2 tools)
- Copy text to system clipboard
- Read current clipboard contents

### Window Management (10 tools)
- List all windows with fuzzy title filtering
- Detailed window info (PID, position, size, state, process name, memory)
- Focus, close, minimize, maximize, restore
- Resize and move to exact coordinates
- Wait for a window to appear (polling with timeout)
- Fuzzy matching with intelligent suggestions on miss

### Process Management (4 tools)
- List processes with filtering, sorting, and pagination
- Graceful termination with force-kill fallback
- Launch applications with optional wait-for-completion
- Wait for a process to become idle (CPU threshold monitoring)

### System (1 tool)
- `health_check` — verify all dependencies, DPI, monitors, Tesseract, and tool count

### UI Automation (6 tools)
- `uia_inspect_window` — get the control tree of a window
- `uia_find_control` — find controls by type, name, or automation ID
- `uia_click_control` — click a control by name (more reliable than coordinates)
- `uia_get_control_value` — read a control's value or text
- `uia_set_control_value` — set a control's value (edit boxes, etc.)
- `uia_get_focused` — get info about the currently focused control

---

## Installation

### Prerequisites

1. **Python 3.10+**
2. **Tesseract OCR** (optional — required only for OCR tools):
   - Download: https://github.com/UB-Mannheim/tesseract/wiki
   - Install and ensure it's on PATH
   - Verify: `tesseract --version`

### Install Package

**From PyPI (recommended):**
```bash
pip install win32-mcp-server
```

**From GitHub (latest unreleased):**
```bash
pip install git+https://github.com/RandyNorthrup/win32-mcp-server.git
```

**From source:**
```bash
git clone https://github.com/RandyNorthrup/win32-mcp-server.git
cd win32-mcp-server
pip install -e .
```

---

## Configuration

### VS Code with GitHub Copilot

Add to your MCP configuration (`%APPDATA%\Code\User\mcp.json`):

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

Or install from the VS Code Marketplace — search **"Windows Automation Inspector"**.

### Claude Desktop

Add to `%APPDATA%\Claude\claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "win32-inspector": {
      "command": "win32-mcp-server"
    }
  }
}
```

### Any MCP Client

The server uses **STDIO transport** and works with any MCP-compatible client.

### CLI Smoke Checks

```powershell
python -m win32_mcp_server --version
python -m win32_mcp_server --list-tools
python -m win32_mcp_server --health-check
```

### Runtime Environment Settings

Set these in your MCP client's server environment when you need production controls:

| Variable | Purpose |
|----------|---------|
| `WIN32_MCP_SECURITY_PROFILE` | `interactive` (default), `read_only`, or `unrestricted` |
| `WIN32_MCP_ALLOWED_TOOLS` / `WIN32_MCP_BLOCKED_TOOLS` | Comma-separated tool allow/block lists |
| `WIN32_MCP_ALLOWED_COMMANDS` / `WIN32_MCP_BLOCKED_COMMANDS` | Comma-separated process command allow/block lists |
| `WIN32_MCP_DRY_RUN` | Simulate mutating tools without touching the OS |
| `WIN32_MCP_CONFIRMATION_TOKEN` | Require a token argument for high-risk tools (`start_process`, `kill_process`, `close_window`) |
| `WIN32_MCP_REDACT_SENSITIVE_OUTPUT` | Redact typed, pasted, and clipboard values in responses/logs |
| `WIN32_MCP_RESULT_ENVELOPE` | Opt into stable JSON dict responses: `{success, tool, data}` |
| `WIN32_MCP_TESSERACT_PATH` | Explicit Tesseract executable path. Usually not needed on Windows; common install paths are auto-detected. |
| `WIN32_MCP_OCR_LANGUAGE` / `WIN32_MCP_OCR_PREPROCESS` | OCR language and preprocessing defaults |
| `WIN32_MCP_CAPTURE_FORMAT` / `WIN32_MCP_CAPTURE_QUALITY` / `WIN32_MCP_CAPTURE_SCALE` | Screenshot defaults |
| `WIN32_MCP_MAX_TIMEOUT_SECONDS` / `WIN32_MCP_MAX_TOOL_RUNTIME_SECONDS` / `WIN32_MCP_MAX_SUBPROCESS_OUTPUT_BYTES` | Runtime resource bounds |
| `WIN32_MCP_COORDINATE_VALIDATION` | Validate coordinates against actual monitor rectangles |
| `WIN32_MCP_PYAUTOGUI_FAILSAFE` | Keep PyAutoGUI fail-safe enabled by default; set `false` only for controlled automation rigs |

---

## Usage Examples

### Smart Automation (Natural Language)
```
"Click the 'Submit' button"
"Wait for 'Loading complete' to appear, then click 'Continue'"
"Fill in the 'Username' field with 'admin@example.com'"
"Take a snapshot of the Chrome window and tell me what you see"
"Right-click the desktop and show me the menu options"
```

### Screen Capture
```
"Capture a screenshot of the entire screen"
"Capture the Notepad window as a compressed JPEG at 50% scale"
"Compare the current screen to this reference image"
```

### OCR
```
"Extract all text from the screen"
"Get structured OCR data from the region at (100, 200) size 800x600"
"Read all text in the Chrome window"
```

### Mouse & Keyboard
```
"Click at (500, 300) with the right mouse button"
"Drag from (100, 100) to (500, 500)"
"Type 'Hello World' — use clipboard paste for Unicode characters"
"Press Ctrl+Shift+S"
```

### Window & Process Management
```
"List all open windows containing 'Visual Studio'"
"Maximize the Chrome window"
"Resize Notepad to 800x600 and move it to (0, 0)"
"Wait for a window titled 'Installation Complete' to appear"
"List the top 20 processes by memory usage"
"Kill process with PID 1234"
```

### Batch Operations
```
"Execute this sequence: click (100,100), wait 500ms, type 'hello', press Enter"
```

---

## All 53 Tools

### Smart Automation
| Tool | Description |
|------|-------------|
| `click_text` | Find text on screen and click it |
| `find_text_on_screen` | Find all text occurrences with coordinates |
| `wait_for_text` | Wait until text appears (polling) |
| `assert_text_visible` | Assert text is/isn't visible |
| `fill_field` | Click labeled field and type value |
| `get_window_snapshot` | Screenshot + OCR in one call |
| `right_click_menu` | Right-click and OCR the menu |
| `execute_sequence` | Batch up to 50 tool calls |

### Screen Capture
| Tool | Description |
|------|-------------|
| `capture_screen` | Full screen screenshot (PNG/JPEG/WebP) |
| `capture_window` | Window screenshot with fuzzy title match |
| `capture_monitor` | Capture specific monitor by index |
| `list_monitors` | List monitors with resolution and DPI |
| `get_pixel_color` | Get RGB/hex color at coordinates |
| `compare_screenshots` | Pixel-level comparison with similarity score |

### OCR
| Tool | Description |
|------|-------------|
| `ocr_screen` | Full screen text extraction |
| `ocr_region` | Region text extraction |
| `ocr_window` | Window text extraction |
| `ocr_screen_structured` | Full screen OCR with bounding boxes |
| `ocr_region_structured` | Region OCR with bounding boxes |

### Mouse
| Tool | Description |
|------|-------------|
| `click` | Click at coordinates (left/right/middle, N clicks) |
| `double_click` | Double-click at coordinates |
| `triple_click` | Triple-click to select line/paragraph |
| `drag` | Drag from start to end with duration |
| `mouse_position` | Get current cursor position |
| `mouse_move` | Move cursor with smooth animation |
| `scroll` | Vertical scroll at position |
| `scroll_horizontal` | Horizontal scroll at position |

### Keyboard
| Tool | Description |
|------|-------------|
| `type_text` | Type text (auto Unicode detection, clipboard fallback) |
| `press_key` | Press key or combo (`ctrl+c`, `alt+f4`) |
| `hotkey` | Hotkey from key array (`["ctrl","shift","s"]`) |

### Clipboard
| Tool | Description |
|------|-------------|
| `clipboard_copy` | Copy text to clipboard |
| `clipboard_paste` | Read clipboard contents |

### Window Management
| Tool | Description |
|------|-------------|
| `list_windows` | List windows with optional title filter |
| `get_window_info` | Detailed window info (PID, process, memory) |
| `focus_window` | Bring window to foreground |
| `close_window` | Close window by title |
| `minimize_window` | Minimize window |
| `maximize_window` | Maximize window |
| `restore_window` | Restore from minimized/maximized |
| `resize_window` | Resize to exact dimensions |
| `move_window` | Move to exact position |
| `wait_for_window` | Wait for window to appear (polling) |

### Process Management
| Tool | Description |
|------|-------------|
| `list_processes` | List processes (filter, sort, paginate) |
| `kill_process` | Terminate process (graceful + force fallback) |
| `start_process` | Launch application with optional wait |
| `wait_for_idle` | Wait for process CPU to drop below threshold |

### System
| Tool | Description |
|------|-------------|
| `health_check` | Full dependency and system status report |

### UI Automation
| Tool | Description |
|------|-------------|
| `uia_inspect_window` | Inspect control tree (buttons, edits, etc.) |
| `uia_find_control` | Find controls by name, automation ID, or type |
| `uia_click_control` | Click a control by name (no coordinates needed) |
| `uia_get_control_value` | Read a control's value/text |
| `uia_set_control_value` | Set a control's value |
| `uia_get_focused` | Get info about the focused control |

---

## Architecture

```
win32-mcp-server/
├── win32_mcp_server/
│   ├── __init__.py              # Package entry, version
│   ├── __main__.py              # python -m support
│   ├── config.py                # Dataclass config, PreprocessMode
│   ├── registry.py              # Decorator-based tool registry + dispatch
│   ├── server.py                # MCP server, stdio transport, health_check
│   ├── utils/
│   │   ├── args.py              # Typed tool argument readers and bounds
│   │   ├── coordinates.py       # DPI awareness, screen geometry, validation
│   │   ├── errors.py            # ToolError with suggestions
│   │   ├── execution.py         # Timeouts, dry-run, serialized mutations
│   │   ├── imaging.py           # Image preprocessing, encoding, diffing
│   │   ├── security.py          # Safety policy, redaction, command controls
│   │   └── window_match.py      # Fuzzy title matching, deduplication, PID
│   └── tools/
│       ├── capture.py           # Screenshot tools (6)
│       ├── ocr.py               # OCR tools (5)
│       ├── mouse.py             # Mouse tools (8)
│       ├── keyboard.py          # Keyboard tools (3)
│       ├── clipboard.py         # Clipboard tools (2)
│       ├── window.py            # Window management tools (10)
│       ├── process.py           # Process management tools (4)
│       ├── smart.py             # Smart automation tools (8)
│       └── uia.py               # UI Automation API tools (6)
├── extension.js                 # VS Code extension bootstrap
├── tests/                       # Safety, registry, and contract tests
├── .github/workflows/ci.yml     # Windows CI quality gate
├── package.json                 # VS Code extension manifest
├── pyproject.toml               # Python package config
└── LICENSE                      # MIT License
```

---

## Security Considerations

> **This server has powerful system control capabilities.** Only use in trusted environments where you control the MCP client.

The server can:
- Capture screenshots of any window or the entire desktop
- Read and write the system clipboard
- Control mouse and keyboard input
- Terminate processes
- Launch applications

### Recommended Practices

1. **Start restrictive** — default `interactive` blocks high-risk process/window actions; use `read_only` for observation-only sessions
2. **Use allow/block lists** — set `WIN32_MCP_ALLOWED_TOOLS`, `WIN32_MCP_BLOCKED_TOOLS`, and command lists for production
3. **Dry-run first** — use `WIN32_MCP_DRY_RUN=true` or per-call `dry_run: true` before granting mutating control
4. **Require confirmation tokens** — set `WIN32_MCP_CONFIRMATION_TOKEN` for high-risk process/window actions
5. **Review automation logs** — tool calls are logged with sensitive fields redacted
6. **Restrict MCP client access** — only trusted clients should invoke this server

---

## Quality Gates

Run the same checks as CI before shipping changes:

```powershell
python -m pip install -e .[dev]
python -m ruff check win32_mcp_server tests
python -m mypy win32_mcp_server
python -m pytest -q
node --check extension.js
```

## Release Publishing

GitHub release assets are published from local build output. PyPI publishing runs through the **Publish PyPI** GitHub Actions workflow using PyPI Trusted Publishing. Configure the PyPI project for repository `RandyNorthrup/win32-mcp-server`, workflow `publish-pypi.yml`, environment `pypi`, then run the workflow with the target release tag.

---

## Troubleshooting

| Problem | Solution |
|---------|----------|
| `TesseractNotFoundError` | Install with `winget install --id tesseract-ocr.tesseract --exact`; PATH is optional for common Windows installs |
| `PermissionError: Access is denied` | Run VS Code / MCP client as Administrator |
| `ModuleNotFoundError: No module named 'mcp'` | `pip install -e .` or `pip install win32-mcp-server` |
| `Window not found: [title]` | Use partial title. Run `list_windows` to see exact titles. Fuzzy matching is automatic. |
| OCR returns empty/garbled text | Try `preprocess: "dark_bg"` or `"high_contrast"` for better results |
| Coordinates are wrong on HiDPI | The server auto-enables DPI awareness. Run `health_check` to verify DPI settings. |

---

## Dependencies

| Package | Purpose |
|---------|---------|
| [mcp](https://github.com/modelcontextprotocol/python-sdk) | Model Context Protocol SDK |
| [mss](https://github.com/BoboTiG/python-mss) | Fast cross-platform screen capture |
| [Pillow](https://github.com/python-pillow/Pillow) | Image processing and encoding |
| [NumPy](https://numpy.org/) | Image preprocessing for OCR |
| [PyAutoGUI](https://github.com/asweigart/pyautogui) | Mouse and keyboard automation |
| [PyGetWindow](https://github.com/asweigart/PyGetWindow) | Window enumeration and control |
| [pyperclip](https://github.com/asweigart/pyperclip) | Clipboard operations |
| [pytesseract](https://github.com/madmaze/pytesseract) | Tesseract OCR wrapper |
| [psutil](https://github.com/giampaolo/psutil) | Process management |
| [RapidFuzz](https://github.com/maxbachmann/RapidFuzz) | Fast fuzzy string matching |
| [uiautomation](https://github.com/yinkaisheng/Python-UIAutomation-for-Windows) | Windows UI Automation API |

---

## Contributing

Contributions welcome!

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/my-feature`)
3. Commit your changes (`git commit -am 'Add my feature'`)
4. Push to the branch (`git push origin feature/my-feature`)
5. Open a Pull Request

---

## License

MIT License — see [LICENSE](LICENSE) file.

## Links

- **Repository**: https://github.com/RandyNorthrup/win32-mcp-server
- **PyPI**: https://pypi.org/project/win32-mcp-server/
- **VS Code Marketplace**: https://marketplace.visualstudio.com/items?itemName=RandyNorthrup.win32-mcp-inspector
- **Issues**: https://github.com/RandyNorthrup/win32-mcp-server/issues
- **MCP Specification**: https://modelcontextprotocol.io/

---

**Author**: [Randy Northrup](https://github.com/RandyNorthrup)
**Built for Windows automation and AI agents**
