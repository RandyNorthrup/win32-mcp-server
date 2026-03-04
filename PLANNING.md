# win32-mcp-server — Enhancement & Fix Planning Document

**Author:** Randy Northrup  
**Date:** March 3, 2026  
**Current Version:** 1.0.1  
**Target Version:** 2.0.0  

---

## Executive Summary

This document outlines a comprehensive plan to transform win32-mcp-server from a basic Windows automation tool into a **reliable, accurate, and fast** desktop application testing server. The improvements focus on five pillars: **accuracy** (better OCR, coordinate handling), **reliability** (error handling, retries, waits), **usability** (high-level tools, smarter matching), **performance** (payload compression, caching), and **power** (UI Automation API, structured output).

---

## Table of Contents

1. [Critical Bugs & Fixes](#1-critical-bugs--fixes)
2. [OCR Accuracy & Intelligence](#2-ocr-accuracy--intelligence)
3. [High-Level Smart Tools](#3-high-level-smart-tools)
4. [Reliability & Robustness](#4-reliability--robustness)
5. [Performance Optimizations](#5-performance-optimizations)
6. [New Capabilities](#6-new-capabilities)
7. [Architecture & Code Quality](#7-architecture--code-quality)
8. [Extension Improvements](#8-extension-improvements)
9. [Implementation Roadmap](#9-implementation-roadmap)

---

## 1. Critical Bugs & Fixes

### 1.1 — No Error Handling in Tool Dispatch
**Problem:** Most tool handlers have no try/except. Any unhandled exception (e.g., window closed mid-operation, invalid coordinates, Tesseract not installed) crashes the server or returns an opaque MCP error.  
**Fix:** Wrap every tool handler in try/except, return structured error messages with actionable context (what failed, why, what to try).

### 1.2 — `type_text` Fails on Non-ASCII Characters
**Problem:** `pyautogui.write()` only handles ASCII. Typing `"café"`, `"über"`, or any unicode text silently fails or throws.  
**Fix:** Detect non-ASCII characters and fall back to clipboard-based typing (`pyperclip.copy(text)` → `pyautogui.hotkey('ctrl', 'v')`). Add a `method` parameter: `"type"` (keystroke) or `"paste"` (clipboard).

### 1.3 — `list_windows` Returns Duplicates
**Problem:** `gw.getAllTitles()` returns duplicate title strings. For each duplicate, the same window object is fetched, resulting in repeated entries.  
**Fix:** Deduplicate by window handle (HWND) instead of title string. Use `gw.getAllWindows()` which returns unique window objects.

### 1.4 — `capture_window` Race Condition
**Problem:** `win.activate()` followed by `asyncio.sleep(0.3)` is unreliable. The window may not be fully rendered/focused in 300ms, especially under load.  
**Fix:** Poll for window foreground state with exponential backoff (max 2s). Better yet, capture by window rect without requiring focus — `mss` can capture any screen region regardless of z-order.

### 1.5 — No DPI/Scaling Awareness
**Problem:** On displays with Windows scaling (125%, 150%, etc.), all coordinates are wrong. Screenshots are the wrong size, clicks land in the wrong place.  
**Fix:** Call `SetProcessDPIAware()` or `SetProcessDpiAwareness()` via ctypes at startup. Detect and report the current DPI scaling factor.

### 1.6 — Extension.js Dependency Check is Wrong
**Problem:** `python -c "import server"` checks for any module named `server`, not the actual package.  
**Fix:** Change to `python -c "import win32_mcp_server"` or `python -m win32_mcp_server --version`, or check `pip show win32-mcp-server`.

### 1.7 — `scroll` Has No Position Parameter
**Problem:** `pyautogui.scroll()` scrolls at the current mouse position, which may not be the intended target.  
**Fix:** Add optional `x`, `y` parameters. If provided, move mouse there first, scroll, then optionally return.

### 1.8 — `press_key` and `hotkey` Are Redundant
**Problem:** `press_key` already handles combos via `+` splitting. `hotkey` does the same thing with an array.  
**Fix:** Keep both for API flexibility but document clearly. Add validation for key names against pyautogui's key list.

### 1.9 — `list_processes` Hardcoded to 50 Results
**Problem:** `processes[:50]` silently truncates. Users can't find processes beyond the first 50.  
**Fix:** Add a `limit` parameter (default 100) and `offset` for pagination. Sort by memory/name. Return total count.

### 1.10 — Window Not Found Returns Useless Error
**Problem:** When `getWindowsWithTitle()` returns nothing, the error is just `"Window not found"`. No help finding the right title.  
**Fix:** Include fuzzy matches (windows with similar titles), and suggest using `list_windows`. Return the searched title for clarity.

---

## 2. OCR Accuracy & Intelligence

### 2.1 — Image Preprocessing Pipeline
**Problem:** Raw screenshots sent to Tesseract produce mediocre results, especially on dark themes, low contrast UIs, or small text.  
**Fix:** Add preprocessing before OCR:
- Convert to grayscale
- Apply adaptive thresholding (Otsu's method)
- Scale up small regions (2x-3x) before OCR
- Optional inversion for light-on-dark text
- Configurable preprocessing modes: `"auto"`, `"light_bg"`, `"dark_bg"`, `"high_contrast"`

### 2.2 — Structured OCR Output with Bounding Boxes
**Problem:** OCR returns flat text with no position info. The AI can't locate where text appears on screen to click on it.  
**Fix:** New tool `ocr_screen_structured` / `ocr_region_structured` using `pytesseract.image_to_data()`. Returns JSON array of `{text, x, y, width, height, confidence}` for each detected word/line.

### 2.3 — OCR Confidence Scoring
**Problem:** No way to know if OCR results are reliable.  
**Fix:** Include per-word confidence from Tesseract. Filter out low-confidence results (configurable threshold, default 60%).

### 2.4 — Window-Specific OCR
**Problem:** No tool to OCR just a specific window — must manually calculate region coordinates.  
**Fix:** New tool `ocr_window` that takes `window_title` and performs OCR on that window's region. Returns text with coordinates relative to the window.

### 2.5 — OCR Language Support
**Problem:** Hardcoded to English.  
**Fix:** Add `lang` parameter (default `"eng"`) to all OCR tools. Support multi-language: `"eng+fra"`.

### 2.6 — OCR Caching
**Problem:** Repeated OCR calls on the same region are expensive.  
**Fix:** Optional short-lived cache (configurable TTL, default 2s) based on screenshot hash. Skip OCR if screen hasn't changed.

---

## 3. High-Level Smart Tools

These tools combine multiple low-level operations into single intelligent actions — the biggest usability win.

### 3.1 — `find_text_on_screen`
**Description:** Find all occurrences of a text string on screen. Returns coordinates of each match.  
**Implementation:** Capture screen → OCR with bounding boxes → fuzzy text matching → return list of `{text, x, y, width, height, confidence}`.  
**Parameters:** `text` (string to find), `window_title` (optional, scope to window), `exact` (bool, default false for fuzzy match).

### 3.2 — `click_text`
**Description:** Find text on screen and click on it. The single most useful tool for UI testing.  
**Implementation:** `find_text_on_screen` → click center of first/best match.  
**Parameters:** `text`, `window_title` (optional), `button` (left/right), `occurrence` (which match to click, default 1).

### 3.3 — `wait_for_text`
**Description:** Wait until specific text appears on screen, with timeout.  
**Implementation:** Poll with `find_text_on_screen` at configurable interval.  
**Parameters:** `text`, `timeout_seconds` (default 10), `poll_interval` (default 0.5), `window_title` (optional).  
**Returns:** Coordinates of found text, or timeout error.

### 3.4 — `wait_for_window`
**Description:** Wait for a window with a given title to appear.  
**Parameters:** `window_title`, `timeout_seconds` (default 10), `poll_interval` (default 0.5).

### 3.5 — `assert_text_visible`
**Description:** Verify that text is (or is not) present on screen. For test assertions.  
**Parameters:** `text`, `should_exist` (bool), `window_title` (optional).  
**Returns:** Pass/fail with details and screenshot.

### 3.6 — `fill_field`
**Description:** Click a text field label, then type into the adjacent input.  
**Implementation:** Find label text → click to the right of it (or use Tab) → clear existing text → type new value.  
**Parameters:** `label_text`, `value`, `window_title` (optional).

### 3.7 — `get_window_snapshot`
**Description:** Comprehensive window state capture: screenshot + OCR text + window position/size + all detected UI text with coordinates.  
**Parameters:** `window_title`.  
**Returns:** Screenshot image + structured JSON with all detected elements.

### 3.8 — `compare_screenshots`
**Description:** Take a screenshot and compare with a previous one to detect changes.  
**Implementation:** Pixel diff with configurable threshold.  
**Parameters:** `reference_image_base64` (or auto-store previous), `region` (optional), `threshold` (default 0.95).

---

## 4. Reliability & Robustness

### 4.1 — Global Error Handling with Context
Wrap all tool calls in a decorator that catches exceptions and returns structured errors:
```python
{
  "error": true, 
  "tool": "click",
  "message": "Coordinates (5000, 3000) are outside screen bounds (1920x1080)",
  "suggestion": "Use capture_screen to see current screen dimensions"
}
```

### 4.2 — Coordinate Validation
Validate all x/y coordinates against actual screen dimensions before executing. Return clear error with screen size info.

### 4.3 — Window Operation Retry Logic
Window operations (`focus`, `activate`, `resize`, etc.) should retry up to 3 times with short delays. Windows can be unresponsive briefly during state transitions.

### 4.4 — Graceful Tesseract Missing
If Tesseract isn't installed, OCR tools should return a helpful error with install link instead of crashing.

### 4.5 — Timeout Support for All Operations
Add optional `timeout` parameter to operations that could hang (especially window activation, process kill).

### 4.6 — Operation Verification
After critical operations, verify the result:
- After `click`: optionally verify mouse position
- After `type_text`: optionally OCR to verify text appeared
- After `focus_window`: verify it's now foreground
- After `kill_process`: verify process is gone

### 4.7 — Rate Limiting & Debounce
Prevent accidental rapid-fire operations (clicking 100 times/second). Add configurable minimum interval between operations.

---

## 5. Performance Optimizations

### 5.1 — Screenshot Compression
**Problem:** Full-screen PNGs are 2-5 MB. Sending these through MCP stdio is slow and bloats context.  
**Fix:** Add `format` parameter (`"png"`, `"jpeg"`, `"webp"`) and `quality` parameter (1-100, default 75 for JPEG). JPEG screenshots at quality 75 are typically 200-500 KB.

### 5.2 — Screenshot Scaling
**Problem:** Full-resolution screenshots are often unnecessary for the AI to understand the UI.  
**Fix:** Add `scale` parameter (0.1-1.0, default 1.0). At 0.5 scale, images are 4x smaller.

### 5.3 — Lazy OCR
Only OCR visible/changed regions. Track what was last captured and skip unchanged areas.

### 5.4 — Connection Pooling / Singleton Resources
Create `mss` instance once and reuse instead of creating per-capture. Same for Tesseract initialization.

### 5.5 — Async Window Operations
Use `asyncio` properly — some operations block the event loop. Run blocking calls (psutil, pyautogui) in `asyncio.to_thread()`.

### 5.6 — Batch Operations
New tool `execute_sequence` — execute multiple operations in order without round-trip overhead:
```json
{
  "steps": [
    {"tool": "click", "args": {"x": 100, "y": 200}},
    {"tool": "type_text", "args": {"text": "hello"}},
    {"tool": "press_key", "args": {"keys": "enter"}}
  ]
}
```

---

## 6. New Capabilities

### 6.1 — `start_process`
Launch an application by path or command. Complements `kill_process`.  
**Parameters:** `command`, `args` (list), `working_directory`, `wait` (bool).

### 6.2 — `get_window_info`
Detailed info about a specific window: title, class name, PID, position, size, state, is_responding, child window count.

### 6.3 — `list_monitors`
List all connected monitors with resolution, position, scaling factor, and primary status.

### 6.4 — `capture_monitor`
Capture a specific monitor by index.

### 6.5 — `right_click_menu`
Right-click at a position and OCR the context menu that appears. Return menu items with coordinates.

### 6.6 — `triple_click`
Select entire line/paragraph. Common for text editing tests.

### 6.7 — `scroll_horizontal`
Horizontal scroll support for wide content.

### 6.8 — `wait_for_idle`
Wait for a window/process to stop being busy (CPU usage drops below threshold).

### 6.9 — `get_pixel_color`
Get the color of a pixel at given coordinates. Useful for state detection (is this button highlighted?).

### 6.10 — `screenshot_diff`
Compare current screen state against a reference image. Return percentage changed and diff visualization.

### 6.11 — UI Automation Integration (Phase 2)
Use Windows UI Automation API via `comtypes` or `uiautomation` package for:
- Proper control tree inspection (buttons, text fields, checkboxes)
- Control-based clicking (click button by name/automation ID)
- Read/write control values directly
- Get control states (enabled, checked, selected)
- Find controls by type, name, or automation properties

This is the **biggest long-term enhancement** — it moves from pixel-based to control-based automation.

### 6.12 — `health_check`
Verify all dependencies are working: Python version, Tesseract installed & version, pyautogui working, screen capture working, DPI settings. Return comprehensive status report.

---

## 7. Architecture & Code Quality

### 7.1 — Modular File Structure
Split the monolithic `server.py` into focused modules:
```
win32_mcp_server/
├── __init__.py
├── __main__.py
├── server.py          # MCP server setup & dispatch
├── tools/
│   ├── __init__.py
│   ├── capture.py     # Screenshot tools
│   ├── ocr.py         # OCR tools
│   ├── mouse.py       # Mouse tools
│   ├── keyboard.py    # Keyboard tools
│   ├── clipboard.py   # Clipboard tools
│   ├── window.py      # Window management
│   ├── process.py     # Process management
│   └── smart.py       # High-level smart tools
├── utils/
│   ├── __init__.py
│   ├── coordinates.py # Coordinate validation, DPI
│   ├── imaging.py     # Image preprocessing
│   ├── errors.py      # Error handling
│   └── window_match.py # Fuzzy window matching
└── config.py          # Configuration management
```

### 7.2 — Tool Registration Decorator
Replace the giant if/elif chain with a decorator-based registration:
```python
@register_tool("click", description="Click at coordinates", schema={...})
async def handle_click(arguments):
    ...
```

### 7.3 — Configuration System
Add server-level configuration for defaults:
- OCR language, preprocessing mode
- Screenshot format and quality
- Default timeouts
- Coordinate validation on/off
- Debug logging level

### 7.4 — Logging
Add proper `logging` module integration with configurable levels. Log all tool calls with timestamps and arguments for debugging test failures.

### 7.5 — Type Hints & Validation
Add Pydantic models for all tool arguments. Validate inputs before execution.

### 7.6 — Tests
Add a test suite:
- Unit tests for coordinate validation, image preprocessing, fuzzy matching
- Integration tests for each tool (with mock screen where possible)
- Smoke test that verifies server starts and lists tools

---

## 8. Extension Improvements

### 8.1 — Fix Dependency Detection
Change `python -c "import server"` to `pip show win32-mcp-server`.

### 8.2 — Tesseract Detection
Check if Tesseract is installed and on PATH. Show install prompt if not.

### 8.3 — Status Bar Indicator
Show server status (running/stopped/error) in VS Code status bar.

### 8.4 — Configuration UI
Add setting for Tesseract path, screenshot format defaults, OCR language.

---

## 9. Implementation Roadmap

### Phase 1 — Critical Fixes (v1.1) — 1-2 days
Priority: Ship-blocking bugs and robustness.

| # | Item | Section |
|---|------|---------|
| 1 | Global error handling for all tools | 1.1, 4.1 |
| 2 | Fix `type_text` unicode support | 1.2 |
| 3 | Fix `list_windows` duplicates | 1.3 |
| 4 | DPI awareness at startup | 1.5 |
| 5 | Coordinate validation | 4.2 |
| 6 | Fix extension dependency check | 1.6 |
| 7 | Add scroll position param | 1.7 |
| 8 | Graceful Tesseract-missing error | 4.4 |
| 9 | Fix `list_processes` pagination | 1.9 |
| 10 | Better "window not found" errors | 1.10 |

### Phase 2 — OCR & Smart Tools (v1.5) — 3-5 days
Priority: The features that make the AI dramatically more effective at testing.

| # | Item | Section |
|---|------|---------|
| 1 | OCR preprocessing pipeline | 2.1 |
| 2 | Structured OCR with bounding boxes | 2.2 |
| 3 | OCR confidence filtering | 2.3 |
| 4 | Window-specific OCR | 2.4 |
| 5 | `find_text_on_screen` | 3.1 |
| 6 | `click_text` | 3.2 |
| 7 | `wait_for_text` | 3.3 |
| 8 | `wait_for_window` | 3.4 |
| 9 | `assert_text_visible` | 3.5 |
| 10 | `get_window_snapshot` | 3.7 |

### Phase 3 — Performance & Polish (v1.8) — 2-3 days
Priority: Speed and developer experience.

| # | Item | Section |
|---|------|---------|
| 1 | Screenshot compression (JPEG/WebP) | 5.1 |
| 2 | Screenshot scaling | 5.2 |
| 3 | Async operations with `to_thread` | 5.5 |
| 4 | `start_process` | 6.1 |
| 5 | `list_monitors` + `capture_monitor` | 6.3, 6.4 |
| 6 | `health_check` | 6.12 |
| 7 | `get_pixel_color` | 6.9 |
| 8 | Batch operations | 5.6 |
| 9 | Logging system | 7.4 |

### Phase 4 — Architecture & UI Automation (v2.0) — 5-7 days
Priority: Long-term quality, maintainability, and advanced features.

| # | Item | Section |
|---|------|---------|
| 1 | Modular file structure | 7.1 |
| 2 | Tool registration decorator | 7.2 |
| 3 | Configuration system | 7.3 |
| 4 | Pydantic validation | 7.5 |
| 5 | UI Automation API integration | 6.11 |
| 6 | `fill_field` | 3.6 |
| 7 | Test suite | 7.6 |
| 8 | Extension improvements | 8.x |

---

## Quick Win Summary

The following changes deliver the **most impact with the least effort**:

1. **`click_text`** — Find text on screen and click it. Eliminates the #1 friction point: figuring out coordinates.
2. **`wait_for_window` / `wait_for_text`** — Eliminates flaky tests caused by timing.
3. **Structured OCR** — Returns text with positions. Enables the AI to reason about UI layout.
4. **Screenshot compression** — 5-10x smaller payloads. Faster round trips.
5. **Global error handling** — Stops the server from crashing on edge cases.
6. **Unicode `type_text`** — Basic correctness fix.

---

## Dependencies to Add

| Package | Purpose | Phase |
|---------|---------|-------|
| `rapidfuzz` | Fuzzy text/window title matching | 2 |
| `numpy` | Image preprocessing (thresholding) | 2 |
| `opencv-python` | Advanced image preprocessing | 2 (optional) |
| `pydantic` | Input validation | 4 |
| `comtypes` or `uiautomation` | Windows UI Automation | 4 |

---

## Metrics for Success

- **OCR accuracy**: >90% on standard Windows UIs (currently ~70%)
- **Tool call success rate**: >99% (currently crashes on edge cases)
- **Screenshot payload size**: <500 KB average (currently 2-5 MB)
- **Time to click a button by text**: <2 seconds end-to-end
- **Zero crashes** from missing Tesseract, bad coordinates, or disappeared windows
