"""
MCP Server for Windows UI Inspection and Control
Full automation: OCR, drag/drop, mouse, keyboard, clipboard, process management

Author: Randy Northrup
GitHub: https://github.com/RandyNorthrup/win32-mcp-server
"""

import asyncio
import base64
import io
import json
import psutil
from typing import Any

from mss import mss
from PIL import Image
import pyautogui
import pygetwindow as gw
import pyperclip
import pytesseract
from mcp.server import Server
from mcp.types import Tool, TextContent, ImageContent

pyautogui.FAILSAFE = False

app = Server("win32-inspector")


@app.list_tools()
async def list_tools() -> list[Tool]:
    return [
        Tool(name="capture_screen", description="Capture full screen screenshot", inputSchema={"type": "object", "properties": {}}),
        Tool(name="capture_window", description="Capture window by title", inputSchema={"type": "object", "properties": {"window_title": {"type": "string"}}, "required": ["window_title"]}),
        Tool(name="list_windows", description="List all open windows with details", inputSchema={"type": "object", "properties": {}}),
        Tool(name="ocr_screen", description="OCR text from full screen", inputSchema={"type": "object", "properties": {}}),
        Tool(name="ocr_region", description="OCR text from region", inputSchema={"type": "object", "properties": {"x": {"type": "number"}, "y": {"type": "number"}, "width": {"type": "number"}, "height": {"type": "number"}}, "required": ["x", "y", "width", "height"]}),
        Tool(name="click", description="Click at coordinates", inputSchema={"type": "object", "properties": {"x": {"type": "number"}, "y": {"type": "number"}, "button": {"type": "string", "enum": ["left", "right", "middle"], "default": "left"}, "clicks": {"type": "number", "default": 1}}, "required": ["x", "y"]}),
        Tool(name="double_click", description="Double click at coordinates", inputSchema={"type": "object", "properties": {"x": {"type": "number"}, "y": {"type": "number"}}, "required": ["x", "y"]}),
        Tool(name="drag", description="Drag from start to end coordinates", inputSchema={"type": "object", "properties": {"start_x": {"type": "number"}, "start_y": {"type": "number"}, "end_x": {"type": "number"}, "end_y": {"type": "number"}, "duration": {"type": "number", "default": 0.5}}, "required": ["start_x", "start_y", "end_x", "end_y"]}),
        Tool(name="type_text", description="Type text at current position", inputSchema={"type": "object", "properties": {"text": {"type": "string"}, "interval": {"type": "number", "default": 0.01}}, "required": ["text"]}),
        Tool(name="press_key", description="Press keyboard key or shortcut", inputSchema={"type": "object", "properties": {"keys": {"type": "string", "description": "Key or combo like 'enter', 'ctrl+c', 'alt+f4'"}}, "required": ["keys"]}),
        Tool(name="hotkey", description="Execute hotkey combination", inputSchema={"type": "object", "properties": {"keys": {"type": "array", "items": {"type": "string"}, "description": "Keys like ['ctrl', 'shift', 's']"}}, "required": ["keys"]}),
        Tool(name="clipboard_copy", description="Copy text to clipboard", inputSchema={"type": "object", "properties": {"text": {"type": "string"}}, "required": ["text"]}),
        Tool(name="clipboard_paste", description="Get clipboard text", inputSchema={"type": "object", "properties": {}}),
        Tool(name="mouse_position", description="Get current mouse position", inputSchema={"type": "object", "properties": {}}),
        Tool(name="mouse_move", description="Move mouse to position", inputSchema={"type": "object", "properties": {"x": {"type": "number"}, "y": {"type": "number"}, "duration": {"type": "number", "default": 0.25}}, "required": ["x", "y"]}),
        Tool(name="scroll", description="Scroll up/down", inputSchema={"type": "object", "properties": {"amount": {"type": "number", "description": "Positive=up, negative=down"}}, "required": ["amount"]}),
        Tool(name="list_processes", description="List running processes with PIDs", inputSchema={"type": "object", "properties": {"filter": {"type": "string", "description": "Optional name filter"}}}),
        Tool(name="kill_process", description="Kill process by PID", inputSchema={"type": "object", "properties": {"pid": {"type": "number"}}, "required": ["pid"]}),
        Tool(name="focus_window", description="Focus/activate window", inputSchema={"type": "object", "properties": {"window_title": {"type": "string"}}, "required": ["window_title"]}),
        Tool(name="close_window", description="Close window by title", inputSchema={"type": "object", "properties": {"window_title": {"type": "string"}}, "required": ["window_title"]}),
        Tool(name="minimize_window", description="Minimize window", inputSchema={"type": "object", "properties": {"window_title": {"type": "string"}}, "required": ["window_title"]}),
        Tool(name="maximize_window", description="Maximize window", inputSchema={"type": "object", "properties": {"window_title": {"type": "string"}}, "required": ["window_title"]}),
        Tool(name="restore_window", description="Restore window", inputSchema={"type": "object", "properties": {"window_title": {"type": "string"}}, "required": ["window_title"]}),
        Tool(name="resize_window", description="Resize window", inputSchema={"type": "object", "properties": {"window_title": {"type": "string"}, "width": {"type": "number"}, "height": {"type": "number"}}, "required": ["window_title", "width", "height"]}),
        Tool(name="move_window", description="Move window position", inputSchema={"type": "object", "properties": {"window_title": {"type": "string"}, "x": {"type": "number"}, "y": {"type": "number"}}, "required": ["window_title", "x", "y"]}),
    ]


@app.call_tool()
async def call_tool(name: str, arguments: Any) -> list[TextContent | ImageContent]:
    
    # Screenshot tools
    if name == "capture_screen":
        with mss() as sct:
            monitor = sct.monitors[1]
            screenshot = sct.grab(monitor)
            img = Image.frombytes('RGB', screenshot.size, screenshot.bgra, 'raw', 'BGRX')
            buffer = io.BytesIO()
            img.save(buffer, format='PNG')
            return [TextContent(type="text", text=f"Screen: {screenshot.size}"), ImageContent(type="image", data=base64.b64encode(buffer.getvalue()).decode(), mimeType="image/png")]
    
    elif name == "capture_window":
        windows = gw.getWindowsWithTitle(arguments["window_title"])
        if not windows: return [TextContent(type="text", text=f"Window not found: {arguments['window_title']}")]
        win = windows[0]
        win.activate()
        await asyncio.sleep(0.3)
        with mss() as sct:
            monitor = {"top": win.top, "left": win.left, "width": win.width, "height": win.height}
            screenshot = sct.grab(monitor)
            img = Image.frombytes('RGB', screenshot.size, screenshot.bgra, 'raw', 'BGRX')
            buffer = io.BytesIO()
            img.save(buffer, format='PNG')
            return [TextContent(type="text", text=f"{win.title}"), ImageContent(type="image", data=base64.b64encode(buffer.getvalue()).decode(), mimeType="image/png")]
    
    # OCR tools
    elif name == "ocr_screen":
        with mss() as sct:
            monitor = sct.monitors[1]
            screenshot = sct.grab(monitor)
            img = Image.frombytes('RGB', screenshot.size, screenshot.bgra, 'raw', 'BGRX')
            text = pytesseract.image_to_string(img)
            return [TextContent(type="text", text=f"OCR Result:\n{text}")]
    
    elif name == "ocr_region":
        with mss() as sct:
            region = {"top": int(arguments["y"]), "left": int(arguments["x"]), "width": int(arguments["width"]), "height": int(arguments["height"])}
            screenshot = sct.grab(region)
            img = Image.frombytes('RGB', screenshot.size, screenshot.bgra, 'raw', 'BGRX')
            text = pytesseract.image_to_string(img)
            return [TextContent(type="text", text=f"OCR: {text}")]
    
    # Mouse tools
    elif name == "click":
        pyautogui.click(int(arguments["x"]), int(arguments["y"]), button=arguments.get("button", "left"), clicks=arguments.get("clicks", 1))
        return [TextContent(type="text", text=f"Clicked {arguments.get('button', 'left')} at ({arguments['x']}, {arguments['y']})")]
    
    elif name == "double_click":
        pyautogui.doubleClick(int(arguments["x"]), int(arguments["y"]))
        return [TextContent(type="text", text=f"Double-clicked ({arguments['x']}, {arguments['y']})")]
    
    elif name == "drag":
        pyautogui.moveTo(int(arguments["start_x"]), int(arguments["start_y"]))
        pyautogui.drag(int(arguments["end_x"]) - int(arguments["start_x"]), int(arguments["end_y"]) - int(arguments["start_y"]), duration=arguments.get("duration", 0.5))
        return [TextContent(type="text", text=f"Dragged from ({arguments['start_x']}, {arguments['start_y']}) to ({arguments['end_x']}, {arguments['end_y']})")]
    
    elif name == "mouse_position":
        pos = pyautogui.position()
        return [TextContent(type="text", text=f"Mouse: ({pos.x}, {pos.y})")]
    
    elif name == "mouse_move":
        pyautogui.moveTo(int(arguments["x"]), int(arguments["y"]), duration=arguments.get("duration", 0.25))
        return [TextContent(type="text", text=f"Moved to ({arguments['x']}, {arguments['y']})")]
    
    elif name == "scroll":
        pyautogui.scroll(int(arguments["amount"]))
        return [TextContent(type="text", text=f"Scrolled {arguments['amount']}")]
    
    # Keyboard tools
    elif name == "type_text":
        pyautogui.write(arguments["text"], interval=arguments.get("interval", 0.01))
        return [TextContent(type="text", text=f"Typed: {arguments['text']}")]
    
    elif name == "press_key":
        keys = arguments["keys"].lower()
        if "+" in keys:
            key_parts = [k.strip() for k in keys.split("+")]
            pyautogui.hotkey(*key_parts)
        else:
            pyautogui.press(keys)
        return [TextContent(type="text", text=f"Pressed: {keys}")]
    
    elif name == "hotkey":
        pyautogui.hotkey(*arguments["keys"])
        return [TextContent(type="text", text=f"Hotkey: {'+'.join(arguments['keys'])}")]
    
    # Clipboard tools
    elif name == "clipboard_copy":
        pyperclip.copy(arguments["text"])
        return [TextContent(type="text", text=f"Copied to clipboard")]
    
    elif name == "clipboard_paste":
        text = pyperclip.paste()
        return [TextContent(type="text", text=f"Clipboard: {text}")]
    
    # Window management
    elif name == "list_windows":
        windows = []
        for title in gw.getAllTitles():
            if title:
                try:
                    win = gw.getWindowsWithTitle(title)[0]
                    windows.append({"title": title, "x": win.left, "y": win.top, "width": win.width, "height": win.height, "visible": win.visible, "minimized": win.isMinimized, "maximized": win.isMaximized})
                except: pass
        return [TextContent(type="text", text=json.dumps(windows, indent=2))]
    
    elif name == "focus_window":
        windows = gw.getWindowsWithTitle(arguments["window_title"])
        if not windows: return [TextContent(type="text", text=f"Window not found")]
        windows[0].activate()
        return [TextContent(type="text", text=f"Focused: {windows[0].title}")]
    
    elif name == "close_window":
        windows = gw.getWindowsWithTitle(arguments["window_title"])
        if not windows: return [TextContent(type="text", text=f"Window not found")]
        windows[0].close()
        return [TextContent(type="text", text=f"Closed: {windows[0].title}")]
    
    elif name == "minimize_window":
        windows = gw.getWindowsWithTitle(arguments["window_title"])
        if not windows: return [TextContent(type="text", text=f"Window not found")]
        windows[0].minimize()
        return [TextContent(type="text", text=f"Minimized")]
    
    elif name == "maximize_window":
        windows = gw.getWindowsWithTitle(arguments["window_title"])
        if not windows: return [TextContent(type="text", text=f"Window not found")]
        windows[0].maximize()
        return [TextContent(type="text", text=f"Maximized")]
    
    elif name == "restore_window":
        windows = gw.getWindowsWithTitle(arguments["window_title"])
        if not windows: return [TextContent(type="text", text=f"Window not found")]
        windows[0].restore()
        return [TextContent(type="text", text=f"Restored")]
    
    elif name == "resize_window":
        windows = gw.getWindowsWithTitle(arguments["window_title"])
        if not windows: return [TextContent(type="text", text=f"Window not found")]
        windows[0].resizeTo(int(arguments["width"]), int(arguments["height"]))
        return [TextContent(type="text", text=f"Resized to {arguments['width']}x{arguments['height']}")]
    
    elif name == "move_window":
        windows = gw.getWindowsWithTitle(arguments["window_title"])
        if not windows: return [TextContent(type="text", text=f"Window not found")]
        windows[0].moveTo(int(arguments["x"]), int(arguments["y"]))
        return [TextContent(type="text", text=f"Moved to ({arguments['x']}, {arguments['y']})")]
    
    # Process management
    elif name == "list_processes":
        filter_name = arguments.get("filter", "").lower()
        processes = []
        for proc in psutil.process_iter(['pid', 'name', 'status', 'memory_info']):
            try:
                if not filter_name or filter_name in proc.info['name'].lower():
                    processes.append({"pid": proc.info['pid'], "name": proc.info['name'], "status": proc.info['status'], "memory_mb": proc.info['memory_info'].rss / 1024 / 1024})
            except: pass
        return [TextContent(type="text", text=json.dumps(processes[:50], indent=2))]
    
    elif name == "kill_process":
        try:
            proc = psutil.Process(int(arguments["pid"]))
            name = proc.name()
            proc.kill()
            return [TextContent(type="text", text=f"Killed process {name} (PID: {arguments['pid']})")]
        except Exception as e:
            return [TextContent(type="text", text=f"Error killing process: {str(e)}")]
    
    return [TextContent(type="text", text=f"Unknown tool: {name}")]


async def async_main():
    from mcp.server.stdio import stdio_server
    async with stdio_server() as (read_stream, write_stream):
        await app.run(read_stream, write_stream, app.create_initialization_options())


def main():
    """Entry point for installed package"""
    asyncio.run(async_main())


if __name__ == "__main__":
    main()
