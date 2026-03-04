"""
Process management tools.

Tools:
  - list_processes   List running processes with filtering and pagination
  - kill_process     Terminate a process by PID
  - start_process    Launch an application
  - wait_for_idle    Wait for a process to become idle (low CPU)
"""

import asyncio
import json
import logging
import subprocess
import time

import psutil
from mcp.types import TextContent

from ..registry import registry
from ..utils.errors import ToolError

logger = logging.getLogger("win32-mcp")


@registry.register("list_processes", "List running processes with PIDs, memory, and CPU info", {
    "type": "object",
    "properties": {
        "filter": {
            "type": "string",
            "description": "Filter by process name (case-insensitive substring match)",
        },
        "limit": {
            "type": "number",
            "description": "Maximum processes to return (default: 100)",
        },
        "offset": {
            "type": "number",
            "description": "Skip this many results for pagination (default: 0)",
        },
        "sort_by": {
            "type": "string",
            "enum": ["name", "pid", "memory", "cpu"],
            "description": "Sort field (default: memory, descending)",
        },
    },
})
async def handle_list_processes(arguments: dict):
    filter_name = arguments.get("filter", "").lower().strip()
    limit = int(arguments.get("limit", 100))
    offset = int(arguments.get("offset", 0))
    sort_by = arguments.get("sort_by", "memory")

    processes: list[dict] = []

    def _gather():
        for proc in psutil.process_iter(["pid", "name", "status", "memory_info", "cpu_percent"]):
            try:
                info = proc.info
                name = info.get("name", "")
                if filter_name and filter_name not in name.lower():
                    continue

                mem_info = info.get("memory_info")
                mem_mb = round(mem_info.rss / 1024 / 1024, 1) if mem_info else 0

                processes.append({
                    "pid": info["pid"],
                    "name": name,
                    "status": info.get("status", "unknown"),
                    "memory_mb": mem_mb,
                    "cpu_percent": info.get("cpu_percent", 0) or 0,
                })
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                pass

    await asyncio.to_thread(_gather)

    # Sort
    sort_keys = {
        "name": lambda p: p["name"].lower(),
        "pid": lambda p: p["pid"],
        "memory": lambda p: p["memory_mb"],
        "cpu": lambda p: p["cpu_percent"],
    }
    reverse = sort_by in ("memory", "cpu")
    processes.sort(key=sort_keys.get(sort_by, sort_keys["memory"]), reverse=reverse)

    total_count = len(processes)
    page = processes[offset:offset + limit]

    return [TextContent(type="text", text=json.dumps({
        "total_count": total_count,
        "showing": f"{offset + 1}–{offset + len(page)}",
        "filter": filter_name or None,
        "sort_by": sort_by,
        "processes": page,
    }, indent=2))]


@registry.register("kill_process", "Terminate a process by PID", {
    "type": "object",
    "properties": {
        "pid": {"type": "number", "description": "Process ID to terminate"},
        "force": {
            "type": "boolean",
            "description": "Force kill (SIGKILL) instead of graceful termination (default: false)",
        },
    },
    "required": ["pid"],
})
async def handle_kill_process(arguments: dict):
    pid = int(arguments["pid"])
    force = arguments.get("force", False)

    try:
        proc = psutil.Process(pid)
        name = proc.name()
    except psutil.NoSuchProcess:
        raise ToolError(
            f"No process with PID {pid}",
            suggestion="Use list_processes to find the correct PID",
        )
    except psutil.AccessDenied:
        raise ToolError(
            f"Access denied to process {pid}",
            suggestion="Run the MCP server as Administrator",
        )

    try:
        if force:
            proc.kill()
        else:
            proc.terminate()
            # Wait up to 5s for graceful shutdown
            try:
                proc.wait(timeout=5)
            except psutil.TimeoutExpired:
                proc.kill()  # Force kill if graceful didn't work
                proc.wait(timeout=3)
    except psutil.AccessDenied:
        raise ToolError(
            f"Access denied when killing process {name} (PID {pid})",
            suggestion="Run the MCP server as Administrator",
        )
    except Exception as exc:
        raise ToolError(f"Failed to kill process {name} (PID {pid}): {exc}")

    return [TextContent(type="text", text=f"Killed process {name} (PID: {pid})")]


@registry.register("start_process", "Launch an application or command", {
    "type": "object",
    "properties": {
        "command": {
            "type": "string",
            "description": "Executable path or command name (e.g. 'notepad', 'C:\\\\Program Files\\\\app.exe')",
        },
        "args": {
            "type": "array",
            "items": {"type": "string"},
            "description": "Command-line arguments",
        },
        "working_directory": {
            "type": "string",
            "description": "Working directory for the process",
        },
        "wait": {
            "type": "boolean",
            "description": "Wait for process to complete (default: false). True=block until done.",
        },
        "timeout_seconds": {
            "type": "number",
            "description": "Timeout in seconds when wait=true (default: 30)",
        },
    },
    "required": ["command"],
})
async def handle_start_process(arguments: dict):
    command = arguments["command"]
    args = arguments.get("args", [])
    cwd = arguments.get("working_directory")
    wait = arguments.get("wait", False)
    timeout = arguments.get("timeout_seconds", 30)

    full_cmd = [command] + args

    if wait:
        try:
            result = await asyncio.to_thread(
                subprocess.run,
                full_cmd,
                cwd=cwd,
                capture_output=True,
                text=True,
                timeout=timeout,
            )
            return {
                "command": command,
                "args": args,
                "returncode": result.returncode,
                "stdout": result.stdout[:10000] if result.stdout else "",
                "stderr": result.stderr[:10000] if result.stderr else "",
                "completed": True,
            }
        except FileNotFoundError:
            raise ToolError(
                f"Command not found: {command}",
                suggestion="Check the executable path or ensure it's on PATH",
            )
        except subprocess.TimeoutExpired:
            raise ToolError(
                f"Command timed out after {timeout}s: {command}",
                suggestion="Increase timeout_seconds or use wait=false",
            )
    else:
        try:
            proc = await asyncio.to_thread(
                subprocess.Popen,
                full_cmd,
                cwd=cwd,
            )
            return {
                "command": command,
                "args": args,
                "pid": proc.pid,
                "started": True,
            }
        except FileNotFoundError:
            raise ToolError(
                f"Command not found: {command}",
                suggestion="Check the executable path or ensure it's on PATH",
            )
        except PermissionError:
            raise ToolError(
                f"Permission denied: {command}",
                suggestion="Run the MCP server as Administrator",
            )


@registry.register("wait_for_idle", "Wait for a process to become idle (CPU usage drops below threshold)", {
    "type": "object",
    "properties": {
        "pid": {"type": "number", "description": "Process ID to monitor"},
        "window_title": {
            "type": "string",
            "description": "Alternative: find PID by window title",
        },
        "timeout_seconds": {
            "type": "number",
            "description": "Maximum wait time (default: 30)",
        },
        "cpu_threshold": {
            "type": "number",
            "description": "CPU percent below which process is considered idle (default: 5.0)",
        },
    },
})
async def handle_wait_for_idle(arguments: dict):
    pid = arguments.get("pid")
    window_title = arguments.get("window_title")
    timeout = arguments.get("timeout_seconds", 30)
    threshold = arguments.get("cpu_threshold", 5.0)

    # Resolve PID from window title if needed
    if pid is None and window_title:
        from ..utils.window_match import find_window_strict, get_window_pid
        win = find_window_strict(window_title)
        hwnd = getattr(win, "_hWnd", None)
        if hwnd:
            pid = get_window_pid(hwnd)

    if pid is None:
        raise ToolError(
            "No PID specified and could not resolve from window_title",
            suggestion="Provide pid or window_title",
        )

    try:
        proc = psutil.Process(int(pid))
    except psutil.NoSuchProcess:
        raise ToolError(f"Process {pid} not found")

    # Initialize CPU measurement
    proc.cpu_percent()
    await asyncio.sleep(0.5)

    start = time.monotonic()
    while True:
        cpu = proc.cpu_percent(interval=0)
        if cpu <= threshold:
            return {
                "idle": True,
                "pid": pid,
                "process_name": proc.name(),
                "final_cpu_percent": round(cpu, 1),
                "elapsed_seconds": round(time.monotonic() - start, 2),
            }

        elapsed = time.monotonic() - start
        if elapsed >= timeout:
            return {
                "idle": False,
                "pid": pid,
                "process_name": proc.name(),
                "final_cpu_percent": round(cpu, 1),
                "elapsed_seconds": round(elapsed, 2),
                "timeout": True,
            }

        await asyncio.sleep(0.5)
