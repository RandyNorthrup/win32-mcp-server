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
import os
import subprocess
import tempfile
import time
from pathlib import Path
from typing import Any

import psutil
from mcp.types import TextContent

from ..config import config
from ..registry import registry
from ..utils.args import get_bool, get_float, get_int, get_str, get_timeout
from ..utils.errors import ToolError
from ..utils.security import enforce_command_allowed

logger = logging.getLogger("win32-mcp")


def _validate_process_request(command: str, args: list[str], cwd: str | None, timeout: float) -> None:
    if not command.strip():
        raise ToolError("Command must not be empty")
    if len(command) > 1_000:
        raise ToolError("Command is too long")
    if len(args) > 128:
        raise ToolError("Too many process arguments")
    too_long = [arg for arg in args if len(arg) > 4_000]
    if too_long:
        raise ToolError("One or more process arguments are too long")
    if cwd is not None and not Path(cwd).is_dir():
        raise ToolError(
            f"Working directory does not exist or is not a directory: {cwd}",
            suggestion="Provide an existing working_directory or omit it.",
        )
    if timeout <= 0 or timeout > config.limits.max_timeout_seconds:
        raise ToolError(f"timeout_seconds must be between 0 and {config.limits.max_timeout_seconds:g}")
    enforce_command_allowed(command)


def _creation_flags() -> int:
    flags = 0
    if os.name == "nt":
        flags |= getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
    return flags


def _read_limited(tmp: Any, max_bytes: int) -> tuple[str, bool]:
    tmp.seek(0)
    data = tmp.read(max_bytes + 1)
    truncated = len(data) > max_bytes
    if truncated:
        data = data[:max_bytes]
    return data.decode("utf-8", errors="replace"), truncated


def _run_process_bounded(
    full_cmd: list[str],
    cwd: str | None,
    timeout: float,
) -> dict[str, Any]:
    max_bytes = config.limits.max_subprocess_output_bytes
    with tempfile.TemporaryFile() as stdout_file, tempfile.TemporaryFile() as stderr_file:
        proc = subprocess.Popen(
            full_cmd,
            cwd=cwd,
            stdin=subprocess.DEVNULL,
            stdout=stdout_file,
            stderr=stderr_file,
            close_fds=True,
            creationflags=_creation_flags(),
        )
        try:
            returncode = proc.wait(timeout=timeout)
        except subprocess.TimeoutExpired:
            proc.kill()
            proc.wait(timeout=5)
            raise

        stdout, stdout_truncated = _read_limited(stdout_file, max_bytes)
        stderr, stderr_truncated = _read_limited(stderr_file, max_bytes)
        return {
            "returncode": returncode,
            "stdout": stdout,
            "stderr": stderr,
            "stdout_truncated": stdout_truncated,
            "stderr_truncated": stderr_truncated,
        }


@registry.register(
    "list_processes",
    "List running processes with PIDs, memory, and CPU info",
    {
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
    },
)
async def handle_list_processes(arguments: dict[str, Any]) -> list[TextContent]:
    filter_name = get_str(arguments, "filter", default="").lower().strip()
    limit = get_int(arguments, "limit", default=100, min_value=1, max_value=500)
    offset = get_int(arguments, "offset", default=0, min_value=0)
    sort_by = get_str(arguments, "sort_by", default="memory")

    processes: list[dict[str, Any]] = []

    def _gather() -> None:
        for proc in psutil.process_iter(["pid", "name", "status", "memory_info", "cpu_percent"]):
            try:
                info = proc.info
                name = info.get("name", "")
                if filter_name and filter_name not in name.lower():
                    continue

                mem_info = info.get("memory_info")
                mem_mb = round(mem_info.rss / 1024 / 1024, 1) if mem_info else 0

                processes.append(
                    {
                        "pid": info["pid"],
                        "name": name,
                        "status": info.get("status", "unknown"),
                        "memory_mb": mem_mb,
                        "cpu_percent": info.get("cpu_percent", 0) or 0,
                    }
                )
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
    page = processes[offset : offset + limit]

    return [
        TextContent(
            type="text",
            text=json.dumps(
                {
                    "total_count": total_count,
                    "showing": f"{offset + 1}-{offset + len(page)}",
                    "filter": filter_name or None,
                    "sort_by": sort_by,
                    "processes": page,
                },
                indent=2,
            ),
        )
    ]


@registry.register(
    "kill_process",
    "Terminate a process by PID",
    {
        "type": "object",
        "properties": {
            "pid": {"type": "number", "description": "Process ID to terminate"},
            "force": {
                "type": "boolean",
                "description": "Force kill (SIGKILL) instead of graceful termination (default: false)",
            },
        },
        "required": ["pid"],
    },
)
async def handle_kill_process(arguments: dict[str, Any]) -> list[TextContent]:
    pid = get_int(arguments, "pid", required=True, min_value=0)
    force = get_bool(arguments, "force", default=False)
    if pid in {0, 4, os.getpid()}:
        raise ToolError("Refusing to terminate a protected or current process")

    try:
        proc = psutil.Process(pid)
        name = proc.name()
    except psutil.NoSuchProcess as exc:
        raise ToolError(
            f"No process with PID {pid}",
            suggestion="Use list_processes to find the correct PID",
        ) from exc
    except psutil.AccessDenied as exc:
        raise ToolError(
            f"Access denied to process {pid}",
            suggestion="Run the MCP server as Administrator",
        ) from exc

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
    except psutil.AccessDenied as exc:
        raise ToolError(
            f"Access denied when killing process {name} (PID {pid})",
            suggestion="Run the MCP server as Administrator",
        ) from exc
    except Exception as exc:
        raise ToolError(f"Failed to kill process {name} (PID {pid}): {exc}") from exc

    # Verify process is gone
    if psutil.pid_exists(pid):
        logger.warning("kill_process: PID %d still exists after kill", pid)

    return [TextContent(type="text", text=f"Killed process {name} (PID: {pid})")]


@registry.register(
    "start_process",
    "Launch an application or command",
    {
        "type": "object",
        "properties": {
            "command": {
                "type": "string",
                "minLength": 1,
                "maxLength": 1000,
                "description": "Executable path or command name (e.g. 'notepad', 'C:\\\\Program Files\\\\app.exe')",
            },
            "args": {
                "type": "array",
                "items": {"type": "string", "maxLength": 4000},
                "maxItems": 128,
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
                "minimum": 0.1,
                "maximum": 120,
                "description": "Timeout in seconds when wait=true (default: 30)",
            },
        },
        "required": ["command"],
    },
)
async def handle_start_process(arguments: dict[str, Any]) -> dict[str, Any]:
    command = get_str(arguments, "command", required=True, min_length=1, max_length=config.limits.max_command_chars)
    args = arguments.get("args", [])
    if not isinstance(args, list) or not all(isinstance(arg, str) for arg in args):
        raise ToolError("args must be an array of strings")
    cwd = get_str(arguments, "working_directory", default="") or None
    wait = get_bool(arguments, "wait", default=False)
    timeout = get_timeout(arguments, default=30.0)

    full_cmd = [command, *args]
    _validate_process_request(command, args, cwd, timeout)

    if wait:
        try:
            result = await asyncio.to_thread(_run_process_bounded, full_cmd, cwd, timeout)
            return {
                "command": command,
                "args": args,
                "returncode": result["returncode"],
                "stdout": result["stdout"],
                "stderr": result["stderr"],
                "stdout_truncated": result["stdout_truncated"],
                "stderr_truncated": result["stderr_truncated"],
                "completed": True,
            }
        except FileNotFoundError as exc:
            raise ToolError(
                f"Command not found: {command}",
                suggestion="Check the executable path or ensure it's on PATH",
            ) from exc
        except subprocess.TimeoutExpired as exc:
            raise ToolError(
                f"Command timed out after {timeout}s: {command}",
                suggestion="Increase timeout_seconds or use wait=false",
            ) from exc
    else:
        try:
            proc = await asyncio.to_thread(
                subprocess.Popen,
                full_cmd,
                cwd=cwd,
                stdin=subprocess.DEVNULL,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                close_fds=True,
                creationflags=_creation_flags(),
            )
            return {
                "command": command,
                "args": args,
                "pid": proc.pid,
                "started": True,
            }
        except FileNotFoundError as exc:
            raise ToolError(
                f"Command not found: {command}",
                suggestion="Check the executable path or ensure it's on PATH",
            ) from exc
        except PermissionError as exc:
            raise ToolError(
                f"Permission denied: {command}",
                suggestion="Run the MCP server as Administrator",
            ) from exc


@registry.register(
    "wait_for_idle",
    "Wait for a process to become idle (CPU usage drops below threshold)",
    {
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
                "minimum": 0,
                "maximum": 100,
                "description": "CPU percent below which process is considered idle (default: 5.0)",
            },
        },
    },
)
async def handle_wait_for_idle(arguments: dict[str, Any]) -> dict[str, Any]:
    pid = get_int(arguments, "pid", min_value=0) if "pid" in arguments and arguments["pid"] is not None else None
    window_title = get_str(arguments, "window_title", default="")
    timeout = get_timeout(arguments, default=30.0)
    threshold = get_float(arguments, "cpu_threshold", default=5.0, min_value=0, max_value=100)

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
        proc = psutil.Process(pid)
    except psutil.NoSuchProcess as exc:
        raise ToolError(f"Process {pid} not found") from exc

    # Initialize CPU measurement
    proc.cpu_percent()
    await asyncio.sleep(0.5)

    start = time.monotonic()
    while True:
        try:
            cpu = proc.cpu_percent(interval=0)
        except psutil.NoSuchProcess:
            return {
                "idle": True,
                "pid": pid,
                "process_name": "terminated",
                "reason": "Process exited during monitoring",
                "elapsed_seconds": round(time.monotonic() - start, 2),
            }

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
