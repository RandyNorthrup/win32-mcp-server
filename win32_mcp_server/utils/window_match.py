"""
Window matching utilities — fuzzy title matching, deduplication, PID lookup.
"""

import ctypes
import ctypes.wintypes
import logging
from typing import Any

import pygetwindow as gw

from .errors import ToolError

logger = logging.getLogger("win32-mcp")

# ---------------------------------------------------------------------------
# Fuzzy matching — prefer rapidfuzz, fall back to difflib
# ---------------------------------------------------------------------------

try:
    from rapidfuzz import fuzz as _rf_fuzz
    from rapidfuzz import process as _rf_process

    def _fuzzy_ratio(s1: str, s2: str) -> int:
        return int(_rf_fuzz.partial_ratio(s1, s2))

    def _fuzzy_best_matches(query: str, choices: list[str], limit: int = 5) -> list[tuple[str, int]]:
        if not choices:
            return []
        results = _rf_process.extract(query, choices, scorer=_rf_fuzz.partial_ratio, limit=limit)
        return [(r[0], int(r[1])) for r in results]

except ImportError:
    from difflib import SequenceMatcher, get_close_matches

    def _fuzzy_ratio(s1: str, s2: str) -> int:
        return int(SequenceMatcher(None, s1.lower(), s2.lower()).ratio() * 100)

    def _fuzzy_best_matches(query: str, choices: list[str], limit: int = 5) -> list[tuple[str, int]]:
        close = get_close_matches(query, choices, n=limit, cutoff=0.3)
        return [(c, _fuzzy_ratio(query, c)) for c in close]


# ---------------------------------------------------------------------------
# Window Finding
# ---------------------------------------------------------------------------

def find_window(
    title: str,
    threshold: int = 60,
) -> tuple[Any | None, list[str]]:
    """Find a window by title with fuzzy matching.

    Strategy:
      1. Exact case-insensitive substring match.
      2. Fuzzy match above threshold.

    Returns:
        (window_object_or_None, list_of_suggestion_titles)
    """
    all_wins = get_all_windows_deduped()
    if not all_wins:
        return None, []

    title_lower = title.lower().strip()

    # Pass 1 — exact substring match (case-insensitive)
    for win in all_wins:
        if title_lower in win.title.lower():
            return win, []

    # Pass 2 — fuzzy
    titles = [w.title for w in all_wins]
    matches = _fuzzy_best_matches(title, titles, limit=5)
    suggestions = [m[0] for m in matches]

    for match_title, score in matches:
        if score >= threshold:
            for win in all_wins:
                if win.title == match_title:
                    return win, suggestions
            break

    return None, suggestions


def find_window_strict(title: str) -> Any:
    """Find a window or raise ToolError with suggestions."""
    win, suggestions = find_window(title)
    if win is None:
        msg = f"Window not found: '{title}'"
        suggestion = None
        if suggestions:
            suggestion = f"Similar windows: {', '.join(suggestions[:3])}. Use list_windows to see all."
        else:
            suggestion = "Use list_windows to see all available windows."
        raise ToolError(msg, suggestion=suggestion)
    return win


# ---------------------------------------------------------------------------
# Deduplication
# ---------------------------------------------------------------------------

def get_all_windows_deduped() -> list[Any]:
    """Return all open windows, deduplicated by HWND, excluding empty titles."""
    seen_handles: set[int] = set()
    result = []
    try:
        for win in gw.getAllWindows():
            if win.title and win.title.strip():
                hwnd = getattr(win, "_hWnd", id(win))
                if hwnd not in seen_handles:
                    seen_handles.add(hwnd)
                    result.append(win)
    except Exception as exc:
        logger.warning("Error enumerating windows: %s", exc)
    return result


# ---------------------------------------------------------------------------
# Window → PID
# ---------------------------------------------------------------------------

def get_window_pid(hwnd: int) -> int | None:
    """Return the PID that owns a window handle, or None."""
    try:
        pid = ctypes.wintypes.DWORD()
        ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        return int(pid.value) if pid.value else None
    except Exception:
        return None


# ---------------------------------------------------------------------------
# Window State Helpers
# ---------------------------------------------------------------------------

def get_window_details(win) -> dict:
    """Extract a detail dict from a pygetwindow window object."""
    hwnd = getattr(win, "_hWnd", None)
    pid = get_window_pid(hwnd) if hwnd else None
    try:
        is_responding = True  # pygetwindow doesn't expose this
        if pid:
            import psutil
            try:
                proc = psutil.Process(pid)
                is_responding = proc.status() != psutil.STATUS_ZOMBIE
            except Exception:
                pass
    except Exception:
        is_responding = None

    return {
        "title": win.title,
        "hwnd": hwnd,
        "pid": pid,
        "x": win.left,
        "y": win.top,
        "width": win.width,
        "height": win.height,
        "visible": win.visible,
        "minimized": win.isMinimized,
        "maximized": win.isMaximized,
        "is_responding": is_responding,
    }
