"""
UI Automation tools using the Windows UI Automation API.

Tools:
  - uia_inspect_window     Get the control tree of a window
  - uia_find_control       Find controls by type, name, or automation ID
  - uia_click_control      Click a control by name or automation ID
  - uia_get_control_value  Read the value/text of a control
  - uia_set_control_value  Set the value/text of a control
  - uia_get_focused        Get info about the currently focused control
"""

import asyncio
import json
import logging
from typing import Any, cast

from mcp.types import TextContent

from ..config import config
from ..registry import registry
from ..utils.args import get_int, get_str
from ..utils.errors import ToolError
from ..utils.security import redact_text

logger = logging.getLogger("win32-mcp")
_MAX_UIA_TEXT_CHARS = 512
_MAX_UIA_RESULTS = 100

# Map user-friendly lowercase names → uiautomation class name suffix.
# e.g. "tabitem" → "TabItemControl", "radiobutton" → "RadioButtonControl"
_CONTROL_TYPE_MAP: dict[str, str] = {
    "button": "ButtonControl",
    "checkbox": "CheckBoxControl",
    "combobox": "ComboBoxControl",
    "custom": "CustomControl",
    "datagrid": "DataGridControl",
    "dataitem": "DataItemControl",
    "document": "DocumentControl",
    "edit": "EditControl",
    "group": "GroupControl",
    "header": "HeaderControl",
    "headeritem": "HeaderItemControl",
    "hyperlink": "HyperlinkControl",
    "list": "ListControl",
    "listitem": "ListItemControl",
    "menu": "MenuControl",
    "menubar": "MenuBarControl",
    "menuitem": "MenuItemControl",
    "pane": "PaneControl",
    "radiobutton": "RadioButtonControl",
    "scrollbar": "ScrollBarControl",
    "separator": "SeparatorControl",
    "slider": "SliderControl",
    "spinner": "SpinnerControl",
    "statusbar": "StatusBarControl",
    "tab": "TabControl",
    "tabitem": "TabItemControl",
    "table": "TableControl",
    "text": "TextControl",
    "titlebar": "TitleBarControl",
    "toolbar": "ToolBarControl",
    "tooltip": "ToolTipControl",
    "tree": "TreeControl",
    "treeitem": "TreeItemControl",
    "window": "WindowControl",
}
_CONTROL_TYPES = set(_CONTROL_TYPE_MAP.keys())
_EDITABLE_CONTROL_TYPES = {"EditControl", "ComboBoxControl", "DocumentControl"}


def _import_uia() -> Any:
    """Import uiautomation, raising ToolError if not available."""
    try:
        import uiautomation

        return uiautomation
    except ImportError as exc:
        raise ToolError(
            "uiautomation package not installed",
            suggestion="pip install uiautomation",
        ) from exc


async def _run_in_thread(func: Any, auto: Any) -> Any:
    """Run *func* in a thread with COM initialized for UI Automation."""

    def _wrapper() -> Any:
        with auto.UIAutomationInitializerInThread():
            return func()

    return await asyncio.to_thread(_wrapper)


def _control_to_dict(ctrl: Any, depth: int = 0, max_depth: int = 1) -> dict[str, Any]:
    """Convert a UI Automation control to a serializable dict."""
    try:
        rect = ctrl.BoundingRectangle
        bounds = {
            "x": rect.left,
            "y": rect.top,
            "width": rect.width(),
            "height": rect.height(),
        }
    except Exception:
        bounds = None

    info: dict[str, Any] = {
        "control_type": ctrl.ControlTypeName,
        "name": ctrl.Name or "",
        "automation_id": ctrl.AutomationId or "",
        "class_name": ctrl.ClassName or "",
    }

    if bounds and bounds["width"] > 0:
        info["bounds"] = bounds
        info["center_x"] = bounds["x"] + bounds["width"] // 2
        info["center_y"] = bounds["y"] + bounds["height"] // 2

    # Include value for editable/readable controls
    try:
        val_pattern = ctrl.GetValuePattern()
        if val_pattern:
            info["value"] = val_pattern.Value
    except Exception:
        pass

    try:
        if ctrl.ControlTypeName in ("CheckBoxControl", "RadioButtonControl"):
            toggle = ctrl.GetTogglePattern()
            if toggle:
                info["checked"] = toggle.ToggleState == 1
    except Exception:
        pass

    try:
        info["is_enabled"] = ctrl.IsEnabled
    except Exception:
        pass

    # Recurse into children
    if depth < max_depth:
        children = []
        try:
            for child in ctrl.GetChildren():
                children.append(_control_to_dict(child, depth + 1, max_depth))
        except Exception:
            pass
        if children:
            info["children"] = children

    return info


def _find_window_control(auto: Any, window_title: str) -> Any:
    """Find a window control by title (exact or partial match)."""
    # Try exact match first
    try:
        win = auto.WindowControl(Name=window_title, searchDepth=1)
        if win.Exists(0, 0):
            return win
    except Exception:
        pass

    # Fall back to substring search
    try:
        win = auto.WindowControl(SubName=window_title, searchDepth=1)
        if win.Exists(0, 0):
            return win
    except Exception:
        pass

    raise ToolError(
        f"Window not found: '{window_title}'",
        suggestion="Use list_windows to see available window titles",
    )


def _get_control_class(auto: Any, control_type: str) -> Any:
    """Get the uiautomation control class for a type name."""
    class_name = _CONTROL_TYPE_MAP.get(control_type.lower())
    if class_name is None:
        raise ToolError(
            f"Unknown control type: '{control_type}'",
            suggestion=f"Valid types: {', '.join(sorted(_CONTROL_TYPES))}",
        )
    cls = getattr(auto, class_name, None)
    if cls is None:
        raise ToolError(
            f"Control class '{class_name}' not found in uiautomation library",
        )
    return cls


def _expected_control_type(control_type: str) -> str:
    """Return UIA ControlTypeName for a friendly type name."""
    expected = _CONTROL_TYPE_MAP.get(control_type.lower())
    if expected is None:
        raise ToolError(
            f"Unknown control type: '{control_type}'",
            suggestion=f"Valid types: {', '.join(sorted(_CONTROL_TYPES))}",
        )
    return expected


def _matches_control(ctrl: Any, name: str, aid: str, ctrl_type: str, *, exact_name: bool) -> bool:
    """Return True if a control matches search criteria."""
    if name:
        ctrl_name = ctrl.Name or ""
        if exact_name:
            if ctrl_name != name:
                return False
        elif name.lower() not in ctrl_name.lower():
            return False
    if aid and aid != (ctrl.AutomationId or ""):
        return False
    return not (ctrl_type and ctrl.ControlTypeName != _expected_control_type(ctrl_type))


def _find_control_in_window(
    win: Any,
    name: str,
    aid: str,
    ctrl_type: str,
    *,
    max_depth: int = 8,
) -> Any:
    """Find one matching control strictly within a window tree."""
    results: list[Any] = []
    _search_tree_objects(
        win,
        name=name,
        aid=aid,
        ctrl_type=ctrl_type,
        results=results,
        max_results=1,
        depth=0,
        max_depth=max_depth,
        exact_name=True,
    )
    if results:
        return results[0]
    raise ToolError(
        f"Control not found: name='{name}', automation_id='{aid}', control_type='{ctrl_type}'",
        suggestion="Use uia_inspect_window or uia_find_control to discover controls.",
    )


async def click_control_impl(
    window_title: str,
    *,
    name: str = "",
    automation_id: str = "",
    control_type: str = "",
    exact_name: bool = True,
) -> dict[str, Any]:
    """Click a UIA control inside a target window."""
    auto = _import_uia()

    def _click() -> dict[str, Any]:
        win = _find_window_control(auto, window_title)
        if control_type:
            _get_control_class(auto, control_type)
        ctrl = _find_control_object_in_window(
            win,
            name=name,
            aid=automation_id,
            ctrl_type=control_type,
            exact_name=exact_name,
        )
        ctrl_info = _control_to_dict(ctrl, depth=0, max_depth=0)
        try:
            invoke = ctrl.GetInvokePattern()
            if invoke:
                invoke.Invoke()
                return {"clicked": True, "method": "uia_invoke", "control": ctrl_info}
        except Exception:
            pass
        ctrl.Click()
        return {"clicked": True, "method": "uia_click", "control": ctrl_info}

    return cast("dict[str, Any]", await _run_in_thread(_click, auto))


async def set_control_value_impl(
    window_title: str,
    value: str,
    *,
    name: str = "",
    automation_id: str = "",
    control_type: str = "",
    exact_name: bool = True,
) -> dict[str, Any]:
    """Set a UIA control value inside a target window."""
    auto = _import_uia()

    def _set() -> dict[str, Any]:
        win = _find_window_control(auto, window_title)
        ctrl = _find_control_object_in_window(
            win,
            name=name,
            aid=automation_id,
            ctrl_type=control_type,
            exact_name=exact_name,
        )
        result = _set_control_value(ctrl, value)
        result["control"] = _control_to_dict(ctrl, depth=0, max_depth=0)
        return result

    return cast("dict[str, Any]", await _run_in_thread(_set, auto))


async def set_labeled_control_value_impl(
    window_title: str,
    label_text: str,
    value: str,
    *,
    direction: str = "right",
    exact_name: bool = False,
) -> dict[str, Any]:
    """Find an editable control near a label and set its value."""
    auto = _import_uia()

    def _set_labeled() -> dict[str, Any]:
        win = _find_window_control(auto, window_title)
        labels: list[Any] = []
        _search_tree_objects(
            win,
            name=label_text,
            aid="",
            ctrl_type="",
            results=labels,
            max_results=20,
            depth=0,
            max_depth=8,
            exact_name=exact_name,
        )
        fields: list[Any] = []
        _collect_editable_controls(win, fields, max_depth=8)
        if not labels or not fields:
            raise ToolError("No matching UIA label/editable field pair found")

        pair = _best_labeled_field(labels, fields, direction)
        if pair is None:
            raise ToolError("No editable UIA field found near label")
        label, field = pair
        result = _set_control_value(field, value)
        result["label_control"] = _control_to_dict(label, depth=0, max_depth=0)
        result["control"] = _control_to_dict(field, depth=0, max_depth=0)
        return result

    return cast("dict[str, Any]", await _run_in_thread(_set_labeled, auto))


def _find_control_object_in_window(
    win: Any,
    *,
    name: str,
    aid: str,
    ctrl_type: str,
    exact_name: bool,
) -> Any:
    results: list[Any] = []
    _search_tree_objects(
        win,
        name=name,
        aid=aid,
        ctrl_type=ctrl_type,
        results=results,
        max_results=1,
        depth=0,
        max_depth=8,
        exact_name=exact_name,
    )
    if results:
        return results[0]
    raise ToolError(
        f"Control not found: name='{name}', automation_id='{aid}', control_type='{ctrl_type}'",
        suggestion="Use uia_inspect_window or uia_find_control to discover controls.",
    )


def _set_control_value(ctrl: Any, value: str) -> dict[str, Any]:
    try:
        vp = ctrl.GetValuePattern()
        if vp:
            vp.SetValue(value)
            return {"set": True, "method": "value_pattern", "value": redact_text(value), "value_length": len(value)}
    except Exception:
        pass

    ctrl.SetFocus()
    ctrl.SendKeys("{Ctrl}a")
    ctrl.SendKeys(value, waitTime=0.05)
    return {"set": True, "method": "sendkeys", "value": redact_text(value), "value_length": len(value)}


def _collect_editable_controls(ctrl: Any, results: list[Any], max_depth: int, depth: int = 0) -> None:
    if depth > max_depth:
        return
    try:
        for child in ctrl.GetChildren():
            if child.ControlTypeName in _EDITABLE_CONTROL_TYPES:
                results.append(child)
            _collect_editable_controls(child, results, max_depth, depth + 1)
    except Exception:
        pass


def _best_labeled_field(labels: list[Any], fields: list[Any], direction: str) -> tuple[Any, Any] | None:
    best: tuple[float, Any, Any] | None = None
    for label in labels:
        label_bounds = _control_bounds(label)
        if label_bounds is None:
            continue
        for field in fields:
            field_bounds = _control_bounds(field)
            if field_bounds is None:
                continue
            score = _label_field_score(label_bounds, field_bounds, direction)
            if score is None:
                continue
            if best is None or score < best[0]:
                best = (score, label, field)
    return (best[1], best[2]) if best else None


def _control_bounds(ctrl: Any) -> tuple[int, int, int, int] | None:
    try:
        rect = ctrl.BoundingRectangle
        width = rect.width()
        height = rect.height()
        if width <= 0 or height <= 0:
            return None
        return int(rect.left), int(rect.top), int(width), int(height)
    except Exception:
        return None


def _label_field_score(
    label: tuple[int, int, int, int],
    field: tuple[int, int, int, int],
    direction: str,
) -> float | None:
    lx, ly, lw, lh = label
    fx, fy, fw, fh = field
    label_cy = ly + lh / 2
    field_cy = fy + fh / 2
    label_cx = lx + lw / 2
    field_cx = fx + fw / 2

    if direction == "below":
        if fy < ly + lh:
            return None
        return abs(field_cx - label_cx) + ((fy - ly - lh) * 2)

    if fx < lx + lw:
        return None
    return abs(field_cy - label_cy) + ((fx - lx - lw) * 2)


# ===================================================================
# MCP Tool Handlers
# ===================================================================


@registry.register(
    "uia_inspect_window",
    "Inspect the UI Automation control tree of a window. "
    "Returns controls with types, names, automation IDs, and bounds.",
    {
        "type": "object",
        "properties": {
            "window_title": {"type": "string", "description": "Full or partial window title"},
            "max_depth": {
                "type": "number",
                "minimum": 0,
                "maximum": 5,
                "description": "How many levels deep to inspect (default: 2, max: 5)",
            },
            "control_type": {
                "type": "string",
                "description": "Filter to only this control type (e.g. 'button', 'edit', 'checkbox')",
            },
        },
        "required": ["window_title"],
    },
)
async def handle_uia_inspect_window(arguments: dict[str, Any]) -> list[TextContent]:
    auto = _import_uia()
    window_title = get_str(arguments, "window_title", required=True, min_length=1, max_length=_MAX_UIA_TEXT_CHARS)
    max_depth = get_int(arguments, "max_depth", default=2, min_value=0, max_value=5)
    filter_type = get_str(arguments, "control_type", default="", max_length=_MAX_UIA_TEXT_CHARS).lower().strip()

    if filter_type and filter_type not in _CONTROL_TYPES:
        raise ToolError(
            f"Unknown control type: '{filter_type}'",
            suggestion=f"Valid types: {', '.join(sorted(_CONTROL_TYPES))}",
        )

    def _inspect() -> dict[str, Any]:
        win = _find_window_control(auto, window_title)
        tree = _control_to_dict(win, depth=0, max_depth=max_depth)

        if filter_type:
            # Flatten and filter
            target_type = _CONTROL_TYPE_MAP[filter_type]
            filtered = _collect_by_type(win, target_type, max_depth)
            tree["filtered_controls"] = filtered
            tree["filtered_count"] = len(filtered)

        return tree

    result = await _run_in_thread(_inspect, auto)
    return [TextContent(type="text", text=json.dumps(result, indent=2, default=str))]


def _collect_by_type(ctrl: Any, type_name: str, max_depth: int, depth: int = 0) -> list[dict[str, Any]]:
    """Recursively collect controls matching a type."""
    results: list[dict[str, Any]] = []
    if depth > max_depth:
        return results

    try:
        for child in ctrl.GetChildren():
            if child.ControlTypeName == type_name:
                results.append(_control_to_dict(child, depth=0, max_depth=0))
            results.extend(_collect_by_type(child, type_name, max_depth, depth + 1))
    except Exception:
        pass

    return results


@registry.register(
    "uia_find_control",
    "Find UI controls by name, automation ID, or type within a window",
    {
        "type": "object",
        "properties": {
            "window_title": {"type": "string", "description": "Full or partial window title"},
            "name": {"type": "string", "description": "Control name to search for (partial match)"},
            "automation_id": {"type": "string", "description": "Automation ID to search for (exact match)"},
            "control_type": {
                "type": "string",
                "description": "Control type filter (e.g. 'button', 'edit')",
            },
            "max_results": {
                "type": "number",
                "minimum": 1,
                "maximum": _MAX_UIA_RESULTS,
                "description": "Maximum results to return (default: 20)",
            },
        },
        "required": ["window_title"],
    },
)
async def handle_uia_find_control(arguments: dict[str, Any]) -> list[TextContent]:
    auto = _import_uia()
    window_title = get_str(arguments, "window_title", required=True, min_length=1, max_length=_MAX_UIA_TEXT_CHARS)
    search_name = get_str(arguments, "name", default="", max_length=_MAX_UIA_TEXT_CHARS).lower().strip()
    search_aid = get_str(arguments, "automation_id", default="", max_length=_MAX_UIA_TEXT_CHARS).strip()
    filter_type = get_str(arguments, "control_type", default="", max_length=_MAX_UIA_TEXT_CHARS).lower().strip()
    max_results = get_int(arguments, "max_results", default=20, min_value=1, max_value=_MAX_UIA_RESULTS)

    if not search_name and not search_aid and not filter_type:
        raise ToolError(
            "At least one of name, automation_id, or control_type must be provided",
        )

    def _find() -> list[dict[str, Any]]:
        win = _find_window_control(auto, window_title)
        results: list[dict[str, Any]] = []
        _search_tree(win, search_name, search_aid, filter_type, results, max_results, depth=0, max_depth=8)
        return results

    found = await _run_in_thread(_find, auto)
    output = {"found": len(found), "controls": found}
    return [TextContent(type="text", text=json.dumps(output, indent=2, default=str))]


def _search_tree(
    ctrl: Any,
    name: str,
    aid: str,
    ctrl_type: str,
    results: list[dict[str, Any]],
    max_results: int,
    depth: int,
    max_depth: int,
) -> None:
    """Recursively search the control tree."""
    if len(results) >= max_results or depth > max_depth:
        return

    try:
        for child in ctrl.GetChildren():
            if len(results) >= max_results:
                return

            if _matches_control(child, name, aid, ctrl_type, exact_name=False):
                results.append(_control_to_dict(child, depth=0, max_depth=0))

            _search_tree(child, name, aid, ctrl_type, results, max_results, depth + 1, max_depth)
    except Exception:
        pass


def _search_tree_objects(
    ctrl: Any,
    name: str,
    aid: str,
    ctrl_type: str,
    results: list[Any],
    max_results: int,
    depth: int,
    max_depth: int,
    *,
    exact_name: bool,
) -> None:
    """Recursively search the control tree and keep live controls."""
    if len(results) >= max_results or depth > max_depth:
        return

    try:
        for child in ctrl.GetChildren():
            if len(results) >= max_results:
                return

            if _matches_control(child, name, aid, ctrl_type, exact_name=exact_name):
                results.append(child)

            _search_tree_objects(
                child,
                name=name,
                aid=aid,
                ctrl_type=ctrl_type,
                results=results,
                max_results=max_results,
                depth=depth + 1,
                max_depth=max_depth,
                exact_name=exact_name,
            )
    except Exception:
        pass


@registry.register(
    "uia_click_control",
    "Click a UI control by name or automation ID. More reliable than coordinate-based clicking.",
    {
        "type": "object",
        "properties": {
            "window_title": {"type": "string", "description": "Full or partial window title"},
            "name": {"type": "string", "description": "Control name to click"},
            "automation_id": {"type": "string", "description": "Automation ID to click"},
            "control_type": {"type": "string", "description": "Control type (e.g. 'button')"},
        },
        "required": ["window_title"],
    },
)
async def handle_uia_click_control(arguments: dict[str, Any]) -> dict[str, Any]:
    window_title = get_str(arguments, "window_title", required=True, min_length=1, max_length=_MAX_UIA_TEXT_CHARS)
    name = get_str(arguments, "name", default="", max_length=_MAX_UIA_TEXT_CHARS).strip()
    aid = get_str(arguments, "automation_id", default="", max_length=_MAX_UIA_TEXT_CHARS).strip()
    ctrl_type = get_str(arguments, "control_type", default="", max_length=_MAX_UIA_TEXT_CHARS).lower().strip()

    if not name and not aid:
        raise ToolError("Either name or automation_id must be provided")

    return await click_control_impl(
        window_title,
        name=name,
        automation_id=aid,
        control_type=ctrl_type,
        exact_name=True,
    )


@registry.register(
    "uia_get_control_value",
    "Read the value or text of a UI control",
    {
        "type": "object",
        "properties": {
            "window_title": {"type": "string", "description": "Full or partial window title"},
            "name": {"type": "string", "description": "Control name"},
            "automation_id": {"type": "string", "description": "Automation ID"},
            "control_type": {"type": "string", "description": "Control type (e.g. 'edit', 'text')"},
        },
        "required": ["window_title"],
    },
)
async def handle_uia_get_control_value(arguments: dict[str, Any]) -> dict[str, Any]:
    auto = _import_uia()
    window_title = get_str(arguments, "window_title", required=True, min_length=1, max_length=_MAX_UIA_TEXT_CHARS)
    name = get_str(arguments, "name", default="", max_length=_MAX_UIA_TEXT_CHARS).strip()
    aid = get_str(arguments, "automation_id", default="", max_length=_MAX_UIA_TEXT_CHARS).strip()
    ctrl_type = get_str(arguments, "control_type", default="", max_length=_MAX_UIA_TEXT_CHARS).lower().strip()

    if not name and not aid:
        raise ToolError("Either name or automation_id must be provided")

    def _get_value() -> dict[str, Any]:
        win = _find_window_control(auto, window_title)
        if ctrl_type:
            _get_control_class(auto, ctrl_type)
        ctrl = _find_control_in_window(win, name, aid, ctrl_type)

        result: dict[str, Any] = _control_to_dict(ctrl, depth=0, max_depth=0)

        # Try to get value via ValuePattern
        try:
            vp = ctrl.GetValuePattern()
            if vp:
                result["value"] = vp.Value
        except Exception:
            pass

        # Try to get text via TextPattern
        try:
            tp = ctrl.GetTextPattern()
            if tp:
                result["text"] = tp.DocumentRange.GetText(-1)
        except Exception:
            pass

        # For text controls, Name is often the value
        if not result.get("value") and not result.get("text"):
            result["value"] = ctrl.Name

        return result

    return cast("dict[str, Any]", await _run_in_thread(_get_value, auto))


@registry.register(
    "uia_set_control_value",
    "Set the value or text of a UI control (edit boxes, combo boxes, etc.)",
    {
        "type": "object",
        "properties": {
            "window_title": {"type": "string", "description": "Full or partial window title"},
            "value": {"type": "string", "description": "Value to set"},
            "name": {"type": "string", "description": "Control name"},
            "automation_id": {"type": "string", "description": "Automation ID"},
            "control_type": {"type": "string", "description": "Control type (e.g. 'edit')"},
        },
        "required": ["window_title", "value"],
    },
)
async def handle_uia_set_control_value(arguments: dict[str, Any]) -> dict[str, Any]:
    window_title = get_str(arguments, "window_title", required=True, min_length=1, max_length=_MAX_UIA_TEXT_CHARS)
    value = get_str(arguments, "value", required=True, max_length=config.limits.max_text_chars)
    name = get_str(arguments, "name", default="", max_length=_MAX_UIA_TEXT_CHARS).strip()
    aid = get_str(arguments, "automation_id", default="", max_length=_MAX_UIA_TEXT_CHARS).strip()
    ctrl_type = get_str(arguments, "control_type", default="", max_length=_MAX_UIA_TEXT_CHARS).lower().strip()

    if not name and not aid:
        raise ToolError("Either name or automation_id must be provided")

    return await set_control_value_impl(
        window_title,
        value,
        name=name,
        automation_id=aid,
        control_type=ctrl_type,
        exact_name=True,
    )


@registry.register(
    "uia_get_focused",
    "Get information about the currently focused UI control",
    {
        "type": "object",
        "properties": {},
    },
)
async def handle_uia_get_focused(arguments: dict[str, Any]) -> dict[str, Any]:
    auto = _import_uia()

    def _get_focused() -> dict[str, Any]:
        ctrl = auto.GetFocusedControl()
        if ctrl is None:
            return {"focused": False}

        info = _control_to_dict(ctrl, depth=0, max_depth=0)
        info["focused"] = True

        # Walk up to find parent window
        try:
            parent = ctrl
            for _ in range(20):
                parent = parent.GetParentControl()
                if parent is None:
                    break
                if parent.ControlTypeName == "WindowControl":
                    info["window_title"] = parent.Name
                    break
        except Exception:
            pass

        return info

    return cast("dict[str, Any]", await _run_in_thread(_get_focused, auto))
