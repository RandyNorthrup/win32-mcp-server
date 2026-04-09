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

from ..registry import registry
from ..utils.errors import ToolError

logger = logging.getLogger("win32-mcp")

# Control type names that map to uiautomation classes
_CONTROL_TYPES = {
    "button",
    "checkbox",
    "combobox",
    "edit",
    "hyperlink",
    "list",
    "listitem",
    "menu",
    "menubar",
    "menuitem",
    "pane",
    "radiobutton",
    "separator",
    "slider",
    "tab",
    "tabitem",
    "text",
    "toolbar",
    "tree",
    "treeitem",
    "window",
    "group",
    "document",
    "custom",
    "dataitem",
    "datagrid",
    "header",
    "headeritem",
    "scrollbar",
    "spinner",
    "statusbar",
    "table",
    "titlebar",
    "tooltip",
}


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
    # Map user-friendly names to class names
    class_name = control_type.capitalize() + "Control"
    cls = getattr(auto, class_name, None)
    if cls is None:
        raise ToolError(
            f"Unknown control type: '{control_type}'",
            suggestion=f"Valid types: {', '.join(sorted(_CONTROL_TYPES))}",
        )
    return cls


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
    window_title = arguments["window_title"]
    max_depth = min(int(arguments.get("max_depth", 2)), 5)
    filter_type = arguments.get("control_type", "").lower().strip()

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
            target_type = filter_type.capitalize() + "Control"
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
            "max_results": {"type": "number", "description": "Maximum results to return (default: 20)"},
        },
        "required": ["window_title"],
    },
)
async def handle_uia_find_control(arguments: dict[str, Any]) -> list[TextContent]:
    auto = _import_uia()
    window_title = arguments["window_title"]
    search_name = arguments.get("name", "").lower().strip()
    search_aid = arguments.get("automation_id", "").strip()
    filter_type = arguments.get("control_type", "").lower().strip()
    max_results = int(arguments.get("max_results", 20))

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

            matches = True
            if name and name not in (child.Name or "").lower():
                matches = False
            if aid and aid != (child.AutomationId or ""):
                matches = False
            if ctrl_type:
                expected_type = ctrl_type.capitalize() + "Control"
                if child.ControlTypeName != expected_type:
                    matches = False

            if matches:
                results.append(_control_to_dict(child, depth=0, max_depth=0))

            _search_tree(child, name, aid, ctrl_type, results, max_results, depth + 1, max_depth)
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
    auto = _import_uia()
    window_title = arguments["window_title"]
    name = arguments.get("name", "").strip()
    aid = arguments.get("automation_id", "").strip()
    ctrl_type = arguments.get("control_type", "").lower().strip()

    if not name and not aid:
        raise ToolError("Either name or automation_id must be provided")

    def _click() -> dict[str, Any]:
        win = _find_window_control(auto, window_title)

        # Build search criteria
        kwargs: dict[str, Any] = {}
        if name:
            kwargs["Name"] = name
        if aid:
            kwargs["AutomationId"] = aid

        if ctrl_type:
            cls = _get_control_class(auto, ctrl_type)
            ctrl = win.Control(**kwargs) if cls is auto.Control else cls(**kwargs)
        else:
            ctrl = win.Control(**kwargs)

        if not ctrl.Exists(3, 0.5):
            raise ToolError(
                f"Control not found: name='{name}', automation_id='{aid}'",
                suggestion="Use uia_inspect_window or uia_find_control to discover controls",
            )

        ctrl_info = _control_to_dict(ctrl, depth=0, max_depth=0)

        try:
            invoke = ctrl.GetInvokePattern()
            if invoke:
                invoke.Invoke()
                return {"clicked": True, "method": "invoke", "control": ctrl_info}
        except Exception:
            pass

        # Fall back to click
        try:
            ctrl.Click()
            return {"clicked": True, "method": "click", "control": ctrl_info}
        except Exception as exc:
            raise ToolError(f"Failed to click control: {exc}") from exc

    return cast("dict[str, Any]", await _run_in_thread(_click, auto))


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
    window_title = arguments["window_title"]
    name = arguments.get("name", "").strip()
    aid = arguments.get("automation_id", "").strip()
    ctrl_type = arguments.get("control_type", "").lower().strip()

    if not name and not aid:
        raise ToolError("Either name or automation_id must be provided")

    def _get_value() -> dict[str, Any]:
        win = _find_window_control(auto, window_title)

        kwargs: dict[str, Any] = {}
        if name:
            kwargs["Name"] = name
        if aid:
            kwargs["AutomationId"] = aid

        if ctrl_type:
            cls = _get_control_class(auto, ctrl_type)
            ctrl = cls(**kwargs) if cls is not auto.Control else win.Control(**kwargs)
        else:
            ctrl = win.Control(**kwargs)

        if not ctrl.Exists(3, 0.5):
            raise ToolError(f"Control not found: name='{name}', automation_id='{aid}'")

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
    auto = _import_uia()
    window_title = arguments["window_title"]
    value = arguments["value"]
    name = arguments.get("name", "").strip()
    aid = arguments.get("automation_id", "").strip()
    ctrl_type = arguments.get("control_type", "").lower().strip()

    if not name and not aid:
        raise ToolError("Either name or automation_id must be provided")

    def _set_value() -> dict[str, Any]:
        win = _find_window_control(auto, window_title)

        kwargs: dict[str, Any] = {}
        if name:
            kwargs["Name"] = name
        if aid:
            kwargs["AutomationId"] = aid

        if ctrl_type:
            cls = _get_control_class(auto, ctrl_type)
            ctrl = cls(**kwargs) if cls is not auto.Control else win.Control(**kwargs)
        else:
            ctrl = win.Control(**kwargs)

        if not ctrl.Exists(3, 0.5):
            raise ToolError(f"Control not found: name='{name}', automation_id='{aid}'")

        # Try ValuePattern first
        try:
            vp = ctrl.GetValuePattern()
            if vp:
                vp.SetValue(value)
                return {"set": True, "method": "value_pattern", "value": value}
        except Exception:
            pass

        # Fall back: focus and type
        try:
            ctrl.SetFocus()
            ctrl.SendKeys("{Ctrl}a")
            ctrl.SendKeys(value, waitTime=0.05)
            return {"set": True, "method": "sendkeys", "value": value}
        except Exception as exc:
            raise ToolError(f"Failed to set control value: {exc}") from exc

    return cast("dict[str, Any]", await _run_in_thread(_set_value, auto))


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
