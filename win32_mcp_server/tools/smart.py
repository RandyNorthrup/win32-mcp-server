"""
High-level smart automation tools.

These tools compose multiple low-level operations into single
intelligent actions — the biggest usability improvement for AI agents.

Tools:
  - find_text_on_screen   Find text anywhere on screen, returns coordinates
  - click_text            Find text and click on it (the #1 most useful tool)
  - wait_for_text         Wait until text appears on screen
  - assert_text_visible   Verify text presence for test assertions
  - fill_field            Click a labeled field and type a value
  - get_window_snapshot   Complete window state: screenshot + OCR + positions
  - right_click_menu      Right-click and OCR the context menu
  - execute_sequence      Run multiple tools in sequence without round-trips
"""

import asyncio
import json
import logging
import time
from typing import Any

import pyautogui
import pyperclip
from mcp.types import ImageContent, TextContent

from ..config import PreprocessMode, config
from ..registry import registry
from ..utils.coordinates import clamp_rect_to_virtual_screen, validate_coordinates
from ..utils.errors import ToolError
from ..utils.imaging import image_to_base64
from ..utils.window_match import _fuzzy_ratio, find_window_strict
from .capture import capture_region_impl, capture_screen_impl, capture_window_impl
from .ocr import ocr_structured_impl

logger = logging.getLogger("win32-mcp")


# ===================================================================
# Text matching engine
# ===================================================================


def _find_text_matches(
    ocr_results: list[dict[str, Any]],
    search_text: str,
    exact: bool = False,
    threshold: int = 75,
) -> list[dict[str, Any]]:
    """Find text matches in OCR results, handling multi-word phrases.

    Strategy:
      1. Try exact substring match within grouped lines.
      2. Fall back to fuzzy matching per-line.
      3. Fall back to fuzzy matching per-word.

    Returns list of match dicts with center_x, center_y for clicking.
    """
    search_lower = search_text.lower().strip()
    if not search_lower:
        return []

    # Group OCR results by line
    lines: dict[tuple[int, int], list[dict[str, Any]]] = {}
    for r in ocr_results:
        key = (r.get("block_num", 0), r.get("line_num", 0))
        lines.setdefault(key, []).append(r)

    matches: list[dict[str, Any]] = []

    for words in lines.values():
        words.sort(key=lambda w: w["x"])
        line_text = " ".join(w["text"] for w in words)
        line_lower = line_text.lower()

        # --- Exact substring match within line ---
        idx = line_lower.find(search_lower)
        if idx >= 0:
            match_words = _extract_matching_words(words, line_text, idx, len(search_text))
            if match_words:
                bbox = _merge_bounding_boxes(match_words)
                bbox["text"] = " ".join(w["text"] for w in match_words)
                bbox["confidence"] = min(w["confidence"] for w in match_words)
                bbox["match_type"] = "exact"
                bbox["match_score"] = 100
                matches.append(bbox)
                continue

        # --- Fuzzy match on full line ---
        if not exact:
            score = _fuzzy_ratio(search_text, line_text)
            if score >= threshold:
                bbox = _merge_bounding_boxes(words)
                bbox["text"] = line_text
                bbox["confidence"] = min(w["confidence"] for w in words)
                bbox["match_type"] = "fuzzy_line"
                bbox["match_score"] = score
                matches.append(bbox)
                continue

    # --- Fuzzy match on individual words (fallback) ---
    if not matches and not exact:
        for r in ocr_results:
            score = _fuzzy_ratio(search_text, r["text"])
            if score >= threshold:
                bbox = {
                    "text": r["text"],
                    "x": r["x"],
                    "y": r["y"],
                    "width": r["width"],
                    "height": r["height"],
                    "confidence": r["confidence"],
                    "match_type": "fuzzy_word",
                    "match_score": score,
                    "center_x": r["x"] + r["width"] // 2,
                    "center_y": r["y"] + r["height"] // 2,
                }
                matches.append(bbox)

    # Sort: exact > fuzzy, higher score > lower, higher confidence > lower
    type_order = {"exact": 3, "fuzzy_line": 2, "fuzzy_word": 1}
    matches.sort(
        key=lambda m: (
            type_order.get(m.get("match_type", ""), 0),
            m.get("match_score", 0),
            m.get("confidence", 0),
        ),
        reverse=True,
    )

    return matches


def _extract_matching_words(
    words: list[dict[str, Any]],
    line_text: str,
    start_idx: int,
    match_len: int,
) -> list[dict[str, Any]]:
    """Given a character-level match position in the line, find which OCR words correspond."""
    result = []
    char_pos = 0
    end_idx = start_idx + match_len

    for w in words:
        word_start = char_pos
        word_end = char_pos + len(w["text"])

        # Check if this word overlaps with the match range
        if word_end > start_idx and word_start < end_idx:
            result.append(w)

        char_pos = word_end + 1  # +1 for space

    return result


def _merge_bounding_boxes(words: list[dict[str, Any]]) -> dict[str, Any]:
    """Merge multiple word bounding boxes into one encompassing box."""
    if not words:
        return {"x": 0, "y": 0, "width": 0, "height": 0, "center_x": 0, "center_y": 0}

    min_x = min(w["x"] for w in words)
    min_y = min(w["y"] for w in words)
    max_x = max(w["x"] + w["width"] for w in words)
    max_y = max(w["y"] + w["height"] for w in words)

    width = max_x - min_x
    height = max_y - min_y

    return {
        "x": min_x,
        "y": min_y,
        "width": width,
        "height": height,
        "center_x": min_x + width // 2,
        "center_y": min_y + height // 2,
    }


# ===================================================================
# Core implementation used by multiple smart tools
# ===================================================================


async def find_text_impl(
    search_text: str,
    window_title: str | None = None,
    exact: bool = False,
    threshold: int = 75,
    lang: str | None = None,
    preprocess: PreprocessMode | None = None,
) -> list[dict[str, Any]]:
    """Find text on screen and return matches with screen coordinates.

    Returns list of match dicts, each with:
      text, x, y, width, height, center_x, center_y,
      screen_x, screen_y, screen_center_x, screen_center_y,
      confidence, match_type, match_score
    """
    offset_x, offset_y = 0, 0

    if window_title:
        win = find_window_strict(window_title)
        # Clamp to virtual screen — handles multi-monitor with negative coordinates
        left, top, width, height = clamp_rect_to_virtual_screen(
            win.left,
            win.top,
            win.width,
            win.height,
        )
        if width <= 0 or height <= 0:
            raise ToolError(
                f"Window '{win.title}' is entirely off-screen",
                suggestion="Move the window on-screen first using move_window",
            )
        img = await capture_region_impl(left, top, width, height)
        offset_x = left
        offset_y = top
    else:
        img = await capture_screen_impl()

    ocr_results = await ocr_structured_impl(
        img,
        lang=lang,
        preprocess=preprocess,
    )

    matches = _find_text_matches(ocr_results, search_text, exact=exact, threshold=threshold)

    # Add screen-space coordinates
    for m in matches:
        m["screen_x"] = m["x"] + offset_x
        m["screen_y"] = m["y"] + offset_y
        m["screen_center_x"] = m["center_x"] + offset_x
        m["screen_center_y"] = m["center_y"] + offset_y

    return matches


# ===================================================================
# MCP Tool Handlers
# ===================================================================


@registry.register(
    "find_text_on_screen",
    "Find all occurrences of text on screen with coordinates",
    {
        "type": "object",
        "properties": {
            "text": {"type": "string", "description": "Text to search for"},
            "window_title": {
                "type": "string",
                "description": "Optional: limit search to this window",
            },
            "exact": {
                "type": "boolean",
                "description": "Require exact substring match (default: false = fuzzy)",
            },
            "threshold": {
                "type": "number",
                "description": "Fuzzy match threshold 0-100 (default: 75)",
            },
        },
        "required": ["text"],
    },
)
async def handle_find_text_on_screen(arguments: dict[str, Any]) -> list[TextContent]:
    text = arguments["text"]
    matches = await find_text_impl(
        text,
        window_title=arguments.get("window_title"),
        exact=arguments.get("exact", False),
        threshold=arguments.get("threshold", 75),
    )

    return [
        TextContent(
            type="text",
            text=json.dumps(
                {
                    "search_text": text,
                    "match_count": len(matches),
                    "matches": matches,
                },
                indent=2,
            ),
        )
    ]


@registry.register(
    "click_text",
    "Find text on screen and click on it",
    {
        "type": "object",
        "properties": {
            "text": {"type": "string", "description": "Text to find and click"},
            "window_title": {
                "type": "string",
                "description": "Optional: limit search to this window",
            },
            "button": {
                "type": "string",
                "enum": ["left", "right", "middle"],
                "description": "Mouse button (default: left)",
            },
            "occurrence": {
                "type": "number",
                "description": "Which occurrence to click: 1=first, 2=second, etc. (default: 1)",
            },
            "exact": {
                "type": "boolean",
                "description": "Require exact match (default: false)",
            },
            "threshold": {
                "type": "number",
                "description": "Fuzzy match threshold 0-100 (default: 75)",
            },
        },
        "required": ["text"],
    },
)
async def handle_click_text(arguments: dict[str, Any]) -> dict[str, Any]:
    text = arguments["text"]
    button = arguments.get("button", "left")
    occurrence = int(arguments.get("occurrence", 1))

    matches = await find_text_impl(
        text,
        window_title=arguments.get("window_title"),
        exact=arguments.get("exact", False),
        threshold=arguments.get("threshold", 75),
    )

    if not matches:
        raise ToolError(
            f"Text '{text}' not found on screen",
            suggestion="Use find_text_on_screen to check what text is visible, or capture_screen to see the UI.",
        )

    if occurrence > len(matches):
        raise ToolError(
            f"Only {len(matches)} occurrence(s) of '{text}' found, but occurrence={occurrence} requested",
            suggestion=f"Use occurrence=1 through {len(matches)}",
        )

    match = matches[occurrence - 1]
    click_x = match["screen_center_x"]
    click_y = match["screen_center_y"]

    await asyncio.to_thread(pyautogui.click, click_x, click_y, button=button)
    await asyncio.sleep(config.automation.click_delay)

    return {
        "clicked_text": match["text"],
        "clicked_at": {"x": click_x, "y": click_y},
        "match_type": match.get("match_type"),
        "match_score": match.get("match_score"),
        "confidence": match.get("confidence"),
        "total_matches": len(matches),
    }


@registry.register(
    "wait_for_text",
    "Wait until specific text appears on screen",
    {
        "type": "object",
        "properties": {
            "text": {"type": "string", "description": "Text to wait for"},
            "window_title": {
                "type": "string",
                "description": "Optional: limit search to this window",
            },
            "timeout_seconds": {
                "type": "number",
                "description": "Maximum wait time in seconds (default: 10)",
            },
            "poll_interval": {
                "type": "number",
                "description": "Seconds between OCR checks (default: 0.5)",
            },
            "exact": {"type": "boolean"},
            "threshold": {"type": "number"},
        },
        "required": ["text"],
    },
)
async def handle_wait_for_text(arguments: dict[str, Any]) -> dict[str, Any]:
    text = arguments["text"]
    timeout = arguments.get("timeout_seconds", config.default_timeout)
    interval = arguments.get("poll_interval", 0.5)
    window_title = arguments.get("window_title")
    exact = arguments.get("exact", False)
    threshold = arguments.get("threshold", 75)

    start = time.monotonic()
    attempts = 0

    while True:
        attempts += 1
        try:
            matches = await find_text_impl(
                text,
                window_title=window_title,
                exact=exact,
                threshold=threshold,
            )
            if matches:
                elapsed = round(time.monotonic() - start, 2)
                best = matches[0]
                return {
                    "found": True,
                    "text": best["text"],
                    "screen_x": best["screen_center_x"],
                    "screen_y": best["screen_center_y"],
                    "confidence": best.get("confidence"),
                    "elapsed_seconds": elapsed,
                    "attempts": attempts,
                    "total_matches": len(matches),
                }
        except ToolError as exc:
            # Re-raise permanent errors (e.g. Tesseract not installed)
            if "not installed" in str(exc).lower():
                raise
            logger.debug("Transient error during text search (attempt %d): %s", attempts, exc)

        elapsed = time.monotonic() - start
        if elapsed >= timeout:
            raise ToolError(
                f"Text '{text}' did not appear within {timeout}s ({attempts} checks)",
                suggestion=(
                    "Increase timeout_seconds, check if the correct window is open, "
                    "or use capture_screen to see what's visible."
                ),
            )

        await asyncio.sleep(interval)


@registry.register(
    "assert_text_visible",
    "Assert that text is or is not visible on screen (for testing)",
    {
        "type": "object",
        "properties": {
            "text": {"type": "string", "description": "Text to check for"},
            "should_exist": {
                "type": "boolean",
                "description": "True=assert visible, False=assert NOT visible (default: true)",
            },
            "window_title": {"type": "string"},
            "exact": {"type": "boolean"},
            "threshold": {"type": "number"},
        },
        "required": ["text"],
    },
)
async def handle_assert_text_visible(arguments: dict[str, Any]) -> dict[str, Any]:
    text = arguments["text"]
    should_exist = arguments.get("should_exist", True)
    window_title = arguments.get("window_title")

    matches = await find_text_impl(
        text,
        window_title=window_title,
        exact=arguments.get("exact", False),
        threshold=arguments.get("threshold", 75),
    )

    found = len(matches) > 0

    if should_exist and not found:
        return {
            "passed": False,
            "assertion": f"Expected '{text}' to be visible, but it was NOT found",
            "match_count": 0,
        }
    if not should_exist and found:
        return {
            "passed": False,
            "assertion": f"Expected '{text}' to NOT be visible, but it WAS found",
            "match_count": len(matches),
            "found_at": [
                {"x": m["screen_center_x"], "y": m["screen_center_y"], "text": m["text"]} for m in matches[:3]
            ],
        }
    return {
        "passed": True,
        "assertion": f"'{text}' {'is' if should_exist else 'is not'} visible as expected",
        "match_count": len(matches),
    }


@registry.register(
    "fill_field",
    "Click a labeled input field and type a value",
    {
        "type": "object",
        "properties": {
            "label_text": {
                "type": "string",
                "description": "Label text next to the input field",
            },
            "value": {
                "type": "string",
                "description": "Text to type into the field",
            },
            "window_title": {"type": "string"},
            "direction": {
                "type": "string",
                "enum": ["right", "below"],
                "description": "Where the input field is relative to the label (default: right)",
            },
            "clear_first": {
                "type": "boolean",
                "description": "Select all and clear existing text before typing (default: true)",
            },
            "offset_px": {
                "type": "number",
                "description": "Pixel offset from label to click point (default: 50)",
            },
        },
        "required": ["label_text", "value"],
    },
)
async def handle_fill_field(arguments: dict[str, Any]) -> dict[str, Any]:
    label = arguments["label_text"]
    value = arguments["value"]
    direction = arguments.get("direction", "right")
    clear_first = arguments.get("clear_first", True)
    offset = int(arguments.get("offset_px", 50))
    window_title = arguments.get("window_title")

    # Find the label
    matches = await find_text_impl(label, window_title=window_title)
    if not matches:
        raise ToolError(
            f"Label '{label}' not found on screen",
            suggestion="Use find_text_on_screen to see what text is visible",
        )

    match = matches[0]

    # Calculate click point based on direction
    if direction == "below":
        click_x = match["screen_center_x"]
        click_y = match["screen_y"] + match["height"] + offset
    else:  # right
        click_x = match["screen_x"] + match["width"] + offset
        click_y = match["screen_center_y"]

    # Click the field
    await asyncio.to_thread(pyautogui.click, click_x, click_y)
    await asyncio.sleep(0.15)

    # Clear existing text if requested
    if clear_first:
        await asyncio.to_thread(pyautogui.hotkey, "ctrl", "a")
        await asyncio.sleep(0.05)
        await asyncio.to_thread(pyautogui.press, "delete")
        await asyncio.sleep(0.05)

    # Type the value — use clipboard for non-ASCII
    if not value.isascii():
        try:
            old_clip = pyperclip.paste()
        except Exception as exc:
            logger.debug("Could not read clipboard: %s", exc)
            old_clip = ""
        pyperclip.copy(value)
        await asyncio.sleep(0.05)
        await asyncio.to_thread(pyautogui.hotkey, "ctrl", "v")
        await asyncio.sleep(0.1)
        try:
            pyperclip.copy(old_clip)
        except Exception as exc:
            logger.debug("Could not restore clipboard: %s", exc)
    else:
        await asyncio.to_thread(
            pyautogui.write,
            value,
            interval=config.automation.type_interval,
        )

    return {
        "label": label,
        "value": value,
        "label_found_at": {"x": match["screen_center_x"], "y": match["screen_center_y"]},
        "clicked_at": {"x": click_x, "y": click_y},
        "direction": direction,
        "cleared": clear_first,
    }


@registry.register(
    "get_window_snapshot",
    "Capture complete window state: screenshot + OCR text with positions",
    {
        "type": "object",
        "properties": {
            "window_title": {"type": "string", "description": "Full or partial window title"},
            "format": {"type": "string", "enum": ["png", "jpeg", "webp"]},
            "quality": {"type": "number"},
            "scale": {"type": "number"},
            "include_ocr": {
                "type": "boolean",
                "description": "Include structured OCR results (default: true)",
            },
        },
        "required": ["window_title"],
    },
)
async def handle_get_window_snapshot(arguments: dict[str, Any]) -> list[TextContent | ImageContent]:
    window_title = arguments["window_title"]
    fmt = arguments.get("format", config.capture.default_format)
    quality = arguments.get("quality", config.capture.default_quality)
    scale = arguments.get("scale", config.capture.default_scale)
    include_ocr = arguments.get("include_ocr", True)

    # Capture window
    img, win_info = await capture_window_impl(window_title)

    # OCR if requested
    ocr_elements = []
    if include_ocr:
        ocr_results = await ocr_structured_impl(img)
        # Offset to screen coordinates — use actual capture position, not window origin
        cap_x = win_info.get("capture_x", win_info["x"])
        cap_y = win_info.get("capture_y", win_info["y"])
        for r in ocr_results:
            r["screen_x"] = r["x"] + cap_x
            r["screen_y"] = r["y"] + cap_y
        ocr_elements = ocr_results

    # Encode image
    data, mime, size = image_to_base64(img, fmt=fmt, quality=quality, scale=scale)

    result_data = {
        "window": win_info,
        "element_count": len(ocr_elements),
        "ocr_elements": ocr_elements,
        "image_format": fmt,
        "image_size_kb": round(size / 1024, 1),
    }

    return [
        TextContent(type="text", text=json.dumps(result_data, indent=2)),
        ImageContent(type="image", data=data, mimeType=mime),
    ]


@registry.register(
    "right_click_menu",
    "Right-click at a position and OCR the context menu",
    {
        "type": "object",
        "properties": {
            "x": {"type": "number", "description": "X coordinate to right-click"},
            "y": {"type": "number", "description": "Y coordinate to right-click"},
            "window_title": {"type": "string", "description": "Optional: window context"},
            "menu_width": {
                "type": "number",
                "description": "Expected menu width in pixels (default: 300)",
            },
            "menu_height": {
                "type": "number",
                "description": "Expected menu height in pixels (default: 400)",
            },
        },
        "required": ["x", "y"],
    },
)
async def handle_right_click_menu(arguments: dict[str, Any]) -> dict[str, Any]:
    x, y = int(arguments["x"]), int(arguments["y"])
    menu_w = int(arguments.get("menu_width", 300))
    menu_h = int(arguments.get("menu_height", 400))

    if config.validate_coordinates:
        validate_coordinates(x, y, "right_click_menu")

    # Right-click
    await asyncio.to_thread(pyautogui.click, x, y, button="right")
    await asyncio.sleep(0.5)  # Wait for menu to appear

    # Capture region around click point (menu appears near click)
    # Menus typically open to the right and below the click point
    region_x, region_y, menu_w, menu_h = clamp_rect_to_virtual_screen(
        x - 10,
        y - 10,
        menu_w + 10,
        menu_h + 10,
    )

    try:
        img = await capture_region_impl(region_x, region_y, menu_w, menu_h)
    except Exception as exc:
        logger.debug(
            "Could not capture menu region (%d,%d %dx%d), falling back to full screen: %s",
            region_x,
            region_y,
            menu_w,
            menu_h,
            exc,
        )
        img = await capture_screen_impl()
        region_x, region_y = 0, 0

    # OCR the menu region
    ocr_results = await ocr_structured_impl(img)

    # Group into menu items (each line = menu item)
    lines: dict[int, list[dict[str, Any]]] = {}
    for r in ocr_results:
        lines.setdefault(r["line_num"], []).append(r)

    menu_items = []
    for line_num in sorted(lines.keys()):
        words = sorted(lines[line_num], key=lambda w: w["x"])
        text = " ".join(w["text"] for w in words).strip()
        if not text or text in {"-", "—"}:
            continue

        # Convert to screen coordinates
        item_x = words[0]["x"] + region_x
        item_y = words[0]["y"] + region_y
        total_w = words[-1]["x"] + words[-1]["width"] - words[0]["x"]
        item_h = max(w["height"] for w in words)

        menu_items.append(
            {
                "text": text,
                "x": item_x,
                "y": item_y,
                "center_x": item_x + total_w // 2,
                "center_y": item_y + item_h // 2,
                "confidence": min(w["confidence"] for w in words),
            }
        )

    return {
        "right_clicked_at": {"x": x, "y": y},
        "menu_item_count": len(menu_items),
        "menu_items": menu_items,
    }


@registry.register(
    "execute_sequence",
    "Execute multiple tools in sequence without round-trip overhead",
    {
        "type": "object",
        "properties": {
            "steps": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "tool": {"type": "string", "description": "Tool name to execute"},
                        "args": {"type": "object", "description": "Tool arguments"},
                        "delay_ms": {
                            "type": "number",
                            "description": "Delay after this step in milliseconds (default: 0)",
                        },
                    },
                    "required": ["tool"],
                },
                "description": "Array of tool invocations to execute in order",
            },
            "stop_on_error": {
                "type": "boolean",
                "description": "Stop execution on first error (default: true)",
            },
        },
        "required": ["steps"],
    },
)
async def handle_execute_sequence(arguments: dict[str, Any]) -> list[TextContent]:
    steps = arguments["steps"]
    stop_on_error = arguments.get("stop_on_error", True)

    if not steps:
        raise ToolError("No steps provided")

    if len(steps) > 50:
        raise ToolError("Maximum 50 steps per sequence")

    results: list[dict[str, Any]] = []

    for i, step in enumerate(steps):
        tool_name = step.get("tool", "")
        tool_args = step.get("args", {})

        handler = registry.get_handler(tool_name)
        if not handler:
            entry = {
                "step": i + 1,
                "tool": tool_name,
                "error": f"Unknown tool: {tool_name}",
                "success": False,
            }
            results.append(entry)
            if stop_on_error:
                break
            continue

        try:
            result = await handler(tool_args)

            # Extract text content from MCP response
            text_parts = []
            has_image = False

            if isinstance(result, list):
                for item in result:
                    if isinstance(item, TextContent):
                        text_parts.append(item.text)
                    elif isinstance(item, ImageContent):
                        has_image = True
            elif isinstance(result, dict):
                text_parts.append(json.dumps(result, default=str))
            elif isinstance(result, TextContent):
                text_parts.append(result.text)
            else:
                text_parts.append(str(result))

            results.append(
                {
                    "step": i + 1,
                    "tool": tool_name,
                    "result": "\n".join(text_parts),
                    "has_image": has_image,
                    "success": True,
                }
            )

        except Exception as exc:
            results.append(
                {
                    "step": i + 1,
                    "tool": tool_name,
                    "error": str(exc),
                    "success": False,
                }
            )
            if stop_on_error:
                break

        # Inter-step delay (capped at 30 s to prevent abuse)
        delay_ms = step.get("delay_ms", 0)
        if delay_ms > 0:
            capped_ms = min(delay_ms, 30_000)
            await asyncio.sleep(capped_ms / 1000.0)

    completed = sum(1 for r in results if r.get("success"))
    failed = sum(1 for r in results if not r.get("success", True))

    return [
        TextContent(
            type="text",
            text=json.dumps(
                {
                    "total_steps": len(steps),
                    "completed": completed,
                    "failed": failed,
                    "results": results,
                },
                indent=2,
            ),
        )
    ]
