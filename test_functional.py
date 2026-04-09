"""
Real functional tests for win32-mcp-server.

These tests actually:
  1. Open Notepad
  2. OCR the screen to find text
  3. Find and click UI elements by text
  4. Type text and verify it appears
  5. Clean up

Requires: Tesseract on PATH, a display, and permission to automate windows.
"""

import asyncio
import json
import subprocess
import sys
import time

# Make sure our package is importable
sys.path.insert(0, ".")

import pyautogui

pyautogui.FAILSAFE = False  # Don't abort on corner moves during testing

from win32_mcp_server.registry import registry

# ---- helpers ----

PASS = 0
FAIL = 0
SKIP = 0


def report(label: str, ok: bool, detail: str = ""):
    global PASS, FAIL
    tag = "PASS" if ok else "FAIL"
    if not ok:
        FAIL += 1
    else:
        PASS += 1
    msg = f"  {tag}: {label}"
    if detail:
        msg += f" — {detail}"
    print(msg)
    return ok


def skip(label: str, reason: str):
    global SKIP
    SKIP += 1
    print(f"  SKIP: {label} — {reason}")


async def call_tool(name: str, args: dict):
    """Call a registered MCP tool by name via dispatch and return parsed result."""
    # Use dispatch (not raw handler) to get consistent TextContent wrapping
    result = await registry.dispatch(name, args)
    # Extract text content
    texts = []
    for item in result:
        if hasattr(item, "text"):
            texts.append(item.text)
    if len(texts) == 1:
        try:
            parsed = json.loads(texts[0])
            # Check if dispatch returned an error
            if isinstance(parsed, dict) and parsed.get("error"):
                raise RuntimeError(parsed.get("message", "Tool error"))
            return parsed
        except (json.JSONDecodeError, TypeError):
            return texts[0]
    return texts


async def call_tool_raw(name: str, args: dict):
    """Call a registered MCP tool by name and return raw MCP content list."""
    handler = registry.get_handler(name)
    if handler is None:
        raise RuntimeError(f"No handler for {name}")
    result = await handler(args)
    if not isinstance(result, list):
        result = [result]
    return result


async def focus_notepad():
    """Focus Notepad window, trying common title variations."""
    for title in ["Notepad", "Untitled", "notepad"]:
        try:
            await call_tool("focus_window", {"window_title": title})
            return True
        except Exception:
            continue
    return False


def launch_notepad(wait: float = 2.0):
    """Launch notepad.exe and wait for the window to appear."""
    proc = subprocess.Popen(["notepad.exe"])
    time.sleep(wait)
    return proc


def kill_notepad(proc, wait: float = 0.5):
    """Kill a notepad process and wait for cleanup."""
    try:
        proc.kill()
    except Exception:
        pass
    time.sleep(wait)


# ---- Test Phases ----


async def test_ocr_reads_screen():
    """Phase 1: Can we OCR text from the current screen at all?"""
    print("\n[Phase 1: OCR reads screen]")
    try:
        result = await call_tool("ocr_screen", {})
        has_text = isinstance(result, str) and len(result.strip()) > 0
        report("OCR returned text", has_text, f"{len(result)} chars")
        return has_text
    except Exception as e:
        report("OCR screen", False, str(e))
        return False


async def test_structured_ocr():
    """Phase 2: Structured OCR returns word positions."""
    print("\n[Phase 2: Structured OCR]")
    try:
        result = await call_tool("ocr_screen_structured", {})
        if isinstance(result, dict):
            elements = result.get("elements", [])
        else:
            elements = []

        report("structured OCR has elements", len(elements) > 0, f"{len(elements)} elements")
        if elements:
            w = elements[0]
            has_coords = all(k in w for k in ("x", "y", "width", "height", "text"))
            report("elements have coordinates", has_coords, f"first: {w.get('text', '?')}")
        return len(elements) > 0
    except Exception as e:
        report("structured OCR", False, str(e))
        return False


async def test_open_notepad_and_find():
    """Phase 3: Open Notepad, find it, type text, find the text via OCR."""
    print("\n[Phase 3: Open Notepad, type, find text via OCR]")

    notepad = launch_notepad()

    try:
        # Verify we can find the notepad window
        result = await call_tool("list_windows", {})
        windows = result.get("windows", []) if isinstance(result, dict) else []
        notepad_found = any(
            "notepad" in w.get("title", "").lower() or "untitled" in w.get("title", "").lower() for w in windows
        )
        report("Notepad window found in list", notepad_found)

        if not notepad_found:
            print("  Can't proceed without Notepad window")
            return False

        # Focus notepad
        focused = await focus_notepad()
        report("focus_window on Notepad", focused)
        if not focused:
            return False

        time.sleep(0.5)

        # Type a unique string we can find via OCR (natural words for best OCR accuracy)
        test_string = "Testing automation carefully"
        await call_tool("type_text", {"text": test_string})
        report("type_text sent", True)
        time.sleep(1)

        # Now OCR the notepad window to find our text
        found_text = False
        for title in ["Notepad", "Untitled"]:
            try:
                ocr_text = await call_tool("ocr_window", {"window_title": title})
                if isinstance(ocr_text, str):
                    # Check for any of the typed words (OCR may not get everything)
                    found_text = any(w in ocr_text.lower() for w in ["testing", "automation", "carefully"])
                    report("OCR found typed text in Notepad", found_text, f"OCR returned: {str(ocr_text)[:200]}")
                    break
            except Exception:
                continue

        if not found_text:
            report("OCR found typed text in Notepad", False, "no matching title found")

        return found_text

    finally:
        kill_notepad(notepad)


async def test_find_text_on_screen():
    """Phase 4: Use find_text_on_screen to locate real UI text."""
    print("\n[Phase 4: find_text_on_screen]")

    notepad = launch_notepad()

    try:
        focused = await focus_notepad()
        if not focused:
            report("focus Notepad", False)
            return False

        time.sleep(0.3)
        unique_text = "Automation"
        await call_tool("type_text", {"text": unique_text})
        time.sleep(1)

        # Use find_text_on_screen (searches full screen by default)
        try:
            result = await call_tool("find_text_on_screen", {"text": unique_text})
            if isinstance(result, dict):
                matches = result.get("matches", [])
                match_count = result.get("match_count", 0)
            else:
                matches = []
                match_count = 0

            report("find_text_on_screen found our text", match_count > 0, f"{match_count} matches")

            if matches:
                m = matches[0]
                has_coords = "screen_center_x" in m and "screen_center_y" in m
                report(
                    "match has screen coordinates",
                    has_coords,
                    f"at ({m.get('screen_center_x')}, {m.get('screen_center_y')})",
                )
                report("match has confidence", "confidence" in m, f"confidence={m.get('confidence')}")
                report(
                    "match has score", "match_score" in m, f"score={m.get('match_score')}, type={m.get('match_type')}"
                )

            return match_count > 0
        except Exception as e:
            report("find_text_on_screen", False, str(e))
            return False
    finally:
        kill_notepad(notepad)


async def test_click_text():
    """Phase 5: Use click_text to actually click on text."""
    print("\n[Phase 5: click_text]")

    notepad = launch_notepad()

    try:
        focused = await focus_notepad()
        if not focused:
            report("focus Notepad", False)
            return False

        time.sleep(0.3)
        await call_tool("type_text", {"text": "Clickable Target"})
        time.sleep(1)

        # Click on the text we just typed (full screen search)
        try:
            result = await call_tool("click_text", {"text": "Clickable Target"})
            if isinstance(result, dict):
                # click_text returns clicked_text and clicked_at: {x, y}
                success = "clicked_text" in result
                at = result.get("clicked_at", {})
                click_x = at.get("x")
                click_y = at.get("y")
            else:
                success = False
                click_x = click_y = None

            report("click_text found and clicked", success, f"at ({click_x}, {click_y})")
            return success
        except Exception as e:
            report("click_text", False, str(e))
            return False
    finally:
        kill_notepad(notepad)


async def test_get_window_snapshot():
    """Phase 6: get_window_snapshot returns screenshot + OCR."""
    print("\n[Phase 6: get_window_snapshot]")

    notepad = launch_notepad()

    try:
        focused = await focus_notepad()
        if not focused:
            report("focus Notepad", False)
            return False

        time.sleep(0.3)
        await call_tool("type_text", {"text": "Snapshot Testing"})
        time.sleep(1)

        # get_window_snapshot should return image + text with positions
        raw_result = await call_tool_raw("get_window_snapshot", {"window_title": "Notepad"})

        has_image = any(hasattr(r, "data") and hasattr(r, "mimeType") for r in raw_result)
        report("snapshot has image", has_image)

        text_items = [r for r in raw_result if hasattr(r, "text")]
        has_text = len(text_items) > 0
        report("snapshot has text content", has_text)

        if has_text:
            try:
                data = json.loads(text_items[0].text)
                has_window = "window" in data
                report("snapshot has window info", has_window)
                if has_window:
                    report("snapshot has window title", "title" in data["window"], data["window"].get("title", "?"))

                ocr_elements = data.get("ocr_elements", [])
                report("snapshot has OCR elements", len(ocr_elements) > 0, f"{len(ocr_elements)} elements")
                if ocr_elements:
                    all_text = " ".join(e.get("text", "") for e in ocr_elements)
                    # Check for any of our typed words in the OCR elements
                    found_our_text = any(w in all_text.lower() for w in ["snapshot", "testing"])
                    report("snapshot OCR found our text", found_our_text, f"OCR text: {all_text[:200]}")
                    # Check that screen coordinates are populated
                    has_screen_coords = all("screen_x" in e and "screen_y" in e for e in ocr_elements[:5])
                    report("OCR elements have screen coordinates", has_screen_coords)
            except json.JSONDecodeError:
                report("snapshot text is valid JSON", False)

        return has_image and has_text
    except Exception as e:
        report("get_window_snapshot", False, str(e))
        return False
    finally:
        kill_notepad(notepad)


async def test_fill_field():
    """Phase 7: fill_field targets a label and types into it."""
    print("\n[Phase 7: fill_field]")

    notepad = launch_notepad()

    try:
        focused = await focus_notepad()
        if not focused:
            report("focus Notepad", False)
            return False

        time.sleep(0.3)

        # fill_field with a nonexistent label should raise ToolError
        try:
            result = await call_tool(
                "fill_field", {"label_text": "NONEXISTENT_LABEL_XYZ", "value": "test_value", "window_title": "Notepad"}
            )
            # If it didn't raise, check the result
            if isinstance(result, dict):
                is_error = not result.get("success", True)
                report("fill_field returns error for missing label", is_error, str(result)[:200])
            else:
                report("fill_field returned result (unexpected)", False, str(result)[:200])
        except Exception as e:
            # ToolError for missing label is expected
            is_expected_error = "not found" in str(e).lower() or "no match" in str(e).lower()
            report("fill_field raises error for missing label", is_expected_error, str(e)[:200])

        return True
    finally:
        kill_notepad(notepad)


async def test_capture_window():
    """Phase 8: capture_window returns actual window screenshot."""
    print("\n[Phase 8: capture_window]")

    notepad = launch_notepad()

    try:
        raw_result = await call_tool_raw("capture_window", {"window_title": "Notepad"})

        has_image = any(hasattr(r, "data") and hasattr(r, "mimeType") for r in raw_result)
        text_items = [r for r in raw_result if hasattr(r, "text")]

        report("capture_window returns image", has_image)
        report("capture_window returns text info", len(text_items) > 0)

        if text_items:
            try:
                data = json.loads(text_items[0].text)
                report("has capture_x", "capture_x" in data)
                report("has capture_y", "capture_y" in data)
                report("has width/height", "capture_width" in data and "capture_height" in data)
                report("has title", "title" in data, f"title={data.get('title', '?')}")
            except json.JSONDecodeError:
                report("capture info is valid JSON", False)

        return has_image
    except Exception as e:
        report("capture_window", False, str(e))
        return False
    finally:
        kill_notepad(notepad)


async def main():
    print("=" * 60)
    print("FUNCTIONAL TESTS — Real Screen Automation")
    print("=" * 60)

    # Phase 1-2: Basic OCR
    ocr_ok = await test_ocr_reads_screen()
    if not ocr_ok:
        print("\nOCR is not working. Cannot proceed with functional tests.")
        print(f"\nResults: {PASS}/{PASS + FAIL} passed, {FAIL} failed")
        return

    await test_structured_ocr()

    # Phase 3-8: Window automation
    await test_open_notepad_and_find()
    await test_find_text_on_screen()
    await test_click_text()
    await test_get_window_snapshot()
    await test_fill_field()
    await test_capture_window()

    print("\n" + "=" * 60)
    total = PASS + FAIL
    print(f"Results: {PASS}/{total} passed, {FAIL} failed, {SKIP} skipped")
    if FAIL == 0:
        print("ALL FUNCTIONAL TESTS PASSED")
    else:
        print("SOME TESTS FAILED — see above for details")
    print("=" * 60)


if __name__ == "__main__":
    asyncio.run(main())
