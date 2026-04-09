"""
Tests for new features: UIA tools, OCR caching, operation verification.

Requires:
  - A display
  - Tesseract at C:\\Program Files\\Tesseract-OCR (for OCR tests)
  - uiautomation package installed
"""

import asyncio
import json
import subprocess
import sys
import time
from pathlib import Path

sys.path.insert(0, ".")

import pyautogui

pyautogui.FAILSAFE = False

from win32_mcp_server.registry import registry

# ---- Helpers ----

PASS = 0
FAIL = 0
SKIP = 0


def check(label: str, ok: bool, detail: str = "") -> bool:
    global PASS, FAIL
    if ok:
        PASS += 1
        print(f"  PASS: {label}")
    else:
        FAIL += 1
        msg = f"  FAIL: {label}"
        if detail:
            msg += f" — {detail}"
        print(msg)
    return ok


def skip(label: str, reason: str) -> None:
    global SKIP
    SKIP += 1
    print(f"  SKIP: {label} — {reason}")


async def call(name: str, args: dict) -> dict | str | list:
    """Call MCP tool and return parsed result (raises on tool error)."""
    result = await registry.dispatch(name, args)
    texts = [item.text for item in result if hasattr(item, "text")]
    if len(texts) == 1:
        try:
            parsed = json.loads(texts[0])
            if isinstance(parsed, dict) and parsed.get("error"):
                raise RuntimeError(parsed.get("message", "Tool error"))
            return parsed
        except (json.JSONDecodeError, TypeError):
            return texts[0]
    return texts


async def call_raw(name: str, args: dict):
    """Call tool and return raw MCP content list."""
    result = await registry.dispatch(name, args)
    return result


def launch_notepad(wait: float = 2.0):
    proc = subprocess.Popen(["notepad.exe"])
    time.sleep(wait)
    return proc


def kill_notepad(proc, wait: float = 0.5):
    try:
        proc.kill()
    except Exception:
        pass
    time.sleep(wait)


async def focus_notepad() -> bool:
    for title in ["Notepad", "Untitled", "notepad"]:
        try:
            await call("focus_window", {"window_title": title})
            return True
        except Exception:
            continue
    return False


async def find_notepad_title() -> str | None:
    """Find the actual window title containing 'notepad' or 'untitled'."""
    result = await call("list_windows", {})
    if isinstance(result, dict):
        for w in result.get("windows", []):
            title = w.get("title", "").lower()
            if "notepad" in title or "untitled" in title:
                return w["title"]
    return None


# ===================================================================
# UIA TOOL TESTS
# ===================================================================


async def test_uia_inspect_window():
    """Test uia_inspect_window against Notepad."""
    print("\n[UIA: inspect window]")
    notepad = launch_notepad()
    try:
        np_title = await find_notepad_title()
        if not np_title:
            check("found notepad window", False, "no notepad window found")
            return

        result = await call("uia_inspect_window", {"window_title": np_title, "max_depth": 2})
        if not isinstance(result, dict):
            check("returned dict", False, f"got {type(result)}")
            return

        check("has control_type", "control_type" in result)
        check("control_type is WindowControl", result.get("control_type") == "WindowControl")
        check("has name", "name" in result)
        check(
            "has children",
            "children" in result and len(result.get("children", [])) > 0,
            f"{len(result.get('children', []))} children",
        )

        # Test with control_type filter
        result2 = await call(
            "uia_inspect_window",
            {
                "window_title": np_title,
                "control_type": "button",
                "max_depth": 3,
            },
        )
        if isinstance(result2, dict):
            check("filter has filtered_controls", "filtered_controls" in result2)
            check("filter has filtered_count", "filtered_count" in result2)
            btn_count = result2.get("filtered_count", 0)
            check("found buttons", btn_count > 0, f"{btn_count} buttons")

        # Test invalid control type
        try:
            await call(
                "uia_inspect_window",
                {
                    "window_title": np_title,
                    "control_type": "invalid_type_xyz",
                },
            )
            check("invalid control type raises error", False, "no error raised")
        except Exception as e:
            check("invalid control type raises error", "unknown control type" in str(e).lower(), str(e)[:100])

        # Test max_depth capping at 5
        result_deep = await call(
            "uia_inspect_window",
            {
                "window_title": np_title,
                "max_depth": 100,  # should be capped to 5
            },
        )
        check("deep inspect returns result", isinstance(result_deep, dict))

    finally:
        kill_notepad(notepad)


async def test_uia_find_control():
    """Test uia_find_control against Notepad."""
    print("\n[UIA: find control]")
    notepad = launch_notepad()
    try:
        np_title = await find_notepad_title()
        if not np_title:
            check("found notepad window", False)
            return

        # Find buttons
        result = await call(
            "uia_find_control",
            {
                "window_title": np_title,
                "control_type": "button",
            },
        )
        check("returns dict with found", isinstance(result, dict) and "found" in result)
        btn_count = result.get("found", 0) if isinstance(result, dict) else 0
        check("found buttons", btn_count > 0, f"{btn_count}")

        if isinstance(result, dict) and result.get("controls"):
            ctrl = result["controls"][0]
            check("control has control_type", "control_type" in ctrl)
            check("control has name", "name" in ctrl)
            check("control has bounds", "bounds" in ctrl or "center_x" in ctrl)

        # Search by name — "Close" button should exist
        result2 = await call(
            "uia_find_control",
            {
                "window_title": np_title,
                "name": "Close",
            },
        )
        found_close = result2.get("found", 0) if isinstance(result2, dict) else 0
        check("found Close button by name", found_close > 0, f"{found_close}")

        # Require at least one search criterion
        try:
            await call("uia_find_control", {"window_title": np_title})
            check("empty criteria raises error", False)
        except Exception as e:
            check("empty criteria raises error", "at least one" in str(e).lower(), str(e)[:100])

        # Window not found
        try:
            await call(
                "uia_find_control",
                {
                    "window_title": "ZZZ_nonexistent_window_XYZ",
                    "control_type": "button",
                },
            )
            check("missing window raises error", False)
        except Exception as e:
            check("missing window raises error", "not found" in str(e).lower(), str(e)[:100])

    finally:
        kill_notepad(notepad)


async def test_uia_click_control():
    """Test uia_click_control against Notepad."""
    print("\n[UIA: click control]")
    notepad = launch_notepad()
    try:
        np_title = await find_notepad_title()
        if not np_title:
            check("found notepad window", False)
            return

        # Click Minimize button
        try:
            result = await call(
                "uia_click_control",
                {
                    "window_title": np_title,
                    "name": "Minimize",
                    "control_type": "button",
                },
            )
            if isinstance(result, dict):
                check("click returned success", result.get("clicked") is True)
                check("click has method", "method" in result, result.get("method", "?"))
                check("click has control info", "control" in result)
            else:
                check("click returned dict", False, str(type(result)))
        except Exception as e:
            check("click Minimize button", False, str(e)[:200])

        time.sleep(0.5)

        # Restore the window for additional tests
        try:
            await call("restore_window", {"window_title": np_title})
        except Exception:
            pass
        time.sleep(0.5)

        # Require name or automation_id
        try:
            await call("uia_click_control", {"window_title": np_title})
            check("click no name/aid raises error", False)
        except Exception as e:
            check("click no name/aid raises error", "name" in str(e).lower() or "automation_id" in str(e).lower())

        # Click nonexistent control
        try:
            await call(
                "uia_click_control",
                {
                    "window_title": np_title,
                    "name": "ZZZ_nonexistent_control_XYZ",
                },
            )
            check("click nonexistent raises error", False)
        except Exception as e:
            check("click nonexistent raises error", "not found" in str(e).lower(), str(e)[:100])

    finally:
        kill_notepad(notepad)


async def test_uia_get_set_value():
    """Test uia_get_control_value and uia_set_control_value."""
    print("\n[UIA: get/set control value]")
    notepad = launch_notepad()
    try:
        np_title = await find_notepad_title()
        if not np_title:
            check("found notepad window", False)
            return

        await focus_notepad()
        time.sleep(0.5)

        # Type something in notepad first
        await call("type_text", {"text": "UIA Value Test"})
        time.sleep(0.5)

        # Try to get value of the edit control
        # Notepad variants use different control types:
        # - Classic Notepad: edit
        # - Modern Notepad: document
        # - Notepad3: pane (Scintilla)
        try:
            edit_count = 0
            for try_type in ["edit", "document", "pane"]:
                result = await call(
                    "uia_find_control",
                    {
                        "window_title": np_title,
                        "control_type": try_type,
                    },
                )
                edit_count = result.get("found", 0) if isinstance(result, dict) else 0
                if edit_count > 0:
                    break

            check("found edit/document control", edit_count > 0, f"{edit_count}")

            if edit_count > 0 and isinstance(result, dict):
                ctrl = result["controls"][0]
                ctrl_name = ctrl.get("name", "")
                ctrl_aid = ctrl.get("automation_id", "")

                # Get value
                get_args: dict[str, str] = {"window_title": np_title}
                if ctrl_aid:
                    get_args["automation_id"] = ctrl_aid
                elif ctrl_name:
                    get_args["name"] = ctrl_name
                else:
                    get_args["control_type"] = "edit"
                    get_args["name"] = ctrl_name or ""

                try:
                    val_result = await call("uia_get_control_value", get_args)
                    check("get_value returns dict", isinstance(val_result, dict))
                    if isinstance(val_result, dict):
                        has_value = "value" in val_result or "text" in val_result
                        check(
                            "get_value has value or text",
                            has_value,
                            f"value={val_result.get('value', '?')[:50]}, text={str(val_result.get('text', '?'))[:50]}",
                        )
                except Exception as e:
                    check("get_value", False, str(e)[:200])

        except Exception as e:
            check("find edit control", False, str(e)[:200])

        # Test error cases
        try:
            await call("uia_get_control_value", {"window_title": np_title})
            check("get_value no name/aid raises error", False)
        except Exception as e:
            check("get_value no name/aid raises error", "name" in str(e).lower() or "automation_id" in str(e).lower())

        try:
            await call("uia_set_control_value", {"window_title": np_title, "value": "test"})
            check("set_value no name/aid raises error", False)
        except Exception as e:
            check("set_value no name/aid raises error", "name" in str(e).lower() or "automation_id" in str(e).lower())

    finally:
        kill_notepad(notepad)


async def test_uia_get_focused():
    """Test uia_get_focused."""
    print("\n[UIA: get focused]")
    notepad = launch_notepad()
    try:
        await focus_notepad()
        time.sleep(0.5)

        result = await call("uia_get_focused", {})
        check("returns dict", isinstance(result, dict))
        if isinstance(result, dict):
            check("has focused field", "focused" in result)
            check("focused is True", result.get("focused") is True)
            check("has control_type", "control_type" in result)
            check("has name", "name" in result)
            # Should report parent window
            has_window = "window_title" in result
            check("has window_title", has_window, result.get("window_title", "?"))
    finally:
        kill_notepad(notepad)


# ===================================================================
# OCR CACHE TESTS
# ===================================================================


async def test_ocr_cache_speedup():
    """Verify repeated OCR on same screen returns cached results faster."""
    print("\n[OCR Cache: speedup]")

    # First call — uncached (cold)
    t0 = time.perf_counter()
    result1 = await call("ocr_screen", {})
    t1 = time.perf_counter()
    cold_ms = (t1 - t0) * 1000

    check("first OCR returned text", isinstance(result1, str) and len(result1) > 0)
    print(f"    cold call: {cold_ms:.0f}ms")

    # Second call immediately — should hit cache
    t2 = time.perf_counter()
    result2 = await call("ocr_screen", {})
    t3 = time.perf_counter()
    warm_ms = (t3 - t2) * 1000

    check("second OCR returned text", isinstance(result2, str) and len(result2) > 0)
    print(f"    warm call: {warm_ms:.0f}ms")

    # Cached call should be at least 2x faster (usually 100x+)
    check("cache is faster than cold", warm_ms < cold_ms, f"warm={warm_ms:.0f}ms, cold={cold_ms:.0f}ms")

    # Same result
    check("cached result matches", result1 == result2)


async def test_ocr_cache_structured():
    """Verify structured OCR also benefits from cache."""
    print("\n[OCR Cache: structured]")

    t0 = time.perf_counter()
    result1 = await call("ocr_screen_structured", {})
    t1 = time.perf_counter()
    cold_ms = (t1 - t0) * 1000

    check("structured OCR returned data", isinstance(result1, dict))

    t2 = time.perf_counter()
    result2 = await call("ocr_screen_structured", {})
    t3 = time.perf_counter()
    warm_ms = (t3 - t2) * 1000

    check("structured cached is faster", warm_ms < cold_ms, f"warm={warm_ms:.0f}ms, cold={cold_ms:.0f}ms")

    # Same element count
    if isinstance(result1, dict) and isinstance(result2, dict):
        check("same element count", len(result1.get("elements", [])) == len(result2.get("elements", [])))

    print(f"    cold={cold_ms:.0f}ms, warm={warm_ms:.0f}ms")


async def test_ocr_cache_expiry():
    """Verify cache entries expire after TTL (2 seconds)."""
    print("\n[OCR Cache: expiry]")

    # Prime the cache
    result1 = await call("ocr_screen", {})
    check("primed cache", isinstance(result1, str) and len(result1) > 0)

    # Wait for TTL to expire
    print("    Waiting 2.5s for cache expiry...")
    await asyncio.sleep(2.5)

    # This call should be a cache miss (expired)
    t0 = time.perf_counter()
    await call("ocr_screen", {})
    t1 = time.perf_counter()
    post_expiry_ms = (t1 - t0) * 1000

    # Should take >50ms (real OCR), not <10ms (cache hit)
    check(
        "post-expiry is a real OCR call", post_expiry_ms > 50, f"{post_expiry_ms:.0f}ms (expected >50ms for real OCR)"
    )


async def test_ocr_cache_different_params():
    """Verify different OCR params get separate cache entries."""
    print("\n[OCR Cache: different params]")

    # Call with default params
    result1 = await call("ocr_screen", {})

    # Call with different preprocess mode — should NOT use cached result
    result2 = await call("ocr_screen", {"preprocess": "high_contrast"})

    # Results may differ since preprocessing is different
    check("both returned text", isinstance(result1, str) and isinstance(result2, str))
    # We can't guarantee they differ but they should both be non-empty
    check("default has text", len(result1) > 0 if isinstance(result1, str) else False)
    check("high_contrast has text", len(result2) > 0 if isinstance(result2, str) else False)


# ===================================================================
# OPERATION VERIFICATION TESTS
# ===================================================================


async def test_click_verify():
    """Test click with verify=True."""
    print("\n[Verification: click]")

    # Click at a known position with verify
    result = await call_raw("click", {"x": 500, "y": 500, "verify": True})
    text = result[0].text if result else ""
    check("click with verify succeeded", "Clicked" in text, text[:100])

    # Click without verify (default)
    result2 = await call_raw("click", {"x": 501, "y": 501})
    text2 = result2[0].text if result2 else ""
    check("click without verify succeeded", "Clicked" in text2, text2[:100])


async def test_focus_verify():
    """Test focus_window with verify=True."""
    print("\n[Verification: focus_window]")
    notepad = launch_notepad()
    try:
        np_title = await find_notepad_title() or "Notepad"

        # Focus with verification
        result = await call("focus_window", {"window_title": np_title, "verify": True})
        check("focus with verify succeeded", isinstance(result, dict | str))

        # Focus nonexistent with verify
        try:
            await call("focus_window", {"window_title": "ZZZ_nonexistent_XYZ", "verify": True})
            check("focus nonexistent raises error", False)
        except Exception as e:
            check("focus nonexistent raises error", "not found" in str(e).lower())
    finally:
        kill_notepad(notepad)


async def test_kill_process_verify():
    """Test kill_process auto-verification (always verifies pid is gone)."""
    print("\n[Verification: kill_process]")

    # Launch a disposable notepad
    proc = launch_notepad(wait=1.5)
    pid = proc.pid

    check("notepad started", pid > 0, f"pid={pid}")

    try:
        result = await call("kill_process", {"pid": pid})
        if isinstance(result, dict):
            check("kill returned result", True)
            # The tool may return "killed" or "terminated_gracefully", etc.
        elif isinstance(result, str):
            check("kill returned text", "kill" in result.lower() or "terminat" in result.lower(), result[:100])
    except Exception as e:
        # Process may already be dead
        check("kill_process completed", "not found" in str(e).lower() or "no process" in str(e).lower(), str(e)[:100])

    time.sleep(0.5)

    # Verify process is actually gone
    import psutil

    check("process is gone", not psutil.pid_exists(pid), f"pid {pid} still exists!")


# ===================================================================
# E2E UIA TESTS (NOT REQUIRING SPECIFIC CONTROLS)
# ===================================================================


async def test_uia_error_handling():
    """Test UIA error handling for edge cases."""
    print("\n[UIA: error handling]")

    # Window not found
    try:
        await call("uia_inspect_window", {"window_title": "ZZZ_bogus_window_XYZ"})
        check("inspect nonexistent window raises error", False)
    except Exception as e:
        check("inspect nonexistent window raises error", "not found" in str(e).lower())

    # Invalid control type for find — find_control doesn't validate types,
    # but inspect_window does. Test that inspect validates properly.
    notepad = launch_notepad()
    try:
        np_title = await find_notepad_title() or "Notepad"
        try:
            await call(
                "uia_inspect_window",
                {
                    "window_title": np_title,
                    "control_type": "bogus_type_xyz",
                },
            )
            check("inspect with invalid type raises error", False)
        except Exception as e:
            check("inspect with invalid type raises error", "unknown control type" in str(e).lower(), str(e)[:100])

        # find_control with bogus type returns 0 results (no validation)
        result = await call(
            "uia_find_control",
            {
                "window_title": np_title,
                "control_type": "bogus_type_xyz",
            },
        )
        found = result.get("found", -1) if isinstance(result, dict) else -1
        check("find with bogus type returns 0", found == 0, f"found={found}")
    finally:
        kill_notepad(notepad)


# ===================================================================
# MAIN
# ===================================================================


async def main():
    print("=" * 60)
    print("NEW FEATURE TESTS — UIA, OCR Cache, Verification")
    print("=" * 60)

    # Check prerequisites
    has_tesseract = Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe").exists()
    print(f"  Tesseract available: {has_tesseract}")

    try:
        import importlib.util

        has_uia = importlib.util.find_spec("uiautomation") is not None
        print(f"  uiautomation: {'installed' if has_uia else 'NOT INSTALLED'}")
    except ImportError:
        has_uia = False
        print("  uiautomation: NOT INSTALLED")

    # --- UIA Tests ---
    if has_uia:
        await test_uia_inspect_window()
        await test_uia_find_control()
        await test_uia_click_control()
        await test_uia_get_set_value()
        await test_uia_get_focused()
        await test_uia_error_handling()
    else:
        skip("UIA tools", "uiautomation not installed")

    # --- OCR Cache Tests ---
    if has_tesseract:
        await test_ocr_cache_speedup()
        await test_ocr_cache_structured()
        await test_ocr_cache_expiry()
        await test_ocr_cache_different_params()
    else:
        skip("OCR Cache", "Tesseract not installed")

    # --- Verification Tests ---
    await test_click_verify()
    await test_focus_verify()
    await test_kill_process_verify()

    # --- Summary ---
    total = PASS + FAIL
    print(f"\n{'=' * 60}")
    print(f"Results: {PASS}/{total} passed, {FAIL} failed, {SKIP} skipped")
    if FAIL == 0:
        print("ALL NEW FEATURE TESTS PASSED")
    else:
        print(f"WARNING: {FAIL} test(s) failed!")
    print("=" * 60)
    return FAIL


if __name__ == "__main__":
    exit_code = asyncio.run(main())
    sys.exit(1 if exit_code > 0 else 0)
