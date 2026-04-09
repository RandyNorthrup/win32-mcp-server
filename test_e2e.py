"""End-to-end functional tests for win32-mcp-server."""

import asyncio
import json

from win32_mcp_server.server import registry


async def test():
    passed = 0
    failed = 0

    def check(name, condition, detail=""):
        nonlocal passed, failed
        if condition:
            print(f"  PASS: {name}")
            passed += 1
        else:
            print(f"  FAIL: {name} — {detail}")
            failed += 1

    # === Test health_check ===
    print("\n[health_check]")
    result = await registry.dispatch("health_check", {})
    data = json.loads(result[0].text)
    check("returns data", "server_version" in data)
    check("version is 2.5.1", data.get("server_version") == "2.5.1", data.get("server_version"))
    check("53 tools registered", data.get("registered_tools") == 53, data.get("registered_tools"))
    check("has dependencies", "dependencies" in data)

    # === Test list_monitors ===
    print("\n[list_monitors]")
    result = await registry.dispatch("list_monitors", {})
    data = json.loads(result[0].text)
    check("has monitors", data.get("monitor_count", 0) > 0)
    check("has dpi", "dpi" in data)
    check("has scaling", "scaling" in data)

    # === Test mouse_position ===
    print("\n[mouse_position]")
    result = await registry.dispatch("mouse_position", {})
    data = json.loads(result[0].text)
    check("returns x", "x" in data)
    check("returns y", "y" in data)

    # === Test list_processes ===
    print("\n[list_processes]")
    result = await registry.dispatch("list_processes", {"limit": 5, "sort_by": "memory"})
    data = json.loads(result[0].text)
    check("has total_count", data.get("total_count", 0) > 0)
    check("respects limit", len(data.get("processes", [])) <= 5)
    check(
        "has process fields",
        all(k in data["processes"][0] for k in ["pid", "name", "memory_mb"]) if data.get("processes") else False,
    )

    # === Test list_windows ===
    print("\n[list_windows]")
    result = await registry.dispatch("list_windows", {})
    data = json.loads(result[0].text)
    check("has count", "count" in data)
    check("has windows", len(data.get("windows", [])) > 0)
    if data.get("windows"):
        w = data["windows"][0]
        check("window has title", "title" in w)
        check("window has hwnd", "hwnd" in w)
        check("window has is_responding", "is_responding" in w)

    # === Test clipboard round-trip ===
    print("\n[clipboard]")
    await registry.dispatch("clipboard_copy", {"text": "mcp_test_value_12345"})
    result = await registry.dispatch("clipboard_paste", {})
    check("clipboard round-trip", "mcp_test_value_12345" in result[0].text)

    # === Test capture_screen ===
    print("\n[capture_screen]")
    result = await registry.dispatch(
        "capture_screen",
        {
            "format": "jpeg",
            "quality": 30,
            "scale": 0.25,
        },
    )
    check("returns 2 items (text + image)", len(result) == 2)
    meta = json.loads(result[0].text)
    check("has screen_size", "screen_size" in meta)
    check("has file_size_kb", "file_size_kb" in meta)
    check("image content present", result[1].type == "image")

    # === Test capture_window ===
    print("\n[capture_window]")
    result = await registry.dispatch(
        "capture_window",
        {
            "window_title": "Visual Studio Code",
            "format": "jpeg",
            "quality": 30,
            "scale": 0.25,
        },
    )
    if result and not json.loads(result[0].text).get("error"):
        meta = json.loads(result[0].text)
        check("has capture_x", "capture_x" in meta)
        check("has capture_y", "capture_y" in meta)
        check("has title", "title" in meta)
    else:
        print("  SKIP: VS Code window not found")

    # === Test get_pixel_color ===
    print("\n[get_pixel_color]")
    result = await registry.dispatch("get_pixel_color", {"x": 100, "y": 100})
    data = json.loads(result[0].text)
    check("has rgb", all(k in data for k in ["r", "g", "b"]))
    check("has hex", "hex" in data)

    # === Test type_text (auto method for ASCII) ===
    print("\n[type_text validation]")
    result = await registry.dispatch("type_text", {"text": ""})
    data = json.loads(result[0].text)
    check("empty text returns error", data.get("error") is True)

    # === Test press_key validation ===
    print("\n[press_key validation]")
    result = await registry.dispatch("press_key", {"keys": ""})
    data = json.loads(result[0].text)
    check("empty keys returns error", data.get("error") is True)

    # === Test hotkey validation ===
    print("\n[hotkey validation]")
    result = await registry.dispatch("hotkey", {"keys": []})
    data = json.loads(result[0].text)
    check("empty key list returns error", data.get("error") is True)

    # === Test scroll partial position validation ===
    print("\n[scroll validation]")
    result = await registry.dispatch("scroll", {"amount": 3, "x": 100})
    data = json.loads(result[0].text)
    check("partial position (x only) returns error", data.get("error") is True)

    result = await registry.dispatch("scroll", {"amount": 3, "y": 100})
    data = json.loads(result[0].text)
    check("partial position (y only) returns error", data.get("error") is True)

    result = await registry.dispatch("scroll_horizontal", {"amount": 3, "x": 100})
    data = json.loads(result[0].text)
    check("h-scroll partial position returns error", data.get("error") is True)

    # === Test unknown tool error ===
    print("\n[error handling]")
    result = await registry.dispatch("nonexistent_tool", {})
    data = json.loads(result[0].text)
    check("unknown tool returns error", data.get("error") is True)
    check("includes available_tools", "available_tools" in data)

    # === Test coordinate validation ===
    print("\n[coordinate validation]")
    result = await registry.dispatch("click", {"x": 99999, "y": 99999})
    data = json.loads(result[0].text)
    check("out-of-bounds click returns error", data.get("error") is True)
    check("error has suggestion", "suggestion" in data)

    # === Test window not found ===
    print("\n[window not found]")
    result = await registry.dispatch(
        "focus_window",
        {
            "window_title": "ZZZ_nonexistent_window_title_XYZ",
        },
    )
    data = json.loads(result[0].text)
    check("missing window returns error", data.get("error") is True)
    check("includes suggestion", "suggestion" in data)

    # === Test execute_sequence ===
    print("\n[execute_sequence]")
    result = await registry.dispatch(
        "execute_sequence",
        {
            "steps": [
                {"tool": "mouse_position", "args": {}},
                {"tool": "list_monitors", "args": {}},
            ],
        },
    )
    data = json.loads(result[0].text)
    check("completes 2 steps", data.get("completed") == 2)
    check("no failures", data.get("failed") == 0)

    # Test execute_sequence with bad tool
    result = await registry.dispatch(
        "execute_sequence",
        {
            "steps": [
                {"tool": "bad_tool", "args": {}},
                {"tool": "mouse_position", "args": {}},
            ],
            "stop_on_error": True,
        },
    )
    data = json.loads(result[0].text)
    check("stops on error", data.get("completed") == 0)
    check("1 failure", data.get("failed") == 1)

    # Test empty steps
    result = await registry.dispatch("execute_sequence", {"steps": []})
    data = json.loads(result[0].text)
    check("empty steps returns error", data.get("error") is True)

    # Test > 50 steps
    result = await registry.dispatch(
        "execute_sequence",
        {
            "steps": [{"tool": "mouse_position"}] * 51,
        },
    )
    data = json.loads(result[0].text)
    check("51 steps returns error", data.get("error") is True)

    # === Test OCR (if Tesseract available) ===
    print("\n[OCR tools]")
    hc = await registry.dispatch("health_check", {})
    hc_data = json.loads(hc[0].text)
    if hc_data.get("tesseract", {}).get("installed"):
        result = await registry.dispatch("ocr_screen", {})
        check("ocr_screen returns text", result[0].type == "text")

        result = await registry.dispatch("ocr_screen_structured", {})
        data = json.loads(result[0].text)
        check("structured OCR has elements", "elements" in data)
    else:
        print("  SKIP: Tesseract not installed")

    # === Test compare_screenshots size limit ===
    print("\n[compare_screenshots validation]")
    huge_b64 = "A" * 70_000_000  # > 67MB limit
    result = await registry.dispatch(
        "compare_screenshots",
        {
            "reference_image": huge_b64,
        },
    )
    data = json.loads(result[0].text)
    check("oversized reference returns error", data.get("error") is True)

    # === Summary ===
    total = passed + failed
    print(f"\n{'=' * 50}")
    print(f"Results: {passed}/{total} passed, {failed} failed")
    if failed == 0:
        print("ALL TESTS PASSED")
    else:
        print(f"WARNING: {failed} test(s) failed!")
    return failed


if __name__ == "__main__":
    exit_code = asyncio.run(test())
    exit(1 if exit_code > 0 else 0)
