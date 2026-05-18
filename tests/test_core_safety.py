from __future__ import annotations

import asyncio
import json
import sys
from typing import Any

import pytest

import win32_mcp_server
from win32_mcp_server import __version__
from win32_mcp_server.config import SecurityConfig, ServerConfig, config
from win32_mcp_server.registry import ToolRegistry, _validate_schema, registry
from win32_mcp_server.server import main
from win32_mcp_server.tools import keyboard, ocr, process, smart, uia
from win32_mcp_server.utils import imaging
from win32_mcp_server.utils.args import get_int, get_poll_interval
from win32_mcp_server.utils.errors import ToolError
from win32_mcp_server.utils.execution import estimate_timeout_seconds
from win32_mcp_server.utils.security import (
    enforce_command_allowed,
    enforce_tool_allowed,
    redact_arguments,
    should_dry_run,
)


def test_schema_validation_rejects_missing_required() -> None:
    schema = {
        "type": "object",
        "properties": {"text": {"type": "string"}},
        "required": ["text"],
    }

    with pytest.raises(ToolError, match="missing required"):
        _validate_schema({}, schema, "demo")


def test_schema_validation_rejects_ranges_and_enums() -> None:
    schema = {
        "type": "object",
        "properties": {
            "button": {"type": "string", "enum": ["left", "right"]},
            "clicks": {"type": "number", "minimum": 1, "maximum": 10},
        },
    }

    with pytest.raises(ToolError, match="one of"):
        _validate_schema({"button": "side"}, schema, "demo")

    with pytest.raises(ToolError, match=">="):
        _validate_schema({"clicks": 0}, schema, "demo")


def test_security_profile_blocks_mutation(monkeypatch: pytest.MonkeyPatch) -> None:
    original = config.security
    monkeypatch.setattr(config, "security", SecurityConfig(profile="read_only"))
    try:
        enforce_tool_allowed("capture_screen")
        with pytest.raises(ToolError, match="read_only"):
            enforce_tool_allowed("click")
    finally:
        monkeypatch.setattr(config, "security", original)


def test_secure_defaults() -> None:
    cfg = ServerConfig()

    assert cfg.security.profile == "interactive"
    assert cfg.automation.pyautogui_failsafe is True
    assert cfg.security.redact_sensitive_output is True


def test_env_can_restore_unrestricted_profile(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("WIN32_MCP_SECURITY_PROFILE", "unrestricted")
    monkeypatch.setenv("WIN32_MCP_PYAUTOGUI_FAILSAFE", "false")

    cfg = ServerConfig.from_env()

    assert cfg.security.profile == "unrestricted"
    assert cfg.automation.pyautogui_failsafe is False


def test_confirmation_token_blocks_high_risk(monkeypatch: pytest.MonkeyPatch) -> None:
    original = config.security
    monkeypatch.setattr(config, "security", SecurityConfig(profile="unrestricted", confirmation_token="ok"))
    try:
        with pytest.raises(ToolError, match="confirmation_token"):
            enforce_tool_allowed("kill_process", {"pid": 123})
        enforce_tool_allowed("kill_process", {"pid": 123, "confirmation_token": "ok"})
    finally:
        monkeypatch.setattr(config, "security", original)


def test_dry_run_detects_mutating_calls(monkeypatch: pytest.MonkeyPatch) -> None:
    original = config.security
    monkeypatch.setattr(config, "security", SecurityConfig(dry_run=True))
    try:
        assert should_dry_run("click", {}) is True
        assert should_dry_run("capture_screen", {}) is False
    finally:
        monkeypatch.setattr(config, "security", original)


def test_dry_run_rejects_non_boolean_override() -> None:
    with pytest.raises(ToolError, match="dry_run"):
        should_dry_run("click", {"dry_run": "false"})


def test_command_allowlist(monkeypatch: pytest.MonkeyPatch) -> None:
    original = config.security
    monkeypatch.setattr(config, "security", SecurityConfig(allowed_commands={"python.exe"}))
    try:
        enforce_command_allowed(r"C:\Python\python.exe")
        with pytest.raises(ToolError, match="not in WIN32_MCP_ALLOWED_COMMANDS"):
            enforce_command_allowed("powershell.exe")
    finally:
        monkeypatch.setattr(config, "security", original)


def test_sensitive_arguments_redacted() -> None:
    redacted = redact_arguments(
        "type_text",
        {"text": "secret-value", "metadata": {"api_key": "abc123"}, "safe": "visible"},
    )

    assert redacted["text"] == "[redacted 12 chars]"
    assert redacted["metadata"]["api_key"] == "[redacted 6 chars]"
    assert redacted["safe"] == "visible"


def test_typed_args_bounds() -> None:
    assert get_int({"x": "7"}, "x", required=True, min_value=1) == 7
    with pytest.raises(ToolError, match=">="):
        get_int({"x": 0}, "x", required=True, min_value=1)
    with pytest.raises(ToolError, match="<="):
        get_poll_interval({"poll_interval": config.limits.max_poll_interval_seconds + 1})


def test_keyboard_rejects_invalid_keys() -> None:
    assert keyboard._parse_key_combo("ctrl+shift+s") == ["ctrl", "shift", "s"]
    assert keyboard._parse_key_combo("control+escape") == ["ctrl", "esc"]
    with pytest.raises(ToolError, match="Invalid key name"):
        keyboard._parse_key_combo("ctrl+definitely-not-a-key")


def test_tesseract_discovery_uses_common_windows_path(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(config.ocr, "tesseract_path", "")
    monkeypatch.setenv("PATH", "")
    monkeypatch.delenv("TESSERACT_CMD", raising=False)
    monkeypatch.setattr(imaging.shutil, "which", lambda _name: None)
    monkeypatch.setattr(
        imaging,
        "_TESSERACT_DEFAULT_PATHS",
        (imaging.Path("C:/Program Files/Tesseract-OCR/tesseract.exe"),),
    )
    monkeypatch.setattr(imaging.Path, "is_file", lambda self: str(self).endswith("tesseract.exe"))

    assert imaging.discover_tesseract_cmd().endswith("Tesseract-OCR\\tesseract.exe")


def test_ocr_cache_returns_deep_copies() -> None:
    key = ("hash", "eng", "auto", "structured:60")
    value = [{"text": "OK", "x": 1}]

    ocr._cache_put(key, value)
    first = ocr._cache_get(key)
    assert isinstance(first, list)
    first[0]["screen_x"] = 99

    second = ocr._cache_get(key)
    assert isinstance(second, list)
    assert "screen_x" not in second[0]


def test_uia_control_type_mapping_handles_compound_names() -> None:
    assert uia._expected_control_type("radiobutton") == "RadioButtonControl"
    assert uia._expected_control_type("listitem") == "ListItemControl"
    assert uia._expected_control_type("datagrid") == "DataGridControl"


class DummyControl:
    Name = "Save As"
    AutomationId = "saveButton"
    ControlTypeName = "ButtonControl"


def test_uia_match_uses_type_map() -> None:
    ctrl = DummyControl()

    assert uia._matches_control(ctrl, "save", "", "button", exact_name=False)
    assert not uia._matches_control(ctrl, "save", "", "checkbox", exact_name=False)


def test_label_field_score_prefers_expected_direction() -> None:
    label = (10, 10, 40, 20)
    right_field = (80, 10, 120, 25)
    below_field = (10, 60, 120, 25)

    assert uia._label_field_score(label, right_field, "right") is not None
    assert uia._label_field_score(label, below_field, "right") is None
    assert uia._label_field_score(label, below_field, "below") is not None


def test_click_text_rejects_zero_occurrence() -> None:
    with pytest.raises(ToolError, match="occurrence"):
        asyncio.run(smart.handle_click_text({"text": "OK", "occurrence": 0}))


def test_run_process_bounded_truncates_output(monkeypatch: pytest.MonkeyPatch) -> None:
    original = config.limits.max_subprocess_output_bytes
    monkeypatch.setattr(config.limits, "max_subprocess_output_bytes", 8)
    try:
        result: dict[str, Any] = process._run_process_bounded(
            [sys.executable, "-c", "print('abcdefghijklmnop')"],
            None,
            10,
        )
    finally:
        monkeypatch.setattr(config.limits, "max_subprocess_output_bytes", original)

    assert result["returncode"] == 0
    assert result["stdout"].startswith("abcdefgh")
    assert result["stdout_truncated"] is True


def test_estimated_sequence_timeout_includes_delays() -> None:
    timeout = estimate_timeout_seconds(
        "execute_sequence",
        {"steps": [{"tool": "click", "delay_ms": 1000}, {"tool": "press_key"}]},
    )

    assert timeout > config.default_timeout


def test_registry_result_envelope_is_opt_in(monkeypatch: pytest.MonkeyPatch) -> None:
    reg = ToolRegistry()

    @reg.register("demo", "Demo", {"type": "object", "properties": {}})
    async def _handler(arguments: dict[str, Any]) -> dict[str, Any]:
        return {"ok": True, "args": arguments}

    monkeypatch.setattr(config, "result_envelope", True)
    content = asyncio.run(reg.dispatch("demo", {"x": 1}))
    payload = json.loads(content[0].text)

    assert payload == {"success": True, "tool": "demo", "data": {"ok": True, "args": {"x": 1}}}


def test_registry_serializes_non_content_lists() -> None:
    reg = ToolRegistry()

    @reg.register("demo_list", "Demo list", {"type": "object", "properties": {}})
    async def _handler(_arguments: dict[str, Any]) -> list[dict[str, int]]:
        return [{"value": 1}]

    content = asyncio.run(reg.dispatch("demo_list", {}))
    payload = json.loads(content[0].text)

    assert payload == {"items": [{"value": 1}]}


def test_dispatch_dry_run_does_not_call_mutating_handler() -> None:
    content = asyncio.run(registry.dispatch("click", {"x": 1, "y": 1, "dry_run": True}))
    payload = json.loads(content[0].text)

    assert payload["dry_run"] is True
    assert payload["tool"] == "click"
    assert payload["would_call"]["dry_run"] is True


def test_registered_tools_have_well_formed_schemas() -> None:
    from win32_mcp_server import server

    tools = server.registry.get_tools()
    names = [tool.name for tool in tools]

    assert len(tools) >= 50
    assert len(names) == len(set(names))

    for tool in tools:
        schema = tool.inputSchema
        assert isinstance(schema, dict), tool.name
        assert schema.get("type") == "object", tool.name
        properties = schema.get("properties", {})
        assert isinstance(properties, dict), tool.name
        for required in schema.get("required", []):
            assert required in properties, f"{tool.name}: required field {required!r} missing schema"


def test_common_security_args_are_advertised() -> None:
    click_schema = registry.get_schema("click")
    start_schema = registry.get_schema("start_process")

    assert click_schema is not None
    assert start_schema is not None
    assert click_schema["properties"]["dry_run"]["type"] == "boolean"
    assert start_schema["properties"]["dry_run"]["type"] == "boolean"
    assert start_schema["properties"]["confirmation_token"]["type"] == "string"


def test_cli_version_exits_without_server(capsys: pytest.CaptureFixture[str]) -> None:
    main(["--version"])

    assert capsys.readouterr().out.strip() == __version__


def test_package_main_version_is_lazy(capsys: pytest.CaptureFixture[str]) -> None:
    win32_mcp_server.main(["--version"])

    assert capsys.readouterr().out.strip() == __version__
