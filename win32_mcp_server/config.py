"""
Server configuration with sensible defaults and environment overrides.

Environment variables use the WIN32_MCP_ prefix, for example:
  WIN32_MCP_SECURITY_PROFILE=interactive
  WIN32_MCP_OCR_LANGUAGE=eng+fra
  WIN32_MCP_CAPTURE_FORMAT=jpeg
"""

import os
from dataclasses import dataclass, field
from typing import Literal

PreprocessMode = Literal["auto", "light_bg", "dark_bg", "high_contrast", "none"]
CaptureFormat = Literal["png", "jpeg", "webp"]
SecurityProfile = Literal["unrestricted", "interactive", "read_only"]

VALID_PREPROCESS_MODES = {"auto", "light_bg", "dark_bg", "high_contrast", "none"}
VALID_CAPTURE_FORMATS = {"png", "jpeg", "webp"}
VALID_SECURITY_PROFILES = {"unrestricted", "interactive", "read_only"}


def _env_str(name: str, default: str) -> str:
    value = os.getenv(name)
    if value is None or not value.strip():
        return default
    return value.strip()


def _env_bool(name: str, default: bool) -> bool:
    value = os.getenv(name)
    if value is None:
        return default
    return value.strip().lower() in {"1", "true", "yes", "on"}


def _env_int(
    name: str,
    default: int,
    min_value: int | None = None,
    max_value: int | None = None,
) -> int:
    value = os.getenv(name)
    if value is None:
        return default
    try:
        parsed = int(value)
    except ValueError:
        return default
    if min_value is not None:
        parsed = max(min_value, parsed)
    if max_value is not None:
        parsed = min(max_value, parsed)
    return parsed


def _env_float(
    name: str,
    default: float,
    min_value: float | None = None,
    max_value: float | None = None,
) -> float:
    value = os.getenv(name)
    if value is None:
        return default
    try:
        parsed = float(value)
    except ValueError:
        return default
    if min_value is not None:
        parsed = max(min_value, parsed)
    if max_value is not None:
        parsed = min(max_value, parsed)
    return parsed


def _env_choice(name: str, default: str, choices: set[str]) -> str:
    value = os.getenv(name)
    if value is None:
        return default
    parsed = value.strip().lower()
    return parsed if parsed in choices else default


def _env_csv(name: str) -> set[str]:
    value = os.getenv(name)
    if value is None:
        return set()
    return {item.strip() for item in value.split(",") if item.strip()}


@dataclass
class OCRConfig:
    """OCR engine configuration."""

    lang: str = "eng"
    tesseract_path: str = ""
    confidence_threshold: int = 60
    preprocess_mode: PreprocessMode = "auto"
    scale_small_images: bool = True
    min_dimension_for_scaling: int = 500
    upscale_factor: int = 2


@dataclass
class CaptureConfig:
    """Screenshot capture configuration."""

    default_format: CaptureFormat = "png"
    default_quality: int = 85
    default_scale: float = 1.0


@dataclass
class AutomationConfig:
    """Mouse/keyboard automation configuration."""

    click_delay: float = 0.05
    type_interval: float = 0.01
    drag_duration: float = 0.5
    move_duration: float = 0.25
    pyautogui_failsafe: bool = True


@dataclass
class RuntimeLimits:
    """Bounds that keep tool calls predictable."""

    max_text_chars: int = 20_000
    max_clipboard_chars: int = 20_000
    max_command_chars: int = 1_000
    max_args: int = 128
    max_arg_chars: int = 4_000
    max_timeout_seconds: float = 120.0
    max_poll_interval_seconds: float = 10.0
    max_subprocess_output_bytes: int = 1_048_576
    max_reference_image_bytes: int = 52_428_800
    max_sequence_steps: int = 50
    max_tool_runtime_seconds: float = 180.0


@dataclass
class SecurityConfig:
    """Security controls for local OS automation."""

    profile: SecurityProfile = "interactive"
    allowed_tools: set[str] = field(default_factory=set)
    blocked_tools: set[str] = field(default_factory=set)
    allowed_commands: set[str] = field(default_factory=set)
    blocked_commands: set[str] = field(default_factory=set)
    dry_run: bool = False
    confirmation_token: str = ""
    redact_sensitive_output: bool = True


@dataclass
class ServerConfig:
    """Top-level server configuration."""

    ocr: OCRConfig = field(default_factory=OCRConfig)
    capture: CaptureConfig = field(default_factory=CaptureConfig)
    automation: AutomationConfig = field(default_factory=AutomationConfig)
    limits: RuntimeLimits = field(default_factory=RuntimeLimits)
    security: SecurityConfig = field(default_factory=SecurityConfig)
    validate_coordinates: bool = True
    default_timeout: float = 10.0
    window_retry_attempts: int = 3
    window_retry_delay: float = 0.3
    min_operation_interval: float = 0.05
    fuzzy_match_threshold: int = 65
    result_envelope: bool = False
    debug: bool = False

    @classmethod
    def from_env(cls) -> "ServerConfig":
        """Create config, applying WIN32_MCP_* environment overrides."""
        cfg = cls()

        cfg.ocr.lang = _env_str("WIN32_MCP_OCR_LANGUAGE", _env_str("WIN32_MCP_OCR_LANG", cfg.ocr.lang))
        cfg.ocr.tesseract_path = _env_str("WIN32_MCP_TESSERACT_PATH", cfg.ocr.tesseract_path)
        cfg.ocr.confidence_threshold = _env_int(
            "WIN32_MCP_OCR_CONFIDENCE_THRESHOLD",
            cfg.ocr.confidence_threshold,
            0,
            100,
        )
        cfg.ocr.preprocess_mode = _env_choice(
            "WIN32_MCP_OCR_PREPROCESS",
            cfg.ocr.preprocess_mode,
            VALID_PREPROCESS_MODES,
        )  # type: ignore[assignment]
        cfg.ocr.scale_small_images = _env_bool("WIN32_MCP_OCR_SCALE_SMALL_IMAGES", cfg.ocr.scale_small_images)
        cfg.ocr.min_dimension_for_scaling = _env_int(
            "WIN32_MCP_OCR_MIN_DIMENSION",
            cfg.ocr.min_dimension_for_scaling,
            50,
            4_000,
        )
        cfg.ocr.upscale_factor = _env_int("WIN32_MCP_OCR_UPSCALE_FACTOR", cfg.ocr.upscale_factor, 1, 8)

        cfg.capture.default_format = _env_choice(
            "WIN32_MCP_CAPTURE_FORMAT",
            cfg.capture.default_format,
            VALID_CAPTURE_FORMATS,
        )  # type: ignore[assignment]
        cfg.capture.default_quality = _env_int("WIN32_MCP_CAPTURE_QUALITY", cfg.capture.default_quality, 1, 100)
        cfg.capture.default_scale = _env_float("WIN32_MCP_CAPTURE_SCALE", cfg.capture.default_scale, 0.1, 1.0)

        cfg.automation.click_delay = _env_float("WIN32_MCP_CLICK_DELAY", cfg.automation.click_delay, 0.0, 5.0)
        cfg.automation.type_interval = _env_float("WIN32_MCP_TYPE_INTERVAL", cfg.automation.type_interval, 0.0, 1.0)
        cfg.automation.drag_duration = _env_float("WIN32_MCP_DRAG_DURATION", cfg.automation.drag_duration, 0.0, 30.0)
        cfg.automation.move_duration = _env_float("WIN32_MCP_MOVE_DURATION", cfg.automation.move_duration, 0.0, 30.0)
        cfg.automation.pyautogui_failsafe = _env_bool(
            "WIN32_MCP_PYAUTOGUI_FAILSAFE",
            cfg.automation.pyautogui_failsafe,
        )

        cfg.validate_coordinates = _env_bool("WIN32_MCP_COORDINATE_VALIDATION", cfg.validate_coordinates)
        cfg.limits.max_timeout_seconds = _env_float(
            "WIN32_MCP_MAX_TIMEOUT_SECONDS",
            cfg.limits.max_timeout_seconds,
            1.0,
        )
        cfg.default_timeout = _env_float(
            "WIN32_MCP_DEFAULT_TIMEOUT",
            cfg.default_timeout,
            0.1,
            cfg.limits.max_timeout_seconds,
        )
        cfg.window_retry_attempts = _env_int("WIN32_MCP_WINDOW_RETRY_ATTEMPTS", cfg.window_retry_attempts, 1, 10)
        cfg.window_retry_delay = _env_float("WIN32_MCP_WINDOW_RETRY_DELAY", cfg.window_retry_delay, 0.0, 10.0)
        cfg.min_operation_interval = _env_float(
            "WIN32_MCP_MIN_OPERATION_INTERVAL",
            cfg.min_operation_interval,
            0.0,
            10.0,
        )
        cfg.fuzzy_match_threshold = _env_int("WIN32_MCP_FUZZY_MATCH_THRESHOLD", cfg.fuzzy_match_threshold, 0, 100)
        cfg.result_envelope = _env_bool("WIN32_MCP_RESULT_ENVELOPE", cfg.result_envelope)
        cfg.debug = _env_bool("WIN32_MCP_DEBUG_LOGGING", _env_bool("WIN32_MCP_DEBUG", cfg.debug))

        cfg.limits.max_text_chars = _env_int("WIN32_MCP_MAX_TEXT_CHARS", cfg.limits.max_text_chars, 1)
        cfg.limits.max_clipboard_chars = _env_int(
            "WIN32_MCP_MAX_CLIPBOARD_CHARS",
            cfg.limits.max_clipboard_chars,
            1,
        )
        cfg.limits.max_command_chars = _env_int("WIN32_MCP_MAX_COMMAND_CHARS", cfg.limits.max_command_chars, 1)
        cfg.limits.max_args = _env_int("WIN32_MCP_MAX_ARGS", cfg.limits.max_args, 0)
        cfg.limits.max_arg_chars = _env_int("WIN32_MCP_MAX_ARG_CHARS", cfg.limits.max_arg_chars, 1)
        cfg.limits.max_subprocess_output_bytes = _env_int(
            "WIN32_MCP_MAX_SUBPROCESS_OUTPUT_BYTES",
            cfg.limits.max_subprocess_output_bytes,
            1024,
        )
        cfg.limits.max_sequence_steps = _env_int(
            "WIN32_MCP_MAX_SEQUENCE_STEPS",
            cfg.limits.max_sequence_steps,
            1,
            500,
        )
        cfg.limits.max_tool_runtime_seconds = _env_float(
            "WIN32_MCP_MAX_TOOL_RUNTIME_SECONDS",
            cfg.limits.max_tool_runtime_seconds,
            1.0,
        )

        cfg.security.profile = _env_choice(
            "WIN32_MCP_SECURITY_PROFILE",
            cfg.security.profile,
            VALID_SECURITY_PROFILES,
        )  # type: ignore[assignment]
        cfg.security.allowed_tools = _env_csv("WIN32_MCP_ALLOWED_TOOLS")
        cfg.security.blocked_tools = _env_csv("WIN32_MCP_BLOCKED_TOOLS")
        cfg.security.allowed_commands = _env_csv("WIN32_MCP_ALLOWED_COMMANDS")
        cfg.security.blocked_commands = _env_csv("WIN32_MCP_BLOCKED_COMMANDS")
        cfg.security.dry_run = _env_bool("WIN32_MCP_DRY_RUN", cfg.security.dry_run)
        cfg.security.confirmation_token = _env_str("WIN32_MCP_CONFIRMATION_TOKEN", cfg.security.confirmation_token)
        cfg.security.redact_sensitive_output = _env_bool(
            "WIN32_MCP_REDACT_SENSITIVE_OUTPUT",
            cfg.security.redact_sensitive_output,
        )

        return cfg


config = ServerConfig.from_env()
