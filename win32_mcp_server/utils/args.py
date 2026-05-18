"""Typed argument readers for tool handlers."""

from typing import Any, TypeVar

from ..config import config
from .errors import ToolError

T = TypeVar("T", bound=str)


def get_str(
    args: dict[str, Any],
    key: str,
    *,
    default: str | None = None,
    required: bool = False,
    min_length: int = 0,
    max_length: int | None = None,
) -> str:
    value = args.get(key, default)
    if value is None:
        if required:
            raise ToolError(f"Missing required argument: {key}")
        return ""
    if not isinstance(value, str):
        raise ToolError(f"Argument '{key}' must be a string")
    if len(value) < min_length:
        raise ToolError(f"Argument '{key}' must be at least {min_length} characters")
    if max_length is not None and len(value) > max_length:
        raise ToolError(f"Argument '{key}' must be at most {max_length} characters")
    return value


def get_text(args: dict[str, Any], key: str, *, required: bool = True) -> str:
    return get_str(
        args,
        key,
        required=required,
        min_length=1 if required else 0,
        max_length=config.limits.max_text_chars,
    )


def get_bool(args: dict[str, Any], key: str, *, default: bool = False) -> bool:
    value = args.get(key, default)
    if not isinstance(value, bool):
        raise ToolError(f"Argument '{key}' must be a boolean")
    return value


def get_int(
    args: dict[str, Any],
    key: str,
    *,
    default: int | None = None,
    required: bool = False,
    min_value: int | None = None,
    max_value: int | None = None,
) -> int:
    value = args.get(key, default)
    if value is None:
        if required:
            raise ToolError(f"Missing required argument: {key}")
        return 0
    if isinstance(value, bool):
        raise ToolError(f"Argument '{key}' must be an integer")
    if isinstance(value, float) and not value.is_integer():
        raise ToolError(f"Argument '{key}' must be an integer")
    try:
        parsed = int(value)
    except (TypeError, ValueError) as exc:
        raise ToolError(f"Argument '{key}' must be an integer") from exc
    _check_range(key, parsed, min_value, max_value)
    return parsed


def get_float(
    args: dict[str, Any],
    key: str,
    *,
    default: float | None = None,
    required: bool = False,
    min_value: float | None = None,
    max_value: float | None = None,
) -> float:
    value = args.get(key, default)
    if value is None:
        if required:
            raise ToolError(f"Missing required argument: {key}")
        return 0.0
    if isinstance(value, bool):
        raise ToolError(f"Argument '{key}' must be a number")
    try:
        parsed = float(value)
    except (TypeError, ValueError) as exc:
        raise ToolError(f"Argument '{key}' must be a number") from exc
    _check_range(key, parsed, min_value, max_value)
    return parsed


def get_enum(args: dict[str, Any], key: str, choices: set[T], *, default: T) -> T:
    value = get_str(args, key, default=default)
    if value not in choices:
        raise ToolError(f"Argument '{key}' must be one of: {', '.join(sorted(choices))}")
    return value


def get_timeout(args: dict[str, Any], key: str = "timeout_seconds", *, default: float | None = None) -> float:
    fallback = config.default_timeout if default is None else default
    return get_float(args, key, default=fallback, min_value=0.1, max_value=config.limits.max_timeout_seconds)


def get_poll_interval(args: dict[str, Any], key: str = "poll_interval", *, default: float = 0.5) -> float:
    return get_float(args, key, default=default, min_value=0.05, max_value=config.limits.max_poll_interval_seconds)


def get_region(args: dict[str, Any]) -> tuple[int, int, int, int]:
    return (
        get_int(args, "x", required=True),
        get_int(args, "y", required=True),
        get_int(args, "width", required=True, min_value=1),
        get_int(args, "height", required=True, min_value=1),
    )


def _check_range(
    key: str,
    value: float,
    min_value: float | None,
    max_value: float | None,
) -> None:
    if min_value is not None and value < min_value:
        raise ToolError(f"Argument '{key}' must be >= {min_value}")
    if max_value is not None and value > max_value:
        raise ToolError(f"Argument '{key}' must be <= {max_value}")
