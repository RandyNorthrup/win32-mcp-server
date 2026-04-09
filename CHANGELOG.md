# Changelog

All notable changes to the **Windows Automation Inspector (MCP)** extension will be documented in this file.

## [2.5.0] — 2026-04-08

### Added
- 6 UI Automation API tools (`uia_inspect_window`, `uia_find_control`, `uia_click_control`, `uia_get_control_value`, `uia_set_control_value`, `uia_get_focused`) for control-based automation via Windows UI Automation.
- OCR result caching with perceptual image hashing and 2-second TTL for faster repeated calls.
- Operation verification: optional `verify` parameter on `click` and `focus_window`; auto-verification on `kill_process`.
- Status bar indicator in VS Code extension (loading/ready/error/disabled states).
- 8 new VS Code configuration settings (Tesseract path, OCR language, preprocess mode, screenshot format/quality/scale, coordinate validation, debug logging).
- Comprehensive test suite: 55 new feature tests, 48 E2E tests, 26 functional tests.

### Fixed
- COM initialization in UI Automation threads (`UIAutomationInitializerInThread` context manager).
- Added `pane`, `menubar`, `separator` to valid UIA control types.
- Strict typing and linting: ruff (40+ rule categories), mypy strict mode, pre-commit hooks.
- 206+ lint violations fixed across all source files.
- All functions annotated with return types and proper `dict[str, Any]` generics.

### Changed
- Total tool count: 53 (up from 47 in v2.0.0).
- Removed unused grid overlay tool.

## [2.0.0] — 2026-03-05

### Added
- 47 enterprise-grade automation tools including `click_text`, `wait_for_text`, structured OCR, batch operations, and more.
- Per-monitor DPI awareness for accurate coordinates on scaled displays.
- Tool registration decorator system with structured error handling.
- Configurable server settings (OCR, capture, automation parameters).
- Fuzzy window title matching with configurable threshold.
- Screenshot compression (PNG, JPEG, WebP) with quality and scale options.
- Multi-monitor support across all capture and coordinate tools.

## [1.0.1] — 2026-03-03

### Fixed
- Initial bug fixes and stability improvements.

## [1.0.0] — 2026-03-01

### Added
- Initial release with core Windows automation tools.
- Screen capture, OCR, mouse, keyboard, clipboard, window, and process tools.
- VS Code extension for automatic MCP server management.
