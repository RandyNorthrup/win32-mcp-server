# Changelog

All notable changes to the **Windows Automation Inspector (MCP)** extension will be documented in this file.

## [Unreleased]

### Added
- Added a manual and release-triggered PyPI publishing workflow using Trusted Publishing.
- Hardened GitHub Actions by pinning action SHAs, reducing permissions, disabling checkout credential persistence, and validating PyPI publish refs.

## [2.6.1] — 2026-05-18

### Added
- Runtime safety profiles via `WIN32_MCP_SECURITY_PROFILE` (`read_only`, `interactive`, `unrestricted`), plus allow/block lists.
- Dry-run mode, high-risk confirmation tokens, and process command allow/block lists.
- Environment-backed server configuration for OCR, capture defaults, timeouts, subprocess output limits, redaction, and coordinate validation.
- Central tool argument validation and redacted audit logging.
- Central execution manager for serialized OS-mutating calls and hard tool timeouts.
- UI Automation-first paths for smart click/fill actions with OCR fallback.
- Optional `WIN32_MCP_RESULT_ENVELOPE` for stable `{success, tool, data}` JSON dict responses.
- Mutating tool schemas now advertise `dry_run`; high-risk tool schemas advertise `confirmation_token`.
- Windows GitHub Actions CI for ruff, mypy, pytest, and VS Code extension syntax checks.
- Dependabot and CodeQL workflows for dependency upkeep and static security scanning.
- `SECURITY.md`, `CONTRIBUTING.md`, `.editorconfig`, `.gitattributes`, and Python `dev` extra for release hygiene.
- Focused safety regression tests for policy, schema validation, OCR cache copying, UIA matching, and bounded subprocess output.
- CLI smoke flags: `--version`, `--list-tools`, and `--health-check`.
- Tesseract auto-discovery for common Windows install paths, so OCR works even when installer does not update PATH.

### Changed
- Default security profile is now `interactive`, blocking high-risk process/window actions unless explicitly set to `unrestricted`.
- PyAutoGUI fail-safe is now enabled by default; set `WIN32_MCP_PYAUTOGUI_FAILSAFE=false` only for isolated automation hosts.
- Package classifier moved to `Production/Stable`.

### Fixed
- Added the `types-psutil` development dependency so strict mypy CI passes across the release matrix.
- Detached started processes from MCP stdio to avoid protocol corruption.
- Bounded waited process output to prevent memory growth.
- Hardened mouse, keyboard, clipboard, window, capture, OCR, process, smart, and UIA argument parsing to reject malformed inputs consistently.
- Validated `press_key` and `hotkey` key names before calling PyAutoGUI.
- Rejected non-boolean per-call `dry_run` overrides instead of treating strings as truthy.
- Made registry safely serialize non-MCP content lists instead of returning invalid content objects.
- `health_check` now reports fail-safe, dry-run, and active allow/block policy counts.
- `python -m win32_mcp_server --version` now exits after printing version instead of starting stdio transport.
- Package import is now lazy and does not import server/tools until the console entry point runs.
- VS Code extension Tesseract check now probes common Windows install paths when no custom path is configured.
- Made VS Code extension detect exact pinned server version and honor custom Tesseract path.
- Scoped UI Automation click/read/set searches to the target window.
- Fixed OCR cache mutation leak from structured OCR callers.
- Rejected invalid `click_text` occurrences before OCR/click work.
- Validated coordinates against actual monitor rectangles instead of only the virtual bounding box.
- Made VS Code install flow use explicit prompt, pinned PyPI package, and shell-free `execFile`.
- Cached Tesseract availability checks to reduce repeated OCR startup overhead.

## [2.5.1] — 2026-04-09

### Fixed
- UIA control type mapping: `_get_control_class()` now correctly maps multi-word types (TabItem, CheckBox, RadioButton, ComboBox, etc.) using explicit PascalCase lookup instead of broken `.capitalize()` logic.
- mss screen capture thread-safety: replaced global singleton with `threading.local()` per-thread instances to fix intermittent `_thread._local srcdc` errors (GDI DCs are thread-affine on Windows).
- Cleaned up mypy type: ignore comments and removed unused pyproject.toml overrides.

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
