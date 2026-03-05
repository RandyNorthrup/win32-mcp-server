# Changelog

All notable changes to the **Windows Automation Inspector (MCP)** extension will be documented in this file.

## [2.0.0] — 2026-03-05

### Added
- **`capture_grid` tool** — Captures the screen with a labelled coordinate grid overlay for precise pixel identification. DPI/zoom-aware, resolution-adaptive, high-contrast dual-colour rendering visible on any background.
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
