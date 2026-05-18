# Security Policy

## Supported Versions

Security fixes target current `main` and latest published release.

## Reporting

Do not open public issues for exploitable behavior. Email the maintainer listed in `pyproject.toml` with:

- affected version or commit
- Windows version
- exact tool call or extension action
- expected impact
- safe reproduction steps

## Operational Guidance

This project can capture screens, read/write clipboard contents, type, click, launch processes, and terminate processes. In production or shared environments:

- keep default `WIN32_MCP_SECURITY_PROFILE=interactive` or use `read_only`
- set `WIN32_MCP_ALLOWED_TOOLS` and `WIN32_MCP_BLOCKED_TOOLS`
- set `WIN32_MCP_ALLOWED_COMMANDS` for process launching
- set `WIN32_MCP_CONFIRMATION_TOKEN` for high-risk actions
- keep `WIN32_MCP_REDACT_SENSITIVE_OUTPUT=true`
- keep default `WIN32_MCP_PYAUTOGUI_FAILSAFE=true` unless the automation host is isolated
- test changes with `WIN32_MCP_DRY_RUN=true` before allowing mutation
