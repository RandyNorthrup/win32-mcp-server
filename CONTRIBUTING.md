# Contributing

## Local Checks

Run this before opening a PR:

```powershell
python -m pip install -e .[dev]
python -m ruff check win32_mcp_server tests
python -m mypy win32_mcp_server
python -m pytest -q
node --check extension.js
```

## Rules

- Keep tool schemas and handler validation in sync.
- Route user-provided text, numbers, booleans, and enums through `win32_mcp_server.utils.args`.
- Keep mutating OS actions behind `registry.dispatch` so safety policy, dry-run, locks, and timeouts apply.
- Redact typed, pasted, clipboard, token, password, and secret-like values in responses and logs.
- Add focused tests for security policy, parsing, schema changes, and regression-prone helpers.
