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

## PyPI Publishing

PyPI release publishing is handled by `.github/workflows/publish-pypi.yml`.

For Trusted Publishing, configure the PyPI project with:

- Owner: `RandyNorthrup`
- Repository: `win32-mcp-server`
- Workflow: `publish-pypi.yml`
- Environment: leave blank

To publish an existing tag after the workflow is on `main`, run the **Publish PyPI** workflow manually and set `ref` to the release tag, for example `v2.6.1`.

## Rules

- Keep tool schemas and handler validation in sync.
- Route user-provided text, numbers, booleans, and enums through `win32_mcp_server.utils.args`.
- Keep mutating OS actions behind `registry.dispatch` so safety policy, dry-run, locks, and timeouts apply.
- Redact typed, pasted, clipboard, token, password, and secret-like values in responses and logs.
- Add focused tests for security policy, parsing, schema changes, and regression-prone helpers.
