# Publishing Guide

## 1. Publish to PyPI

### Prerequisites
```bash
pip install build twine
```

### Create Account
1. Create account at https://pypi.org/account/register/
2. Verify email
3. Enable 2FA (required for uploads)
4. Generate API token at https://pypi.org/manage/account/token/
   - Scope: "Entire account" or specific to "win32-mcp-server"
   - Save the token securely (starts with `pypi-`)

### Build Package
```bash
# Clean previous builds
rm -rf dist/ build/ *.egg-info

# Build distribution packages
python -m build
```

This creates:
- `dist/win32_mcp_server-1.0.0-py3-none-any.whl` (wheel)
- `dist/win32-mcp-server-1.0.0.tar.gz` (source distribution)

### Test Upload (TestPyPI - Optional)
```bash
# Upload to test repository
python -m twine upload --repository testpypi dist/*

# Test installation
pip install --index-url https://test.pypi.org/simple/ win32-mcp-server
```

### Production Upload
```bash
# Upload to PyPI
python -m twine upload dist/*

# When prompted:
# Username: __token__
# Password: [paste your pypi-xxx token]
```

### Configure PyPI Token (Recommended)
Create `~/.pypirc`:
```ini
[pypi]
username = __token__
password = pypi-YourActualTokenHere

[testpypi]
username = __token__
password = pypi-YourTestTokenHere
```

Then upload without prompts:
```bash
python -m twine upload dist/*
```

---

## 2. Publish to VS Code MCP Extension Marketplace

VS Code MCP extensions are distributed through the **VS Code Marketplace** or **GitHub releases**.

### Option A: GitHub Release (Simpler)

1. **Push to GitHub**:
   ```bash
   git remote add origin https://github.com/RandyNorthrup/win32-mcp-server.git
   git branch -M main
   git push -u origin main
   ```

2. **Create GitHub Release**:
   - Go to https://github.com/RandyNorthrup/win32-mcp-server/releases/new
   - Tag: `v1.0.0`
   - Title: `Windows Automation Inspector v1.0.0`
   - Description: Copy from README features section
   - Attach `mcp.json` file
   - Publish release

3. **VS Code Discovery**:
   - Users install via PyPI: `pip install win32-mcp-server`
   - Add to VS Code `mcp.json` manually
   - Or use `mcp.json` from GitHub release

### Option B: VS Code Marketplace (Official Extension)

1. **Create VS Code Extension Package**:
   
   Create `package.json`:
   ```json
   {
     "name": "win32-mcp-inspector",
     "displayName": "Windows Automation Inspector (MCP)",
     "version": "1.0.0",
     "publisher": "YourVSCodePublisherId",
     "description": "Comprehensive Windows automation MCP server",
     "categories": ["Other"],
     "keywords": ["mcp", "windows", "automation"],
     "repository": {
       "type": "git",
       "url": "https://github.com/RandyNorthrup/win32-mcp-server"
     },
     "license": "MIT",
     "engines": {
       "vscode": "^1.80.0"
     },
     "main": "./extension.js",
     "contributes": {
       "configuration": {
         "title": "Windows MCP Inspector",
         "properties": {
           "win32-mcp.enabled": {
             "type": "boolean",
             "default": true,
             "description": "Enable Windows Automation Inspector MCP server"
           }
         }
       }
     },
     "scripts": {
       "postinstall": "pip install win32-mcp-server"
     }
   }
   ```

2. **Create Publisher Account**:
   - Go to https://marketplace.visualstudio.com/manage
   - Create publisher ID (e.g., "randy-northrup")

3. **Package Extension**:
   ```bash
   npm install -g @vscode/vsce
   vsce package
   ```

4. **Publish Extension**:
   ```bash
   vsce publish
   ```

### Option C: MCP Server Registry (Emerging)

The MCP ecosystem is new. Monitor these resources:
- https://github.com/modelcontextprotocol/servers - Official server list
- https://github.com/punkpeye/awesome-mcp-servers - Community registry

Submit PR to add your server:
```markdown
## win32-mcp-server
Comprehensive Windows automation with 25+ tools for screen capture, OCR, mouse/keyboard control, window management, and process control.

**Installation**: `pip install win32-mcp-server`
**GitHub**: https://github.com/RandyNorthrup/win32-mcp-server
**Platform**: Windows only
```

---

## 3. Post-Publication

### Update README Badges
Add to top of README.md:
```markdown
[![PyPI version](https://badge.fury.io/py/win32-mcp-server.svg)](https://badge.fury.io/py/win32-mcp-server)
[![Downloads](https://pepy.tech/badge/win32-mcp-server)](https://pepy.tech/project/win32-mcp-server)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
```

### Announce Release
- Reddit: r/Python, r/vscode, r/automation
- Twitter/X: #Python #VSCode #MCP #Automation
- Hacker News: Show HN
- Dev.to: Write tutorial article

### Monitor Issues
- GitHub Issues: https://github.com/RandyNorthrup/win32-mcp-server/issues
- PyPI project page: https://pypi.org/project/win32-mcp-server/

---

## Version Updates

When releasing new versions:

1. **Update version** in `pyproject.toml`:
   ```toml
   version = "1.0.1"
   ```

2. **Update `mcp.json`** version

3. **Commit and tag**:
   ```bash
   git add pyproject.toml mcp.json
   git commit -m "Bump version to 1.0.1"
   git tag v1.0.1
   git push origin main --tags
   ```

4. **Rebuild and upload**:
   ```bash
   rm -rf dist/
   python -m build
   python -m twine upload dist/*
   ```

5. **Create GitHub release** for the new tag

---

## Quick Checklist

- [ ] PyPI account created with 2FA
- [ ] API token generated and saved
- [ ] Package built: `python -m build`
- [ ] Uploaded to PyPI: `twine upload dist/*`
- [ ] GitHub repository created
- [ ] Code pushed to GitHub
- [ ] GitHub release created with `mcp.json`
- [ ] README badges added
- [ ] Announced on social media
- [ ] Submitted to MCP server registries
