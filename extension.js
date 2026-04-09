const vscode = require('vscode');
const { exec } = require('child_process');
const { promisify } = require('util');
const execAsync = promisify(exec);

const GITHUB_INSTALL_URL = 'git+https://github.com/RandyNorthrup/win32-mcp-server.git';

/**
 * Resolve the pip command — prefers `python -m pip` for reliability.
 * Falls back to bare `pip` if Python can't be found.
 */
async function resolvePip() {
    // Try python -m pip first (most reliable on Windows)
    for (const py of ['python', 'python3', 'py']) {
        try {
            await execAsync(`${py} -m pip --version`);
            return `${py} -m pip`;
        } catch { /* try next */ }
    }
    // Bare pip as last resort
    try {
        await execAsync('pip --version');
        return 'pip';
    } catch {
        return null;
    }
}

/**
 * @param {vscode.ExtensionContext} context
 */
async function activate(context) {
    console.log('Windows MCP Inspector extension v2.5 is now active');

    // --- Status bar indicator ---
    const statusBar = vscode.window.createStatusBarItem(vscode.StatusBarAlignment.Right, 100);
    statusBar.command = 'win32-mcp.openSettings';
    context.subscriptions.push(statusBar);

    function setStatus(state, tooltip) {
        const icons = { ready: '$(check)', error: '$(error)', disabled: '$(circle-slash)', loading: '$(sync~spin)' };
        statusBar.text = `${icons[state] || '$(info)'} Win32 MCP`;
        statusBar.tooltip = tooltip || 'Windows MCP Inspector';
        statusBar.show();
    }

    const config = vscode.workspace.getConfiguration('win32-mcp');

    if (!config.get('enabled')) {
        console.log('Windows MCP Inspector is disabled in settings');
        setStatus('disabled', 'Win32 MCP Inspector — disabled in settings');
        return;
    }

    setStatus('loading', 'Checking win32-mcp-server installation…');

    // Resolve pip command
    const pip = await resolvePip();
    if (!pip) {
        vscode.window.showErrorMessage(
            'Python is not installed or not on PATH. win32-mcp-server requires Python 3.10+.',
            'Download Python'
        ).then(selection => {
            if (selection === 'Download Python') {
                vscode.env.openExternal(vscode.Uri.parse('https://www.python.org/downloads/'));
            }
        });
        setStatus('error', 'Win32 MCP Inspector — Python not found');
        return;
    }

    // Check if win32-mcp-server package is installed
    try {
        await execAsync(`${pip} show win32-mcp-server`);
        console.log('win32-mcp-server is already installed');
        setStatus('ready', 'Win32 MCP Inspector — ready');
    } catch (error) {
        if (config.get('autoInstall')) {
            const outputChannel = vscode.window.createOutputChannel('Windows MCP Inspector');
            vscode.window.showInformationMessage(
                'Installing Windows MCP Inspector and all dependencies...',
                'Show Output'
            ).then(selection => {
                if (selection === 'Show Output') {
                    outputChannel.show();
                }
            });

            outputChannel.appendLine(`Installing win32-mcp-server from GitHub...`);
            outputChannel.appendLine(`Running: ${pip} install ${GITHUB_INSTALL_URL}`);

            try {
                const { stdout, stderr } = await execAsync(
                    `${pip} install "${GITHUB_INSTALL_URL}"`,
                    { timeout: 120000 }
                );
                if (stdout) outputChannel.appendLine(stdout);
                if (stderr) outputChannel.appendLine(stderr);
                console.log('win32-mcp-server installed successfully');
                setStatus('ready', 'Win32 MCP Inspector — ready');

                // Check for Tesseract
                try {
                    await execAsync('tesseract --version');
                } catch (tessError) {
                    vscode.window.showWarningMessage(
                        'Tesseract OCR is not installed. OCR tools will not work without it.',
                        'Download Tesseract',
                        'Ignore'
                    ).then(selection => {
                        if (selection === 'Download Tesseract') {
                            vscode.env.openExternal(vscode.Uri.parse('https://github.com/UB-Mannheim/tesseract/wiki'));
                        }
                    });
                }

                vscode.window.showInformationMessage(
                    'Windows MCP Inspector installed successfully! Restart VS Code to activate.',
                    'Restart Now'
                ).then(selection => {
                    if (selection === 'Restart Now') {
                        vscode.commands.executeCommand('workbench.action.reloadWindow');
                    }
                });
            } catch (installError) {
                outputChannel.appendLine(`Installation failed: ${installError.message}`);
                outputChannel.show();
                vscode.window.showErrorMessage(
                    'Failed to install win32-mcp-server. See Output panel for details.',
                    'Open Terminal'
                ).then(selection => {
                    if (selection === 'Open Terminal') {
                        const terminal = vscode.window.createTerminal('MCP Install');
                        terminal.show();
                        terminal.sendText(`${pip} install "${GITHUB_INSTALL_URL}"`);
                    }
                });
                console.error('Installation error:', installError);
                setStatus('error', 'Win32 MCP Inspector — install failed');
            }
        } else {
            vscode.window.showWarningMessage(
                'win32-mcp-server is not installed. Install it to use Windows MCP Inspector.',
                'Install Now',
                'Show Instructions'
            ).then(selection => {
                if (selection === 'Install Now') {
                    const terminal = vscode.window.createTerminal('MCP Install');
                    terminal.show();
                    terminal.sendText(`${pip} install "${GITHUB_INSTALL_URL}"`);
                } else if (selection === 'Show Instructions') {
                    vscode.env.openExternal(vscode.Uri.parse('https://github.com/RandyNorthrup/win32-mcp-server#readme'));
                }
            });
            setStatus('error', 'Win32 MCP Inspector — not installed');
        }
    }

    // Register commands
    let disposable = vscode.commands.registerCommand('win32-mcp.openSettings', () => {
        vscode.commands.executeCommand('workbench.action.openSettings', 'win32-mcp');
    });

    context.subscriptions.push(disposable);
}

function deactivate() {}

module.exports = {
    activate,
    deactivate
};
