const vscode = require('vscode');
const { execFile } = require('child_process');
const { promisify } = require('util');

const execFileAsync = promisify(execFile);
const PACKAGE_NAME = 'win32-mcp-server';
const PACKAGE_VERSION = '2.6.0';
const INSTALL_SPEC = `${PACKAGE_NAME}==${PACKAGE_VERSION}`;

async function run(cmd, args, options = {}) {
    return execFileAsync(cmd, args, {
        timeout: options.timeout || 30000,
        windowsHide: true,
        maxBuffer: 1024 * 1024,
    });
}

async function resolvePython() {
    for (const candidate of ['python', 'python3', 'py']) {
        try {
            await run(candidate, ['--version']);
            await run(candidate, ['-m', 'pip', '--version']);
            return candidate;
        } catch {
            // Try next candidate.
        }
    }
    return null;
}

async function installedPackageVersion(pythonCmd) {
    try {
        const { stdout } = await run(pythonCmd, [
            '-c',
            `import importlib.metadata as m; print(m.version("${PACKAGE_NAME}"))`,
        ]);
        return stdout.trim();
    } catch {
        return null;
    }
}

async function installPackage(pythonCmd, outputChannel) {
    outputChannel.appendLine(`Installing ${INSTALL_SPEC} from PyPI...`);
    const { stdout, stderr } = await run(
        pythonCmd,
        ['-m', 'pip', 'install', INSTALL_SPEC],
        { timeout: 120000 }
    );
    if (stdout) outputChannel.appendLine(stdout);
    if (stderr) outputChannel.appendLine(stderr);
}

async function checkTesseract(tesseractPath) {
    const candidates = [];
    if (tesseractPath && tesseractPath.trim()) candidates.push(tesseractPath.trim());
    candidates.push(
        'tesseract',
        'C:\\Program Files\\Tesseract-OCR\\tesseract.exe',
        'C:\\Program Files (x86)\\Tesseract-OCR\\tesseract.exe'
    );

    for (const executable of candidates) {
        try {
            await run(executable, ['--version']);
            return true;
        } catch {
            // Try next common install location.
        }
    }
    return false;
}

/**
 * @param {vscode.ExtensionContext} context
 */
async function activate(context) {
    console.log('Windows MCP Inspector extension v2.6 is now active');

    const statusBar = vscode.window.createStatusBarItem(vscode.StatusBarAlignment.Right, 100);
    statusBar.command = 'win32-mcp.openSettings';
    context.subscriptions.push(statusBar);

    function setStatus(state, tooltip) {
        const icons = { ready: '$(check)', error: '$(error)', disabled: '$(circle-slash)', loading: '$(sync~spin)' };
        statusBar.text = `${icons[state] || '$(info)'} Win32 MCP`;
        statusBar.tooltip = tooltip || 'Windows MCP Inspector';
        statusBar.show();
    }

    const disposable = vscode.commands.registerCommand('win32-mcp.openSettings', () => {
        vscode.commands.executeCommand('workbench.action.openSettings', 'win32-mcp');
    });
    context.subscriptions.push(disposable);

    const config = vscode.workspace.getConfiguration('win32-mcp');
    if (!config.get('enabled')) {
        console.log('Windows MCP Inspector is disabled in settings');
        setStatus('disabled', 'Win32 MCP Inspector disabled in settings');
        return;
    }

    setStatus('loading', 'Checking win32-mcp-server installation');

    const pythonCmd = await resolvePython();
    if (!pythonCmd) {
        vscode.window.showErrorMessage(
            'Python is not installed or not on PATH. win32-mcp-server requires Python 3.10+.',
            'Download Python'
        ).then((selection) => {
            if (selection === 'Download Python') {
                vscode.env.openExternal(vscode.Uri.parse('https://www.python.org/downloads/'));
            }
        });
        setStatus('error', 'Win32 MCP Inspector: Python not found');
        return;
    }

    const installedVersion = await installedPackageVersion(pythonCmd);
    if (installedVersion === PACKAGE_VERSION) {
        setStatus('ready', 'Win32 MCP Inspector ready');
        return;
    }

    const installMessage = installedVersion
        ? `${PACKAGE_NAME} ${installedVersion} is installed. Install pinned ${INSTALL_SPEC} to match this extension.`
        : `${PACKAGE_NAME} is not installed. Install ${INSTALL_SPEC} to use Windows MCP Inspector.`;

    if (!config.get('autoInstall')) {
        vscode.window.showWarningMessage(
            installMessage,
            'Install Now',
            'Show Instructions'
        ).then((selection) => {
            if (selection === 'Install Now') {
                const terminal = vscode.window.createTerminal('MCP Install');
                terminal.show();
                terminal.sendText(`${pythonCmd} -m pip install ${INSTALL_SPEC}`);
            } else if (selection === 'Show Instructions') {
                vscode.env.openExternal(vscode.Uri.parse('https://github.com/RandyNorthrup/win32-mcp-server#readme'));
            }
        });
        setStatus('error', 'Win32 MCP Inspector not installed');
        return;
    }

    const outputChannel = vscode.window.createOutputChannel('Windows MCP Inspector');
    const selection = await vscode.window.showInformationMessage(
        `Install ${INSTALL_SPEC} from PyPI?`,
        { modal: true },
        'Install',
        'Cancel'
    );
    if (selection !== 'Install') {
        setStatus('error', 'Win32 MCP Inspector install skipped');
        return;
    }

    try {
        await installPackage(pythonCmd, outputChannel);
        setStatus('ready', 'Win32 MCP Inspector ready');

        if (!(await checkTesseract(config.get('tesseractPath')))) {
            vscode.window.showWarningMessage(
                'Tesseract OCR is not installed. OCR tools will not work without it.',
                'Download Tesseract',
                'Ignore'
            ).then((tessSelection) => {
                if (tessSelection === 'Download Tesseract') {
                    vscode.env.openExternal(vscode.Uri.parse('https://github.com/UB-Mannheim/tesseract/wiki'));
                }
            });
        }

        vscode.window.showInformationMessage(
            'Windows MCP Inspector installed successfully. Restart VS Code to activate.',
            'Restart Now'
        ).then((restartSelection) => {
            if (restartSelection === 'Restart Now') {
                vscode.commands.executeCommand('workbench.action.reloadWindow');
            }
        });
    } catch (installError) {
        outputChannel.appendLine(`Installation failed: ${installError.message}`);
        outputChannel.show();
        vscode.window.showErrorMessage('Failed to install win32-mcp-server. See Output panel for details.');
        console.error('Installation error:', installError);
        setStatus('error', 'Win32 MCP Inspector install failed');
    }
}

function deactivate() {}

module.exports = {
    activate,
    deactivate,
};
