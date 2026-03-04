const vscode = require('vscode');
const { exec } = require('child_process');
const { promisify } = require('util');
const execAsync = promisify(exec);

/**
 * @param {vscode.ExtensionContext} context
 */
async function activate(context) {
    console.log('Windows MCP Inspector extension v2.0 is now active');

    const config = vscode.workspace.getConfiguration('win32-mcp');
    
    if (!config.get('enabled')) {
        console.log('Windows MCP Inspector is disabled in settings');
        return;
    }

    // Check if win32-mcp-server package is installed
    try {
        await execAsync('pip show win32-mcp-server');
        console.log('win32-mcp-server is already installed');
    } catch (error) {
        if (config.get('autoInstall')) {
            vscode.window.showInformationMessage(
                'Installing Windows MCP Inspector dependencies...',
                'Show Output'
            ).then(selection => {
                if (selection === 'Show Output') {
                    const outputChannel = vscode.window.createOutputChannel('Windows MCP Inspector');
                    outputChannel.show();
                    outputChannel.appendLine('Installing win32-mcp-server...');
                }
            });

            try {
                const { stdout, stderr } = await execAsync('pip install win32-mcp-server');
                console.log('win32-mcp-server installed successfully:', stdout);
                
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
                vscode.window.showErrorMessage(
                    'Failed to install win32-mcp-server. Please run: pip install win32-mcp-server',
                    'Open Terminal'
                ).then(selection => {
                    if (selection === 'Open Terminal') {
                        const terminal = vscode.window.createTerminal('MCP Install');
                        terminal.show();
                        terminal.sendText('pip install win32-mcp-server');
                    }
                });
                console.error('Installation error:', installError);
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
                    terminal.sendText('pip install win32-mcp-server');
                } else if (selection === 'Show Instructions') {
                    vscode.env.openExternal(vscode.Uri.parse('https://github.com/RandyNorthrup/win32-mcp-server#readme'));
                }
            });
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
