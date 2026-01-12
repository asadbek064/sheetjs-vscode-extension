import * as vscode from 'vscode';
import { ExcelEditorProvider } from './excelEditorProvider';

export function activate(context: vscode.ExtensionContext) {
  console.log('SheetJS Spreadsheet Viewer extension activating...');

  const provider = ExcelEditorProvider.register(context);
  context.subscriptions.push(provider);

  // Command to disable viewer for current file extension
  const disableCommand = vscode.commands.registerCommand('sheetjs.disableForExtension', async (uri?: vscode.Uri) => {
    // Try to get URI from parameter, active text editor, or prompt user
    if (!uri) {
      const activeTab = vscode.window.tabGroups.activeTabGroup.activeTab;
      if (activeTab?.label) {
        // Try to find the file by label in workspace
        const files = await vscode.workspace.findFiles(`**/${activeTab.label}`);
        if (files.length > 0) {
          uri = files[0];
        }
      }
    }

    if (!uri) {
      vscode.window.showWarningMessage('No file is currently open or could not determine file path');
      return;
    }

    const ext = uri.fsPath.split('.').pop()?.toLowerCase();
    if (!ext) {
      vscode.window.showWarningMessage('Could not determine file extension');
      return;
    }

    // Update workbench.editorAssociations to use default editor
    const workbenchConfig = vscode.workspace.getConfiguration('workbench');
    const associations = workbenchConfig.get<Record<string, string>>('editorAssociations') || {};

    if (associations[`*.${ext}`] === 'default') {
      vscode.window.showInformationMessage(`SheetJS viewer is already disabled for .${ext} files`);
      return;
    }

    associations[`*.${ext}`] = 'default';
    await workbenchConfig.update('editorAssociations', associations, vscode.ConfigurationTarget.Global);

    // Close the current tab and reopen with default editor
    await vscode.commands.executeCommand('workbench.action.closeActiveEditor');
    await vscode.commands.executeCommand('vscode.open', uri);

    vscode.window.showInformationMessage(`SheetJS viewer disabled for .${ext} files`);
  });

  // Command to enable viewer for current file extension
  const enableCommand = vscode.commands.registerCommand('sheetjs.enableForExtension', async (uri?: vscode.Uri) => {
    // Try to get URI from parameter or active tab
    if (!uri) {
      uri = vscode.window.activeTextEditor?.document.uri;
      if (!uri) {
        const activeTab = vscode.window.tabGroups.activeTabGroup.activeTab;
        if (activeTab?.label) {
          const files = await vscode.workspace.findFiles(`**/${activeTab.label}`);
          if (files.length > 0) {
            uri = files[0];
          }
        }
      }
    }

    if (!uri) {
      vscode.window.showWarningMessage('No file is currently open or could not determine file path');
      return;
    }

    const ext = uri.fsPath.split('.').pop()?.toLowerCase();
    if (!ext) {
      vscode.window.showWarningMessage('Could not determine file extension');
      return;
    }

    // Remove from workbench.editorAssociations
    const workbenchConfig = vscode.workspace.getConfiguration('workbench');
    const associations = workbenchConfig.get<Record<string, string>>('editorAssociations') || {};

    // If not set to 'default', it's already enabled
    if (associations[`*.${ext}`] !== 'default') {
      vscode.window.showInformationMessage(`SheetJS viewer is already enabled for .${ext} files`);
      return;
    }

    // Create a new object without the extension (to avoid proxy issues)
    const newAssociations: Record<string, string> = {};
    for (const key in associations) {
      if (key !== `*.${ext}`) {
        newAssociations[key] = associations[key];
      }
    }
    await workbenchConfig.update('editorAssociations', newAssociations, vscode.ConfigurationTarget.Global);

    // Close the current tab and reopen with SheetJS viewer
    await vscode.commands.executeCommand('workbench.action.closeActiveEditor');
    await vscode.commands.executeCommand('vscode.openWith', uri, 'excelViewer.spreadsheet');

    vscode.window.showInformationMessage(`SheetJS viewer enabled for .${ext} files`);
  });

  // Command to open with viewer
  const openCommand = vscode.commands.registerCommand('sheetjs.openWithViewer', async () => {
    const uris = await vscode.window.showOpenDialog({
      canSelectMany: false,
      openLabel: 'Open with SheetJS Viewer',
      filters: {
        'Spreadsheets': ['xlsx', 'xls', 'csv', 'ods', 'xlsm', 'xlsb', 'numbers']
      }
    });

    if (uris && uris[0]) {
      await vscode.commands.executeCommand('vscode.openWith', uris[0], 'excelViewer.spreadsheet');
    }
  });

  context.subscriptions.push(disableCommand, enableCommand, openCommand);

  console.log('SheetJS Spreadsheet Viewer extension is now active');
  vscode.window.showInformationMessage('SheetJS Spreadsheet Viewer is ready!');
}

export function deactivate() {}