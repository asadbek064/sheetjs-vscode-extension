import * as vscode from 'vscode';
import { ExcelEditorProvider } from './excelEditorProvider';

export function activate(context: vscode.ExtensionContext) {
  console.log('Excel Viewer extension activating...');
  
  const provider = ExcelEditorProvider.register(context);
  context.subscriptions.push(provider);
  
  console.log('Excel Viewer extension is now active');
  vscode.window.showInformationMessage('Excel Viewer is ready!');
}

export function deactivate() {}