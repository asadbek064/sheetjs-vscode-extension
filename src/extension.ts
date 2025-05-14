import * as vscode from 'vscode';
import { ExcelEditorProvider } from './excelEditorProvider';

export function activate(context: vscode.ExtensionContext) {
  console.log('SheetJS Spreadsheet Viewer extension activating...');
  
  const provider = ExcelEditorProvider.register(context);
  context.subscriptions.push(provider);
  
  console.log('SheetJS Spreadsheet Viewer extension is now active');
  vscode.window.showInformationMessage('SheetJS Spreadsheet Viewer is ready!');
}

export function deactivate() {}