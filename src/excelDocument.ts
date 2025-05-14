import * as vscode from 'vscode';

export class ExcelDocument implements vscode.CustomDocument {
  constructor(public readonly uri: vscode.Uri) {}
  dispose() {}
}