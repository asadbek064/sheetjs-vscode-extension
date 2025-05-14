import * as vscode from 'vscode';
import * as XLSX from 'xlsx';
import { ExcelDocument } from './excelDocument';
import { getLoadingViewHtml, getErrorViewHtml, getExcelViewerHtml } from './webviewContent';
import { parseRange, colLetterToNum, numToColLetter } from './excelUtils';

export class ExcelEditorProvider implements vscode.CustomReadonlyEditorProvider<ExcelDocument> {
  // cache workbooks in memory to avoid re-parsing
  private workbookCache = new Map<string, XLSX.WorkBook>();
  private sheetCache = new Map<string, string>();

  public static register(context: vscode.ExtensionContext): vscode.Disposable {
    return vscode.window.registerCustomEditorProvider(
      'excelViewer.spreadsheet', 
      new ExcelEditorProvider(context),
      { webviewOptions: { retainContextWhenHidden: true } } // keep webview state when hidden
    );
  }

  constructor(private readonly context: vscode.ExtensionContext) {}

  async openCustomDocument(uri: vscode.Uri): Promise<ExcelDocument> {
    console.log(`Opening document: ${uri.fsPath}`);
    return new ExcelDocument(uri);
  }

  async resolveCustomEditor(document: ExcelDocument, webviewPanel: vscode.WebviewPanel): Promise<void> {
    console.log(`Resolving editor for: ${document.uri.fsPath}`);
    
    webviewPanel.webview.options = { enableScripts: true };
    
    this.setLoadingView(webviewPanel);
    
    //  small timeout to ensure the loading view is rendered before heavy processing begins
    await new Promise(resolve => setTimeout(resolve, 50));
    
    try {
      // process file in a non-blocking way
      setImmediate(async () => {
        try {
          await this.processExcelFile(document, webviewPanel);
        } catch (error) {
          console.error('Error processing file:', error);
          this.setErrorView(webviewPanel, error);
        }
      });
    } catch (error) {
      console.error('Error setting up Excel viewer:', error);
      this.setErrorView(webviewPanel, error);
    }
  }

  private setLoadingView(webviewPanel: vscode.WebviewPanel): void {
    webviewPanel.webview.html = getLoadingViewHtml();
  }

  private setErrorView(webviewPanel: vscode.WebviewPanel, error: any): void {
    webviewPanel.webview.html = getErrorViewHtml(error);
  }

  private updateLoadingProgress(webviewPanel: vscode.WebviewPanel, message: string): void {
    webviewPanel.webview.postMessage({
      type: 'loadingProgress',
      message
    });
  }

  private async processExcelFile(document: ExcelDocument, webviewPanel: vscode.WebviewPanel): Promise<void> {
    // check if we have a cached workbook for this file
    let workbook: XLSX.WorkBook;
    const cacheKey = document.uri.toString();
    
    if (this.workbookCache.has(cacheKey)) {
      console.log('Using cached workbook');
      workbook = this.workbookCache.get(cacheKey)!;
      this.updateLoadingProgress(webviewPanel, 'Using cached workbook...');
    } else {
      workbook = await this.loadWorkbook(document, webviewPanel);
    }
    
    // setup the initial view with just the sheet selector
    const sheetNames = workbook.SheetNames;
    this.setupWebviewContent(document, webviewPanel, workbook, sheetNames);
  }

private async loadWorkbook(document: ExcelDocument, webviewPanel: vscode.WebviewPanel): Promise<XLSX.WorkBook> {
  this.updateLoadingProgress(webviewPanel, 'Reading file...');
  
  try {
    const data = await vscode.workspace.fs.readFile(document.uri);
    console.log(`Read ${data.byteLength} bytes`);
    
    // very large files, update the progress message
    if (data.byteLength > 5 * 1024 * 1024) { // 5MB
      this.updateLoadingProgress(webviewPanel, 'Parsing large file...');
    }
    
    // get file extension to optimize parsing options
    const fileExtension = document.uri.fsPath.split('.').pop()?.toLowerCase() || '';
    
    // parse with SheetJS with options optimized for the file type
    console.log(`Parsing ${fileExtension} file...`);
    this.updateLoadingProgress(webviewPanel, `Parsing ${fileExtension} data...`);
    
    // timeout to ensure UI updates before heavy parsing
    await new Promise(resolve => setTimeout(resolve, 10));
    
    const options: XLSX.ParsingOptions = { 
      type: 'array',
      cellStyles: true,
      cellDates: true,
    };
    
    // specific format handling
    if (fileExtension === 'csv') {
      options.cellDates = false;
      options.cellStyles = false;
    }
    
    const workbook = XLSX.read(new Uint8Array(data), options);
    
    this.updateLoadingProgress(webviewPanel, 'Preparing view...');
    
    // cache the workbook
    const cacheKey = document.uri.toString();
    this.workbookCache.set(cacheKey, workbook);
    
    return workbook;
  } catch (error) {
    console.error(`Error reading file: ${error}`);
    this.setErrorView(webviewPanel, error);
    throw error;
  }
}

 private setupWebviewContent(
  document: ExcelDocument, 
  webviewPanel: vscode.WebviewPanel, 
  workbook: XLSX.WorkBook, 
  sheetNames: string[]
): void {
  // exit early if there are no sheets
  if (sheetNames.length === 0) {
    webviewPanel.webview.html = getErrorViewHtml(new Error("No sheets or tables found in this file."));
    return;
  }

  // create a dropdown for sheet selection
  let sheetSelector = `
    <div class="sheet-selector">
      <label for="sheet-select">Select worksheet: </label>
      <select id="sheet-select">
  `;
  
  sheetNames.forEach((name, index) => {
    sheetSelector += `<option value="${index}">${name}</option>`;
  });
  
  sheetSelector += `
      </select>
    </div>
  `;
  
  // setup up the HTML with JS to handle sheet switching
  webviewPanel.webview.html = getExcelViewerHtml(sheetNames, sheetSelector);
  
  // handle messages from the webview
  this.setupMessageHandlers(document, webviewPanel, workbook);
}

  private setupMessageHandlers(
    document: ExcelDocument, 
    webviewPanel: vscode.WebviewPanel, 
    workbook: XLSX.WorkBook
  ): void {
    const baseCacheKey = document.uri.toString();
    
    webviewPanel.webview.onDidReceiveMessage(async message => {
      if (message.type === 'getSheetPage') {
        await this.handleGetSheetPage(
          baseCacheKey,
          workbook,
          webviewPanel,
          message
        );
      }
    });
  }

  private async handleGetSheetPage(
    baseCacheKey: string,
    workbook: XLSX.WorkBook,
    webviewPanel: vscode.WebviewPanel,
    message: any
  ): Promise<void> {
    const { sheetName, page, rowsPerPage, maxColumns } = message;
    const cacheKey = `${baseCacheKey}-${sheetName}-page-${page}`;
    
    this.updateLoadingProgress(webviewPanel, `Preparing page ${page + 1} of sheet: ${sheetName}`);
    
    // setTimeout to ensure loading message gets displayed
    setTimeout(async () => {
      try {
        let sheetHtml: string;
        let rangeInfo: any = null;
        
        // check if this page is already cached
        if (this.sheetCache.has(cacheKey)) {
          sheetHtml = this.sheetCache.get(cacheKey)!;
        } else {
          const sheet = workbook.Sheets[sheetName];
          
          // get the range of the sheet (e.g., A1:Z100)
          const range = sheet['!ref'] || '';
          rangeInfo = parseRange(range);
          
          if (!rangeInfo) {
            sheetHtml = '<p>No data in this sheet</p>';
          } else {
            sheetHtml = this.processSheetPage(sheet, rangeInfo, page, rowsPerPage, maxColumns);
            
            // cache the result
            this.sheetCache.set(cacheKey, sheetHtml);
          }
        }
        
        // send the sheet data back to the webview
        webviewPanel.webview.postMessage({
          type: 'sheetData',
          sheetName: sheetName,
          page: page,
          html: sheetHtml,
          range: rangeInfo
        });
      } catch (error) {
        console.error(`Error processing sheet ${sheetName}:`, error);
        webviewPanel.webview.postMessage({
          type: 'error',
          message: `Error loading sheet: ${error}`
        });
      }
    }, 10);
  }

  private processSheetPage(
    sheet: XLSX.WorkSheet, 
    rangeInfo: any, 
    page: number, 
    rowsPerPage: number, 
    maxColumns: number
  ): string {
    // calc the range for this page
    const pageStartRow = Math.min(rangeInfo.startRow + (page * rowsPerPage), rangeInfo.endRow);
    const pageEndRow = Math.min(pageStartRow + rowsPerPage - 1, rangeInfo.endRow);
    
    // limit columns if there are too many
    const effectiveEndColNum = Math.min(rangeInfo.endColNum, rangeInfo.startColNum + maxColumns - 1);
    const effectiveEndCol = numToColLetter(effectiveEndColNum);
    
    // create a new range for this page
    const pageRange = `${rangeInfo.startCol}${pageStartRow}:${effectiveEndCol}${pageEndRow}`;
    
    // create a new sheet with just this page's data
    const newPageSheet: XLSX.WorkSheet = { '!ref': pageRange };
    
    // preserve important sheet properties
    if (sheet['!cols']) newPageSheet['!cols'] = sheet['!cols'];
    if (sheet['!rows']) newPageSheet['!rows'] = sheet['!rows'];
    if (sheet['!merges']) {
      // filter merges that are in this page's range
      newPageSheet['!merges'] = sheet['!merges'].filter(merge => {
        return merge.s.r >= pageStartRow - 1 && 
                merge.e.r <= pageEndRow - 1 &&
                merge.s.c >= rangeInfo.startColNum - 1 &&
                merge.e.c <= effectiveEndColNum - 1;
      });
    }
    
    // copy cells in range
    Object.keys(sheet).forEach(key => {
      if (key.charAt(0) !== '!') {
        // check if cell is in our page range
        const cellCol = key.replace(/[0-9]/g, '');
        const cellRow = parseInt(key.replace(/[^0-9]/g, ''));
        const cellColNum = colLetterToNum(cellCol);
        
        if (cellRow >= pageStartRow && 
            cellRow <= pageEndRow && 
            cellColNum >= rangeInfo.startColNum && 
            cellColNum <= effectiveEndColNum) {
          newPageSheet[key] = sheet[key];
        }
      }
    });
    
    // convert to HTML
    return XLSX.utils.sheet_to_html(newPageSheet);
  }
}