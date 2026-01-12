const commonStyles = `
  body { 
    font-family: var(--vscode-font-family);
    background-color: var(--vscode-editor-background);
    color: var(--vscode-foreground);
    margin: 0;
    padding: 0;
  }
`;

const loadingStyles = `
  body {
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    height: 100vh;
    text-align: center;
  }
  .loader-container {
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
  }
  .loader {
    border: 6px solid var(--vscode-panel-border);
    border-radius: 50%;
    border-top: 6px solid var(--vscode-button-background);
    width: 60px;
    height: 60px;
    animation: spin 1.5s linear infinite;
    margin: 30px auto;
  }
  @keyframes spin {
    0% { transform: rotate(0deg); }
    100% { transform: rotate(360deg); }
  }
  h2 { margin-bottom: 20px; }
  .message { margin-top: 20px; max-width: 80%; }
`;

const errorStyles = `
  body { padding: 20px; }
  .error-container {
    border: 1px solid var(--vscode-inputValidation-errorBorder);
    background-color: var(--vscode-inputValidation-errorBackground);
    padding: 20px;
    border-radius: 4px;
    margin-top: 20px;
  }
  h1 { color: var(--vscode-errorForeground); }
  pre {
    background-color: var(--vscode-editor-background);
    padding: 10px;
    overflow: auto;
    border-radius: 4px;
  }
`;

const excelViewerStyles = `
  body { 
    padding: 0;
    overflow-x: auto;
  }
  .container {
    display: flex;
    flex-direction: column;
    height: 100vh;
  }
  .sheet-selector { 
    flex: 0 0 auto;
    padding: 10px 20px;
    margin-bottom: 0; 
    background-color: var(--vscode-editor-background);
    position: sticky;
    top: 0;
    z-index: 10;
    border-bottom: 1px solid var(--vscode-panel-border);
  }
  .sheet-content {
    flex: 1 1 auto;
    overflow: auto;
    padding: 0 20px;
  }
  table { 
    border-collapse: collapse; 
    margin-top: 15px;
    table-layout: fixed;
  }
  th, td { 
    border: 1px solid var(--vscode-panel-border); 
    padding: 6px; 
    text-align: left;
    min-width: 80px;
    max-width: 300px;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
  }
  th { 
    background-color: var(--vscode-list-hoverBackground); 
    font-weight: bold;
    position: sticky;
    top: 0;
    z-index: 5;
  }
  .status-bar {
    flex: 0 0 auto;
    padding: 5px 20px;
    background-color: var(--vscode-editor-background);
    border-top: 1px solid var(--vscode-panel-border);
    font-size: 12px;
    color: var(--vscode-descriptionForeground);
  }
  select { 
    padding: 5px; 
    background-color: var(--vscode-dropdown-background);
    color: var(--vscode-dropdown-foreground);
    border: 1px solid var(--vscode-dropdown-border);
    border-radius: 2px;
    min-width: 200px;
  }
  label { margin-right: 10px; }
  .loader {
    border: 4px solid var(--vscode-panel-border);
    border-radius: 50%;
    border-top: 4px solid var(--vscode-button-background);
    width: 20px;
    height: 20px;
    animation: spin 1s linear infinite;
    display: inline-block;
    vertical-align: middle;
    margin-left: 10px;
  }
  .loading-indicator {
    display: none;
    align-items: center;
    margin-left: 15px;
  }
  .loading-text { margin-left: 8px; }
  .controls { display: flex; align-items: center; }
  .pagination {
    margin-left: auto;
    display: flex;
    align-items: center;
  }
  button {
    background-color: var(--vscode-button-background);
    color: var(--vscode-button-foreground);
    border: none;
    padding: 4px 10px;
    margin: 0 5px;
    border-radius: 2px;
    cursor: pointer;
  }
  button:disabled {
    opacity: 0.5;
    cursor: not-allowed;
  }
  .range-indicator {
    margin: 0 10px;
    font-size: 12px;
  }
  .nav-controls {
    display: flex;
    justify-content: space-between;
    align-items: center;
    width: 100%;
    margin-bottom: 10px;
  }
`;

const loadingScript = `
  // set up message listener to update loading status
  window.addEventListener('message', event => {
    const message = event.data;
    if (message && message.type === 'loadingProgress') {
      document.getElementById('loading-status').textContent = message.message;
    }
  });
`;

function getExcelViewerScript(sheetNames: string[], maxRows: number, maxColumns: number) {
  return `
    // config
    const config = {
      rowsPerPage: ${maxRows},
      maxColumns: ${maxColumns}
    };
    
    // Store sheet names and state
    const sheetNames = ${JSON.stringify(sheetNames)};
    const vscode = acquireVsCodeApi();
    const cache = {};
    let currentPage = 0;
    let currentSheetName = '';
    let currentSheetRange = null;
    
    // DOM elements
    const currentSheetElement = document.getElementById('current-sheet');
    const prevButton = document.getElementById('prev-page');
    const nextButton = document.getElementById('next-page');
    const rangeIndicator = document.getElementById('range-indicator');
    const statusBar = document.getElementById('status-bar');
    const sheetSelector = document.getElementById('sheet-select');
    
    // loading indicator functions
    function showLoading(message = "Loading...") {
      const indicator = document.getElementById('loading-indicator');
      const loadingText = indicator.querySelector('.loading-text');
      loadingText.textContent = message;
      indicator.style.display = 'flex';
    }
    
    function hideLoading() {
      document.getElementById('loading-indicator').style.display = 'none';
    }
    
    function parseSheetRange(rangeStr) {
      if (!rangeStr || !rangeStr.includes(':')) return null;
      
      const parts = rangeStr.split(':');
      const startCell = parts[0];
      const endCell = parts[1];
      
      // extract column letters and row numbers
      const startCol = startCell.replace(/[0-9]/g, '');
      const endCol = endCell.replace(/[0-9]/g, '');
      const startRow = parseInt(startCell.replace(/[^0-9]/g, ''));
      const endRow = parseInt(endCell.replace(/[^0-9]/g, ''));
      
      // calc dimensions
      const totalRows = endRow - startRow + 1;
      const startColNum = colLetterToNum(startCol);
      const endColNum = colLetterToNum(endCol);
      const totalCols = endColNum - startColNum + 1;
      
      return {
        startRow, endRow, startCol, endCol, 
        startColNum, endColNum, totalRows, totalCols
      };
    }
    
    function colLetterToNum(col) {
      let result = 0;
      for (let i = 0; i < col.length; i++) {
        result = result * 26 + (col.charCodeAt(i) - 64);
      }
      return result;
    }
    
    function numToColLetter(num) {
      let result = '';
      while (num > 0) {
        const modulo = (num - 1) % 26;
        result = String.fromCharCode(65 + modulo) + result;
        num = Math.floor((num - modulo) / 26);
      }
      return result || 'A';
    }
    
    // UI update functions
    function updatePaginationControls(page, totalRows) {
      const totalPages = Math.ceil(totalRows / config.rowsPerPage);
      prevButton.disabled = page <= 0;
      nextButton.disabled = page >= totalPages - 1;
      
      const startRow = page * config.rowsPerPage + 1;
      const endRow = Math.min((page + 1) * config.rowsPerPage, totalRows);
      
      rangeIndicator.textContent = \`Rows \${startRow}-\${endRow} of \${totalRows}\`;
    }
    
    function updateStatusBar(sheetName, range) {
      if (!range) {
        statusBar.textContent = \`Sheet: \${sheetName}\`;
        return;
      }
      
      statusBar.textContent = \`Sheet: \${sheetName} | Total: \${range.totalRows} rows × \${range.totalCols} columns\`;
    }
    
    // Sheet loading function
    function loadSheetPage(sheetName, page) {
      currentPage = page;
      currentSheetName = sheetName;
      
      showLoading(\`Loading page \${page + 1} of sheet \${sheetName}...\`);
      
      vscode.postMessage({
        type: 'getSheetPage',
        sheetName: sheetName,
        page: page,
        rowsPerPage: config.rowsPerPage,
        maxColumns: config.maxColumns
      });
    }
    
    // Event handling
    window.addEventListener('message', event => {
      const message = event.data;
      
      if (message.type === 'sheetData') {
        // add the sheet to the cache
        const cacheKey = \`\${message.sheetName}-page-\${message.page}\`;
        cache[cacheKey] = message.html;
        
        // update range info if provided
        if (message.range) {
          currentSheetRange = message.range;
          updatePaginationControls(message.page, message.range.totalRows);
          updateStatusBar(message.sheetName, message.range);
        }
        
        // update the view if this is still the current sheet and page
        if (currentSheetName === message.sheetName && currentPage === message.page) {
          currentSheetElement.innerHTML = \`<h2>\${message.sheetName}</h2>\${message.html}\`;
          hideLoading();
        }
      } 
      else if (message.type === 'loadingProgress') {
        showLoading(message.message);
      }
      else if (message.type === 'error') {
        currentSheetElement.innerHTML = \`<h2>Error</h2><p>\${message.message}</p>\`;
        hideLoading();
      }
    });
    
    // UI event listeners
    sheetSelector.addEventListener('change', function(e) {
      const selectedIndex = parseInt(e.target.value);
      const selectedSheet = sheetNames[selectedIndex];
      
      // reset page to 0 when changing sheets
      currentPage = 0;
      loadSheetPage(selectedSheet, 0);
    });
    
    prevButton.addEventListener('click', function() {
      if (currentPage > 0) {
        loadSheetPage(currentSheetName, currentPage - 1);
      }
    });
    
    nextButton.addEventListener('click', function() {
      loadSheetPage(currentSheetName, currentPage + 1);
    });
    
    // init with first sheet
    if (sheetNames.length > 0) {
      loadSheetPage(sheetNames[0], 0);
    }
  `;
}

// create HTML document with given components
function createHtmlDocument(styles: string, body: string, script: string = ''): string {
  return `
    <!DOCTYPE html>
    <html>
      <head>
        <style>
          ${commonStyles}
          ${styles}
        </style>
      </head>
      <body>
        ${body}
        
        <script>
          ${script}
        </script>
      </body>
    </html>
  `;
}

// HTML for loading view
export function getLoadingViewHtml(): string {
  const body = `
    <div class="loader-container">
      <h2>Loading File</h2>
      <div class="loader"></div>
      <p class="message">Please wait while we process the file. Large files may take a moment to load...</p>
      <p id="loading-status">Initializing...</p>
    </div>
  `;
  
  return createHtmlDocument(loadingStyles, body, loadingScript);
}

// HTML for error view
export function getErrorViewHtml(error: any): string {
  const body = `
    <h1>Error Opening File</h1>
    <p>An error occurred while trying to open the file:</p>
    <div class="error-container">
      <pre>${error.toString()}</pre>
    </div>
  `;
  
  return createHtmlDocument(errorStyles, body);
}

// HTML for Excel viewer
export function getExcelViewerHtml(
  sheetNames: string[],
  sheetSelector: string,
  maxRows: number = 1000,
  maxColumns: number = 100
): string {
  const body = `
    <div class="container">
      <div class="sheet-selector">
        <div class="nav-controls">
          <div class="controls">
            ${sheetSelector}
            <div id="loading-indicator" class="loading-indicator">
              <div class="loader"></div>
              <span class="loading-text">Loading...</span>
            </div>
          </div>
          <div class="pagination">
            <button id="prev-page" disabled>← Previous</button>
            <span class="range-indicator" id="range-indicator">Loading...</span>
            <button id="next-page">Next →</button>
          </div>
        </div>
      </div>
      
      <div class="sheet-content">
        <div id="current-sheet">
          <p>Loading sheet data...</p>
        </div>
      </div>
      
      <div class="status-bar" id="status-bar">
        Ready
      </div>
    </div>
  `;
  
  return createHtmlDocument(excelViewerStyles, body, getExcelViewerScript(sheetNames, maxRows, maxColumns));
}