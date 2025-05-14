// parse range like A1:Z100 and return information about dimensions
export function parseRange(rangeStr: string) {
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
  const startColNum = colLetterToNum(startCol);
  const endColNum = colLetterToNum(endCol);
  const totalRows = endRow - startRow + 1;
  const totalCols = endColNum - startColNum + 1;
  
  return {
    startRow,
    endRow,
    startCol,
    endCol,
    startColNum,
    endColNum,
    totalRows,
    totalCols
  };
}

// (A=1, B=2, etc.)
export function colLetterToNum(col: string): number {
  let result = 0;
  for (let i = 0; i < col.length; i++) {
    result = result * 26 + (col.charCodeAt(i) - 64);
  }
  return result;
}

// (1=A, 2=B, etc.)
export function numToColLetter(num: number): string {
  let result = '';
  while (num > 0) {
    const modulo = (num - 1) % 26;
    result = String.fromCharCode(65 + modulo) + result;
    num = Math.floor((num - modulo) / 26);
  }
  return result || 'A';
}