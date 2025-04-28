/**
 * Parse a cell reference (e.g., "A1") to its row and column indices.
 * @param cellRef The cell reference (e.g., "A1").
 * @returns The row and column indices (1-based).
 */
export function cellRefToIndices(cellRef: string): { row: number; col: number } {
  const match = cellRef.match(/^([A-Z]+)(\d+)$/);
  if (!match) {
    throw new Error(`Invalid cell reference: ${cellRef}`);
  }
  
  const colStr = match[1];
  const row = parseInt(match[2], 10);
  
  // Convert column letter(s) to number
  let col = 0;
  for (let i = 0; i < colStr.length; i++) {
    col = col * 26 + (colStr.charCodeAt(i) - 64);
  }
  
  return { row, col };
}

/**
 * Parse a cell range (e.g., "A1:C10") to its row and column indices.
 * @param range The cell range (e.g., "A1:C10").
 * @returns The start and end row and column indices (1-based).
 */
export function cellRangeToIndices(range: string): { 
  startRow: number; 
  startCol: number; 
  endRow: number; 
  endCol: number; 
} {
  const parts = range.split(":");
  
  // If only one cell reference is provided, use it as both start and end
  const startCellRef = parts[0];
  const endCellRef = parts.length > 1 ? parts[1] : startCellRef;
  
  const start = cellRefToIndices(startCellRef);
  const end = cellRefToIndices(endCellRef);
  
  return {
    startRow: start.row,
    startCol: start.col,
    endRow: end.row,
    endCol: end.col
  };
}

/**
 * Convert column index to column letter(s) (e.g., 1 -> "A", 27 -> "AA").
 * @param colIndex The column index (1-based).
 * @returns The column letter(s) (e.g., "A", "B", "AA").
 */
export function colIndexToLetter(colIndex: number): string {
  let temp = colIndex;
  let letter = '';
  
  while (temp > 0) {
    const remainder = (temp - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    temp = Math.floor((temp - 1) / 26);
  }
  
  return letter;
}

/**
 * Convert row and column indices to a cell reference (e.g., { row: 1, col: 1 } -> "A1").
 * @param row The row index (1-based).
 * @param col The column index (1-based).
 * @returns The cell reference (e.g., "A1").
 */
export function indicesToCellRef(row: number, col: number): string {
  return `${colIndexToLetter(col)}${row}`;
}

/**
 * Convert start and end row and column indices to a cell range (e.g., { startRow: 1, startCol: 1, endRow: 10, endCol: 3 } -> "A1:C10").
 * @param startRow The start row index (1-based).
 * @param startCol The start column index (1-based).
 * @param endRow The end row index (1-based).
 * @param endCol The end column index (1-based).
 * @returns The cell range (e.g., "A1:C10").
 */
export function indicesToCellRange(startRow: number, startCol: number, endRow: number, endCol: number): string {
  return `${indicesToCellRef(startRow, startCol)}:${indicesToCellRef(endRow, endCol)}`;
}
