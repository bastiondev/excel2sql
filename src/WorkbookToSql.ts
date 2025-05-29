import * as XLSX from 'xlsx';

/**
 * Regular expression to match cell references in templates
 * Matches patterns like:
 * - <SheetName>!A1
 * - <SheetName>!A1:A10
 * - <SheetName>!A1:
 */
const CELL_REFERENCE_REGEX = /<([^>]+)>!([A-Z]+\d+)(?::([A-Z]+\d+)?)?/g;

/**
 * Interface for parsed cell reference
 */
interface CellReference {
  sheetName: string;
  startCell: string;
  endCell?: string;
  isOpenRange: boolean;
}

/**
 * Interface for a range of cells
 */
interface CellRange {
  startRow: number;
  startCol: number;
  endRow?: number;
  endCol?: number;
  isOpenRange: boolean;
}

/**
 * Converts Excel workbook and SQL templates to SQL queries
 * 
 * @param workbook - The Excel workbook
 * @param templates - Array of SQL template strings with cell references
 * @returns Array of generated SQL queries
 */
export function workbookToSql(workbook: XLSX.WorkBook, templates: string[]): string[] {
  const results: string[] = [];
  
  // Process each template
  for (const template of templates) {
    // Extract all cell references from the template
    const cellReferences = extractCellReferences(template);
    
    // If no cell references found, add the template as is
    if (cellReferences.length === 0) {
      results.push(template);
      continue;
    }
    
    // Check if the template contains any open ranges
    const hasOpenRanges = cellReferences.some(ref => ref.isOpenRange);
    
    if (hasOpenRanges) {
      // Process templates with open ranges (multiple SQL statements)
      const rangeQueries = processRangeTemplate(workbook, template, cellReferences);
      results.push(...rangeQueries);
    } else {
      // Process templates with only single cell references (one SQL statement)
      const singleQuery = processSingleCellTemplate(workbook, template, cellReferences);
      results.push(singleQuery);
    }
  }
  
  return results;
}

/**
 * Extracts cell references from a template string
 * 
 * @param template - SQL template string
 * @returns Array of parsed cell references
 */
function extractCellReferences(template: string): CellReference[] {
  const references: CellReference[] = [];
  let match;
  
  // Reset regex to start from the beginning
  CELL_REFERENCE_REGEX.lastIndex = 0;
  
  while ((match = CELL_REFERENCE_REGEX.exec(template)) !== null) {
    const [, sheetName, startCell, endCell] = match;
    const isOpenRange = match[0].endsWith(':');
    
    references.push({
      sheetName,
      startCell,
      endCell: endCell || undefined,
      isOpenRange
    });
  }
  
  return references;
}

/**
 * Processes a template with only single cell references
 * 
 * @param workbook - The Excel workbook
 * @param template - SQL template string
 * @param cellReferences - Array of cell references
 * @returns Generated SQL query
 */
function processSingleCellTemplate(
  workbook: XLSX.WorkBook,
  template: string,
  cellReferences: CellReference[]
): string {
  let result = template;
  
  // Replace each cell reference with its value
  for (const ref of cellReferences) {
    const value = getCellValue(workbook, ref.sheetName, ref.startCell);
    const cellRef = `<${ref.sheetName}>!${ref.startCell}${ref.endCell ? `:${ref.endCell}` : ''}${ref.isOpenRange ? ':' : ''}`;
    
    // Replace the cell reference with its value
    result = result.replace(cellRef, formatValueForSql(value));
  }
  
  return result;
}

/**
 * Processes a template with range references
 * 
 * @param workbook - The Excel workbook
 * @param template - SQL template string
 * @param cellReferences - Array of cell references
 * @returns Array of generated SQL queries
 */
function processRangeTemplate(
  workbook: XLSX.WorkBook,
  template: string,
  cellReferences: CellReference[]
): string[] {
  const results: string[] = [];
  
  // Find all ranges and their lengths
  const ranges: CellRange[] = [];
  let maxRowCount = 0;
  
  for (const ref of cellReferences) {
    if (ref.isOpenRange || ref.endCell) {
      const range = parseCellRange(ref);
      ranges.push(range);
      
      // Calculate the row count for this range
      const sheet = workbook.Sheets[ref.sheetName];
      if (!sheet) continue;
      
      if (ref.isOpenRange) {
        // For open ranges, find the last row with data
        const lastRow = findLastRowInColumn(sheet, range.startCol);
        range.endRow = lastRow;
        const rowCount = lastRow - range.startRow + 1;
        maxRowCount = Math.max(maxRowCount, rowCount);
      } else if (range.endRow) {
        // For closed ranges, use the specified end row
        const rowCount = range.endRow - range.startRow + 1;
        maxRowCount = Math.max(maxRowCount, rowCount);
      }
    }
  }
  
  // Generate a query for each row in the range
  for (let rowIndex = 0; rowIndex < maxRowCount; rowIndex++) {
    let rowTemplate = template;
    
    // Replace each cell reference with its value for this row
    for (const ref of cellReferences) {
      const cellRef = `<${ref.sheetName}>!${ref.startCell}${ref.endCell ? `:${ref.endCell}` : ''}${ref.isOpenRange ? ':' : ''}`;
      
      if (ref.isOpenRange || ref.endCell) {
        // Handle range references
        const range = parseCellRange(ref);
        const currentRow = range.startRow + rowIndex;
        
        // Skip if we're past the end of a closed range
        if (range.endRow && currentRow > range.endRow) continue;
        
        // Get the cell at the current row in the range
        const cellId = getCellId(range.startCol, currentRow);
        const value = getCellValue(workbook, ref.sheetName, cellId);
        
        // Replace the range reference with the value from the current row
        rowTemplate = rowTemplate.replace(cellRef, formatValueForSql(value));
      } else {
        // Handle single cell references
        const value = getCellValue(workbook, ref.sheetName, ref.startCell);
        rowTemplate = rowTemplate.replace(cellRef, formatValueForSql(value));
      }
    }
    
    results.push(rowTemplate);
  }
  
  return results;
}

/**
 * Parses a cell reference into a range
 * 
 * @param ref - Cell reference
 * @returns Cell range with row and column indices
 */
function parseCellRange(ref: CellReference): CellRange {
  const startCoords = parseCellCoordinates(ref.startCell);
  let endCoords = undefined;
  
  if (ref.endCell) {
    endCoords = parseCellCoordinates(ref.endCell);
  }
  
  return {
    startRow: startCoords.row,
    startCol: startCoords.col,
    endRow: endCoords?.row,
    endCol: endCoords?.col,
    isOpenRange: ref.isOpenRange
  };
}

/**
 * Parses cell coordinates from a cell ID (e.g., 'A1' -> {col: 0, row: 1})
 * 
 * @param cellId - Cell ID (e.g., 'A1')
 * @returns Object with row and column indices (0-based)
 */
function parseCellCoordinates(cellId: string): { col: number; row: number } {
  // Extract column letters and row number
  const match = cellId.match(/([A-Z]+)(\d+)/);
  if (!match) throw new Error(`Invalid cell ID: ${cellId}`);
  
  const colLetters = match[1];
  const rowNumber = parseInt(match[2], 10);
  
  // Convert column letters to 0-based index (A=0, B=1, ..., Z=25, AA=26, ...)
  let colIndex = 0;
  for (let i = 0; i < colLetters.length; i++) {
    colIndex = colIndex * 26 + (colLetters.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
  }
  colIndex--; // Adjust to 0-based index
  
  // Return coordinates (0-based)
  return {
    col: colIndex,
    row: rowNumber - 1 // Convert to 0-based index
  };
}

/**
 * Converts column index and row index to cell ID (e.g., 0,0 -> 'A1')
 * 
 * @param colIndex - Column index (0-based)
 * @param rowIndex - Row index (0-based)
 * @returns Cell ID (e.g., 'A1')
 */
function getCellId(colIndex: number, rowIndex: number): string {
  let colId = '';
  let col = colIndex + 1; // Convert to 1-based for the calculation
  
  while (col > 0) {
    const remainder = (col - 1) % 26;
    colId = String.fromCharCode('A'.charCodeAt(0) + remainder) + colId;
    col = Math.floor((col - 1) / 26);
  }
  
  return `${colId}${rowIndex + 1}`; // Convert row back to 1-based
}

/**
 * Finds the last row with data in a specific column
 * 
 * @param sheet - Excel worksheet
 * @param colIndex - Column index (0-based)
 * @returns Last row index with data (0-based)
 */
function findLastRowInColumn(sheet: XLSX.WorkSheet, colIndex: number): number {
  let lastRow = 0;
  const colLetter = getCellId(colIndex, 0).replace(/\d+/, '');
  
  // Iterate through all cells in the sheet
  const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1:A1');
  for (let r = range.s.r; r <= range.e.r; r++) {
    const cellAddress = `${colLetter}${r + 1}`;
    const cell = sheet[cellAddress];
    
    if (cell && cell.v !== undefined && cell.v !== null) {
      lastRow = Math.max(lastRow, r);
    }
  }
  
  return lastRow;
}

/**
 * Gets the value of a cell from the workbook
 * 
 * @param workbook - Excel workbook
 * @param sheetName - Name of the sheet
 * @param cellId - Cell ID (e.g., 'A1')
 * @returns Cell value as string
 */
function getCellValue(workbook: XLSX.WorkBook, sheetName: string, cellId: string): any {
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) return '';
  
  const cell = sheet[cellId];
  if (!cell) return '';
  
  return cell.v;
}

/**
 * Converts a cell value to a string representation without any additional formatting
 * 
 * @param value - The cell value
 * @returns String representation of the cell value
 */
function formatValueForSql(value: any): string {
  if (value === null || value === undefined) {
    return '';
  }
  
  // Simply convert the value to a string without adding quotes or any formatting
  // The template will handle all quoting and formatting
  return String(value);
}
