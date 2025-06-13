import * as XLSX from 'xlsx';

/**
 * Regular expression to match query references in templates
 * Matches patterns like:
 * - ?queryName.columnName
 * - ?queryName[index].columnName
 */
const QUERY_REFERENCE_REGEX = /\?(\w+)(?:\[(\d+)\])?\.([\w_]+)/g;

/**
 * Interface for query results mapping
 */
interface QueryResults {
  [queryName: string]: any[];
}

/**
 * Populates an Excel workbook template with SQL query results
 * 
 * @param template - The template workbook
 * @param queryResults - Map of query names to their results
 * @returns Populated Excel workbook
 */
export function sqlToWorkbook(template: XLSX.WorkBook, queryResults: QueryResults): XLSX.WorkBook {
  // Create a copy of the template workbook
  const workbook = JSON.parse(JSON.stringify(template));
  
  // Process each sheet in the workbook
  for (const sheetName of Object.keys(workbook.Sheets)) {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet['!ref']) continue;
    
    // Get the range of cells in the sheet
    const range = XLSX.utils.decode_range(sheet['!ref']);
    
    // First pass: collect all cells and their query references
    const cellRefs = new Map<string, { queryName: string; index?: number; columnName: string }>();
    const iterativeRows = new Map<number, { queryName: string; columns: Map<number, string> }>();
    
    // Scan all cells for query references
    for (let r = range.s.r; r <= range.e.r; r++) {
      for (let c = range.s.c; c <= range.e.c; c++) {
        const addr = XLSX.utils.encode_cell({ r, c });
        const cell = sheet[addr];
        
        if (cell?.v && typeof cell.v === 'string') {
          QUERY_REFERENCE_REGEX.lastIndex = 0;
          const match = QUERY_REFERENCE_REGEX.exec(cell.v);
          
          if (match) {
            const [, queryName, index, columnName] = match;
            
            if (!index) {
              // Iterative reference
              if (!iterativeRows.has(r)) {
                iterativeRows.set(r, { queryName, columns: new Map() });
              }
              iterativeRows.get(r)!.columns.set(c, columnName);
            } else {
              // Direct reference
              cellRefs.set(addr, { queryName, index: +index, columnName });
            }
          }
        }
      }
    }
    
    // Process direct references first
    for (const [addr, ref] of cellRefs) {
      const results = queryResults[ref.queryName];
      if (results?.[ref.index!]?.[ref.columnName] !== undefined) {
        const value = results[ref.index!][ref.columnName];
        sheet[addr] = { t: typeof value === 'number' ? 'n' : 's', v: value };
      }
    }
    
    // Process iterative rows in order
    const sortedRows = Array.from(iterativeRows.entries()).sort(([a], [b]) => a - b);
    let totalRowsInserted = 0;
    
    for (const [templateRow, { queryName, columns }] of sortedRows) {
      const results = queryResults[queryName];
      if (!results?.length) continue;
      
      // Calculate rows to insert
      const rowsToInsert = results.length - 1; // -1 because we'll use the template row
      
      if (rowsToInsert > 0) {
        // Shift all cells below this row down
        const cellsToMove = new Map<string, XLSX.CellObject>();
        
        // Collect cells to move
        for (const [addr, cell] of Object.entries(sheet)) {
          if (addr === '!ref' || !cell || typeof cell !== 'object') continue;
          
          const { r, c } = XLSX.utils.decode_cell(addr);
          if (r > templateRow) {
            cellsToMove.set(addr, cell as XLSX.CellObject);
            delete sheet[addr];
          }
        }
        
        // Move cells down
        for (const [addr, cell] of cellsToMove) {
          const { r, c } = XLSX.utils.decode_cell(addr);
          const newAddr = XLSX.utils.encode_cell({ r: r + rowsToInsert, c });
          sheet[newAddr] = cell;
        }
        
        totalRowsInserted += rowsToInsert;
      }
      
      // Get cells from the template row that should be copied (formulas and styles)
      const templateCells = new Map<number, XLSX.CellObject>();
      for (let c = range.s.c; c <= range.e.c; c++) {
        const addr = XLSX.utils.encode_cell({ r: templateRow, c });
        const cell = sheet[addr];
        if ((cell?.f || cell?.s) && !columns.has(c)) {
          templateCells.set(c, cell);
        }
      }
      
      // Get styles from data cells in template row
      const dataStyles = new Map<number, any>();
      for (const [col] of columns) {
        const addr = XLSX.utils.encode_cell({ r: templateRow, c: col });
        const cell = sheet[addr];
        if (cell?.s) {
          dataStyles.set(col, cell.s);
        }
      }
      
      // Fill in the data and copy formulas/styles
      for (let i = 0; i < results.length; i++) {
        const row = templateRow + i;
        
        // Fill in query data with styles
        for (const [col, columnName] of columns) {
          const value = results[i][columnName];
          const addr = XLSX.utils.encode_cell({ r: row, c: col });
          const style = dataStyles.get(col);
          sheet[addr] = { 
            t: typeof value === 'number' ? 'n' : 's', 
            v: value,
            ...(style ? { s: style } : {})
          };
        }
        
        // Copy formulas and styles
        if (i > 0) { // Skip template row
          for (const [col, cell] of templateCells) {
            const newAddr = XLSX.utils.encode_cell({ r: row, c: col });
            const newCell: XLSX.CellObject = { t: 's' }; // Default to string type
            
            // Copy formula if present
            if (cell.f) {
              const updatedFormula = cell.f.replace(/([A-Z]+)(\d+)/g, (_, col, rowNum) => {
                return col + (parseInt(rowNum) + i);
              });
              newCell.t = 'n';
              newCell.f = updatedFormula;
            }
            
            // Copy style if present
            if (cell.s) {
              newCell.s = cell.s;
            }
            
            sheet[newAddr] = newCell;
          }
        }
      }
    }
    
    // Update sheet range
    if (totalRowsInserted > 0) {
      range.e.r += totalRowsInserted;
      sheet['!ref'] = XLSX.utils.encode_range(range);
    }
  }
  
  return workbook;
}