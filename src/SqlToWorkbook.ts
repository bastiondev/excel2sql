import { Workbook, Worksheet, Row, Cell } from 'exceljs';

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
export async function sqlToWorkbook(template: Workbook, queryResults: QueryResults): Promise<Workbook> {
  // Process each sheet in the template
  for (const sheet of template.worksheets) {
    // First pass: collect all cells and their query references
    const cellRefs = new Map<string, { queryName: string; index?: number; columnName: string }>();
    const iterativeRows = new Map<number, { queryName: string; columns: Map<number, string> }>();
    
    // First pass: scan for query references
    sheet.eachRow((row, rowNumber) => {
      row.eachCell((cell, colNumber) => {
        const value = cell.value;
        if (value && typeof value === 'string') {
          QUERY_REFERENCE_REGEX.lastIndex = 0;
          const match = QUERY_REFERENCE_REGEX.exec(value);
          
          if (match) {
            const [, queryName, index, columnName] = match;
            
            if (!index) {
              // Iterative reference
              if (!iterativeRows.has(rowNumber)) {
                iterativeRows.set(rowNumber, { queryName, columns: new Map() });
              }
              iterativeRows.get(rowNumber)!.columns.set(colNumber, columnName);
            } else {
              // Direct reference
              cellRefs.set(`${rowNumber},${colNumber}`, { queryName, index: +index, columnName });
            }
          }
        }
      });
    });
    
    // Process direct references first
    for (const [cellRef, { queryName, index, columnName }] of cellRefs) {
      const [row, col] = cellRef.split(',').map(Number);
      const cell = sheet.getRow(row).getCell(col);
      const results = queryResults[queryName];
      
      if (results && index !== undefined && results[index]) {
        const value = results[index][columnName];
        if (value !== undefined) {
          // Set value directly, type will be inferred
          cell.value = value;
        }
      }
    }
    
    // Process iterative references
    const sortedRows = Array.from(iterativeRows.entries())
      .sort(([a], [b]) => a - b);
    
    for (const [templateRowNum, { queryName, columns }] of sortedRows) {
      const results = queryResults[queryName];
      if (!results?.length) continue;
      
      const templateRow = sheet.getRow(templateRowNum);
      
      // Insert rows for the results (except the template row)
      if (results.length > 1) {
        sheet.spliceRows(templateRowNum + 1, 0, ...Array(results.length - 1).fill(null));
      }
      
      // For each result row
      for (let i = 0; i < results.length; i++) {
        const targetRowNum = templateRowNum + i;
        const targetRow = sheet.getRow(targetRowNum);

        // First set the height of the row to the first row's height if it's not the first:
        if (i > 0) {
          targetRow.height = templateRow.height;
        }
        
        // Copy all cells from template row first
        templateRow.eachCell((cell, colNumber) => {
          const targetCell = targetRow.getCell(colNumber);
          
          if (!columns.has(colNumber)) {
            if (cell.formula) {
              // Update row references in formula
              const updatedFormula = cell.formula.replace(/([A-Z]+)(\d+)/g, (_, col, rowNum) => {
                const offset = targetRowNum - templateRowNum;
                return col + (parseInt(rowNum) + offset);
              });
              targetCell.value = { formula: updatedFormula };
            } else {
              targetCell.value = cell.value;
            }
          } else {
            // Fill in query data
            const columnName = columns.get(colNumber)!;
            const value = results[i][columnName];
            if (value !== undefined) {
              // Set value directly, type will be inferred
              targetCell.value = value;
            }
          }
          // console.log(JSON.stringify(cell.style))
          targetCell.style = JSON.parse(JSON.stringify(cell.style));
        });
      }
    }
  }
  
  return template;
}