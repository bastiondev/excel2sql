/**
 * Central test runner for sqlToWorkbook tests
 * Loads all test cases from the cases directory and runs them against
 * the corresponding workbooks in the workbooks directory
 */

import * as fs from 'fs';
import * as path from 'path';
import { Workbook, Cell } from 'exceljs';
import { sqlToWorkbook } from '../src/SqlToWorkbook';

// Define the directory paths
const CASES_DIR = path.join(__dirname, 'cases/toExcel');
const WORKBOOKS_DIR = path.join(__dirname, 'workbooks/toExcel');

// Helper function to load a workbook
async function loadWorkbook(filename: string): Promise<Workbook> {
  const filePath = path.join(WORKBOOKS_DIR, filename);
  const workbook = new Workbook();
  await workbook.xlsx.readFile(filePath);
  return workbook;
}

// Helper function to compare workbooks
function compareWorkbooks(actual: Workbook, expected: Workbook): void {
  // Compare sheet names
  expect(actual.worksheets.map(ws => ws.name)).toEqual(expected.worksheets.map(ws => ws.name));
  
  // Compare each sheet
  for (const actualSheet of actual.worksheets) {
    const expectedSheet = expected.getWorksheet(actualSheet.name);
    expect(expectedSheet).toBeTruthy();
    

    
    // Get the range of cells to compare
    const actualRows = actualSheet.getRows(1, actualSheet.rowCount) || [];
    const expectedRows = expectedSheet!.getRows(1, expectedSheet!.rowCount) || [];
    expect(actualRows.length).toEqual(expectedRows.length);
    
    // Compare each row
    for (let r = 0; r < actualRows.length; r++) {
      const actualRow = actualRows[r];
      const expectedRow = expectedRows[r];
      
      // Compare each cell in the row
      for (let c = 1; c <= actualSheet.columnCount; c++) {
        const actualCell = actualRow.getCell(c);
        const expectedCell = expectedRow.getCell(c);
        
        // Compare formulas if present
        if (actualCell.formula || expectedCell.formula) {
          expect(actualCell.formula).toEqual(expectedCell.formula);
        }
        
        // Compare values if no formula
        if (!actualCell.formula && !expectedCell.formula) {
          expect(actualCell.value).toEqual(expectedCell.value);
          expect(actualCell.type).toEqual(expectedCell.type);
        }
        
        // Compare styles
        expect(actualCell.style).toEqual(expectedCell.style);
      }
    }
  }
}

// Dynamically load and run all test cases
describe('sqlToWorkbook', () => {
  // Get all test case files
  const caseFiles = fs.readdirSync(CASES_DIR)
    .filter(file => file.endsWith('.test.ts') || file.endsWith('.test.js'));
  
  // For each test case file
  for (const caseFile of caseFiles) {
    // Import the test case
    const testCase = require(path.join(CASES_DIR, caseFile));
    const testName = testCase.testName;
    
    // Create a test for this case
    test(`should correctly process ${testName} template`, async () => {
      // Load the template and expected workbooks
      const templatePath = `${testName}.template.xlsx`;
      const expectedPath = `${testName}.expected.xlsx`;
      const template = await loadWorkbook(templatePath);
      const expected = await loadWorkbook(expectedPath);
      
      // Run sqlToWorkbook with the query results
      const result = await sqlToWorkbook(template, testCase.queryResults);
      
      // Write debug output
      const debugPath = path.join(WORKBOOKS_DIR, `${testName}.debug.xlsx`);
      await result.xlsx.writeFile(debugPath);
      
      // Compare with expected workbook
      compareWorkbooks(result, expected);
    });
  }
});
