/**
 * Central test runner for sqlToWorkbook tests
 * Loads all test cases from the cases directory and runs them against
 * the corresponding workbooks in the workbooks directory
 */

import * as fs from 'fs';
import * as path from 'path';
import * as XLSX from 'xlsx';
import { sqlToWorkbook } from '../src/SqlToWorkbook';

// Define the directory paths
const CASES_DIR = path.join(__dirname, 'cases/toExcel');
const WORKBOOKS_DIR = path.join(__dirname, 'workbooks/toExcel');

// Helper function to load a workbook
function loadWorkbook(filename: string): XLSX.WorkBook {
  const filePath = path.join(WORKBOOKS_DIR, filename);
  const fileData = fs.readFileSync(filePath);
  return XLSX.read(fileData, { type: 'buffer', cellStyles: true });
}

// Helper function to compare workbooks
function compareWorkbooks(actual: XLSX.WorkBook, expected: XLSX.WorkBook): void {
  // Compare sheets
  expect(Object.keys(actual.Sheets)).toEqual(Object.keys(expected.Sheets));

  // Compare each sheet
  for (const sheetName of Object.keys(actual.Sheets)) {
    const actualSheet = actual.Sheets[sheetName];
    const expectedSheet = expected.Sheets[sheetName];

    // Compare ranges
    expect(actualSheet['!ref']).toBe(expectedSheet['!ref']);

    // Compare column widths if present
    if (actualSheet['!cols'] || expectedSheet['!cols']) {
      expect(actualSheet['!cols']).toEqual(expectedSheet['!cols']);
    }

    // Compare cell values
    const range = XLSX.utils.decode_range(actualSheet['!ref'] || 'A1');
    for (let r = range.s.r; r <= range.e.r; r++) {
      for (let c = range.s.c; c <= range.e.c; c++) {
        const cellAddress = XLSX.utils.encode_cell({ r, c });
        const actualCell = actualSheet[cellAddress];
        const expectedCell = expectedSheet[cellAddress];

        if (!actualCell && !expectedCell) continue;

        // Compare formulas if present
        if (actualCell?.f || expectedCell?.f) {
          expect(actualCell?.f).toEqual(expectedCell?.f);
        }
        
        // Compare values if no formula
        if (!actualCell?.f && !expectedCell?.f) {
          expect(actualCell?.v).toEqual(expectedCell?.v);
          expect(actualCell?.t).toEqual(expectedCell?.t);
        }
        
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
      const template = loadWorkbook(templatePath);
      const expected = loadWorkbook(expectedPath);
      
      // Run sqlToWorkbook with the query results
      const result = sqlToWorkbook(template, testCase.queryResults);
      
      // Write debug output
      const debugPath = path.join(WORKBOOKS_DIR, `${testName}.debug.xlsx`);
      XLSX.writeFile(result, debugPath, {cellStyles: true});
      
      // Compare with expected workbook
      compareWorkbooks(result, expected);
    });
  }
});
