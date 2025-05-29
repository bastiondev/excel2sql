/**
 * Central test runner for workbookToSql tests
 * Loads all test cases from the cases directory and runs them against
 * the corresponding workbooks in the workbooks directory
 */

import * as fs from 'fs';
import * as path from 'path';
import * as XLSX from 'xlsx';
import { workbookToSql } from '../src/WorkbookToSql';

// Define the directory paths
const CASES_DIR = path.join(__dirname, 'cases');
const WORKBOOKS_DIR = path.join(__dirname, 'workbooks');

// Helper function to load a workbook
function loadWorkbook(filename: string): XLSX.WorkBook {
  const filePath = path.join(WORKBOOKS_DIR, filename);
  const fileData = fs.readFileSync(filePath);
  return XLSX.read(fileData, { type: 'buffer' });
}

// Dynamically load and run all test cases
describe('workbookToSql', () => {
  // Get all test case files
  const caseFiles = fs.readdirSync(CASES_DIR)
    .filter(file => file.endsWith('.test.ts') || file.endsWith('.test.js'));
  
  // For each test case file
  for (const caseFile of caseFiles) {
    // Import the test case
    const testCase = require(path.join(CASES_DIR, caseFile));
    const testName = testCase.testName;
    
    // Create a test for this case
    test(`should correctly process ${testName} workbook`, async () => {
      // Load the workbook
      const workbookPath = `${testName}.xlsx`;
      const workbook = loadWorkbook(workbookPath);
      
      // Run workbookToSql with the templates
      const results = workbookToSql(workbook, testCase.templates);
      
      // Compare with expected results
      expect(results).toHaveLength(testCase.expectedResults.length);
      expect(results).toEqual(testCase.expectedResults);
    });
  }
});
