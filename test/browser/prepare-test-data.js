/**
 * This script prepares test data for browser tests by:
 * 1. Loading all test cases from the cases directory
 * 2. Loading template workbooks and converting them to base64
 * 3. Loading expected workbooks and converting them to base64
 * 4. Writing all this data to a JSON file that can be loaded in the browser
 */

// Register ts-node to handle TypeScript files
require('ts-node').register({
  transpileOnly: true,
  compilerOptions: {
    module: 'commonjs'
  }
});

const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');

// Define the directory paths
const CASES_DIR = path.join(__dirname, '../cases/toExcel');
const WORKBOOKS_DIR = path.join(__dirname, '../workbooks/toExcel');
const OUTPUT_FILE = path.join(__dirname, 'test-data.json');

async function loadWorkbookAsBase64(filename) {
  const filePath = path.join(WORKBOOKS_DIR, filename);
  const buffer = await fs.promises.readFile(filePath);
  return buffer.toString('base64');
}

async function prepareTestData() {
  // Get all test case files
  const caseFiles = fs.readdirSync(CASES_DIR)
    .filter(file => file.endsWith('.test.ts') || file.endsWith('.test.js'));
  
  const testData = {};
  
  // For each test case file
  for (const caseFile of caseFiles) {
    try {
      // Import the test case
      const testCase = require(path.join(CASES_DIR, caseFile));
      const testName = testCase.testName;
      
      if (!testName) {
        console.warn(`Skipping ${caseFile}: No testName found`);
        continue;
      }
    
    // Load the template and expected workbooks as base64
    const templatePath = `${testName}.template.xlsx`;
    const expectedPath = `${testName}.expected.xlsx`;
    
    const templateBase64 = await loadWorkbookAsBase64(templatePath);
    const expectedBase64 = await loadWorkbookAsBase64(expectedPath);
    
      // Store the test data
      testData[testName] = {
        queryResults: testCase.queryResults,
        templateBase64,
        expectedBase64
      };
    } catch (error) {
      console.error(`Error processing test case ${caseFile}:`, error);
    }
  }
  
  // Write the test data as a JavaScript file that defines a global variable
  const jsContent = `window.__testData = ${JSON.stringify(testData, null, 2)};`;
  await fs.promises.writeFile(OUTPUT_FILE.replace('.json', '.js'), jsContent);
  console.log(`Test data written to ${OUTPUT_FILE.replace('.json', '.js')}`);
  
  // Also write as JSON for reference
  await fs.promises.writeFile(OUTPUT_FILE, JSON.stringify(testData, null, 2));
  console.log(`Test data also written as JSON to ${OUTPUT_FILE}`);
}

prepareTestData().catch(console.error);
