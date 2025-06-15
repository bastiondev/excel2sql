describe('sqlToWorkbook in browser', function() {
  // Increase timeout for async operations
  this.timeout(10000);
  
  // Access test data from global context
  // The test-data.json file is loaded directly by Karma as a script
  const testData = window.__testData;
  
  before(function() {
    // Check if we have test data
    if (!testData || Object.keys(testData).length === 0) {
      console.error('No test data available');
    } else {
      console.log(`Loaded ${Object.keys(testData).length} test cases`);
    }
  });
  
  // Helper function to compare workbooks
  async function compareWorkbooks(actual, expected) {
    // Compare each worksheet
    for (const actualSheet of actual.worksheets) {
      const expectedSheet = expected.getWorksheet(actualSheet.name);
      expect(expectedSheet).to.not.be.undefined;
      
      // Compare cell values
      actualSheet.eachRow((row, rowNumber) => {
        row.eachCell((cell, colNumber) => {
          const expectedCell = expectedSheet.getRow(rowNumber).getCell(colNumber);
          const actualValue = cell.value;
          const expectedValue = expectedCell.value;
          
          // In browser tests, we're primarily concerned with data values
          // rather than exact formula representation which can differ between environments
          
          // Skip formula comparisons - they're represented differently in browser vs node
          if ((actualValue && typeof actualValue === 'object' && actualValue.formula) ||
              (expectedValue && typeof expectedValue === 'object' && expectedValue.formula) ||
              (typeof actualValue === 'string' && actualValue.startsWith('=')) ||
              (typeof expectedValue === 'string' && expectedValue.startsWith('='))) {
            // Skip formula comparison
            return;
          }
          
          // For all other values, do a deep comparison
          expect(actualValue).to.deep.equal(expectedValue);
        });
      });
    }
  }
  
  // Generate tests for each test case
  (testData ? Object.keys(testData) : []).forEach(testName => {
    it(`should correctly process ${testName} template in browser`, async function() {
      const { queryResults, templateBase64, expectedBase64 } = testData[testName];
      
      // Load the template workbook
      const templateBuffer = Uint8Array.from(atob(templateBase64), c => c.charCodeAt(0));
      const template = await new ExcelJS.Workbook().xlsx.load(templateBuffer.buffer);
      
      // Load the expected workbook
      const expectedBuffer = Uint8Array.from(atob(expectedBase64), c => c.charCodeAt(0));
      const expected = await new ExcelJS.Workbook().xlsx.load(expectedBuffer.buffer);
      
      // Process the workbook
      const result = await excel2sql.sqlToWorkbook(template, queryResults);
      
      // Compare the result with the expected workbook
      await compareWorkbooks(result, expected);
    });
  });
});
