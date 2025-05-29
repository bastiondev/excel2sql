/**
 * Test case for SimpleStringInt.xlsx workbook
 */

export const testName = 'SimpleStringInt';

export const templates = [
  "INSERT INTO test_table (string_col, int_col) VALUES ('<Sheet1>!A1:', <Sheet1>!B1:);"
];

export const expectedResults = [
  "INSERT INTO test_table (string_col, int_col) VALUES ('One', 1);",
  "INSERT INTO test_table (string_col, int_col) VALUES ('Two', 2);",
  "INSERT INTO test_table (string_col, int_col) VALUES ('Three', 3);"
];
