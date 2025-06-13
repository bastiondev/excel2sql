const XLSX = require('xlsx-style');

// Create template workbook
const template = { SheetNames: [], Sheets: {} };

// Create base sheet
const templateSheet = {
  '!ref': 'A1:E7',
  A1: { v: 'Product Inventory', t: 's' },
  A2: { v: 'ID', t: 's' },
  B2: { v: 'Product Name', t: 's' },
  C2: { v: 'Price', t: 's' },
  D2: { v: 'Stock', t: 's' },
  E2: { v: 'Total Value', t: 's' },
  A3: { v: '?products.id', t: 's' },
  B3: { v: '?products.name', t: 's' },
  C3: { v: '?products.price', t: 's' },
  D3: { v: '?products.stock', t: 's' },
  E3: { f: 'C2*D2', t: 'n' },
  A6: { v: 'Total Inventory Value:', t: 's' },
  B6: { v: '?summary[0].total_value', t: 's' }
};

// Set column widths
templateSheet['!cols'] = [
  { width: 8 },  // ID
  { width: 15 }, // Name
  { width: 10 }, // Price
  { width: 10 }, // Stock
  { width: 12 }  // Total
];

// Define styles
const titleStyle = {
  font: { bold: true, sz: 14 },
  alignment: { horizontal: 'left' }
};

const headerStyle = {
  font: { bold: true, color: { rgb: "FFFFFF" } },
  fill: { fgColor: { rgb: "4472C4" }, patternType: "solid" },
  border: {
    top: { style: "thin", color: { rgb: "000000" } },
    bottom: { style: "thin", color: { rgb: "000000" } },
    left: { style: "thin", color: { rgb: "000000" } },
    right: { style: "thin", color: { rgb: "000000" } }
  },
  alignment: { horizontal: 'center' }
};

const dataBorderStyle = {
  border: {
    top: { style: "thin", color: { rgb: "000000" } },
    bottom: { style: "thin", color: { rgb: "000000" } },
    left: { style: "thin", color: { rgb: "000000" } },
    right: { style: "thin", color: { rgb: "000000" } }
  }
};

const currencyStyle = {
  ...dataBorderStyle,
  numFmt: '"$"#,##0.00'
};

const numberStyle = {
  ...dataBorderStyle,
  alignment: { horizontal: 'right' }
};

// Apply styles to template
templateSheet.A1.s = titleStyle;

// Header row styles
['A2', 'B2', 'C2', 'D2', 'E2'].forEach(cell => {
  templateSheet[cell].s = headerStyle;
});

// Data row styles
templateSheet.A3.s = numberStyle;  // ID
templateSheet.B3.s = dataBorderStyle;  // Name
templateSheet.C3.s = currencyStyle;  // Price
templateSheet.D3.s = numberStyle;  // Stock
templateSheet.E3.s = currencyStyle;  // Total Value

// Add sheet to workbook
template.SheetNames.push('Sheet1');
template.Sheets['Sheet1'] = templateSheet;

// Write template workbook
XLSX.writeFile(template, 'StyledQuery.template.xlsx');

// Create expected workbook
const expected = { SheetNames: [], Sheets: {} };
const expectedSheet = {
  '!ref': 'A1:E7',
  A1: { v: 'Product Inventory', t: 's' },
  A2: { v: 'ID', t: 's' },
  B2: { v: 'Product Name', t: 's' },
  C2: { v: 'Price', t: 's' },
  D2: { v: 'Stock', t: 's' },
  E2: { v: 'Total Value', t: 's' },
  A3: { v: 1, t: 'n' },
  B3: { v: 'Widget', t: 's' },
  C3: { v: 19.99, t: 'n' },
  D3: { v: 150, t: 'n' },
  E3: { f: 'C2*D2', t: 'n' },
  A4: { v: 2, t: 'n' },
  B4: { v: 'Gadget', t: 's' },
  C4: { v: 24.99, t: 'n' },
  D4: { v: 75, t: 'n' },
  E4: { f: 'C3*D3', t: 'n' },
  A5: { v: 3, t: 'n' },
  B5: { v: 'Doohickey', t: 's' },
  C5: { v: 14.99, t: 'n' },
  D5: { v: 200, t: 'n' },
  E5: { f: 'C4*D4', t: 'n' },
  A7: { v: 'Total Inventory Value:', t: 's' },
  B7: { v: 7246.75, t: 'n' }
};

// Copy column widths
expectedSheet['!cols'] = templateSheet['!cols'];

// Apply styles to expected
expectedSheet.A1.s = titleStyle;

// Header row styles
['A2', 'B2', 'C2', 'D2', 'E2'].forEach(cell => {
  expectedSheet[cell].s = headerStyle;
});

// Data row styles (for all data rows)
for (let row = 2; row <= 4; row++) {
  expectedSheet[`A${row}`].s = numberStyle;  // ID
  expectedSheet[`B${row}`].s = dataBorderStyle;  // Name
  expectedSheet[`C${row}`].s = currencyStyle;  // Price
  expectedSheet[`D${row}`].s = numberStyle;  // Stock
  expectedSheet[`E${row}`].s = currencyStyle;  // Total Value
}

// Add sheet to workbook
expected.SheetNames.push('Sheet1');
expected.Sheets['Sheet1'] = expectedSheet;

// Write expected workbook
XLSX.writeFile(expected, 'StyledQuery.expected.xlsx');
