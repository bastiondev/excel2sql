# Excel2SQL

Excel2SQL is a transformer for populating Excel workbooks from SQL queries and generating SQL queries from Excel workbooks.

## Workbooks to SQL

The workbook to SQL function utilizes as set of templated queries to generate SQL queries from the data in an Excel workbook.  The template refers to cells or ranges in a workbook to generate 1 or more SQL queries.  The syntax is:

* cell: `<SheetName>!CellRow`, e.g. `Sheet1!A1`
* closed range: `<SheetName>!CellRow:CellRow`, e.g. `Sheet1!A1:A10`
* open range: `<SheetName>!CellRow:`, e.g. `Sheet1!A1:`


A templated query query can either refer to single cells or a range of cells.  For single cell reference the template is translated into a single SQL query.  For a range of cells the template is translated into multiple SQL queries, one for each cell in the range.  This means that the workbook's ranges must be the same length for all range references.

### Single Cell Reference Template example

Example workbook:

|   | A  | B  |
|---|----|----|
| 1 | 10 | 20 |
| 2 | 30 | 40 |
| 3 | 50 | 60 |

Example templated query:

```
INSERT INTO table1 (col1, col2, col3) VALUES ('<Sheet1>!A1', '<Sheet1>!B1', '<Sheet1>!A3');
```

Resulting SQL:

```
INSERT INTO table1 (col1, col2, col3) VALUES ('10', '20', '50');
```

### Open Range Templated Query example

Example workbook:

|   | A  | B  |
|---|----|----|
| 1 | 10 | 20 |
| 2 | 30 | 40 |
| 3 | 50 | 60 |

Example templated query:

```
INSERT INTO table1 (col1, col2) VALUES ('<Sheet1>!A1:', '<Sheet1>!B1:');
```

Resulting SQL:

```
INSERT INTO table1 (col1, col2) VALUES ('10', '20');
INSERT INTO table1 (col1, col2) VALUES ('30', '40');
INSERT INTO table1 (col1, col2) VALUES ('50', '60');
```

### Closed Range Templated Query example

Example workbook:

|   | A  | B  |
|---|----|----|
| 1 | 10 | 20 |
| 2 | 30 | 40 |
| 3 | 50 | 60 |

Example templated query:

```
INSERT INTO table1 (col1, col2) VALUES ('<Sheet1>!A1:A2', '<Sheet1>!B1:B2');
```

Resulting SQL:

```
INSERT INTO table1 (col1, col2) VALUES ('10', '20');
INSERT INTO table1 (col1, col2) VALUES ('30', '40');
```

Multi-Sheet Example:

Example workbook:

| Sheet1 | A  | B  |
|--------|----|----|
| 1      | 10 | 20 |
| 2      | 30 | 40 |
| 3      | 50 | 60 |

| Sheet2 | A  | B  |
|--------|----|----|
| 1      | 15 | 25 |
| 2      | 35 | 45 |
| 3      | 55 | 65 |

Example templated query:

```
INSERT INTO table1 (col1, col2) VALUES ('<Sheet1>!A1:', '<Sheet2>!B1:');
```

Resulting SQL:

```
INSERT INTO table1 (col1, col2) VALUES ('10', '25');
INSERT INTO table1 (col1, col2) VALUES ('30', '45');
INSERT INTO table1 (col1, col2) VALUES ('50', '65');
```



