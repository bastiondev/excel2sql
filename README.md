# Excel2SQL

Excel2SQL is a TypeScript library for translating Excel workbooks to and from SQL queries. It works in both Node.js and browser environments.

## Installation

```bash
npm install excel2sql
```

## Quick Start

```typescript
import { workbookToSql, sqlToWorkbook } from 'excel2sql';
```

### Workbook → SQL

Generate SQL INSERT statements from an Excel workbook using a templated query:

```typescript
import { Workbook } from 'exceljs';

const workbook = new Workbook();
await workbook.xlsx.readFile('path/to/workbook.xlsx');

const sql = workbookToSql(workbook, [
  "INSERT INTO table1 (col1, col2) VALUES ('<Sheet1>!A1:', '<Sheet1>!B1:');"
]);
```

### SQL → Workbook

Populate an Excel template with data from SQL queries:

```typescript
import { Workbook } from 'exceljs';

const template = new Workbook();
await template.xlsx.readFile('path/to/template.xlsx');

const result = await sqlToWorkbook(template, {
  data: [{ id: 1, name: 'name1' }, { id: 2, name: 'name2' }]
});

await result.xlsx.writeFile('path/to/output.xlsx');
```

---

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


## SQL to Workbooks

The SQL to Workbooks functionality utilizes a set of parameterized queries and references to queried data in the workbooks to populate an Excel workbook with data from the queries.

### Iterative query reference:

The iterative query reference allows for populating a range of cells with data from a query.  The query is executed and each row is copied down for the number of rows in the result.

Example query:

`data`:
```sql
SELECT id, name from table1;
```

Excel Template:

|   | A  | B  |
|---|----|----|
| 1 | ?data.id | ?data.name |

Given this data:

| id | name |
|----|----|
| 1 | name1 |
| 2 | name2 |
| 3 | name3 |

Resulting in:

|   | A  | B  |
|---|----|----|
| 1 | 1  | name1 |
| 2 | 2  | name2 |
| 3 | 3  | name3 |

### Direct query reference:

The direct query reference allows for populating a cell with a specific cell of data from a query.  The query is executed and the result is written to the cell.

Example query:

`data`:
```sql
SELECT id, name from table1;
```
Given this data:

| id | name |
|----|----|
| 1 | name1 |
| 2 | name2 |
| 3 | name3 |

Excel Template:

|   | A  | B  |
|---|----|----|
| 1 | ?data[0].id | ?data[0].name |
| 2 | ?data[2].id | ?data[2].name |

Resulting in:

|   | A  | B  |
|---|----|----|
| 1 | 1  | name1 |
| 2 | 3  | name3 |

### Multiple queries:

Multiple queries can be executed and the results can be referenced in the template.

Example queries:

`data_one`:
```sql
SELECT id, name from table1;
```

`data_two`:
```sql
SELECT id, name from table2;
```

Excel Template:

|   | A  | B  |
|---|----|----|
| 1 | ?data_one[0].id | ?data_one[0].name |
| 2 | ?data_two[2].id | ?data_two[2].name |

Given this data:

`data_one`:
| id | name |
|----|----|
| 1 | name1 |
| 2 | name2 |
| 3 | name3 |

`data_two`:
| id | name |
|----|----|
| 1 | name1 |
| 2 | name2 |
| 3 | name3 |

Resulting in:

|   | A  | B  |
|---|----|----|
| 1 | 1  | name1 |
| 2 | 3  | name3 |

### Multiple sheets with iterative rows:

Multiple sheets can have also reference to iterative rows of data.

Example queries:

`data_one`:
```sql
SELECT id, name from table1;
```

`data_two`:
```sql
SELECT id, name from table2;
```

Excel Template:

Sheet1:
|   | A  | B  |
|---|----|----|
| 1 | ?data_one.id | ?data_one.name |

Sheet2:
|   | A  | B  |
|---|----|----|
| 1 | ?data_two.id | ?data_two.name |

Given this data:

`data_one`:
| id | name |
|----|----|
| 1 | name1 |
| 2 | name2 |
| 3 | name3 |

`data_two`:
| id | name |
|----|----|
| 1 | name2-1 |
| 2 | name2-2 |
| 3 | name2-3 |

Resulting in:

Sheet1:
|   | A  | B  |
|---|----|----|
| 1 | 1  | name1 |
| 2 | 2  | name2 |
| 3 | 3  | name3 |

Sheet2:
|   | A  | B  |
|---|----|----|
| 1 | 1  | name2-1 |
| 2 | 2  | name2-2 |
| 3 | 3  | name2-3 |

### Iterative rows with formulas and styles

When rows are copied down, all styles and formulas are copied down as well.  Query references will be copied down as above, and formulas will be copied down with references copied down relatively.  This is done as an insert operation, so rows below the copied down rows will be pushed down.

Example Query:
`data`:
```sql
SELECT id, name from table1;
```

Excel Template:

|   | A  | B  | C  |
|---|----|----|----|
| 1 | ?data.id | ?data.name | =A1+B1 |

Given this data:

| id | name |
|----|----|
| 1 | name1 |
| 2 | name2 |
| 3 | name3 |

Resulting in:

|   | A  | B  | C  |
|---|----|----|----|
| 1 | 1  | name1 | =A1+B1 |
| 2 | 2  | name2 | =A2+B2 |
| 3 | 3  | name3 | =A3+B3 |


### Variablized Queries

Variables can be used in queries to parameterize the queries.  This allows for dynamic queries that can be used to populate the workbook with data from the queries.

Example Query:
`data`:
```sql
SELECT id, name from table1 where id = ?id;
```

Excel Template:

|   | A  | B  |
|---|----|----|
| 1 | ?data.id | ?data.name |

Given this data:

| id | name |
|----|----|
| 1 | name1 |
| 2 | name2 |
| 3 | name3 |

Resulting in:

|   | A  | B  |
|---|----|----|
| 1 | 1  | name1 |
| 2 | 2  | name2 |
| 3 | 3  | name3 |
