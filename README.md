# Excel2SQL Mapping Language

Excel2SQL is a bi-directional mapping language for transforming data between Excel workbooks and SQL databases. It provides a template-driven approach where:

- **Excel → SQL**: Templated SQL queries reference workbook cells to generate executable SQL statements.
- **SQL → Excel**: Named SQL queries populate workbook cells via reference expressions.

---

## Part 1: Excel → SQL (Workbook to SQL)

Templated queries extract data from Excel cells and embed them into SQL statement templates. The template engine scans for cell references enclosed in angle brackets and substitutes them with the corresponding cell values.

### 1.1 Cell Reference Syntax

All cell references use standard Excel notation wrapped in angle brackets:

| Reference Type       | Syntax                          | Description                                                 |
|----------------------|---------------------------------|-------------------------------------------------------------|
| Single cell          | `<SheetName>!CellRow`           | References exactly one cell                                 |
| Closed row range     | `<SheetName>!CellRow:CellRow`   | References a fixed range of rows (inclusive), same column    |
| Open row range       | `<SheetName>!CellRow:`          | References from the start cell down to the sheet's max row  |
| Closed column range  | `<SheetName>!CellRow\|CellRow`  | References a fixed range of columns (inclusive), same row    |
| Open column range    | `<SheetName>!CellRow\|`         | References from the start cell rightward to the sheet's max column |

- **SheetName** — The name of the worksheet tab. If the sheet name contains a literal `>`, escape it as `\>` (e.g., `<Sales \> 2026>!A1`). No other characters require escaping.
- **CellRow** — A column letter followed by a row number (e.g., `A1`, `B12`, `AA3`).
- **Delimiter determines direction** — `:` (colon) iterates over rows (vertical); `|` (pipe) iterates over columns (horizontal).

**Validation**: The delimiter must match the axis of variation. A colon range requires matching column letters (e.g., `A1:A10`); a pipe range requires matching row numbers (e.g., `A1|D1`). A mismatch (e.g., `A1:D1` or `A1|A10`) is an error.

### 1.2 Template Expansion Rules

1. **Single-cell references only** — The template is expanded into exactly **one** SQL statement. Each `<Sheet>!Cell` placeholder is replaced with the literal cell value.

2. **Row range references present** — When the template contains `:` (colon) ranges, it is expanded into **one SQL statement per row** in the range. On each iteration, all row range references advance to the next row together.

3. **Column range references present** — When the template contains `|` (pipe) ranges, it is expanded into **one SQL statement per column** in the range. On each iteration, all column range references advance to the next column together.

4. **Row and column ranges must not be mixed in a single template.** A template may contain row ranges or column ranges, but not both. (Single-cell references can be mixed with either type.)

5. **All range references in a single template must span the same count; a mismatch is a runtime error.** For row ranges, a closed range defines an explicit row count; an open range extends from the start cell to the sheet's max row, including any blank rows in between. For column ranges, a closed range defines an explicit column count; an open range extends from the start cell to the sheet's max column, including any blank columns. The **max row** (or max column) is a sheet-global value — the highest row (or rightmost column) containing any data on that sheet, regardless of which column (or row) the data is in. Because each sheet determines its own max independently, open ranges referencing different sheets may resolve to different counts; if so, this is a runtime error. Blank cells produce statements with empty/NULL values per rule 8.

6. **Single-cell and range references can be mixed** in one template. The single-cell values remain constant across all expanded statements while the range values iterate.

7. Cell values are substituted as literal strings. If a cell contains a formula, the **computed value** is used, not the formula text. The template author is responsible for quoting and SQL type handling. Implementations must properly escape values or use parameterized queries to prevent SQL injection.

8. **Empty and NULL cells** — An empty cell (zero-length string) is substituted as a zero-length string (`''`). A cell with no value (NULL) produces no value — the placeholder is removed, yielding a bare `NULL` in the SQL output. The template author should account for both cases in their SQL (e.g., using `NULLIF` or conditional logic if the distinction matters).

9. **Any valid SQL statement can be templated** — not just `INSERT`. `UPDATE`, `DELETE`, and `INSERT ... ON CONFLICT ... UPDATE` (upsert) patterns all work. The upsert pattern is particularly powerful: the workbook becomes the source of truth, and each sync idempotently converges the database to match the spreadsheet's current state.

### 1.3 Examples

#### 1.3.1 Single Cell Reference

Given **Sheet1**:

|   | A  | B  |
|---|----|----|
| 1 | 10 | 20 |
| 2 | 30 | 40 |
| 3 | 50 | 60 |

Template:
```sql
INSERT INTO table1 (col1, col2, col3) VALUES ('<Sheet1>!A1', '<Sheet1>!B1', '<Sheet1>!A3');
```

Result:
```sql
INSERT INTO table1 (col1, col2, col3) VALUES ('10', '20', '50');
```

#### 1.3.2 Open Range

Same **Sheet1** as above.

Template:
```sql
INSERT INTO table1 (col1, col2) VALUES ('<Sheet1>!A1:', '<Sheet1>!B1:');
```

Result (one statement per row, from row 1 to the sheet's max row):
```sql
INSERT INTO table1 (col1, col2) VALUES ('10', '20');
INSERT INTO table1 (col1, col2) VALUES ('30', '40');
INSERT INTO table1 (col1, col2) VALUES ('50', '60');
```

#### 1.3.3 Closed Range

Same **Sheet1** as above.

Template:
```sql
INSERT INTO table1 (col1, col2) VALUES ('<Sheet1>!A1:A2', '<Sheet1>!B1:B2');
```

Result (rows 1 through 2 only):
```sql
INSERT INTO table1 (col1, col2) VALUES ('10', '20');
INSERT INTO table1 (col1, col2) VALUES ('30', '40');
```

#### 1.3.4 Multi-Sheet References

**Sheet1**:

|   | A  | B  |
|---|----|----|
| 1 | 10 | 20 |
| 2 | 30 | 40 |
| 3 | 50 | 60 |

**Sheet2**:

|   | A  | B  |
|---|----|----|
| 1 | 15 | 25 |
| 2 | 35 | 45 |
| 3 | 55 | 65 |

Template:
```sql
INSERT INTO table1 (col1, col2) VALUES ('<Sheet1>!A1:', '<Sheet2>!B1:');
```

Result:
```sql
INSERT INTO table1 (col1, col2) VALUES ('10', '25');
INSERT INTO table1 (col1, col2) VALUES ('30', '45');
INSERT INTO table1 (col1, col2) VALUES ('50', '65');
```

#### 1.3.5 Closed Column Range

**Sheet1**:

|   | A    | B    | C    | D    |
|---|------|------|------|------|
| 1 | Jan  | Feb  | Mar  | Apr  |
| 2 | 1000 | 1200 | 950  | 1100 |

Template:
```sql
INSERT INTO monthly_sales (month, amount) VALUES ('<Sheet1>!A1|D1', '<Sheet1>!A2|D2');
```

Result (one statement per column, A through D):
```sql
INSERT INTO monthly_sales (month, amount) VALUES ('Jan', '1000');
INSERT INTO monthly_sales (month, amount) VALUES ('Feb', '1200');
INSERT INTO monthly_sales (month, amount) VALUES ('Mar', '950');
INSERT INTO monthly_sales (month, amount) VALUES ('Apr', '1100');
```

#### 1.3.6 Open Column Range

Same **Sheet1** as above (max column is D).

Template:
```sql
INSERT INTO monthly_sales (month, amount) VALUES ('<Sheet1>!A1|', '<Sheet1>!A2|');
```

Result (identical to the closed range example — open range extends to max column):
```sql
INSERT INTO monthly_sales (month, amount) VALUES ('Jan', '1000');
INSERT INTO monthly_sales (month, amount) VALUES ('Feb', '1200');
INSERT INTO monthly_sales (month, amount) VALUES ('Mar', '950');
INSERT INTO monthly_sales (month, amount) VALUES ('Apr', '1100');
```

#### 1.3.7 Real-World: Employee Roster Management

**Employees**:

|   | A (EmpID) | B (FirstName) | C (LastName) | D (Email)              | E (StartDate) | F (Status) |
|---|-----------|---------------|--------------|------------------------|----------------|------------|
| 1 | E1001     | Jane          | Doe          | jane.doe@acme.com      | 2026-03-15     | active     |
| 2 | E1002     | Carlos        | Reyes        | carlos.reyes@acme.com  | 2026-03-15     | active     |
| 3 | E1003     | Aisha         | Patel        | aisha.patel@acme.com   | 2026-04-01     | active     |

**Assignments**:

|   | A (EmpID) | B (DeptCode) | C (ManagerID) | D (LocationCode) |
|---|-----------|--------------|---------------|------------------|
| 1 | E1001     | ENG          | M200          | NYC              |
| 2 | E1002     | MKT          | M305          | CHI              |
| 3 | E1003     | ENG          | M200          | NYC              |

Template 1 — Upsert employees (conflict on the natural key `emp_id`):
```sql
INSERT INTO employees (emp_id, first_name, last_name, email, start_date, status)
VALUES ('<Employees>!A1:', '<Employees>!B1:', '<Employees>!C1:', '<Employees>!D1:', '<Employees>!E1:', '<Employees>!F1:')
ON CONFLICT (emp_id) DO UPDATE
SET first_name  = EXCLUDED.first_name,
    last_name   = EXCLUDED.last_name,
    email       = EXCLUDED.email,
    start_date  = EXCLUDED.start_date,
    status      = EXCLUDED.status;
```

Result (3 statements — one per row):
```sql
INSERT INTO employees (emp_id, first_name, last_name, email, start_date, status)
VALUES ('E1001', 'Jane', 'Doe', 'jane.doe@acme.com', '2026-03-15', 'active')
ON CONFLICT (emp_id) DO UPDATE
SET first_name  = EXCLUDED.first_name,
    last_name   = EXCLUDED.last_name,
    email       = EXCLUDED.email,
    start_date  = EXCLUDED.start_date,
    status      = EXCLUDED.status;
-- ... (2 more statements for E1002, E1003)
```

Template 2 — Upsert department assignments (conflict on `emp_id`):
```sql
INSERT INTO dept_assignments (emp_id, dept_code, manager_id, location_code)
VALUES ('<Assignments>!A1:', '<Assignments>!B1:', '<Assignments>!C1:', '<Assignments>!D1:')
ON CONFLICT (emp_id) DO UPDATE
SET dept_code     = EXCLUDED.dept_code,
    manager_id    = EXCLUDED.manager_id,
    location_code = EXCLUDED.location_code;
```

Result (3 statements — one per row):
```sql
INSERT INTO dept_assignments (emp_id, dept_code, manager_id, location_code)
VALUES ('E1001', 'ENG', 'M200', 'NYC')
ON CONFLICT (emp_id) DO UPDATE
SET dept_code     = EXCLUDED.dept_code,
    manager_id    = EXCLUDED.manager_id,
    location_code = EXCLUDED.location_code;
-- ... (2 more statements for E1002, E1003)
```

#### 1.3.8 Real-World: Mixed Single-Cell and Range References (Budget Management)

**Config**:

|   | A              | B          |
|---|----------------|------------|
| 1 | FY2026         | Q2         |
| 2 | 2026-04-01     | 2026-06-30 |

**LineItems**:

|   | A (CostCenter) | B (GLCode) | C (Amount)  | D (Description)       |
|---|-----------------|------------|-------------|-----------------------|
| 1 | CC-100          | 5100       | 25000.00    | Cloud hosting         |
| 2 | CC-100          | 5200       | 8500.00     | Software licenses     |
| 3 | CC-200          | 5100       | 12000.00    | Dev environment       |
| 4 | CC-200          | 6100       | 3200.00     | Travel & training     |

Template — single-cell values for the fiscal context, range for line items, upsert on the composite key:
```sql
INSERT INTO budget_entries (fiscal_year, quarter, period_start, period_end, cost_center, gl_code, amount, description)
VALUES ('<Config>!A1', '<Config>!B1', '<Config>!A2', '<Config>!B2',
        '<LineItems>!A1:', '<LineItems>!B1:', '<LineItems>!C1:', '<LineItems>!D1:')
ON CONFLICT (fiscal_year, quarter, cost_center, gl_code) DO UPDATE
SET period_start = EXCLUDED.period_start,
    period_end   = EXCLUDED.period_end,
    amount       = EXCLUDED.amount,
    description  = EXCLUDED.description;
```

Result (4 statements — one per line item, single-cell values repeated in each):
```sql
INSERT INTO budget_entries (fiscal_year, quarter, period_start, period_end, cost_center, gl_code, amount, description)
VALUES ('FY2026', 'Q2', '2026-04-01', '2026-06-30', 'CC-100', '5100', '25000.00', 'Cloud hosting')
ON CONFLICT (fiscal_year, quarter, cost_center, gl_code) DO UPDATE
SET period_start = EXCLUDED.period_start,
    period_end   = EXCLUDED.period_end,
    amount       = EXCLUDED.amount,
    description  = EXCLUDED.description;
-- ... (3 more statements for CC-100/5200, CC-200/5100, CC-200/6100)
```

#### 1.3.9 Real-World: Product Catalog Management

**Products**:

|   | A (SKU)     | B (Name)              | C (Price) | D (EffectiveDate) | E (Category) | F (Active) |
|---|-------------|----------------------|-----------|---------------------|--------------|------------|
| 1 | SKU-44821   | Wireless Mouse       | 29.99     | 2026-04-01          | Peripherals  | true       |
| 2 | SKU-10053   | USB-C Hub            | 14.50     | 2026-04-01          | Peripherals  | true       |
| 3 | SKU-88712   | Standing Desk        | 199.00    | 2026-05-01          | Furniture    | true       |
| 4 | SKU-55190   | Noise-Cancel Headset | 89.95     | 2026-04-01          | Audio        | true       |

Template:
```sql
INSERT INTO products (sku, name, price, effective_date, category, active)
VALUES ('<Products>!A1:', '<Products>!B1:', '<Products>!C1:', '<Products>!D1:', '<Products>!E1:', '<Products>!F1:')
ON CONFLICT (sku) DO UPDATE
SET name           = EXCLUDED.name,
    price          = EXCLUDED.price,
    effective_date = EXCLUDED.effective_date,
    category       = EXCLUDED.category,
    active         = EXCLUDED.active;
```

Result (4 statements — one per product row):
```sql
INSERT INTO products (sku, name, price, effective_date, category, active)
VALUES ('SKU-44821', 'Wireless Mouse', '29.99', '2026-04-01', 'Peripherals', 'true')
ON CONFLICT (sku) DO UPDATE
SET name           = EXCLUDED.name,
    price          = EXCLUDED.price,
    effective_date = EXCLUDED.effective_date,
    category       = EXCLUDED.category,
    active         = EXCLUDED.active;
-- ... (3 more statements for SKU-10053, SKU-88712, SKU-55190)
```

#### 1.3.10 Real-World: Customer Account Management

**Accounts**:

|   | A (AccountID) | B (CompanyName)   | C (ContactEmail)         | D (Tier)   | E (Status)  |
|---|---------------|-------------------|--------------------------|------------|-------------|
| 1 | ACC-9001      | Globex Corp       | admin@globex.com         | enterprise | active      |
| 2 | ACC-9045      | Initech LLC       | billing@initech.com      | standard   | suspended   |
| 3 | ACC-9102      | Umbrella Inc      | ops@umbrella.com         | enterprise | active      |
| 4 | ACC-9200      | Stark Industries  | tony@stark.com           | premium    | active      |

Template:
```sql
INSERT INTO accounts (account_id, company_name, contact_email, tier, status)
VALUES ('<Accounts>!A1:', '<Accounts>!B1:', '<Accounts>!C1:', '<Accounts>!D1:', '<Accounts>!E1:')
ON CONFLICT (account_id) DO UPDATE
SET company_name  = EXCLUDED.company_name,
    contact_email = EXCLUDED.contact_email,
    tier          = EXCLUDED.tier,
    status        = EXCLUDED.status;
```

Result (4 statements — one per account row):
```sql
INSERT INTO accounts (account_id, company_name, contact_email, tier, status)
VALUES ('ACC-9001', 'Globex Corp', 'admin@globex.com', 'enterprise', 'active')
ON CONFLICT (account_id) DO UPDATE
SET company_name  = EXCLUDED.company_name,
    contact_email = EXCLUDED.contact_email,
    tier          = EXCLUDED.tier,
    status        = EXCLUDED.status;
-- ... (3 more statements for ACC-9045, ACC-9102, ACC-9200)
```

#### 1.3.11 Real-World: Monthly Metrics Pivot (Column Iteration)

**Metrics**:

|   | A (Metric)     | B (Jan)  | C (Feb)  | D (Mar)  | E (Apr)  | F (May)  | G (Jun)  |
|---|----------------|----------|----------|----------|----------|----------|----------|
| 1 | month          | 2026-01  | 2026-02  | 2026-03  | 2026-04  | 2026-05  | 2026-06  |
| 2 | revenue        | 42000    | 45000    | 51000    | 48000    | 53000    | 57000    |
| 3 | new_customers  | 120      | 135      | 142      | 128      | 150      | 165      |
| 4 | churn_rate     | 0.03     | 0.025    | 0.028    | 0.031    | 0.022    | 0.019    |

Template — column ranges iterate across months; row 1 provides the month key, rows 2–4 provide the metric values:
```sql
INSERT INTO monthly_kpis (month, revenue, new_customers, churn_rate)
VALUES ('<Metrics>!B1|', '<Metrics>!B2|', '<Metrics>!B3|', '<Metrics>!B4|')
ON CONFLICT (month) DO UPDATE
SET revenue       = EXCLUDED.revenue,
    new_customers = EXCLUDED.new_customers,
    churn_rate    = EXCLUDED.churn_rate;
```

Result (6 statements — one per column, B through G):
```sql
INSERT INTO monthly_kpis (month, revenue, new_customers, churn_rate)
VALUES ('2026-01', '42000', '120', '0.03')
ON CONFLICT (month) DO UPDATE
SET revenue       = EXCLUDED.revenue,
    new_customers = EXCLUDED.new_customers,
    churn_rate    = EXCLUDED.churn_rate;
-- ... (5 more statements for 2026-02 through 2026-06)
```

---

## Part 2: SQL → Excel (SQL to Workbook)

Named SQL queries are executed and their result sets are mapped into Excel cells using reference expressions. The reference syntax uses a `?` prefix to denote a query binding.

### 2.1 Reference Syntax

| Reference Type              | Syntax                   | Description                                                  |
|-----------------------------|--------------------------|--------------------------------------------------------------|
| Iterative (rows)            | `?queryName.column`      | Fills downward — one cell per result row, copying the template row for each row in the result set |
| Iterative (columns)         | `?queryName.column\|`    | Fills rightward — one cell per result row, copying the template column for each row in the result set |
| Direct (indexed)            | `?queryName[n].column`   | Fills a single cell with column value from result row `n` (zero-indexed) |

- **queryName** — An alias assigned to the SQL query in the mapping configuration.
- **column** — A column name from the query's result set.
- **n** — A zero-based row index into the result set.
- **`|` suffix** — The trailing pipe signals columnar (horizontal) expansion, mirroring the `|` delimiter in Excel → SQL column ranges.

> **Note on `?` syntax**: The `?` prefix appears in two distinct contexts. Inside SQL query strings, `?name` denotes a runtime parameter bound at execution time (see [2.3.7 Variablized Queries](#237-variablized-queries)). Inside Excel cell values, `?name.column`, `?name.column|`, or `?name[n].column` denotes a query result reference. The forms are unambiguous: query references always include a `.column` accessor; parameters never do. The two contexts (SQL string vs. Excel cell) never overlap.

### 2.2 Expansion Rules

1. **Iterative row references** — The template row containing `?queryName.column` is replicated once per row in the result set. All iterative references in the same row must come from the same query (so the row count is unambiguous). The new rows are **inserted**, pushing existing rows below them downward.

2. **Iterative column references** — The template column containing `?queryName.column|` is replicated once per row in the result set, expanding rightward. All columnar iterative references in the same column must come from the same query (so the column count is unambiguous). The new columns are **inserted**, pushing existing columns to the right. Row and column iterative references must not be mixed on the same sheet.

3. **Direct references** — Each `?queryName[n].column` is replaced in-place with the specific value. No row or column replication occurs. Direct references can coexist with either row or column iterative references on the same sheet. If the index `n` is out of bounds (i.e., the result set has fewer than `n+1` rows), this is a runtime error.

4. **Empty result sets** — If an iterative query returns zero rows, the template row (or template column, for columnar references) is deleted from the sheet. This ensures the output contains no rows/columns for that region rather than leaving unresolved references in the workbook.

5. **Processing order** — References on a sheet are processed top-to-bottom in row order (or left-to-right in column order for columnar references). Each iterative expansion inserts rows or columns, shifting all content beyond it. Direct references and subsequent iterative regions are resolved at their shifted positions. The template author must account for this when designing sheets with multiple data regions.

6. **Formulas and styles** — When iterative rows are copied down (or columns copied right), Excel formulas have their cell references adjusted relatively (e.g., `=A1+B1` becomes `=A2+B2` on the next row, or `=A1+B1` becomes `=B1+C1` on the next column). Styles (fonts, borders, number formats) are preserved.

7. **Multiple queries** — A mapping can define any number of named queries. Different sheets or different regions of the same sheet can reference different queries.

8. **Variablized queries** — Query SQL can include `?variable` placeholders that are bound at execution time, allowing the same mapping definition to serve different runtime contexts.

9. **Data types, NULL, and empty values** — Query result values are written to Excel cells with their native data types preserved: numbers as numbers, strings as strings, dates as dates, booleans as booleans. When a query returns `NULL` for a column, the corresponding Excel cell is left empty (no value). When a query returns a zero-length string (`''`), the cell is set to a zero-length string. The data round-trips faithfully: the value and type in the cell match exactly what the database returned.

### 2.3 Examples

#### 2.3.1 Iterative Query Reference

Query `data`:
```sql
SELECT id, name FROM table1;
```

Result set:

| id | name  |
|----|-------|
| 1  | name1 |
| 2  | name2 |
| 3  | name3 |

Excel template:

|   | A        | B          |
|---|----------|------------|
| 1 | ?data.id | ?data.name |

Output:

|   | A | B     |
|---|---|-------|
| 1 | 1 | name1 |
| 2 | 2 | name2 |
| 3 | 3 | name3 |

#### 2.3.2 Iterative Column Reference

Query `data`:
```sql
SELECT id, name FROM table1;
```

Result set:

| id | name  |
|----|-------|
| 1  | name1 |
| 2  | name2 |
| 3  | name3 |

Excel template:

|   | A          |
|---|------------|
| 1 | ?data.id|  |
| 2 | ?data.name| |

Output (one column per result row, expanding rightward):

|   | A | B | C |
|---|---|---|---|
| 1 | 1 | 2 | 3 |
| 2 | name1 | name2 | name3 |

#### 2.3.3 Direct (Indexed) Query Reference

Query `data`:
```sql
SELECT id, name FROM table1;
```

Result set:

| id | name  |
|----|-------|
| 1  | name1 |
| 2  | name2 |
| 3  | name3 |

Excel template:

|   | A             | B               |
|---|---------------|-----------------|
| 1 | ?data[0].id   | ?data[0].name   |
| 2 | ?data[2].id   | ?data[2].name   |

Output:

|   | A | B     |
|---|---|-------|
| 1 | 1 | name1 |
| 2 | 3 | name3 |

#### 2.3.4 Multiple Queries

Query `data_one`:
```sql
SELECT id, name FROM table1;
```

Query `data_two`:
```sql
SELECT id, name FROM table2;
```

Excel template:

|   | A                 | B                   |
|---|-------------------|---------------------|
| 1 | ?data_one[0].id   | ?data_one[0].name   |
| 2 | ?data_two[2].id   | ?data_two[2].name   |

Given both queries return their respective data, the output substitutes each reference independently.

#### 2.3.5 Multiple Sheets with Iterative Rows

Query `data_one`:
```sql
SELECT id, name FROM table1;
```

Query `data_two`:
```sql
SELECT id, name FROM table2;
```

**Sheet1** template:

|   | A            | B              |
|---|--------------|----------------|
| 1 | ?data_one.id | ?data_one.name |

**Sheet2** template:

|   | A            | B              |
|---|--------------|----------------|
| 1 | ?data_two.id | ?data_two.name |

Each sheet expands independently based on its respective query's result set.

#### 2.3.6 Iterative Rows with Formulas and Styles

Query `data`:
```sql
SELECT id, name FROM table1;
```

Excel template:

|   | A        | B          | C       |
|---|----------|------------|---------|
| 1 | ?data.id | ?data.name | =A1+B1  |

Output:

|   | A | B     | C       |
|---|---|-------|---------|
| 1 | 1 | name1 | =A1+B1  |
| 2 | 2 | name2 | =A2+B2  |
| 3 | 3 | name3 | =A3+B3  |

Formulas adjust their row references relative to the new row position. Styles on the template row are cloned to each new row.

#### 2.3.7 Variablized Queries

Query `data`:
```sql
SELECT id, name FROM table1 WHERE id = ?id;
```

The `?id` parameter is supplied at execution time. If `?id = 2`, only the matching row is returned and mapped into the template.

#### 2.3.8 Real-World: Sales Report with Summary and Detail

Query `summary`:
```sql
SELECT
    COUNT(*)        AS total_orders,
    SUM(amount)     AS total_revenue,
    AVG(amount)     AS avg_order_value,
    MIN(order_date) AS period_start,
    MAX(order_date) AS period_end
FROM orders
WHERE order_date BETWEEN ?start_date AND ?end_date;
```

Query `detail`:
```sql
SELECT
    o.order_id,
    c.company_name,
    o.order_date,
    o.amount,
    o.status
FROM orders o
JOIN customers c ON c.customer_id = o.customer_id
WHERE o.order_date BETWEEN ?start_date AND ?end_date
ORDER BY o.order_date;
```

**Summary** sheet template:

|   | A                  | B                         |
|---|--------------------|---------------------------|
| 1 | Report Period      |                           |
| 2 | From:              | ?summary[0].period_start  |
| 3 | To:                | ?summary[0].period_end    |
| 4 |                    |                           |
| 5 | Total Orders       | ?summary[0].total_orders  |
| 6 | Total Revenue      | ?summary[0].total_revenue |
| 7 | Avg Order Value    | ?summary[0].avg_order_value |

**Detail** sheet template:

|   | A               | B                    | C                  | D              | E              | F                  |
|---|-----------------|----------------------|--------------------|----------------|----------------|--------------------|
| 1 | Order ID        | Customer             | Date               | Amount         | Status         | Running Total      |
| 2 | ?detail.order_id | ?detail.company_name | ?detail.order_date | ?detail.amount | ?detail.status | =SUM($D$2:D2)     |

**Detail** expands row 2 for each transaction. The `=SUM($D$2:D2)` formula adjusts to `=SUM($D$2:D3)`, `=SUM($D$2:D4)`, etc., producing a running total. Row 1 (the header) is static.

#### 2.3.9 Real-World: Multi-Sheet Inventory Report

Query `inv_nyc`:
```sql
SELECT sku, product_name, qty_on_hand, reorder_point, unit_cost
FROM inventory
WHERE warehouse_code = 'NYC'
ORDER BY product_name;
```

Query `inv_chi`:
```sql
SELECT sku, product_name, qty_on_hand, reorder_point, unit_cost
FROM inventory
WHERE warehouse_code = 'CHI'
ORDER BY product_name;
```

**NYC** sheet template:

|   | A            | B                    | C                   | D                     | E                | F                          |
|---|--------------|----------------------|---------------------|-----------------------|------------------|----------------------------|
| 1 | SKU          | Product              | On Hand             | Reorder Point         | Unit Cost        | Inventory Value            |
| 2 | ?inv_nyc.sku | ?inv_nyc.product_name | ?inv_nyc.qty_on_hand | ?inv_nyc.reorder_point | ?inv_nyc.unit_cost | =C2*E2                   |

**CHI** sheet template:

|   | A            | B                    | C                   | D                     | E                | F                          |
|---|--------------|----------------------|---------------------|-----------------------|------------------|----------------------------|
| 1 | SKU          | Product              | On Hand             | Reorder Point         | Unit Cost        | Inventory Value            |
| 2 | ?inv_chi.sku | ?inv_chi.product_name | ?inv_chi.qty_on_hand | ?inv_chi.reorder_point | ?inv_chi.unit_cost | =C2*E2                   |

Each sheet independently expands its iterative row. The `=C2*E2` formula adjusts to match the new row number (`=C3*E3`, `=C4*E4`, etc.).

#### 2.3.10 Real-World: Mixed Direct and Iterative References on One Sheet

Query `customer`:
```sql
SELECT customer_id, company_name, contact_email, balance_due
FROM customers
WHERE customer_id = ?cust_id;
```

Query `transactions`:
```sql
SELECT txn_date, description, debit, credit
FROM ledger
WHERE customer_id = ?cust_id
ORDER BY txn_date;
```

**Statement** sheet template:

|   | A                         | B                            | C     | D     |
|---|---------------------------|------------------------------|-------|-------|
| 1 | Customer Statement        |                              |       |       |
| 2 | ?customer[0].company_name | ?customer[0].contact_email   |       |       |
| 3 | Balance Due:              | ?customer[0].balance_due     |       |       |
| 4 |                           |                              |       |       |
| 5 | Date                      | Description                  | Debit | Credit |
| 6 | ?transactions.txn_date    | ?transactions.description    | ?transactions.debit | ?transactions.credit |

Rows 1–5 are static (direct references fill in-place). Row 6 is iterative — it expands downward once per result row.

#### 2.3.11 Real-World: Parameterized Query for Dynamic Filtering

Query `team_members`:
```sql
SELECT emp_id, full_name, role, hire_date
FROM employees
WHERE dept_code = ?dept AND location_code = ?location
ORDER BY hire_date;
```

With parameters `?dept = 'ENG'` and `?location = 'NYC'`, the query returns engineering staff in New York.

**Roster** sheet template:

|   | A                   | B                      | C                  | D                      |
|---|---------------------|------------------------|--------------------|------------------------|
| 1 | Employee ID         | Name                   | Role               | Hire Date              |
| 2 | ?team_members.emp_id | ?team_members.full_name | ?team_members.role | ?team_members.hire_date |

#### 2.3.12 Real-World: Processing Order and Footer Rows

Query `line_items`:
```sql
SELECT description, qty, unit_price
FROM invoice_lines
WHERE invoice_id = ?inv_id
ORDER BY line_num;
```

Query `invoice`:
```sql
SELECT invoice_id, customer_name, invoice_date, total_amount, tax, grand_total
FROM invoices
WHERE invoice_id = ?inv_id;
```

**Invoice** sheet template:

|   | A                            | B                         | C                       |
|---|------------------------------|---------------------------|-------------------------|
| 1 | Invoice:                     | ?invoice[0].invoice_id    | ?invoice[0].invoice_date |
| 2 | Customer:                    | ?invoice[0].customer_name |                         |
| 3 |                              |                           |                         |
| 4 | Description                  | Qty                       | Unit Price              |
| 5 | ?line_items.description      | ?line_items.qty           | ?line_items.unit_price  |
| 6 |                              |                           |                         |
| 7 |                              | Subtotal:                 | ?invoice[0].total_amount |
| 8 |                              | Tax:                      | ?invoice[0].tax         |
| 9 |                              | Grand Total:              | ?invoice[0].grand_total |

Processing sequence:
1. Rows 1–2: Direct references resolved in-place (invoice header).
2. Row 5: Iterative region expands — if the query returns 4 rows, rows 5–8 become detail lines. The original rows 6–9 (spacer and footer) shift down by 3 (4 inserted rows minus the 1 template row).
3. Rows 9–12 (shifted): Direct references in the footer resolve at their new positions.

The footer lands immediately below the last detail row regardless of how many rows are inserted — the engine handles shifting during top-to-bottom processing.

#### 2.3.13 Real-World: Monthly KPI Dashboard (Column Iteration)

Query `monthly_kpis`:
```sql
SELECT month_label, revenue, new_customers, churn_rate
FROM monthly_kpis
WHERE fiscal_year = ?fy
ORDER BY month_seq;
```

Result set:

| month_label | revenue | new_customers | churn_rate |
|-------------|---------|---------------|------------|
| Jan         | 42000   | 120           | 0.030      |
| Feb         | 45000   | 135           | 0.025      |
| Mar         | 51000   | 142           | 0.028      |
| Apr         | 48000   | 128           | 0.031      |

**Dashboard** sheet template:

|   | A              | B                           |
|---|----------------|-----------------------------|
| 1 | Month          | ?monthly_kpis.month_label|  |
| 2 | Revenue        | ?monthly_kpis.revenue|      |
| 3 | New Customers  | ?monthly_kpis.new_customers| |
| 4 | Churn Rate     | ?monthly_kpis.churn_rate|   |

Output (column B is the template column; it expands rightward into B–E):

|   | A              | B     | C     | D     | E     |
|---|----------------|-------|-------|-------|-------|
| 1 | Month          | Jan   | Feb   | Mar   | Apr   |
| 2 | Revenue        | 42000 | 45000 | 51000 | 48000 |
| 3 | New Customers  | 120   | 135   | 142   | 128   |
| 4 | Churn Rate     | 0.030 | 0.025 | 0.028 | 0.031 |

Column A (static labels) remains untouched — only the template column (B) is replicated. Formulas or styles on the template column are copied rightward with relative column adjustment.

---

## Appendix: Quick Reference

### Excel → SQL Cell Reference Cheatsheet

| Pattern                       | Example                  | Produces                                  |
|-------------------------------|--------------------------|-------------------------------------------|
| `<Sheet>!A1`                  | `<Config>!A1`            | Single value                              |
| `<Sheet>!A1:A5`              | `<Data>!A1:A5`           | 5 rows (colon = row iteration)            |
| `<Sheet>!A1:`                 | `<Data>!A1:`             | All rows from A1 to the sheet's max row   |
| `<Sheet>!A1\|D1`             | `<Data>!A1\|D1`          | 4 columns (pipe = column iteration)       |
| `<Sheet>!A1\|`               | `<Data>!A1\|`            | All columns from A1 to the sheet's max column |

### SQL → Excel Reference Cheatsheet

| Pattern                       | Example                  | Behavior                                  |
|-------------------------------|--------------------------|-------------------------------------------|
| `?query.column`               | `?data.name`             | Iterative — row per result (downward)     |
| `?query.column\|`             | `?data.name\|`           | Iterative — column per result (rightward) |
| `?query[n].column`            | `?data[0].name`          | Direct — single cell                      |
| `?variable` (in SQL)        | `WHERE id = ?id`         | Parameter binding at runtime |
