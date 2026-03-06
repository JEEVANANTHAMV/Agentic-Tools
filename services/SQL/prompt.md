# Guidelines for SQL to Excel Converter API

## Tool Name - sql_to_excel

## Basic Structure

The SQL to Excel converter allows you to execute SQL queries and export the results directly to Excel format. It supports various SQL operations and provides formatted Excel output with proper headers and data types.

## Input Format

### SQL Query

Provide your SQL query as a string:

```sql
SELECT id, name, email, created_at
FROM users
WHERE status = 'active'
ORDER BY created_at DESC
```

## Supported SQL Operations

### SELECT Queries

Basic select:
```sql
SELECT * FROM table_name
```

With columns:
```sql
SELECT id, name, email FROM users
```

With aliases:
```sql
SELECT u.name, d.department_name
FROM users u
JOIN departments d ON u.department_id = d.id
```

### WHERE Clauses

```sql
SELECT * FROM products
WHERE price > 100
AND category = 'electronics'
```

### JOIN Operations

```sql
SELECT o.order_id, c.name, o.total
FROM orders o
INNER JOIN customers c ON o.customer_id = c.id
```

### Aggregations

```sql
SELECT department, COUNT(*) as employee_count, AVG(salary) as avg_salary
FROM employees
GROUP BY department
HAVING COUNT(*) > 5
```

### ORDER BY

```sql
SELECT * FROM products
ORDER BY price DESC, name ASC
```

### LIMIT

```sql
SELECT * FROM users
LIMIT 100
```

## Excel Output Options

### Sheet Name

Specify the sheet name for the output:

```
[SHEET:Query Results]
```

### Header Formatting

```
[HEADERS:BOLD]
[HEADERS:COLOR:0000FF]
[HEADERS:BACKGROUND:EEEEEE]
```

### Auto-Size Columns

```
[AUTO_SIZE:enabled]
```

### Freeze Panes

```
[FREEZE:1]  # Freeze first row
[FREEZE:1,1]  # Freeze first row and column
```

## Content Format with Options

The `query` parameter contains the SQL query, and options can be specified:

```sql
SELECT id, name, email, department, salary
FROM employees
WHERE salary > 50000
ORDER BY salary DESC

[SHEET:Employee Report]
[HEADERS:BOLD]
[AUTO_SIZE:enabled]
```

## Multiple Queries

You can execute multiple queries, each creating a separate sheet:

```sql
# Sheet: Active Users
SELECT id, name, email FROM users WHERE status = 'active'

---

# Sheet: User Statistics
SELECT COUNT(*) as total, AVG(age) as avg_age FROM users

---

# Sheet: Recent Activity
SELECT * FROM activity_log ORDER BY timestamp DESC LIMIT 100
```

## Best Practices

1. Use parameterized queries to prevent SQL injection
2. Include LIMIT for large result sets
3. Use meaningful column aliases
4. Handle NULL values appropriately
5. Test queries before exporting
6. Consider performance for complex queries
7. Use indexes for frequently queried columns
8. Validate data types for Excel compatibility

## Example Queries

### Example 1: Simple Select

```sql
SELECT id, first_name, last_name, email, phone
FROM customers
WHERE country = 'USA'
ORDER BY last_name
```

### Example 2: Aggregation Report

```sql
SELECT
    DATE_TRUNC('month', order_date) as month,
    COUNT(*) as order_count,
    SUM(total_amount) as total_revenue,
    AVG(total_amount) as avg_order_value
FROM orders
WHERE order_date >= '2024-01-01'
GROUP BY DATE_TRUNC('month', order_date)
ORDER BY month
```

### Example 3: Join Query

```sql
SELECT
    c.customer_id,
    c.customer_name,
    o.order_id,
    o.order_date,
    o.total_amount,
    p.product_name,
    oi.quantity,
    oi.unit_price
FROM customers c
INNER JOIN orders o ON c.customer_id = o.customer_id
INNER JOIN order_items oi ON o.order_id = oi.order_id
INNER JOIN products p ON oi.product_id = p.product_id
WHERE o.order_date >= '2024-01-01'
ORDER BY o.order_date DESC
```

### Example 4: Multiple Sheets

```sql
# Sheet: Summary
SELECT
    COUNT(*) as total_orders,
    SUM(total_amount) as total_revenue,
    AVG(total_amount) as avg_order
FROM orders

---

# Sheet: Top Products
SELECT
    p.product_name,
    SUM(oi.quantity) as total_sold,
    SUM(oi.quantity * oi.unit_price) as revenue
FROM products p
INNER JOIN order_items oi ON p.product_id = oi.product_id
GROUP BY p.product_id
ORDER BY total_sold DESC
LIMIT 10

---

# Sheet: Monthly Trends
SELECT
    DATE_TRUNC('month', order_date) as month,
    COUNT(*) as orders,
    SUM(total_amount) as revenue
FROM orders
GROUP BY DATE_TRUNC('month', order_date)
ORDER BY month
```

## API Call Format

To execute a SQL query and generate Excel, make a POST request to the endpoint with the following JSON structure:

```json
{
  "query": "Your SQL query string",
  "filename": "query_results.xlsx"
}
```

### Example cURL Request

```bash
curl -X 'POST' \
  'http://101.53.140.44:8002/api/v1/execute-sql-excel' \
  -H 'accept: application/json' \
  -H 'Content-Type: application/json' \
  -d '{
  "query": "SELECT id, name, email, created_at FROM users WHERE status = \"active\" ORDER BY created_at DESC LIMIT 100",
  "filename": "active_users.xlsx"
}'
```

## Tool Call Integration

When integrating this tool into your application, use the following format:

```javascript
{
  "tool_name": "sql_to_excel",
  "parameters": {
    "query": "[Your SQL query string]",
    "filename": "output_filename.xlsx"
  }
}
```

## Error Handling

Common errors and their solutions:

| Error | Solution |
|-------|----------|
| Syntax error | Check SQL syntax and table/column names |
| Permission denied | Verify database user permissions |
| Connection failed | Check database connection settings |
| Query timeout | Add LIMIT or optimize query |
| Data type mismatch | Ensure data types are Excel-compatible |

By following these guidelines, you can effectively execute SQL queries and export results to Excel using the sql_to_excel tool.
