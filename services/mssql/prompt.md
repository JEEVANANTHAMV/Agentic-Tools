# Guidelines for the MS SQL Connector API

## Tool Name - mssql_connector

A single endpoint to connect to **any** Microsoft SQL Server database and run
actions against it. Database credentials are always supplied **in the request
payload** (never from `.env`), so the same endpoint can serve many different
databases and users at once.

## Endpoint

```
POST /api/v1/mssql
```

The endpoint has **two modes**:

| mode   | purpose                                                                 |
|--------|-------------------------------------------------------------------------|
| `list` | Discover the available actions and how to call them (no DB connection). |
| `call` | Run one `action` against the database described by `connection`.        |

### 1. Discover actions (how to use)

```json
{ "mode": "list" }
```

Returns the catalogue of actions, their required/optional parameters, and a
ready-to-use example payload for each. Call this first if you are unsure how to
build a request.

### 2. Call an action

```json
{
  "mode": "call",
  "action": "<action_name>",
  "connection": { ...credentials... },
  "params": { ...action-specific... }
}
```

## Connection object

Supplied on every `call`. Credentials are per-request — nothing is read from the
environment.

| field                      | required | default  | description                                  |
|----------------------------|----------|----------|----------------------------------------------|
| `host`                     | yes      | —        | Server hostname or IP                        |
| `port`                     | no       | `1433`   | TCP port                                     |
| `user`                     | yes      | —        | Login / username                             |
| `password`                 | yes      | —        | Login password                               |
| `database`                 | no       | `master` | Initial database to connect to               |
| `instance`                 | no       | —        | Named instance, e.g. `SQLEXPRESS`            |
| `driver`                   | no       | auto     | Explicit ODBC driver name (auto-detected)    |
| `encrypt`                  | no       | `true`   | Encrypt the connection                       |
| `trust_server_certificate` | no       | `true`   | Trust a self-signed server certificate       |
| `login_timeout`            | no       | `30`     | Connection timeout (seconds)                 |

## Actions

### `test_connection`
Verify credentials and return the SQL Server version.
- Required: none · Optional: `params.database`

### `list_databases`
List every database on the server (each flagged `is_system`).
- Required: none

### `list_tables`
List tables (and views) inside a database. Filtering and paging run **in SQL Server**,
so it scales to thousands of tables.
- Required: `params.database`
- Optional: `params.schema`, `params.include_views`, `params.search` (name substring),
  `params.limit` (default 200, max 5000), `params.offset`
- Response `meta` includes `total_matched` and `has_more` for paging.

### `list_columns`
List columns of **one** table (type, length, nullability, default).
- Required: `params.database`, `params.table`
- Optional: `params.schema`, `params.search` (column-name substring), `params.data_type`

### `search_columns`
Search columns **across every table** in a database — the "grep a column name" case for
huge schemas. Filtering/paging happen in SQL, so a DB with millions of columns is never
fully transferred.
- Required: `params.database`
- Optional: `params.search` (column substring), `params.table` (table substring),
  `params.data_type`, `params.schema`, `params.limit`, `params.offset`
- Response `meta` includes `total_matched` and `has_more`.

### `execute_query`
Run **any** SQL. SELECT-style statements return `columns` + `rows`; write
statements (`INSERT` / `UPDATE` / `DELETE` / DDL) are committed and return
`affected_rows`.
- Required: `params.query`
- Optional: `params.database`, `params.read_only` (block writes), `params.max_rows` (default 1000)

## Example requests

### List databases

```bash
curl -X POST 'http://101.53.140.44:8002/api/v1/mssql' \
  -H 'Content-Type: application/json' \
  -d '{
    "mode": "call",
    "action": "list_databases",
    "connection": { "host": "10.0.0.5", "user": "sa", "password": "Pass@word1" }
  }'
```

### List tables inside a database

```bash
curl -X POST 'http://101.53.140.44:8002/api/v1/mssql' \
  -H 'Content-Type: application/json' \
  -d '{
    "mode": "call",
    "action": "list_tables",
    "connection": { "host": "10.0.0.5", "user": "sa", "password": "Pass@word1" },
    "params": { "database": "Northwind", "schema": "dbo" }
  }'
```

### Search a huge schema (filter + paginate in SQL)

Find tables whose name contains `order`, 200 at a time:

```bash
curl -X POST 'http://101.53.140.44:8002/api/v1/mssql' \
  -H 'Content-Type: application/json' \
  -d '{
    "mode": "call",
    "action": "list_tables",
    "connection": { "host": "10.0.0.5", "user": "sa", "password": "Pass@word1" },
    "params": { "database": "Northwind", "search": "order", "limit": 200, "offset": 0 }
  }'
```

Find every column named like `customer_id` across **all** tables in the database:

```bash
curl -X POST 'http://101.53.140.44:8002/api/v1/mssql' \
  -H 'Content-Type: application/json' \
  -d '{
    "mode": "call",
    "action": "search_columns",
    "connection": { "host": "10.0.0.5", "user": "sa", "password": "Pass@word1" },
    "params": { "database": "Northwind", "search": "customer_id", "limit": 200, "offset": 0 }
  }'
```

### Read-only query

```bash
curl -X POST 'http://101.53.140.44:8002/api/v1/mssql' \
  -H 'Content-Type: application/json' \
  -d '{
    "mode": "call",
    "action": "execute_query",
    "connection": { "host": "10.0.0.5", "user": "sa", "password": "Pass@word1" },
    "params": {
      "database": "Northwind",
      "query": "SELECT TOP 10 * FROM dbo.Customers",
      "read_only": true
    }
  }'
```

### Write query

```bash
curl -X POST 'http://101.53.140.44:8002/api/v1/mssql' \
  -H 'Content-Type: application/json' \
  -d '{
    "mode": "call",
    "action": "execute_query",
    "connection": { "host": "10.0.0.5", "user": "sa", "password": "Pass@word1" },
    "params": {
      "database": "Northwind",
      "query": "UPDATE dbo.Customers SET City = '\''Berlin'\'' WHERE CustomerID = '\''ALFKI'\''"
    }
  }'
```

## Response envelope

```json
{
  "status": "success | error",
  "mode": "call",
  "action": "list_tables",
  "message": "Human readable summary",
  "target":  { "host": "10.0.0.5", "port": 1433, "database": "Northwind" },
  "data":    { "...action specific result..." },
  "meta":    { "driver": "pyodbc (ODBC Driver 18 for SQL Server)", "count": 42 },
  "usage":   null,
  "created_at": "2026-06-24T10:00:00"
}
```

On failure `status` is `"error"` and `message` explains what went wrong (bad
credentials, unreachable host, SQL syntax error, read-only violation, etc.).

## Driver / connectivity notes

- The service uses **pyodbc** when a Microsoft *ODBC Driver for SQL Server* is
  installed on the host (auto-detected; or set `connection.driver` explicitly).
- If no ODBC driver is present, it automatically falls back to **pymssql**,
  whose pip wheel bundles FreeTDS — so `pip install -r requirements.txt` is
  enough to connect on a fresh clone with no system packages.
- For the best experience install the Microsoft ODBC Driver 18 for SQL Server.

## Working with very large schemas (2000+ tables / columns)

Don't fetch everything and grep/jq it client-side — that still ships the whole
schema across the wire first. Instead, push the search **into SQL Server**:

- `list_tables` and `search_columns` accept `params.search` (a `LIKE '%...%'`
  substring match) plus `params.limit` / `params.offset`. The database does the
  filtering; only matches come back.
- Every paged response returns `meta.total_matched` and `meta.has_more`, so you
  can loop pages (`offset += limit`) until `has_more` is `false`.
- `search_columns` searches column names across **all** tables at once — use it
  instead of calling `list_columns` for every table.
- For anything beyond substring matching (regex, joins on catalog views,
  `sys.*` metadata), use `execute_query`. The catalog views are the
  fully-flexible "grep" of the schema, e.g.:

  ```sql
  SELECT TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, DATA_TYPE
  FROM INFORMATION_SCHEMA.COLUMNS
  WHERE COLUMN_NAME LIKE '%amount%' AND DATA_TYPE IN ('decimal','money')
  ORDER BY TABLE_NAME
  ```

  The structured JSON responses are also easy to post-process with `jq` if you
  want to filter a page further on the client.

## Playbook: how to explore a huge database (smart, structured workflow)

When the database has thousands of tables you will **never** "list everything".
Work top-down, cheapest call first, narrowing at each step. Follow this order:

```
list_databases                      -> which database?
   |
   v
search_columns (by concept)         -> WHICH tables hold the data I need?
   |   e.g. search "email", "amount", "customer_id"
   v
list_tables (search by keyword)     -> confirm the candidate table names
   |   e.g. search "invoice", "order"
   v
list_columns (one candidate table)  -> exact column names + data types
   |
   v
profile the table (row count, PK, FK, indexes via execute_query)
   |
   v
execute_query (read_only=true, TOP N) -> pull a small, precise sample
```

Key idea: **you usually know the *concept* (a column like `customer_id`,
`amount`, `email`) but not the table.** So lead with `search_columns`, not
`list_tables`. It pinpoints the few tables worth inspecting instead of scrolling
2000 names.

### Golden rules
1. **Never `SELECT *` without a row cap.** Always `SELECT TOP 100 ...`, and keep
   `params.read_only: true` while exploring so you can't change data by accident.
2. **Filter on the server, page the results.** Use `params.search` + `limit` /
   `offset`; loop until `meta.has_more` is `false`. Don't pull all rows to filter
   locally.
3. **Find the *important* tables first** — sort by row count (recipe below), not
   alphabetically. The big tables are usually the ones you want.
4. **Discover joins before writing them** — read the foreign keys (recipe below)
   so you join on the right columns instead of guessing.
5. **Check columns/types before querying** so your `WHERE` clause matches the
   real data type (string vs. int vs. date).

### Recipe catalog (run each via `execute_query`, `read_only: true`)

These query SQL Server's catalog views and are cheap even on massive schemas.

**A. Biggest tables first (row counts, no full scan):**
```sql
SELECT TOP 50 s.name AS [schema], t.name AS [table], SUM(p.rows) AS row_count
FROM sys.tables t
JOIN sys.schemas s ON t.schema_id = s.schema_id
JOIN sys.partitions p ON t.object_id = p.object_id AND p.index_id IN (0,1)
GROUP BY s.name, t.name
ORDER BY row_count DESC
```

**B. Which tables contain a concept (column name search across the DB):**
```sql
SELECT TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, DATA_TYPE
FROM INFORMATION_SCHEMA.COLUMNS
WHERE COLUMN_NAME LIKE '%customer%'
ORDER BY TABLE_NAME
```
(or just use the `search_columns` action, which does exactly this with paging.)

**C. How does this table join to others (foreign keys both directions):**
```sql
SELECT fk.name AS fk_name,
       tp.name AS from_table, cp.name AS from_column,
       tr.name AS to_table,   cr.name AS to_column
FROM sys.foreign_keys fk
JOIN sys.foreign_key_columns fkc ON fk.object_id = fkc.constraint_object_id
JOIN sys.tables  tp ON fkc.parent_object_id     = tp.object_id
JOIN sys.columns cp ON fkc.parent_object_id     = cp.object_id AND fkc.parent_column_id     = cp.column_id
JOIN sys.tables  tr ON fkc.referenced_object_id = tr.object_id
JOIN sys.columns cr ON fkc.referenced_object_id = cr.object_id AND fkc.referenced_column_id = cr.column_id
WHERE tp.name = 'Orders' OR tr.name = 'Orders'
```

**D. Primary key columns of a table:**
```sql
SELECT kcu.COLUMN_NAME
FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS tc
JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE kcu ON tc.CONSTRAINT_NAME = kcu.CONSTRAINT_NAME
WHERE tc.CONSTRAINT_TYPE = 'PRIMARY KEY' AND tc.TABLE_NAME = 'Orders'
ORDER BY kcu.ORDINAL_POSITION
```

**E. Indexes on a table (so your WHERE/JOIN hits an index):**
```sql
SELECT i.name AS index_name, i.type_desc, c.name AS column_name, ic.key_ordinal
FROM sys.indexes i
JOIN sys.index_columns ic ON i.object_id = ic.object_id AND i.index_id = ic.index_id
JOIN sys.columns c ON ic.object_id = c.object_id AND ic.column_id = c.column_id
WHERE i.object_id = OBJECT_ID('dbo.Orders')
ORDER BY i.name, ic.key_ordinal
```

**F. Safe sample of a table's data:**
```sql
SELECT TOP 10 * FROM dbo.Orders
```

**G. Tables matching a keyword, ranked by size (B + A combined):**
```sql
SELECT TOP 50 s.name AS [schema], t.name AS [table], SUM(p.rows) AS row_count
FROM sys.tables t
JOIN sys.schemas s ON t.schema_id = s.schema_id
JOIN sys.partitions p ON t.object_id = p.object_id AND p.index_id IN (0,1)
WHERE t.name LIKE '%order%'
GROUP BY s.name, t.name
ORDER BY row_count DESC
```

### Worked example — "get the 10 most recent orders for a customer email"

1. `search_columns` `{ "database": "Sales", "search": "email" }`
   → finds `Sales.dbo.Customers.Email`.
2. `execute_query` recipe **C** with `... WHERE tp.name='Customers' OR tr.name='Customers'`
   → finds `Orders.CustomerId -> Customers.Id`.
3. `list_columns` `{ "database": "Sales", "table": "Orders" }`
   → confirms `OrderDate`, `CustomerId`.
4. `execute_query` (read_only):
   ```sql
   SELECT TOP 10 o.OrderId, o.OrderDate, o.Total
   FROM dbo.Orders o
   JOIN dbo.Customers c ON o.CustomerId = c.Id
   WHERE c.Email = 'jane@example.com'
   ORDER BY o.OrderDate DESC
   ```

Each step is a small, bounded call — you never load the whole schema or table.

## Read-only vs. write

- `params.read_only = true` permits only `SELECT` / `WITH` statements; anything
  else is rejected before it reaches the database.
- `params.read_only = false` (default) allows full read **and** write access,
  committing writes automatically.

## Error handling

| Error                       | Solution                                                |
|-----------------------------|---------------------------------------------------------|
| Login failed for user       | Check `user` / `password`                               |
| Cannot open database        | Check `params.database` exists and the user has access  |
| Could not connect / timeout | Check `host`, `port`, firewall, and `encrypt` settings  |
| read_only violation         | Remove the write statement or set `read_only: false`    |
| No usable MS SQL driver     | `pip install -r requirements.txt` (installs pymssql)    |
