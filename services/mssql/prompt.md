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
