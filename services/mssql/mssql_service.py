"""
MS SQL Server connector service.

Credentials arrive per-request (in the API payload) so a single endpoint can
serve many databases / users. Connection is resilient:

  1. pyodbc  (Microsoft recommended)  -- used when an ODBC Driver for SQL Server
     is installed and importable. The driver is auto-detected unless one is
     explicitly supplied in the payload.
  2. pymssql (FreeTDS, bundled in the pip wheel) -- automatic fallback when no
     ODBC driver is available, so `pip install -r requirements.txt` alone is
     enough to talk to SQL Server on a fresh clone.
"""

import re
import uuid
import urllib.parse
import datetime as _dt
import decimal
from typing import Any, Dict, List, Optional, Tuple

from sqlalchemy import create_engine, text
from sqlalchemy.engine import Engine

from models.mssql_model import MSSQLConnection, MSSQLActionParams, MSSQL_ACTIONS


class MSSQLError(Exception):
    """Raised for connection / execution failures (returned to the caller as status='error')."""


class ReadOnlyViolation(Exception):
    """Raised when a non-SELECT statement is submitted while read_only=True."""


# ODBC drivers we will try (best first) when none is explicitly requested.
PREFERRED_ODBC_DRIVERS = [
    "ODBC Driver 18 for SQL Server",
    "ODBC Driver 17 for SQL Server",
    "ODBC Driver 13 for SQL Server",
    "ODBC Driver 13.1 for SQL Server",
    "ODBC Driver 11 for SQL Server",
    "SQL Server Native Client 11.0",
    "SQL Server",
]

# Statement keywords allowed when read_only=True.
_READ_ONLY_ALLOWED = {"select", "with"}


class MSSQLService:
    # ---------------------------------------------------------------- engine

    def _resolve_odbc_driver(self, requested: Optional[str]) -> Optional[str]:
        """Pick an installed ODBC driver. Returns None if pyodbc isn't usable."""
        try:
            import pyodbc  # noqa: F401
        except Exception:
            return None

        try:
            available = [d for d in pyodbc.drivers()]
        except Exception:
            available = []

        if requested:
            for d in available:
                if d.lower() == requested.lower():
                    return d
            # Honour an explicit request even if not detected; surfaces a clear error.
            return requested

        for pref in PREFERRED_ODBC_DRIVERS:
            for d in available:
                if d.lower() == pref.lower():
                    return d
        for d in available:
            if "sql server" in d.lower():
                return d
        return None

    def _pymssql_available(self) -> bool:
        try:
            import pymssql  # noqa: F401
            return True
        except Exception:
            return False

    def _build_engine(self, conn: MSSQLConnection, database: Optional[str] = None) -> Tuple[Engine, str]:
        """Build a short-lived SQLAlchemy engine for the requested database."""
        db = database or conn.database or "master"
        server = f"{conn.host}\\{conn.instance}" if conn.instance else conn.host

        # 1) Preferred path: pyodbc with a detected/explicit driver.
        driver = self._resolve_odbc_driver(conn.driver)
        if driver:
            odbc_parts = {
                "DRIVER": "{" + driver + "}",
                "SERVER": f"{server},{conn.port}",
                "DATABASE": db,
                "UID": conn.user,
                "PWD": conn.password,
                "Encrypt": "yes" if conn.encrypt else "no",
                "TrustServerCertificate": "yes" if conn.trust_server_certificate else "no",
                "Connection Timeout": str(conn.login_timeout),
            }
            connstr = ";".join(f"{k}={v}" for k, v in odbc_parts.items())
            url = "mssql+pyodbc:///?odbc_connect=" + urllib.parse.quote_plus(connstr)
            engine = create_engine(
                url,
                pool_pre_ping=True,
                connect_args={"timeout": conn.login_timeout},
            )
            return engine, f"pyodbc ({driver})"

        # 2) Fallback: pymssql (FreeTDS bundled in the wheel).
        if self._pymssql_available():
            usr = urllib.parse.quote_plus(conn.user)
            pwd = urllib.parse.quote_plus(conn.password)
            host = f"{server}:{conn.port}"
            url = f"mssql+pymssql://{usr}:{pwd}@{host}/{db}"
            engine = create_engine(
                url,
                pool_pre_ping=True,
                # pass timeouts as typed ints (not URL query strings) so the driver gets the right types
                connect_args={"timeout": conn.login_timeout, "login_timeout": conn.login_timeout},
            )
            return engine, "pymssql"

        raise MSSQLError(
            "No usable MS SQL driver found. Install an ODBC Driver for SQL Server "
            "(used via pyodbc) or rely on the bundled 'pymssql' wheel "
            "(pip install -r requirements.txt)."
        )

    # ----------------------------------------------------------- serialisation

    @staticmethod
    def _json_safe(value: Any) -> Any:
        """Convert DB values into JSON-serialisable primitives."""
        if value is None or isinstance(value, (str, int, float, bool)):
            return value
        if isinstance(value, decimal.Decimal):
            # keep precision for large/precise numbers, else float
            f = float(value)
            return f if f == value else str(value)
        if isinstance(value, (_dt.datetime, _dt.date, _dt.time)):
            return value.isoformat()
        if isinstance(value, _dt.timedelta):
            return str(value)
        if isinstance(value, (bytes, bytearray, memoryview)):
            return bytes(value).hex()
        if isinstance(value, uuid.UUID):
            return str(value)
        return str(value)

    @classmethod
    def _rows_to_records(cls, columns: List[str], rows) -> List[Dict[str, Any]]:
        return [{col: cls._json_safe(val) for col, val in zip(columns, row)} for row in rows]

    @staticmethod
    def _clamp_paging(limit: int, offset: int) -> Tuple[int, int]:
        """Clamp paging to safe bounds. Values are ints from the validated payload, so
        they are safe to inline into OFFSET/FETCH (no injection risk)."""
        limit = max(1, min(int(limit), 5000))
        offset = max(0, int(offset))
        return limit, offset

    # ----------------------------------------------------------- read-only guard

    @staticmethod
    def _statements(sql: str) -> List[str]:
        """Split SQL into individual statements, with comments stripped (used for the read-only check)."""
        # strip block comments and line comments for keyword detection only
        cleaned = re.sub(r"/\*.*?\*/", " ", sql, flags=re.S)
        cleaned = re.sub(r"--[^\n]*", " ", cleaned)
        return [s.strip() for s in cleaned.split(";") if s.strip()]

    @classmethod
    def _assert_read_only(cls, sql: str) -> None:
        statements = cls._statements(sql)
        if not statements:
            raise ReadOnlyViolation("No executable statement found.")
        for stmt in statements:
            first = stmt.split(None, 1)[0].lower() if stmt.split() else ""
            if first not in _READ_ONLY_ALLOWED:
                raise ReadOnlyViolation(
                    f"read_only mode is enabled; statement starting with '{first.upper() or '?'}' "
                    f"is not permitted. Only {', '.join(sorted(_READ_ONLY_ALLOWED)).upper()} are allowed."
                )

    # ----------------------------------------------------------------- actions

    def test_connection(self, conn: MSSQLConnection, params: MSSQLActionParams) -> Tuple[Dict[str, Any], Dict[str, Any]]:
        engine, driver_used = self._build_engine(conn, params.database)
        try:
            with engine.connect() as c:
                version = c.execute(text("SELECT @@VERSION")).scalar()
                current_db = c.execute(text("SELECT DB_NAME()")).scalar()
            data = {"connected": True, "server_version": version, "current_database": current_db}
            return data, {"driver": driver_used}
        finally:
            engine.dispose()

    def list_databases(self, conn: MSSQLConnection, params: MSSQLActionParams) -> Tuple[Dict[str, Any], Dict[str, Any]]:
        engine, driver_used = self._build_engine(conn, conn.database or "master")
        sql = text(
            "SELECT name, database_id, "
            "CONVERT(varchar, create_date, 126) AS create_date, "
            "state_desc, "
            "CASE WHEN database_id <= 4 THEN 1 ELSE 0 END AS is_system "
            "FROM sys.databases ORDER BY name"
        )
        try:
            with engine.connect() as c:
                result = c.execute(sql)
                cols = list(result.keys())
                records = self._rows_to_records(cols, result.fetchall())
            for r in records:
                r["is_system"] = bool(r.get("is_system"))
            return {"databases": records}, {"driver": driver_used, "count": len(records)}
        finally:
            engine.dispose()

    def list_tables(self, conn: MSSQLConnection, params: MSSQLActionParams) -> Tuple[Dict[str, Any], Dict[str, Any]]:
        db = params.database or conn.database
        if not db:
            raise MSSQLError("A target 'database' is required for list_tables (set params.database).")

        engine, driver_used = self._build_engine(conn, db)
        limit, offset = self._clamp_paging(params.limit, params.offset)

        # Build a WHERE clause so filtering happens in SQL Server, not after transfer.
        where = ["1=1"]
        binds: Dict[str, Any] = {}
        if not params.include_views:
            where.append("TABLE_TYPE = 'BASE TABLE'")
        if params.db_schema:
            where.append("TABLE_SCHEMA = :schema")
            binds["schema"] = params.db_schema
        if params.search:
            where.append("TABLE_NAME LIKE :search")          # substring "grep" pushed to the DB
            binds["search"] = f"%{params.search}%"
        where_sql = " AND ".join(where)

        count_sql = text(f"SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES WHERE {where_sql}")
        page_sql = text(
            "SELECT TABLE_CATALOG, TABLE_SCHEMA, TABLE_NAME, TABLE_TYPE "
            "FROM INFORMATION_SCHEMA.TABLES "
            f"WHERE {where_sql} "
            "ORDER BY TABLE_SCHEMA, TABLE_NAME "
            f"OFFSET {offset} ROWS FETCH NEXT {limit} ROWS ONLY"
        )
        try:
            with engine.connect() as c:
                total = int(c.execute(count_sql, binds).scalar() or 0)
                result = c.execute(page_sql, binds)
                cols = list(result.keys())
                records = self._rows_to_records(cols, result.fetchall())
            meta = {
                "driver": driver_used,
                "database": db,
                "returned": len(records),
                "total_matched": total,
                "limit": limit,
                "offset": offset,
                "has_more": offset + len(records) < total,
                "search": params.search,
            }
            return {"database": db, "tables": records}, meta
        finally:
            engine.dispose()

    def search_columns(self, conn: MSSQLConnection, params: MSSQLActionParams) -> Tuple[Dict[str, Any], Dict[str, Any]]:
        """Search columns ACROSS every table in a database (the 'grep a column name' use case).

        Filtering and paging are pushed to SQL Server so a schema with millions of
        columns never has to be transferred or held in memory."""
        db = params.database or conn.database
        if not db:
            raise MSSQLError("A target 'database' is required for search_columns (set params.database).")

        engine, driver_used = self._build_engine(conn, db)
        limit, offset = self._clamp_paging(params.limit, params.offset)

        where = ["1=1"]
        binds: Dict[str, Any] = {}
        if params.search:
            where.append("COLUMN_NAME LIKE :search")
            binds["search"] = f"%{params.search}%"
        if params.db_schema:
            where.append("TABLE_SCHEMA = :schema")
            binds["schema"] = params.db_schema
        if params.table:
            where.append("TABLE_NAME LIKE :table")           # substring filter on table name
            binds["table"] = f"%{params.table}%"
        if params.data_type:
            where.append("DATA_TYPE = :dtype")
            binds["dtype"] = params.data_type
        where_sql = " AND ".join(where)

        count_sql = text(f"SELECT COUNT(*) FROM INFORMATION_SCHEMA.COLUMNS WHERE {where_sql}")
        page_sql = text(
            "SELECT TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, ORDINAL_POSITION, "
            "DATA_TYPE, CHARACTER_MAXIMUM_LENGTH, IS_NULLABLE "
            "FROM INFORMATION_SCHEMA.COLUMNS "
            f"WHERE {where_sql} "
            "ORDER BY TABLE_SCHEMA, TABLE_NAME, ORDINAL_POSITION "
            f"OFFSET {offset} ROWS FETCH NEXT {limit} ROWS ONLY"
        )
        try:
            with engine.connect() as c:
                total = int(c.execute(count_sql, binds).scalar() or 0)
                result = c.execute(page_sql, binds)
                cols = list(result.keys())
                records = self._rows_to_records(cols, result.fetchall())
            meta = {
                "driver": driver_used,
                "database": db,
                "returned": len(records),
                "total_matched": total,
                "limit": limit,
                "offset": offset,
                "has_more": offset + len(records) < total,
                "search": params.search,
            }
            return {"database": db, "matches": records}, meta
        finally:
            engine.dispose()

    def list_columns(self, conn: MSSQLConnection, params: MSSQLActionParams) -> Tuple[Dict[str, Any], Dict[str, Any]]:
        db = params.database or conn.database
        if not db:
            raise MSSQLError("A target 'database' is required for list_columns (set params.database).")
        if not params.table:
            raise MSSQLError("A 'table' is required for list_columns (set params.table).")

        engine, driver_used = self._build_engine(conn, db)
        schema_filter = " AND TABLE_SCHEMA = :schema" if params.db_schema else ""
        search_filter = " AND COLUMN_NAME LIKE :search" if params.search else ""
        type_filter = " AND DATA_TYPE = :dtype" if params.data_type else ""
        sql = text(
            "SELECT TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, ORDINAL_POSITION, "
            "DATA_TYPE, CHARACTER_MAXIMUM_LENGTH, NUMERIC_PRECISION, NUMERIC_SCALE, "
            "IS_NULLABLE, COLUMN_DEFAULT "
            "FROM INFORMATION_SCHEMA.COLUMNS "
            f"WHERE TABLE_NAME = :table{schema_filter}{search_filter}{type_filter} "
            "ORDER BY ORDINAL_POSITION"
        )
        binds: Dict[str, Any] = {"table": params.table}
        if params.db_schema:
            binds["schema"] = params.db_schema
        if params.search:
            binds["search"] = f"%{params.search}%"
        if params.data_type:
            binds["dtype"] = params.data_type
        try:
            with engine.connect() as c:
                result = c.execute(sql, binds)
                cols = list(result.keys())
                records = self._rows_to_records(cols, result.fetchall())
            # Only treat an empty result as an error when no filters were applied,
            # since a search/type filter matching nothing is a valid empty result.
            if not records and not params.search and not params.data_type:
                raise MSSQLError(
                    f"No columns found for table '{params.table}' in database '{db}'. "
                    "Check the table name (and schema)."
                )
            return {"database": db, "table": params.table, "columns": records}, {
                "driver": driver_used,
                "count": len(records),
            }
        finally:
            engine.dispose()

    def execute_query(self, conn: MSSQLConnection, params: MSSQLActionParams) -> Tuple[Dict[str, Any], Dict[str, Any]]:
        if not params.query or not params.query.strip():
            raise MSSQLError("A 'query' is required for execute_query (set params.query).")
        if params.read_only:
            self._assert_read_only(params.query)

        db = params.database or conn.database or "master"
        engine, driver_used = self._build_engine(conn, db)
        try:
            # begin() opens a transaction and commits on success -> handles writes.
            with engine.begin() as c:
                # exec_driver_sql sends the raw SQL to the DBAPI so colons (e.g. time
                # literals) and other tokens are never treated as bind parameters.
                result = c.exec_driver_sql(params.query)

                if result.returns_rows:
                    columns = list(result.keys())
                    fetched = result.fetchmany(params.max_rows + 1)
                    truncated = len(fetched) > params.max_rows
                    rows = fetched[: params.max_rows]
                    records = self._rows_to_records(columns, rows)
                    data = {
                        "statement_type": "read",
                        "columns": columns,
                        "rows": records,
                        "truncated": truncated,
                    }
                    meta = {
                        "driver": driver_used,
                        "database": db,
                        "row_count": len(records),
                        "max_rows": params.max_rows,
                        "truncated": truncated,
                    }
                    return data, meta

                affected = result.rowcount if result.rowcount is not None else -1
                data = {"statement_type": "write", "affected_rows": affected}
                meta = {"driver": driver_used, "database": db, "affected_rows": affected}
                return data, meta
        finally:
            engine.dispose()

    # ----------------------------------------------------------------- dispatch

    DISPATCH = {
        "test_connection": "test_connection",
        "list_databases": "list_databases",
        "list_tables": "list_tables",
        "list_columns": "list_columns",
        "search_columns": "search_columns",
        "execute_query": "execute_query",
    }

    def run_action(
        self, action: str, conn: MSSQLConnection, params: MSSQLActionParams
    ) -> Tuple[Dict[str, Any], Dict[str, Any]]:
        """Dispatch to an action. Raises MSSQLError / ReadOnlyViolation on failure."""
        method_name = self.DISPATCH.get(action)
        if not method_name:
            raise ValueError(f"Unknown action '{action}'. Available: {', '.join(MSSQL_ACTIONS)}")
        method = getattr(self, method_name)
        try:
            return method(conn, params)
        except (MSSQLError, ReadOnlyViolation, ValueError):
            raise
        except Exception as e:  # wrap driver/SQL errors in a clean message
            raise MSSQLError(self._clean_db_error(e)) from e

    @staticmethod
    def _clean_db_error(exc: Exception) -> str:
        msg = str(exc)
        # SQLAlchemy wraps driver errors with the SQL appended after a newline; trim noise.
        msg = msg.split("\n[SQL:")[0].strip()
        return msg or exc.__class__.__name__

    # ------------------------------------------------------------------- usage

    @staticmethod
    def get_usage() -> Dict[str, Any]:
        """Self-documenting catalogue returned when mode='list'."""
        connection_example = {
            "host": "10.0.0.5",
            "port": 1433,
            "user": "sa",
            "password": "YourStrong!Passw0rd",
            "database": "master",
            "encrypt": True,
            "trust_server_certificate": True,
        }
        actions = [
            {
                "action": "test_connection",
                "description": "Verify credentials and return the SQL Server version.",
                "required_params": [],
                "optional_params": ["database"],
                "example": {"mode": "call", "action": "test_connection", "connection": connection_example},
            },
            {
                "action": "list_databases",
                "description": "List all databases on the server (with a system-database flag).",
                "required_params": [],
                "optional_params": [],
                "example": {"mode": "call", "action": "list_databases", "connection": connection_example},
            },
            {
                "action": "list_tables",
                "description": "List tables (and views) inside a database. Filtering and paging happen in SQL, "
                               "so this scales to schemas with thousands of tables. Use params.search for a "
                               "name substring, and params.limit/params.offset to page; the response meta returns "
                               "total_matched and has_more.",
                "required_params": ["params.database"],
                "optional_params": ["params.schema", "params.include_views", "params.search", "params.limit", "params.offset"],
                "example": {
                    "mode": "call",
                    "action": "list_tables",
                    "connection": connection_example,
                    "params": {"database": "Northwind", "search": "order", "schema": "dbo", "limit": 200, "offset": 0},
                },
            },
            {
                "action": "list_columns",
                "description": "List columns (type, nullability, default) for ONE table. "
                               "Optionally narrow with params.search (column-name substring) or params.data_type.",
                "required_params": ["params.database", "params.table"],
                "optional_params": ["params.schema", "params.search", "params.data_type"],
                "example": {
                    "mode": "call",
                    "action": "list_columns",
                    "connection": connection_example,
                    "params": {"database": "Northwind", "table": "Customers", "schema": "dbo"},
                },
            },
            {
                "action": "search_columns",
                "description": "Search columns ACROSS every table in a database ('grep a column name'). "
                               "Filter by params.search (column substring), params.table (table substring), "
                               "params.data_type, params.schema; page with params.limit/params.offset. "
                               "Built for huge schemas - filtering/paging run in SQL Server.",
                "required_params": ["params.database"],
                "optional_params": ["params.search", "params.table", "params.data_type", "params.schema", "params.limit", "params.offset"],
                "example": {
                    "mode": "call",
                    "action": "search_columns",
                    "connection": connection_example,
                    "params": {"database": "Northwind", "search": "customer_id", "limit": 200, "offset": 0},
                },
            },
            {
                "action": "execute_query",
                "description": "Run any SQL against a database. Reads return rows; writes (INSERT/UPDATE/DELETE/DDL) "
                               "are committed and return affected-row counts. Set params.read_only=true to allow "
                               "only SELECT/WITH.",
                "required_params": ["params.query"],
                "optional_params": ["params.database", "params.read_only", "params.max_rows"],
                "example": {
                    "mode": "call",
                    "action": "execute_query",
                    "connection": connection_example,
                    "params": {
                        "database": "Northwind",
                        "query": "SELECT TOP 10 * FROM dbo.Customers",
                        "read_only": True,
                        "max_rows": 1000,
                    },
                },
            },
        ]
        return {
            "endpoint": "/api/v1/mssql",
            "method": "POST",
            "modes": {
                "list": "Send {\"mode\": \"list\"} to get this catalogue (no DB connection needed).",
                "call": "Send {\"mode\": \"call\", \"action\": \"<action>\", \"connection\": {...}, \"params\": {...}}.",
            },
            "notes": [
                "Credentials are passed per-request in 'connection' (never from .env), so many DBs/users are supported.",
                "Connection uses pyodbc when an ODBC Driver for SQL Server is installed, otherwise the bundled pymssql.",
                "execute_query supports both read-only and write statements; use params.read_only to restrict to reads.",
                "For huge schemas (1000s of tables/columns), filtering and paging are pushed into SQL: use "
                "params.search + params.limit/params.offset on list_tables and search_columns. The response "
                "'meta' returns total_matched and has_more so you can page instead of fetching everything.",
                "For fully custom search, run execute_query against the catalog views, e.g. "
                "SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE COLUMN_NAME LIKE '%amount%' - this is the "
                "unlimited-power equivalent of grep over the schema.",
            ],
            "recommended_workflow": [
                "1. list_databases -> pick the database.",
                "2. search_columns by concept (e.g. 'email', 'amount', 'customer_id') to find WHICH tables hold "
                "the data. You usually know the column concept, not the table - so lead with this, not list_tables.",
                "3. list_tables with params.search to confirm candidate table names (paged).",
                "4. list_columns on a candidate table for exact column names + data types.",
                "5. Profile via execute_query on catalog views: row counts (sys.partitions), primary/foreign keys "
                "(sys.foreign_keys), indexes (sys.indexes) - so you join correctly and hit indexes.",
                "6. execute_query with read_only=true and SELECT TOP N to pull a small, precise sample.",
                "Golden rules: never SELECT * without a TOP cap; filter/page on the server; find the biggest tables "
                "(by row count) and the foreign keys before writing joins. See the prompt.md playbook for copy-paste "
                "catalog-view recipes.",
            ],
            "available_actions": MSSQL_ACTIONS,
            "actions": actions,
        }
