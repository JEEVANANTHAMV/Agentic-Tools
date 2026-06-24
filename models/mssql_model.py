from pydantic import BaseModel, Field, ConfigDict
from typing import Optional, Any, Dict, List, Literal
from datetime import datetime


# Actions supported by the single /mssql endpoint.
# Kept here so both the service and the router/docs share one source of truth.
MSSQL_ACTIONS = [
    "test_connection",
    "list_databases",
    "list_tables",
    "list_columns",
    "search_columns",
    "execute_query",
]


class MSSQLConnection(BaseModel):
    """
    MS SQL Server connection credentials.

    Credentials are always supplied per-request in the payload (never read from
    .env) so the same endpoint can serve many different databases / users.
    """
    host: str = Field(..., description="SQL Server hostname or IP address")
    port: int = Field(1433, description="SQL Server TCP port (default 1433)")
    user: str = Field(..., description="Login / username")
    password: str = Field(..., description="Login password")
    database: Optional[str] = Field("master", description="Initial database to connect to (defaults to 'master')")
    instance: Optional[str] = Field(None, description="Named instance, e.g. 'SQLEXPRESS' (optional)")
    driver: Optional[str] = Field(None, description="Explicit ODBC driver name. Auto-detected if omitted.")
    encrypt: bool = Field(True, description="Encrypt the connection (ODBC 'Encrypt')")
    trust_server_certificate: bool = Field(True, description="Trust a self-signed server certificate")
    login_timeout: int = Field(30, description="Connection/login timeout in seconds")


class MSSQLActionParams(BaseModel):
    """Per-action parameters. Only the fields relevant to the chosen action are used."""
    model_config = ConfigDict(populate_by_name=True)

    database: Optional[str] = Field(None, description="Database to target (list_tables / list_columns / search_columns / execute_query)")
    db_schema: Optional[str] = Field(None, alias="schema", description="Schema filter, e.g. 'dbo' (optional)")
    table: Optional[str] = Field(None, description="Table name. Required (exact) for list_columns; a substring filter for search_columns")
    query: Optional[str] = Field(None, description="SQL to run (required for execute_query)")
    read_only: bool = Field(False, description="If true, only SELECT/WITH statements are permitted")
    max_rows: int = Field(1000, description="Maximum rows returned by execute_query")
    include_views: bool = Field(True, description="Include views in list_tables output")

    # --- search / pagination (for huge schemas: 2000+ tables / columns) ---
    search: Optional[str] = Field(None, description="Case-insensitive substring match (SQL LIKE) on table/column names")
    data_type: Optional[str] = Field(None, description="Filter by SQL data type, e.g. 'int' (search_columns / list_columns)")
    limit: int = Field(200, description="Max rows per page for list_tables / search_columns (1-5000)")
    offset: int = Field(0, description="Row offset for pagination (list_tables / search_columns)")


class MSSQLRequest(BaseModel):
    """
    Single-endpoint request.

    - mode="list"  -> returns the catalogue of actions and how to use them (no DB call).
    - mode="call"  -> runs `action` against the database described by `connection`.
    """
    mode: Literal["list", "call"] = Field("list", description="'list' to discover actions, 'call' to run one")
    action: Optional[str] = Field(None, description=f"Action to run when mode='call'. One of: {', '.join(MSSQL_ACTIONS)}")
    connection: Optional[MSSQLConnection] = Field(None, description="DB credentials (required when mode='call')")
    params: MSSQLActionParams = Field(default_factory=MSSQLActionParams, description="Action-specific parameters")


class MSSQLResponse(BaseModel):
    """Flexible response that adapts to each action while keeping a stable envelope."""
    status: str
    mode: str
    action: Optional[str] = None
    message: str
    target: Optional[Dict[str, Any]] = None   # connection target (host/port/database) WITHOUT credentials
    data: Optional[Any] = None                # action result payload (rows, databases, tables, columns...)
    meta: Optional[Dict[str, Any]] = None     # execution metadata (driver used, row counts, etc.)
    usage: Optional[Any] = None               # how-to-use catalogue (only for mode='list')
    created_at: datetime
