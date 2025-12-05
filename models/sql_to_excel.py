from pydantic import BaseModel
from typing import Optional, List
from datetime import datetime

class SQLQueryRequest(BaseModel):
    query: str
    filename: Optional[str] = None

class SQLQueryResponse(BaseModel):
    status: str
    message: str
    filename: str
    object_name: str
    download_url: str
    created_at: datetime