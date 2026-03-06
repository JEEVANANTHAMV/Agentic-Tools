from pydantic import BaseModel, Field
from typing import Optional, Any
from datetime import datetime

class JSONRequest(BaseModel):
    content: str
    json_schema: Optional[Any] = Field(default=None, alias="schema")
    filename: Optional[str] = None

class JSONResponse(BaseModel):
    status: str
    message: str
    filename: str
    object_name: str
    download_url: str
    created_at: datetime
