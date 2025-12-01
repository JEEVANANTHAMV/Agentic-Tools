from pydantic import BaseModel
from typing import Optional
from datetime import datetime

class DocumentRequest(BaseModel):
    content: str
    filename: Optional[str] = None

class DocumentResponse(BaseModel):
    status: str
    message: str
    filename: str
    object_name: str
    download_url: str
    created_at: datetime

class DocumentListResponse(BaseModel):
    documents: list
    count: int