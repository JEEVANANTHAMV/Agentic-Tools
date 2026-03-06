from pydantic import BaseModel
from typing import Optional
from datetime import datetime

class PDFRequest(BaseModel):
    content: str
    filename: Optional[str] = None

class PDFResponse(BaseModel):
    status: str
    message: str
    filename: str
    object_name: str
    download_url: str
    created_at: datetime
