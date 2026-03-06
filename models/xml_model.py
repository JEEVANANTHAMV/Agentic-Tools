from pydantic import BaseModel
from typing import Optional
from datetime import datetime

class XMLRequest(BaseModel):
    content: str
    transform: Optional[str] = None
    filename: Optional[str] = None

class XMLResponse(BaseModel):
    status: str
    message: str
    filename: str
    object_name: str
    download_url: str
    created_at: datetime
