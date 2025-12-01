from pydantic import BaseModel
from typing import Optional
from datetime import datetime

class ExcelRequest(BaseModel):
    content: str
    filename: Optional[str] = None

class ExcelResponse(BaseModel):
    status: str
    message: str
    filename: str
    object_name: str
    download_url: str
    created_at: datetime
