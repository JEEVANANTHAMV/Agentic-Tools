from pydantic import BaseModel
from typing import Optional, List
from datetime import datetime

class CSVRequest(BaseModel):
    content: str
    operations: Optional[List[dict]] = None
    filename: Optional[str] = None

class CSVResponse(BaseModel):
    status: str
    message: str
    filename: str
    object_name: str
    download_url: str
    created_at: datetime
