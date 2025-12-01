from pydantic import BaseModel
from typing import Optional
from datetime import datetime

class PresentationRequest(BaseModel):
    content: str  # HTML content string
    filename: Optional[str] = None

class PresentationResponse(BaseModel):
    status: str
    message: str
    filename: str
    object_name: str
    download_url: str
    created_at: datetime