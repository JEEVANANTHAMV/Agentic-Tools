from pydantic import BaseModel
from typing import Optional
from datetime import datetime

class MarkdownRequest(BaseModel):
    content: str
    output_format: str  # html, pdf, docx
    filename: Optional[str] = None

class MarkdownResponse(BaseModel):
    status: str
    message: str
    filename: str
    object_name: str
    download_url: str
    created_at: datetime
