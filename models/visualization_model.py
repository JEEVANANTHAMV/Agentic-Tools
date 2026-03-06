from pydantic import BaseModel
from typing import Optional
from datetime import datetime

class VisualizationRequest(BaseModel):
    data: str
    chart_type: str
    filename: Optional[str] = None

class VisualizationResponse(BaseModel):
    status: str
    message: str
    filename: str
    object_name: str
    download_url: str
    created_at: datetime
