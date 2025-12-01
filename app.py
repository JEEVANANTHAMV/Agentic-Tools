from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from router import router as document_router
from config import settings

# Create FastAPI app
app = FastAPI(
    title=settings.APP_TITLE,
    version=settings.APP_VERSION,
    description="API for generating and managing Word documents with MinIO storage"
)

# Add CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Include routers
app.include_router(document_router, prefix="/api/v1")

@app.on_event("startup")
async def startup_event():
    print(f"Starting {settings.APP_TITLE} v{settings.APP_VERSION}")
    print(f"Server running on http://{settings.HOST}:{settings.PORT}")
