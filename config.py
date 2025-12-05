import os
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

class Settings:
    # FastAPI Settings
    APP_TITLE = os.getenv("APP_TITLE", "Forjinn Tools")
    APP_VERSION = os.getenv("APP_VERSION", "2.0.0")
    PORT = int(os.getenv("PORT", 19801))
    HOST = os.getenv("HOST", "0.0.0.0")
    
    # MinIO Settings
    MINIO_ENDPOINT = os.getenv("MINIO_ENDPOINT", "localhost:9000")
    MINIO_ACCESS_KEY = os.getenv("MINIO_ACCESS_KEY", "minioadmin")
    MINIO_SECRET_KEY = os.getenv("MINIO_SECRET_KEY", "minioadmin")
    MINIO_SECURE = os.getenv("MINIO_SECURE", "False").lower() == "true"
    MINIO_BUCKET_NAME = os.getenv("MINIO_BUCKET_NAME", "documents")
    
    # Document Settings
    DEFAULT_FONT_NAME = os.getenv("DEFAULT_FONT_NAME", "Calibri")
    DEFAULT_FONT_SIZE = int(os.getenv("DEFAULT_FONT_SIZE", 11))

    DOCUMENT_LOCATION = os.getenv("DOCUMENT_LOCATION","generated_documents")
    DB_HOST = os.getenv("DB_HOST", "localhost")
    DB_PORT = os.getenv("DB_PORT", "3306")
    DB_USER = os.getenv("DB_USER", "root")
    DB_PASSWORD = os.getenv("DB_PASSWORD", "")
    DB_NAME = os.getenv("DB_NAME", "your_database")

settings = Settings()