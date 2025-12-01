from minio import Minio
from minio.error import S3Error
from fastapi import HTTPException
from io import BytesIO
from datetime import datetime, timedelta
from typing import List, Dict, Any
import uuid
from config import settings

class MinioHandler:
    def __init__(self):
        self.client = Minio(
            endpoint=settings.MINIO_ENDPOINT,
            access_key=settings.MINIO_ACCESS_KEY,
            secret_key=settings.MINIO_SECRET_KEY,
            secure=settings.MINIO_SECURE
        )
        self.bucket_name = settings.MINIO_BUCKET_NAME
        self._ensure_bucket_exists()
    
    def _ensure_bucket_exists(self):
        """Create bucket if it doesn't exist"""
        try:
            if not self.client.bucket_exists(self.bucket_name):
                self.client.make_bucket(self.bucket_name)
                print(f"Bucket '{self.bucket_name}' created successfully")
        except S3Error as e:
            print(f"Error creating bucket: {e}")
            raise HTTPException(status_code=500, detail=f"Error connecting to MinIO: {e}")
    
    def upload_document(self, file_stream: BytesIO, object_name: str, content_type: str = "application/vnd.openxmlformats-officedocument.wordprocessingml.document") -> str:
        """Upload a document to MinIO and return the object name"""
        try:
            # Reset stream position
            file_stream.seek(0)
            
            # Upload to MinIO
            result = self.client.put_object(
                bucket_name=self.bucket_name,
                object_name=object_name,
                data=file_stream,
                length=file_stream.getbuffer().nbytes,
                content_type=content_type
            )
            
            return object_name
        except S3Error as e:
            print(f"Error uploading file: {e}")
            raise HTTPException(status_code=500, detail=f"Error uploading file: {e}")
    
    def download_document(self, object_name: str) -> BytesIO:
        """Download a document from MinIO"""
        try:
            # Get object from MinIO
            response = self.client.get_object(self.bucket_name, object_name)
            
            # Read data into BytesIO
            file_data = BytesIO(response.read())
            response.close()
            response.release_conn()
            
            return file_data
        except S3Error as e:
            print(f"Error downloading file: {e}")
            raise HTTPException(status_code=404, detail=f"File not found: {e}")
    
    def get_presigned_url(self, object_name: str, expires: timedelta = timedelta(days=7)) -> str:
        """Generate a presigned URL for downloading a document"""
        try:
            url = self.client.presigned_get_object(
                bucket_name=self.bucket_name,
                object_name=object_name,
                expires=expires
            )
            return url
        except S3Error as e:
            print(f"Error generating presigned URL: {e}")
            raise HTTPException(status_code=500, detail=f"Error generating download URL: {e}")
    
    def list_documents(self, prefix: str = "") -> List[Dict[str, Any]]:
        """List all documents in the bucket, optionally filtered by prefix"""
        try:
            objects = self.client.list_objects(self.bucket_name, prefix=prefix, recursive=True)
            
            documents = []
            for obj in objects:
                documents.append({
                    "name": obj.object_name,
                    "size": obj.size,
                    "last_modified": obj.last_modified,
                    "etag": obj.etag,
                    "content_type": obj.content_type,
                    "download_url": self.get_presigned_url(obj.object_name)
                })
            
            return documents
        except S3Error as e:
            print(f"Error listing documents: {e}")
            raise HTTPException(status_code=500, detail=f"Error listing documents: {e}")
    
    def delete_document(self, object_name: str) -> bool:
        """Delete a document from MinIO"""
        try:
            self.client.remove_object(self.bucket_name, object_name)
            return True
        except S3Error as e:
            print(f"Error deleting file: {e}")
            raise HTTPException(status_code=500, detail=f"Error deleting file: {e}")