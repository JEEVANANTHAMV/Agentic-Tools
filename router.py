from fastapi import APIRouter, Depends, HTTPException, Response
from fastapi.responses import FileResponse
from models.document_models import DocumentRequest, DocumentResponse, DocumentListResponse
from models.excel_model import ExcelRequest, ExcelResponse
from models.presentation_model import PresentationResponse, PresentationRequest
from services.docx.docx_creator import DocxCreator
from services.minio_handler import MinioHandler
from services.excel.excel_creator import ExcelCreator
from services.powerpoint.ppt_creator import PresentationCreator
from datetime import datetime
from typing import Optional
import socket, os
from config import settings
from models.sql_to_excel import SQLQueryRequest, SQLQueryResponse
from services.SQL.sql_to_excel import SQLToExcelService
from models.pdf_model import PDFRequest, PDFResponse
from models.csv_model import CSVRequest, CSVResponse
from models.json_model import JSONRequest, JSONResponse
from models.xml_model import XMLRequest, XMLResponse
from models.markdown_model import MarkdownRequest, MarkdownResponse
from models.visualization_model import VisualizationRequest, VisualizationResponse
from services.pdf.pdf_creator import PDFCreator
from services.csv.csv_processor import CSVProcessor
from services.json.json_formatter import JSONFormatter
from services.xml.xml_parser import XMLParser
from services.markdown.markdown_converter import MarkdownConverter
from services.visualization.visualization_creator import VisualizationCreator

router = APIRouter()

def get_docx_creator():
    return DocxCreator()

def get_minio_handler():
    return MinioHandler()

def get_excel_creator():
    return ExcelCreator()

def get_presentation_creator():
    return PresentationCreator()

def get_sql_service():
    return SQLToExcelService()

def get_pdf_creator():
    return PDFCreator()

def get_csv_processor():
    return CSVProcessor()

def get_json_formatter():
    return JSONFormatter()

def get_xml_parser():
    return XMLParser()

def get_markdown_converter():
    return MarkdownConverter()

def get_visualization_creator():
    return VisualizationCreator()

def get_server_ip():
    """Get server IP address"""
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except:
        return "localhost"

@router.post("/generate-document", response_model=DocumentResponse)
async def generate_document(
    request: DocumentRequest,
    docx_creator: DocxCreator = Depends(get_docx_creator)
):
    try:
        # Create date-based folder structure
        today = datetime.now()
        folder_path = os.path.join(
            settings.DOCUMENT_LOCATION,
            today.strftime('%Y'),
            today.strftime('%m'),
            today.strftime('%d')
        )
        os.makedirs(folder_path, exist_ok=True)
        
        # Generate filename
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = request.filename or f"document_{timestamp}"
        if not filename.endswith('.docx'):
            filename += '.docx'
        
        # Full file path
        filepath = os.path.join(folder_path, filename)
        
        # Create document
        doc_stream = docx_creator.create_document(request.content, filename)
        
        # Save to local file system
        with open(filepath, 'wb') as f:
            f.write(doc_stream.read())
        
        # Generate download URL (include the /api/v1 prefix)
        server_ip = get_server_ip()
        relative_path = filepath.replace(settings.DOCUMENT_LOCATION + os.sep, '').replace(os.sep, '/')
        download_url = f"http://{server_ip}:{settings.PORT}/api/v1/download/{relative_path}"
        
        return DocumentResponse(
            status="success",
            message="Document generated successfully",
            filename=filename,
            object_name=relative_path,
            download_url=download_url,
            created_at=datetime.now()
        )
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@router.get("/download/{path:path}")
async def download_file(path: str):
    """Download generated document using path structure YYYY/MM/DD/filename"""
    try:
        # Reconstruct the full file path
        filepath = os.path.join(settings.DOCUMENT_LOCATION, path)
        
        if not os.path.exists(filepath):
            raise HTTPException(status_code=421, detail="File not found")
        
        # Extract filename from path
        filename = path.split('/')[-1]
        
        return FileResponse(
            filepath,
            media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            filename=filename
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@router.get("/list-documents", response_model=DocumentListResponse)
async def list_documents(
    prefix: Optional[str] = None,
    minio_handler: MinioHandler = Depends(get_minio_handler)
):
    """List all documents in MinIO"""
    try:
        documents = minio_handler.list_documents(prefix or "")
        return DocumentListResponse(documents=documents, count=len(documents))
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@router.delete("/delete-document/{object_name:path}")
async def delete_document(
    object_name: str,
    minio_handler: MinioHandler = Depends(get_minio_handler)
):
    """Delete a document from MinIO"""
    try:
        success = minio_handler.delete_document(object_name)
        return {"status": "success", "message": "Document deleted successfully" if success else "Failed to delete document"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@router.post("/generate-excel", response_model=ExcelResponse)
async def generate_excel(
    request: ExcelRequest,
    excel_creator: ExcelCreator = Depends(get_excel_creator)
):
    try:
        # Create date-based folder structure
        today = datetime.now()
        folder_path = os.path.join(
            settings.DOCUMENT_LOCATION,
            today.strftime('%Y'),
            today.strftime('%m'),
            today.strftime('%d')
        )
        os.makedirs(folder_path, exist_ok=True)
        
        # Generate filename
        filename = excel_creator.generate_filename(request.filename)
        
        # Full file path
        filepath = os.path.join(folder_path, filename)
        
        # Create Excel file from content
        excel_stream = excel_creator.create_excel_from_content(request.content, filename)
        
        # Save to local file system
        with open(filepath, 'wb') as f:
            f.write(excel_stream.read())
        
        # Generate download URL
        server_ip = get_server_ip()
        relative_path = filepath.replace(settings.DOCUMENT_LOCATION + os.sep, '').replace(os.sep, '/')
        download_url = f"http://{server_ip}:{settings.PORT}/api/v1/download/{relative_path}"
        
        return ExcelResponse(
            status="success",
            message="Excel file generated successfully",
            filename=filename,
            object_name=relative_path,
            download_url=download_url,
            created_at=datetime.now()
        )
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@router.post("/generate-presentation", response_model=PresentationResponse)
async def generate_presentation(
    request: PresentationRequest,
    presentation_creator: PresentationCreator = Depends(get_presentation_creator)
):
    try:
        # Create date-based folder structure
        today = datetime.now()
        folder_path = os.path.join(
            "generated_presentations",
            today.strftime('%Y'),
            today.strftime('%m'),
            today.strftime('%d')
        )
        os.makedirs(folder_path, exist_ok=True)
        
        # Generate filename
        filename = presentation_creator.generate_filename(request.filename)
        
        # Full file path
        filepath = os.path.join(folder_path, filename)
        
        # Create presentation
        presentation_stream = presentation_creator.create_presentation(request.content, filename)
        
        # Save to local file system
        with open(filepath, 'wb') as f:
            f.write(presentation_stream.read())
        
        # Generate download URL
        server_ip = get_server_ip()
        relative_path = filepath.replace("generated_presentations" + os.sep, '').replace(os.sep, '/')
        download_url = f"http://{server_ip}:{settings.PORT}/api/v1/download/{relative_path}"
        
        return PresentationResponse(
            status="success",
            message="Presentation generated successfully",
            filename=filename,
            object_name=relative_path,
            download_url=download_url,
            created_at=datetime.now()
        )
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@router.post("/execute-sql-excel", response_model=SQLQueryResponse)
async def execute_sql_query(
    request: SQLQueryRequest,
    sql_service: SQLToExcelService = Depends(get_sql_service)
):
    try:
        # Create date-based folder structure
        today = datetime.now()
        folder_path = os.path.join(
            settings.DOCUMENT_LOCATION,
            today.strftime('%Y'),
            today.strftime('%m'),
            today.strftime('%d')
        )
        os.makedirs(folder_path, exist_ok=True)
        
        # Generate filename
        filename = sql_service.generate_filename(request.filename)
        
        # Full file path
        filepath = os.path.join(folder_path, filename)
        
        # Execute query and create Excel file
        excel_stream = sql_service.execute_query_to_excel(request.query, filename)
        
        # Save to local file system
        with open(filepath, 'wb') as f:
            f.write(excel_stream.read())
        
        # Generate download URL
        server_ip = get_server_ip()
        relative_path = filepath.replace(settings.DOCUMENT_LOCATION + os.sep, '').replace(os.sep, '/')
        download_url = f"http://{server_ip}:{settings.PORT}/api/v1/download/{relative_path}"
        
        return SQLQueryResponse(
            status="success",
            message="SQL query executed and Excel file generated successfully",
            filename=filename,
            object_name=relative_path,
            download_url=download_url,
            created_at=datetime.now()
        )
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@router.get("/")
async def root():
    """API information"""
    return {
        "name": "API for forjinn tools.",
        "version": "2.0.0",
        "endpoints": {
            "generate": "/generate-document (POST)",
            "download": "/download/{object_name:path} (GET)",
            "list": "/list-documents (GET)",
            "delete": "/delete-document/{object_name:path} (DELETE)"
        },
        "server_ip": get_server_ip()
    }

# PDF Converter Endpoints
@router.post("/convert-to-pdf", response_model=PDFResponse, tags=["PDF"])
async def convert_to_pdf(
    request: PDFRequest,
    pdf_creator: PDFCreator = Depends(get_pdf_creator)
):
    """Convert content to PDF format"""
    try:
        today = datetime.now()
        folder_path = os.path.join(
            settings.DOCUMENT_LOCATION,
            today.strftime('%Y'),
            today.strftime('%m'),
            today.strftime('%d')
        )
        os.makedirs(folder_path, exist_ok=True)
        
        filename = pdf_creator.generate_filename(request.filename)
        filepath = os.path.join(folder_path, filename)
        
        pdf_stream = pdf_creator.create_pdf_from_content(request.content, filename)
        
        with open(filepath, 'wb') as f:
            f.write(pdf_stream.read())
        
        server_ip = get_server_ip()
        relative_path = filepath.replace(settings.DOCUMENT_LOCATION + os.sep, '').replace(os.sep, '/')
        download_url = f"http://{server_ip}:{settings.PORT}/api/v1/download/{relative_path}"
        
        return PDFResponse(
            status="success",
            message="PDF generated successfully",
            filename=filename,
            object_name=relative_path,
            download_url=download_url,
            created_at=datetime.now()
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# CSV Processor Endpoints
@router.post("/process-csv", response_model=CSVResponse, tags=["CSV"])
async def process_csv(
    request: CSVRequest,
    csv_processor: CSVProcessor = Depends(get_csv_processor)
):
    """Process and transform CSV data"""
    try:
        today = datetime.now()
        folder_path = os.path.join(
            settings.DOCUMENT_LOCATION,
            today.strftime('%Y'),
            today.strftime('%m'),
            today.strftime('%d')
        )
        os.makedirs(folder_path, exist_ok=True)
        
        filename = csv_processor.generate_filename(request.filename)
        filepath = os.path.join(folder_path, filename)
        
        csv_stream = csv_processor.process_csv(request.content, request.operations, filename)
        
        with open(filepath, 'wb') as f:
            f.write(csv_stream.read())
        
        server_ip = get_server_ip()
        relative_path = filepath.replace(settings.DOCUMENT_LOCATION + os.sep, '').replace(os.sep, '/')
        download_url = f"http://{server_ip}:{settings.PORT}/api/v1/download/{relative_path}"
        
        return CSVResponse(
            status="success",
            message="CSV processed successfully",
            filename=filename,
            object_name=relative_path,
            download_url=download_url,
            created_at=datetime.now()
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# JSON Formatter Endpoints
@router.post("/format-json", response_model=JSONResponse, tags=["JSON"])
async def format_json(
    request: JSONRequest,
    json_formatter: JSONFormatter = Depends(get_json_formatter)
):
    """Format and validate JSON data"""
    try:
        today = datetime.now()
        folder_path = os.path.join(
            settings.DOCUMENT_LOCATION,
            today.strftime('%Y'),
            today.strftime('%m'),
            today.strftime('%d')
        )
        os.makedirs(folder_path, exist_ok=True)
        
        filename = json_formatter.generate_filename(request.filename)
        filepath = os.path.join(folder_path, filename)
        
        json_stream = json_formatter.format_json(request.content, request.json_schema, filename)
        
        with open(filepath, 'wb') as f:
            f.write(json_stream.read())
        
        server_ip = get_server_ip()
        relative_path = filepath.replace(settings.DOCUMENT_LOCATION + os.sep, '').replace(os.sep, '/')
        download_url = f"http://{server_ip}:{settings.PORT}/api/v1/download/{relative_path}"
        
        return JSONResponse(
            status="success",
            message="JSON formatted successfully",
            filename=filename,
            object_name=relative_path,
            download_url=download_url,
            created_at=datetime.now()
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# XML Parser Endpoints
@router.post("/parse-xml", response_model=XMLResponse, tags=["XML"])
async def parse_xml(
    request: XMLRequest,
    xml_parser: XMLParser = Depends(get_xml_parser)
):
    """Parse and transform XML data"""
    try:
        today = datetime.now()
        folder_path = os.path.join(
            settings.DOCUMENT_LOCATION,
            today.strftime('%Y'),
            today.strftime('%m'),
            today.strftime('%d')
        )
        os.makedirs(folder_path, exist_ok=True)
        
        filename = xml_parser.generate_filename(request.filename)
        filepath = os.path.join(folder_path, filename)
        
        xml_stream = xml_parser.parse_xml(request.content, request.transform, filename)
        
        with open(filepath, 'wb') as f:
            f.write(xml_stream.read())
        
        server_ip = get_server_ip()
        relative_path = filepath.replace(settings.DOCUMENT_LOCATION + os.sep, '').replace(os.sep, '/')
        download_url = f"http://{server_ip}:{settings.PORT}/api/v1/download/{relative_path}"
        
        return XMLResponse(
            status="success",
            message="XML parsed successfully",
            filename=filename,
            object_name=relative_path,
            download_url=download_url,
            created_at=datetime.now()
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# Markdown Converter Endpoints
@router.post("/convert-markdown", response_model=MarkdownResponse, tags=["Markdown"])
async def convert_markdown(
    request: MarkdownRequest,
    markdown_converter: MarkdownConverter = Depends(get_markdown_converter)
):
    """Convert Markdown to various formats"""
    try:
        today = datetime.now()
        folder_path = os.path.join(
            settings.DOCUMENT_LOCATION,
            today.strftime('%Y'),
            today.strftime('%m'),
            today.strftime('%d')
        )
        os.makedirs(folder_path, exist_ok=True)
        
        filename = markdown_converter.generate_filename(request.filename, request.output_format)
        filepath = os.path.join(folder_path, filename)
        
        output_stream = markdown_converter.convert_markdown(request.content, request.output_format, filename)
        
        with open(filepath, 'wb') as f:
            f.write(output_stream.read())
        
        server_ip = get_server_ip()
        relative_path = filepath.replace(settings.DOCUMENT_LOCATION + os.sep, '').replace(os.sep, '/')
        download_url = f"http://{server_ip}:{settings.PORT}/api/v1/download/{relative_path}"
        
        return MarkdownResponse(
            status="success",
            message=f"Markdown converted to {request.output_format} successfully",
            filename=filename,
            object_name=relative_path,
            download_url=download_url,
            created_at=datetime.now()
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# Data Visualization Endpoints
@router.post("/create-visualization", response_model=VisualizationResponse, tags=["Visualization"])
async def create_visualization(
    request: VisualizationRequest,
    visualization_creator: VisualizationCreator = Depends(get_visualization_creator)
):
    """Create charts and visualizations from data"""
    try:
        today = datetime.now()
        folder_path = os.path.join(
            settings.DOCUMENT_LOCATION,
            today.strftime('%Y'),
            today.strftime('%m'),
            today.strftime('%d')
        )
        os.makedirs(folder_path, exist_ok=True)
        
        filename = visualization_creator.generate_filename(request.filename)
        filepath = os.path.join(folder_path, filename)
        
        viz_stream = visualization_creator.create_visualization(request.data, request.chart_type, filename)
        
        with open(filepath, 'wb') as f:
            f.write(viz_stream.read())
        
        server_ip = get_server_ip()
        relative_path = filepath.replace(settings.DOCUMENT_LOCATION + os.sep, '').replace(os.sep, '/')
        download_url = f"http://{server_ip}:{settings.PORT}/api/v1/download/{relative_path}"
        
        return VisualizationResponse(
            status="success",
            message=f"{request.chart_type} chart created successfully",
            filename=filename,
            object_name=relative_path,
            download_url=download_url,
            created_at=datetime.now()
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))