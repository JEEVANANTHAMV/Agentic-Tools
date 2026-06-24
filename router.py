from fastapi import APIRouter, Depends, HTTPException, Response, Request
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
from models.mssql_model import MSSQLRequest, MSSQLResponse
from services.mssql.mssql_service import MSSQLService, MSSQLError, ReadOnlyViolation

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

def get_mssql_service():
    return MSSQLService()

def get_server_ip(request: Request):
    """Get server IP address from request"""
    return request.url.hostname or "127.0.0.1"

@router.post("/generate-document", response_model=DocumentResponse)
async def generate_document(
    request: DocumentRequest,
    docx_creator: DocxCreator = Depends(get_docx_creator),
    server_ip: str = Depends(get_server_ip)
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
        # Try multiple locations for the file
        possible_locations = [
            settings.DOCUMENT_LOCATION,  # generated_documents
            "generated_presentations",   # generated_presentations
            "generated_html_presentations"  # generated_html_presentations
        ]
        
        filepath = None
        for location in possible_locations:
            test_path = os.path.join(location, path)
            if os.path.exists(test_path):
                filepath = test_path
                break
        
        if filepath is None:
            raise HTTPException(status_code=421, detail="File not found")
        
        # Extract filename from path
        filename = path.split('/')[-1]
        
        # Determine media type based on file extension
        if filename.endswith('.pptx'):
            media_type = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
        elif filename.endswith('.docx'):
            media_type = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        elif filename.endswith('.xlsx'):
            media_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        elif filename.endswith('.pdf'):
            media_type = 'application/pdf'
        elif filename.endswith('.html'):
            media_type = 'text/html'
        else:
            media_type = 'application/octet-stream'
        
        content_disposition = "inline" if filename.endswith('.html') else "attachment"
        if filename.endswith('.html'):
            return FileResponse(
                filepath,
                media_type=media_type,
                content_disposition_type=content_disposition
            )
        else:
            return FileResponse(
                filepath,
                media_type=media_type,
                filename=filename,
                content_disposition_type=content_disposition
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
    excel_creator: ExcelCreator = Depends(get_excel_creator),
    server_ip: str = Depends(get_server_ip)
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
    presentation_creator: PresentationCreator = Depends(get_presentation_creator),
    server_ip: str = Depends(get_server_ip)
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
        presentation_stream = await presentation_creator.create_presentation(request.content, filename)
        
        # Save to local file system
        with open(filepath, 'wb') as f:
            f.write(presentation_stream.read())
            
        # Create presentation HTML wrapper
        html_folder_path = os.path.join(
            "generated_html_presentations",
            today.strftime('%Y'),
            today.strftime('%m'),
            today.strftime('%d')
        )
        os.makedirs(html_folder_path, exist_ok=True)
        html_filename = filename.replace('.pptx', '.html')
        html_filepath = os.path.join(html_folder_path, html_filename)
        
        # We need to apply Tailwind injected CSS to html
        from services.powerpoint.ppt_creator import _inject_css
        styled_html = _inject_css(request.content)
        
        # Wrap it in our presentation viewer
        presentation_html = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>Presentation Preview</title>
    <style>
        body {{ margin: 0; background: #111; height: 100vh; width: 100vw; overflow: hidden; position: relative; }}
        .slide-container {{ position: absolute; left: 50%; top: 50%; width: 1280px; height: 720px; box-shadow: 0 0 30px rgba(0,0,0,0.8); background: white; transform-origin: center center; }}
        .slide {{ display: none !important; width: 100%; height: 100%; }}
        .ppt-slide {{ display: none !important; width: 100%; height: 100%; }}
        .slide.active, .ppt-slide.active {{ display: block !important; }}
    </style>
</head>
<body>
    <div class="slide-container" id="presentation-container">
        {styled_html}
    </div>
    <script>
        const slides = document.querySelectorAll('.slide, .ppt-slide');
        let currentSlide = 0;
        if(slides.length > 0) slides[currentSlide].classList.add('active');
        
        function resize() {{
            const container = document.getElementById('presentation-container');
            const scale = Math.min(window.innerWidth / 1280, window.innerHeight / 720);
            container.style.transform = `translate(-50%, -50%) scale(${{scale}})`;
        }}
        window.addEventListener('resize', resize);
        resize();

        document.addEventListener('keydown', (e) => {{
            if (e.key === 'Enter' || e.key === 'ArrowRight' || e.key === ' ') {{
                if (currentSlide < slides.length - 1) {{
                    slides[currentSlide].classList.remove('active');
                    currentSlide++;
                    slides[currentSlide].classList.add('active');
                }}
            }} else if (e.key === 'ArrowLeft') {{
                if (currentSlide > 0) {{
                    slides[currentSlide].classList.remove('active');
                    currentSlide--;
                    slides[currentSlide].classList.add('active');
                }}
            }}
        }});
        
        document.addEventListener('click', (e) => {{
            if (currentSlide < slides.length - 1) {{
                slides[currentSlide].classList.remove('active');
                currentSlide++;
                slides[currentSlide].classList.add('active');
            }}
        }});
    </script>
</body>
</html>"""
        
        with open(html_filepath, 'w', encoding='utf-8') as f:
            f.write(presentation_html)
            
        # Generate download URL
        html_relative_path = html_filepath.replace("generated_html_presentations" + os.sep, '').replace(os.sep, '/')
        preview_url = f"http://{server_ip}:{settings.PORT}/api/v1/download/{html_relative_path}?view=presentation"
        
        relative_path = filepath.replace("generated_presentations" + os.sep, '').replace(os.sep, '/')
        download_url = f"http://{server_ip}:{settings.PORT}/api/v1/download/{relative_path}"
        
        return PresentationResponse(
            status="success",
            message="Presentation generated successfully",
            filename=filename,
            object_name=relative_path,
            download_url=download_url,
            preview_url=preview_url,
            created_at=datetime.now()
        )
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@router.post("/execute-sql-excel", response_model=SQLQueryResponse)
async def execute_sql_query(
    request: SQLQueryRequest,
    sql_service: SQLToExcelService = Depends(get_sql_service),
    server_ip: str = Depends(get_server_ip)
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
async def root(server_ip: str = Depends(get_server_ip)):
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
        "server_ip": server_ip
    }

# PDF Converter Endpoints
@router.post("/convert-to-pdf", response_model=PDFResponse, tags=["PDF"])
async def convert_to_pdf(
    request: PDFRequest,
    pdf_creator: PDFCreator = Depends(get_pdf_creator),
    server_ip: str = Depends(get_server_ip)
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
    csv_processor: CSVProcessor = Depends(get_csv_processor),
    server_ip: str = Depends(get_server_ip)
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
    json_formatter: JSONFormatter = Depends(get_json_formatter),
    server_ip: str = Depends(get_server_ip)
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
    xml_parser: XMLParser = Depends(get_xml_parser),
    server_ip: str = Depends(get_server_ip)
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
    markdown_converter: MarkdownConverter = Depends(get_markdown_converter),
    server_ip: str = Depends(get_server_ip)
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
    visualization_creator: VisualizationCreator = Depends(get_visualization_creator),
    server_ip: str = Depends(get_server_ip)
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

# MS SQL Connector Endpoint (single endpoint, two modes: list actions / call action)
@router.post("/mssql", response_model=MSSQLResponse, tags=["MS SQL"])
async def mssql_connector(
    request: MSSQLRequest,
    mssql_service: MSSQLService = Depends(get_mssql_service)
):
    """
    Connect to ANY MS SQL Server database and run actions against it.

    Two modes in one endpoint:
    - mode="list": returns the catalogue of actions and how to use them (no DB call).
    - mode="call": runs `action` against the database described by `connection`.

    Credentials are passed per-request in `connection` (never read from .env),
    so many different databases/users can be served from this one endpoint.
    """
    # ---- mode: list -> self-documenting "how to use" catalogue ----
    if request.mode == "list":
        return MSSQLResponse(
            status="success",
            mode="list",
            action=None,
            message="Available MS SQL actions and how to call them. Use mode='call' to run one.",
            usage=mssql_service.get_usage(),
            created_at=datetime.now()
        )

    # ---- mode: call -> validate then dispatch ----
    if not request.connection:
        raise HTTPException(
            status_code=400,
            detail="'connection' is required when mode='call' (host, user, password, ...)."
        )
    if not request.action:
        raise HTTPException(
            status_code=400,
            detail="'action' is required when mode='call'. Send {\"mode\":\"list\"} to see available actions."
        )

    # Target info echoed back on every action (no credentials included).
    target = {
        "host": request.connection.host,
        "port": request.connection.port,
        "database": request.params.database or request.connection.database,
    }

    try:
        data, meta = mssql_service.run_action(request.action, request.connection, request.params)
        return MSSQLResponse(
            status="success",
            mode="call",
            action=request.action,
            message=f"Action '{request.action}' completed successfully.",
            target=target,
            data=data,
            meta=meta,
            created_at=datetime.now()
        )
    except ValueError as e:
        # Unknown action / bad arguments -> client error.
        raise HTTPException(status_code=400, detail=str(e))
    except ReadOnlyViolation as e:
        return MSSQLResponse(
            status="error",
            mode="call",
            action=request.action,
            message=str(e),
            target=target,
            created_at=datetime.now()
        )
    except MSSQLError as e:
        # Connection / SQL execution failures -> return a clean, readable error body.
        return MSSQLResponse(
            status="error",
            mode="call",
            action=request.action,
            message=str(e),
            target=target,
            created_at=datetime.now()
        )