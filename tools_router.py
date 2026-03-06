from fastapi import APIRouter, HTTPException
from pydantic import BaseModel
from typing import List, Optional, Dict, Any
import os
import json

router = APIRouter()

# Define the available tools with their metadata
AVAILABLE_TOOLS = [
    {
        "tool_name": "document_writer",
        "description": "Generate Word documents with rich text formatting, fonts, tables, and lists",
        "endpoint": "/api/v1/generate-document",
        "method": "POST",
        "service_path": "services/docx",
        "parameters": {
            "content": {"type": "string", "description": "Formatted content for the document using markdown and font tags", "required": True},
            "filename": {"type": "string", "description": "Optional filename for the document", "required": False}
        },
        "response_model": {
            "status": "string",
            "message": "string",
            "filename": "string",
            "object_name": "string",
            "download_url": "string",
            "created_at": "datetime"
        }
    },
    {
        "tool_name": "excel_generator",
        "description": "Generate Excel files with multiple sheets, formatted cells, colors, and borders",
        "endpoint": "/api/v1/generate-excel",
        "method": "POST",
        "service_path": "services/excel",
        "parameters": {
            "content": {"type": "string", "description": "Formatted content for the Excel file with sheet definitions and table syntax", "required": True},
            "filename": {"type": "string", "description": "Optional filename for the Excel file", "required": False}
        },
        "response_model": {
            "status": "string",
            "message": "string",
            "filename": "string",
            "object_name": "string",
            "download_url": "string",
            "created_at": "datetime"
        }
    },
    {
        "tool_name": "presentation_creator",
        "description": "Generate PowerPoint presentations with slides, layouts, and content",
        "endpoint": "/api/v1/generate-presentation",
        "method": "POST",
        "service_path": "services/powerpoint",
        "parameters": {
            "content": {"type": "string", "description": "Formatted content for the presentation with slide definitions", "required": True},
            "filename": {"type": "string", "description": "Optional filename for the presentation", "required": False}
        },
        "response_model": {
            "status": "string",
            "message": "string",
            "filename": "string",
            "object_name": "string",
            "download_url": "string",
            "created_at": "datetime"
        }
    },
    {
        "tool_name": "sql_to_excel",
        "description": "Execute SQL queries and export results to Excel format",
        "endpoint": "/api/v1/execute-sql-excel",
        "method": "POST",
        "service_path": "services/SQL",
        "parameters": {
            "query": {"type": "string", "description": "SQL query to execute", "required": True},
            "filename": {"type": "string", "description": "Optional filename for the Excel file", "required": False}
        },
        "response_model": {
            "status": "string",
            "message": "string",
            "filename": "string",
            "object_name": "string",
            "download_url": "string",
            "created_at": "datetime"
        }
    },
    {
        "tool_name": "pdf_converter",
        "description": "Convert documents to PDF format with customizable settings",
        "endpoint": "/api/v1/convert-to-pdf",
        "method": "POST",
        "service_path": "services/pdf",
        "parameters": {
            "content": {"type": "string", "description": "Content to convert to PDF", "required": True},
            "filename": {"type": "string", "description": "Optional filename for the PDF", "required": False}
        },
        "response_model": {
            "status": "string",
            "message": "string",
            "filename": "string",
            "object_name": "string",
            "download_url": "string",
            "created_at": "datetime"
        }
    },
    {
        "tool_name": "csv_processor",
        "description": "Process and transform CSV data with filtering, sorting, and formatting",
        "endpoint": "/api/v1/process-csv",
        "method": "POST",
        "service_path": "services/csv",
        "parameters": {
            "content": {"type": "string", "description": "CSV content or file reference to process", "required": True},
            "operations": {"type": "array", "description": "List of operations to perform (filter, sort, transform)", "required": False},
            "filename": {"type": "string", "description": "Optional filename for the output", "required": False}
        },
        "response_model": {
            "status": "string",
            "message": "string",
            "filename": "string",
            "object_name": "string",
            "download_url": "string",
            "created_at": "datetime"
        }
    },
    {
        "tool_name": "json_formatter",
        "description": "Format, validate, and transform JSON data with schema support",
        "endpoint": "/api/v1/format-json",
        "method": "POST",
        "service_path": "services/json",
        "parameters": {
            "content": {"type": "string", "description": "JSON content to format", "required": True},
            "schema": {"type": "object", "description": "Optional JSON schema for validation", "required": False},
            "filename": {"type": "string", "description": "Optional filename for the output", "required": False}
        },
        "response_model": {
            "status": "string",
            "message": "string",
            "filename": "string",
            "object_name": "string",
            "download_url": "string",
            "created_at": "datetime"
        }
    },
    {
        "tool_name": "xml_parser",
        "description": "Parse, transform, and validate XML documents with XSD support",
        "endpoint": "/api/v1/parse-xml",
        "method": "POST",
        "service_path": "services/xml",
        "parameters": {
            "content": {"type": "string", "description": "XML content to parse", "required": True},
            "transform": {"type": "string", "description": "Optional XSLT transformation", "required": False},
            "filename": {"type": "string", "description": "Optional filename for the output", "required": False}
        },
        "response_model": {
            "status": "string",
            "message": "string",
            "filename": "string",
            "object_name": "string",
            "download_url": "string",
            "created_at": "datetime"
        }
    },
    {
        "tool_name": "markdown_converter",
        "description": "Convert Markdown to various formats (HTML, PDF, DOCX) with custom styling",
        "endpoint": "/api/v1/convert-markdown",
        "method": "POST",
        "service_path": "services/markdown",
        "parameters": {
            "content": {"type": "string", "description": "Markdown content to convert", "required": True},
            "output_format": {"type": "string", "description": "Target format: html, pdf, docx", "required": True},
            "filename": {"type": "string", "description": "Optional filename for the output", "required": False}
        },
        "response_model": {
            "status": "string",
            "message": "string",
            "filename": "string",
            "object_name": "string",
            "download_url": "string",
            "created_at": "datetime"
        }
    },
    {
        "tool_name": "data_visualizer",
        "description": "Create charts and visualizations from data and export to various formats",
        "endpoint": "/api/v1/create-visualization",
        "method": "POST",
        "service_path": "services/visualization",
        "parameters": {
            "data": {"type": "string", "description": "Data in JSON or CSV format", "required": True},
            "chart_type": {"type": "string", "description": "Type of chart: bar, line, pie, scatter, etc.", "required": True},
            "filename": {"type": "string", "description": "Optional filename for the output", "required": False}
        },
        "response_model": {
            "status": "string",
            "message": "string",
            "filename": "string",
            "object_name": "string",
            "download_url": "string",
            "created_at": "datetime"
        }
    }
]


class ToolInfo(BaseModel):
    tool_name: str
    description: str
    endpoint: str
    method: str
    parameters: Dict[str, Any]
    response_model: Dict[str, str]


class ToolListResponse(BaseModel):
    total_tools: int
    tools: List[ToolInfo]


class ToolPromptResponse(BaseModel):
    tool_name: str
    endpoint: str
    method: str
    prompt: str
    sample_curl: str
    parameters: Dict[str, Any]


def read_prompt_file(service_path: str) -> str:
    """Read the prompt.md file for a given service"""
    prompt_file_path = os.path.join(service_path, "prompt.md")
    try:
        with open(prompt_file_path, 'r', encoding='utf-8') as f:
            return f.read()
    except FileNotFoundError:
        return f"Prompt documentation not found for {service_path}"
    except Exception as e:
        return f"Error reading prompt file: {str(e)}"


def generate_sample_curl(tool: Dict[str, Any]) -> str:
    """Generate a sample curl command for a given tool"""
    base_url = "http://101.53.140.44:8002"
    endpoint = tool["endpoint"]
    method = tool["method"]
    params = tool["parameters"]
    
    # Build the JSON payload based on required parameters
    payload_parts = []
    for param_name, param_info in params.items():
        if param_info.get("required", False):
            if param_name == "content":
                payload_parts.append(f'    "content": "Sample content for {tool["tool_name"]}"')
            elif param_name == "query":
                payload_parts.append('    "query": "SELECT * FROM table LIMIT 10"')
            elif param_name == "output_format":
                payload_parts.append('    "output_format": "html"')
            elif param_name == "chart_type":
                payload_parts.append('    "chart_type": "bar"')
            elif param_name == "data":
                payload_parts.append('    "data": "{\\"name\\": \\"test\\", \\"value\\": 100}"')
            else:
                payload_parts.append(f'    "{param_name}": "sample_value"')
    
    # Add optional filename parameter
    payload_parts.append(f'    "filename": "{tool["tool_name"]}_output"')
    
    payload = "{\n" + ",\n".join(payload_parts) + "\n  }"
    
    curl_command = f"""curl -X '{method}' \
  '{base_url}{endpoint}' \
  -H 'accept: application/json' \
  -H 'Content-Type: application/json' \
  -d '{payload}'"""
    
    return curl_command


@router.get("/tools", response_model=ToolListResponse, tags=["Tools"])
async def list_all_tools():
    """
    List all available tools in the system.
    
    This endpoint returns a comprehensive list of all available tools
    with their descriptions, endpoints, methods, and parameter information.
    
    Returns:
        ToolListResponse: A response containing the total number of tools and their details
    """
    tools_info = [
        ToolInfo(
            tool_name=tool["tool_name"],
            description=tool["description"],
            endpoint=tool["endpoint"],
            method=tool["method"],
            parameters=tool["parameters"],
            response_model=tool["response_model"]
        )
        for tool in AVAILABLE_TOOLS
    ]
    
    return ToolListResponse(
        total_tools=len(tools_info),
        tools=tools_info
    )


@router.get("/tools/{tool_name}/prompt", response_model=ToolPromptResponse, tags=["Tools"])
async def get_tool_prompt(tool_name: str):
    """
    Get the prompt documentation and sample curl command for a specific tool.
    
    This endpoint returns:
    - The full prompt.md documentation for the tool
    - A sample curl command to use the tool
    - Parameter information
    
    Args:
        tool_name: The name of the tool (e.g., "document_writer", "excel_generator")
    
    Returns:
        ToolPromptResponse: A response containing the prompt, sample curl, and parameters
    
    Raises:
        HTTPException: If the tool is not found
    """
    # Find the tool in the available tools list
    tool = None
    for t in AVAILABLE_TOOLS:
        if t["tool_name"] == tool_name:
            tool = t
            break
    
    if not tool:
        raise HTTPException(
            status_code=404,
            detail=f"Tool '{tool_name}' not found. Available tools: {', '.join(t['tool_name'] for t in AVAILABLE_TOOLS)}"
        )
    
    # Read the prompt file
    prompt = read_prompt_file(tool["service_path"])
    
    # Generate sample curl command
    sample_curl = generate_sample_curl(tool)
    
    return ToolPromptResponse(
        tool_name=tool["tool_name"],
        endpoint=tool["endpoint"],
        method=tool["method"],
        prompt=prompt,
        sample_curl=sample_curl,
        parameters=tool["parameters"]
    )


@router.get("/tools/{tool_name}", response_model=ToolInfo, tags=["Tools"])
async def get_tool_info(tool_name: str):
    """
    Get detailed information about a specific tool.
    
    Args:
        tool_name: The name of the tool
    
    Returns:
        ToolInfo: Detailed information about the tool
    
    Raises:
        HTTPException: If the tool is not found
    """
    for tool in AVAILABLE_TOOLS:
        if tool["tool_name"] == tool_name:
            return ToolInfo(
                tool_name=tool["tool_name"],
                description=tool["description"],
                endpoint=tool["endpoint"],
                method=tool["method"],
                parameters=tool["parameters"],
                response_model=tool["response_model"]
            )
    
    raise HTTPException(
        status_code=404,
        detail=f"Tool '{tool_name}' not found. Available tools: {', '.join(t['tool_name'] for t in AVAILABLE_TOOLS)}"
    )
