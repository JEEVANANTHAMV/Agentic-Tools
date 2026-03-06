import json
import re
from datetime import datetime
from io import BytesIO
from config import settings

class JSONFormatter:
    def __init__(self):
        self.default_font_name = settings.DEFAULT_FONT_NAME
        self.default_font_size = settings.DEFAULT_FONT_SIZE
    
    def format_json(self, content: str, schema: dict = None, filename: str = None) -> BytesIO:
        """Format and validate JSON content and return as BytesIO"""
        try:
            # Parse JSON
            data = json.loads(content)
            
            # Validate against schema if provided
            if schema:
                self.validate_schema(data, schema)
            
            # Format with indentation
            formatted = json.dumps(data, indent=2, ensure_ascii=False)
            
            # Convert to BytesIO
            output = BytesIO(formatted.encode('utf-8'))
            output.seek(0)
            
            return output
            
        except json.JSONDecodeError as e:
            raise ValueError(f"Invalid JSON: {str(e)}")
    
    def validate_schema(self, data: dict, schema: dict):
        """Validate data against JSON schema"""
        schema_type = schema.get('type')
        required = schema.get('required', [])
        properties = schema.get('properties', {})
        
        # Check required fields
        if isinstance(data, dict):
            for field in required:
                if field not in data:
                    raise ValueError(f"Missing required field: {field}")
            
            # Validate property types
            for prop, prop_schema in properties.items():
                if prop in data:
                    self.validate_property(data[prop], prop_schema)
    
    def validate_property(self, value, prop_schema):
        """Validate a single property against its schema"""
        prop_type = prop_schema.get('type')
        
        if prop_type == 'string':
            if not isinstance(value, str):
                raise ValueError(f"Expected string, got {type(value).__name__}")
        elif prop_type == 'integer':
            if not isinstance(value, int):
                raise ValueError(f"Expected integer, got {type(value).__name__}")
        elif prop_type == 'number':
            if not isinstance(value, (int, float)):
                raise ValueError(f"Expected number, got {type(value).__name__}")
        elif prop_type == 'boolean':
            if not isinstance(value, bool):
                raise ValueError(f"Expected boolean, got {type(value).__name__}")
        elif prop_type == 'array':
            if not isinstance(value, list):
                raise ValueError(f"Expected array, got {type(value).__name__}")
        elif prop_type == 'object':
            if not isinstance(value, dict):
                raise ValueError(f"Expected object, got {type(value).__name__}")
    
    def transform_json(self, data: dict, operations: list = None) -> dict:
        """Apply transformations to JSON data"""
        if not operations:
            return data
        
        for op in operations:
            op_type = op.get('operation')
            
            if op_type == 'extract':
                data = self.extract_path(data, op.get('path'))
            elif op_type == 'filter':
                data = self.filter_array(data, op)
            elif op_type == 'rename':
                data = self.rename_key(data, op.get('old_key'), op.get('new_key'))
            elif op_type == 'add_field':
                data = self.add_field(data, op)
            elif op_type == 'remove_field':
                data = self.remove_field(data, op.get('field'))
        
        return data
    
    def extract_path(self, data: dict, path: str):
        """Extract value using JSONPath-like syntax"""
        # Simple path extraction (e.g., $.address.city)
        path = path.replace('$.', '')
        parts = path.split('.')
        
        result = data
        for part in parts:
            if isinstance(result, dict):
                result = result.get(part)
            elif isinstance(result, list) and part == '*':
                break
            else:
                return None
        
        return result
    
    def filter_array(self, data: dict, operation: dict):
        """Filter array elements"""
        path = operation.get('path', '').replace('$.', '')
        field = operation.get('field')
        value = operation.get('value')
        
        parts = path.split('.')
        result = data
        for part in parts[:-1]:
            if isinstance(result, dict):
                result = result.get(part)
        
        array = result.get(parts[-1], [])
        filtered = [item for item in array if item.get(field) == value]
        
        result[parts[-1]] = filtered
        return data
    
    def rename_key(self, data: dict, old_key: str, new_key: str):
        """Rename a key in the data"""
        if isinstance(data, dict) and old_key in data:
            data[new_key] = data.pop(old_key)
        return data
    
    def add_field(self, data: dict, operation: dict):
        """Add a new field to the data"""
        field = operation.get('field')
        value = operation.get('value')
        
        if value == 'CURRENT_DATE':
            value = datetime.now().isoformat()
        
        data[field] = value
        return data
    
    def remove_field(self, data: dict, field: str):
        """Remove a field from the data"""
        if isinstance(data, dict) and field in data:
            data.pop(field)
        return data
    
    def generate_filename(self, filename: str = None) -> str:
        """Generate a filename with timestamp if not provided"""
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = filename or f"json_{timestamp}"
        if not filename.endswith('.json'):
            filename += '.json'
        return filename
