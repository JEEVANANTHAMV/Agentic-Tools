import xml.etree.ElementTree as ET
from xml.dom import minidom
from datetime import datetime
from io import BytesIO
from config import settings

class XMLParser:
    def __init__(self):
        self.default_font_name = settings.DEFAULT_FONT_NAME
        self.default_font_size = settings.DEFAULT_FONT_SIZE
    
    def parse_xml(self, content: str, transform: str = None, filename: str = None) -> BytesIO:
        """Parse and optionally transform XML content and return as BytesIO"""
        try:
            # Parse XML
            root = ET.fromstring(content)
            
            # Apply transformation if provided
            if transform:
                root = self.apply_transform(root, transform)
            
            # Pretty print
            xml_str = self.pretty_print(root)
            
            # Convert to BytesIO
            output = BytesIO(xml_str.encode('utf-8'))
            output.seek(0)
            
            return output
            
        except ET.ParseError as e:
            raise ValueError(f"Invalid XML: {str(e)}")
    
    def pretty_print(self, elem):
        """Pretty print XML element"""
        rough_string = ET.tostring(elem, encoding='unicode')
        reparsed = minidom.parseString(rough_string)
        return reparsed.toprettyxml(indent="  ", encoding='utf-8').decode('utf-8')
    
    def apply_transform(self, root, transform: str):
        """Apply XSLT-like transformation"""
        # Simple transformation support
        if transform.startswith('[EXTRACT:'):
            path = transform[10:-1]
            return self.extract_path(root, path)
        elif transform.startswith('[FILTER:'):
            parts = transform[7:-1].split(':')
            if len(parts) >= 3:
                return self.filter_elements(root, parts[0], parts[1], parts[2])
        
        return root
    
    def extract_path(self, root, path: str):
        """Extract elements using XPath-like syntax"""
        # Convert simple path to XPath
        xpath = path.replace('/', './/')
        
        try:
            elements = root.findall(xpath)
            if elements:
                # Create new root with extracted elements
                new_root = ET.Element('results')
                for elem in elements:
                    new_root.append(elem)
                return new_root
        except:
            pass
        
        return root
    
    def filter_elements(self, root, tag: str, field: str, value: str):
        """Filter elements based on field value"""
        # Find all elements with the given tag
        elements = root.findall(f'.//{tag}')
        
        # Filter based on field value
        filtered = []
        for elem in elements:
            # Check if element has the field as attribute or child
            if elem.get(field) == value:
                filtered.append(elem)
            else:
                child = elem.find(field)
                if child is not None and child.text == value:
                    filtered.append(elem)
        
        # Create new root with filtered elements
        new_root = ET.Element('filtered_results')
        for elem in filtered:
            new_root.append(elem)
        
        return new_root
    
    def validate_xml(self, content: str, schema: str = None):
        """Validate XML against XSD schema"""
        try:
            root = ET.fromstring(content)
            
            if schema:
                # XSD validation would go here
                # For now, just validate that it's well-formed
                pass
            
            return True
        except ET.ParseError:
            return False
    
    def extract_value(self, root, path: str):
        """Extract a single value using XPath"""
        xpath = path.replace('/', './/')
        
        try:
            elem = root.find(xpath)
            if elem is not None:
                return elem.text
        except:
            pass
        
        return None
    
    def generate_filename(self, filename: str = None) -> str:
        """Generate a filename with timestamp if not provided"""
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = filename or f"xml_{timestamp}"
        if not filename.endswith('.xml'):
            filename += '.xml'
        return filename
