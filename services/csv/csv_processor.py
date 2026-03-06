import csv
import io
import re
import json
from datetime import datetime
from io import BytesIO
from config import settings

class CSVProcessor:
    def __init__(self):
        self.default_font_name = settings.DEFAULT_FONT_NAME
        self.default_font_size = settings.DEFAULT_FONT_SIZE
    
    def process_csv(self, content: str, operations: list = None, filename: str = None) -> BytesIO:
        """Process CSV content with optional operations and return as BytesIO"""
        # Parse content to extract CSV data and inline operations
        csv_data, inline_ops = self.parse_content(content)
        
        # Combine operations
        all_operations = (operations or []) + (inline_ops or [])
        
        # Read CSV data
        reader = csv.DictReader(io.StringIO(csv_data))
        rows = list(reader)
        fieldnames = reader.fieldnames
        
        if not rows:
            raise ValueError("No data found in CSV content")
        
        # Apply operations
        for op in all_operations:
            rows, fieldnames = self.apply_operation(op, rows, fieldnames)
        
        # Write to BytesIO
        output = BytesIO()
        output_str = io.StringIO()
        writer = csv.DictWriter(output_str, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)
        output.write(output_str.getvalue().encode('utf-8'))
        output.seek(0)
        
        return output
    
    def parse_content(self, content: str):
        """Parse content to extract CSV data and inline operations"""
        lines = content.split('\n')
        csv_lines = []
        operations = []
        
        for line in lines:
            line = line.strip()
            if line.startswith('#'):
                continue  # Skip comments
            elif line.startswith('[') and line.endswith(']'):
                # Parse inline operation
                op = self.parse_inline_operation(line)
                if op:
                    operations.append(op)
            elif line:
                csv_lines.append(line)
        
        csv_data = '\n'.join(csv_lines)
        return csv_data, operations if operations else None
    
    def parse_inline_operation(self, line: str):
        """Parse inline operation syntax"""
        # [FILTER:column:condition:value]
        if line.startswith('[FILTER:'):
            parts = line[8:-1].split(':')
            if len(parts) >= 4:
                return {
                    'operation': 'filter',
                    'column': parts[0],
                    'condition': parts[1],
                    'value': parts[2]
                }
        
        # [SORT:column:order]
        elif line.startswith('[SORT:'):
            parts = line[6:-1].split(':')
            if len(parts) >= 2:
                return {
                    'operation': 'sort',
                    'column': parts[0],
                    'order': parts[1]
                }
        
        # [SELECT:col1,col2,col3]
        elif line.startswith('[SELECT:'):
            cols = line[7:-1].split(',')
            return {
                'operation': 'select',
                'columns': [c.strip() for c in cols]
            }
        
        # [TRANSFORM:column:function]
        elif line.startswith('[TRANSFORM:'):
            parts = line[10:-1].split(':')
            if len(parts) >= 2:
                return {
                    'operation': 'transform',
                    'column': parts[0],
                    'function': parts[1]
                }
        
        return None
    
    def apply_operation(self, operation: dict, rows: list, fieldnames: list):
        """Apply a single operation to the data"""
        op_type = operation.get('operation')
        
        if op_type == 'filter':
            return self.filter_rows(operation, rows, fieldnames)
        elif op_type == 'sort':
            return self.sort_rows(operation, rows, fieldnames)
        elif op_type == 'select':
            return self.select_columns(operation, rows, fieldnames)
        elif op_type == 'transform':
            return self.transform_column(operation, rows, fieldnames)
        elif op_type == 'add_column':
            return self.add_column(operation, rows, fieldnames)
        elif op_type == 'remove_column':
            return self.remove_column(operation, rows, fieldnames)
        
        return rows, fieldnames
    
    def filter_rows(self, operation: dict, rows: list, fieldnames: list):
        """Filter rows based on condition"""
        column = operation.get('column')
        condition = operation.get('condition')
        value = operation.get('value')
        
        if column not in fieldnames:
            return rows, fieldnames
        
        filtered = []
        for row in rows:
            cell_value = row.get(column, '')
            if self.evaluate_condition(cell_value, condition, value):
                filtered.append(row)
        
        return filtered, fieldnames
    
    def evaluate_condition(self, cell_value: str, condition: str, value: str) -> bool:
        """Evaluate a condition"""
        try:
            # Try numeric comparison
            cell_num = float(cell_value)
            val_num = float(value)
            
            if condition == 'equals':
                return cell_num == val_num
            elif condition == 'not_equals':
                return cell_num != val_num
            elif condition == 'greater_than':
                return cell_num > val_num
            elif condition == 'less_than':
                return cell_num < val_num
            elif condition == 'greater_than_or_equal':
                return cell_num >= val_num
            elif condition == 'less_than_or_equal':
                return cell_num <= val_num
        except ValueError:
            # String comparison
            if condition == 'equals':
                return cell_value == value
            elif condition == 'not_equals':
                return cell_value != value
            elif condition == 'contains':
                return value in cell_value
            elif condition == 'starts_with':
                return cell_value.startswith(value)
            elif condition == 'ends_with':
                return cell_value.endswith(value)
        
        return False
    
    def sort_rows(self, operation: dict, rows: list, fieldnames: list):
        """Sort rows by column"""
        column = operation.get('column')
        order = operation.get('order', 'ascending')
        
        if column not in fieldnames:
            return rows, fieldnames
        
        reverse = order == 'descending'
        
        try:
            rows = sorted(rows, key=lambda x: float(x.get(column, 0)), reverse=reverse)
        except ValueError:
            rows = sorted(rows, key=lambda x: x.get(column, ''), reverse=reverse)
        
        return rows, fieldnames
    
    def select_columns(self, operation: dict, rows: list, fieldnames: list):
        """Select specific columns"""
        columns = operation.get('columns', [])
        new_fieldnames = [c for c in columns if c in fieldnames]
        new_rows = [{c: row.get(c, '') for c in new_fieldnames} for row in rows]
        
        return new_rows, new_fieldnames
    
    def transform_column(self, operation: dict, rows: list, fieldnames: list):
        """Transform column values"""
        column = operation.get('column')
        function = operation.get('function')
        
        if column not in fieldnames:
            return rows, fieldnames
        
        for row in rows:
            value = row.get(column, '')
            if function == 'uppercase':
                row[column] = value.upper()
            elif function == 'lowercase':
                row[column] = value.lower()
            elif function == 'trim':
                row[column] = value.strip()
            elif function == 'round':
                try:
                    row[column] = str(round(float(value)))
                except ValueError:
                    pass
            elif function == 'format_currency':
                try:
                    row[column] = f"${float(value):,.2f}"
                except ValueError:
                    pass
        
        return rows, fieldnames
    
    def add_column(self, operation: dict, rows: list, fieldnames: list):
        """Add a new column"""
        name = operation.get('name')
        formula = operation.get('formula', '')
        
        fieldnames = list(fieldnames) + [name]
        
        for row in rows:
            row[name] = ''
            if formula:
                # Simple formula evaluation
                try:
                    # Replace column names with values
                    eval_formula = formula
                    for col in fieldnames[:-1]:
                        eval_formula = eval_formula.replace(col, str(row.get(col, 0)))
                    row[name] = str(eval(eval_formula))
                except:
                    row[name] = ''
        
        return rows, fieldnames
    
    def remove_column(self, operation: dict, rows: list, fieldnames: list):
        """Remove columns"""
        columns = operation.get('columns', [])
        new_fieldnames = [c for c in fieldnames if c not in columns]
        
        for row in rows:
            for col in columns:
                row.pop(col, None)
        
        return rows, new_fieldnames
    
    def generate_filename(self, filename: str = None) -> str:
        """Generate a filename with timestamp if not provided"""
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = filename or f"csv_{timestamp}"
        if not filename.endswith('.csv'):
            filename += '.csv'
        return filename
