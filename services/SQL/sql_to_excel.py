import pandas as pd
from sqlalchemy import create_engine, text
from io import BytesIO
from datetime import datetime
from typing import List
from config import settings

class SQLToExcelService:
    def __init__(self):
        # Initialize database connection
        self.db_connection_string = (
            f"mysql+mysqlconnector://{settings.DB_USER}:{settings.DB_PASSWORD}@"
            f"{settings.DB_HOST}:{settings.DB_PORT}/{settings.DB_NAME}"
        )
        self.engine = create_engine(self.db_connection_string)
    
    def execute_query_to_excel(self, query: str, filename: str = None) -> BytesIO:
        """
        Execute SQL query and return results as Excel file in BytesIO format
        """
        try:
            # Execute the query and get results
            df = pd.read_sql_query(text(query), self.engine)
            
            # Create a BytesIO object to store the Excel file
            excel_stream = BytesIO()
            
            # Use ExcelWriter to write to the BytesIO object
            with pd.ExcelWriter(excel_stream, engine='openpyxl') as writer:
                # Write the query text first
                df_query_text = pd.DataFrame({'Query Executed': [query]})
                df_query_text.to_excel(
                    writer, 
                    sheet_name='Results', 
                    startrow=0, 
                    startcol=0, 
                    index=False, 
                    header=True
                )
                
                # Write the results table below the query text
                df.to_excel(
                    writer, 
                    sheet_name='Results', 
                    startrow=2,  # Start 2 rows below the query text
                    startcol=0, 
                    index=False, 
                    header=True
                )
            
            # Reset the stream position to the beginning
            excel_stream.seek(0)
            
            return excel_stream
            
        except Exception as e:
            raise Exception(f"Error executing SQL query: {str(e)}")
    
    def execute_multiple_queries_to_excel(self, queries: List[str], filename: str = None) -> BytesIO:
        """
        Execute multiple SQL queries and return results as Excel file with proper spacing
        """
        try:
            # Create a BytesIO object to store the Excel file
            excel_stream = BytesIO()
            
            # Use ExcelWriter to write to the BytesIO object
            with pd.ExcelWriter(excel_stream, engine='openpyxl') as writer:
                row_offset = 0
                
                for i, query in enumerate(queries):
                    # Execute the query and get results
                    df = pd.read_sql_query(text(query), self.engine)
                    
                    # Write the query text first
                    df_query_text = pd.DataFrame({'Query Executed': [query]})
                    df_query_text.to_excel(
                        writer, 
                        sheet_name='Results', 
                        startrow=row_offset, 
                        startcol=0, 
                        index=False, 
                        header=True
                    )
                    
                    # Write the results table below the query text
                    table_start_row = row_offset + 2
                    df.to_excel(
                        writer, 
                        sheet_name='Results', 
                        startrow=table_start_row, 
                        startcol=0, 
                        index=False, 
                        header=True
                    )
                    
                    # Update the row offset for the next query
                    # Current offset + query text row (1) + blank line (1) + results table + header (1) + 10 empty rows
                    row_offset = table_start_row + len(df) + 1 + 10
            
            # Reset the stream position to the beginning
            excel_stream.seek(0)
            
            return excel_stream
            
        except Exception as e:
            raise Exception(f"Error executing SQL queries: {str(e)}")
    
    def generate_filename(self, filename: str = None) -> str:
        """Generate a filename with timestamp if not provided"""
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = filename or f"sql_results_{timestamp}"
        if not filename.endswith('.xlsx'):
            filename += '.xlsx'
        return filename
    
    def generate_object_name(self, filename: str) -> str:
        """Generate object name with date-based folder structure"""
        today = datetime.now()
        return f"{today.strftime('%Y')}/{today.strftime('%m')}/{today.strftime('%d')}/{filename}"