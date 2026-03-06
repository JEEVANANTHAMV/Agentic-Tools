from reportlab.lib.pagesizes import A4, letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.colors import HexColor, black
from reportlab.pdfgen import canvas
import re
from datetime import datetime
from io import BytesIO
from config import settings

class PDFCreator:
    def __init__(self):
        self.default_font_name = settings.DEFAULT_FONT_NAME
        self.default_font_size = settings.DEFAULT_FONT_SIZE
    
    def create_pdf_from_content(self, content: str, filename: str = None) -> BytesIO:
        """Create a PDF from string content and return as BytesIO"""
        # Create BytesIO buffer
        pdf_buffer = BytesIO()
        
        # Create document
        doc = SimpleDocTemplate(
            pdf_buffer,
            pagesize=A4,
            rightMargin=0.75*inch,
            leftMargin=0.75*inch,
            topMargin=0.75*inch,
            bottomMargin=0.75*inch
        )
        
        # Get styles
        styles = getSampleStyleSheet()
        
        # Build story (content)
        story = []
        
        # Parse and add content
        self.parse_and_format_content(story, content, styles)
        
        # Build PDF
        doc.build(story)
        
        # Reset buffer position
        pdf_buffer.seek(0)
        
        return pdf_buffer
    
    def generate_filename(self, filename: str = None) -> str:
        """Generate a filename with timestamp if not provided"""
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = filename or f"document_{timestamp}"
        if not filename.endswith('.pdf'):
            filename += '.pdf'
        return filename
    
    def parse_and_format_content(self, story, content, styles):
        """
        Parse content string and create PDF elements
        Supports:
        # Heading 1
        ## Heading 2
        ### Heading 3
        [BOLD]Bold text[/BOLD]
        [ITALIC]Italic text[/ITALIC]
        [COLOR:RRGGBB]Colored text[/COLOR]
        - Bullet point
        1. Numbered list
        |Header1|Header2| for tables
        [PAGE:orientation,size,margins]
        [PAGEBREAK]
        """
        lines = content.split('\n')
        i = 0
        
        while i < len(lines):
            line = lines[i].strip()
            
            if not line:
                story.append(Spacer(1, 12))
                i += 1
                continue
            
            # Check for page settings
            if line.startswith('[PAGE:'):
                i += 1
                continue
            
            # Check for page break
            if line == '[PAGEBREAK]':
                story.append(PageBreak())
                i += 1
                continue
            
            # Check if this is a table
            if line.startswith('|') and '|' in line[1:]:
                # Collect all table lines
                table_lines = [line]
                j = i + 1
                while j < len(lines) and lines[j].strip().startswith('|'):
                    table_lines.append(lines[j].strip())
                    j += 1
                
                # Create table
                table = self.create_table_from_markdown(table_lines, styles)
                if table:
                    story.append(table)
                    story.append(Spacer(1, 12))
                i = j
                continue
            
            # Handle headings
            if line.startswith('###'):
                text = line.replace('###', '').strip()
                text = self.process_inline_formatting(text)
                p = Paragraph(text, styles['Heading3'])
                story.append(p)
                story.append(Spacer(1, 12))
            elif line.startswith('##'):
                text = line.replace('##', '').strip()
                text = self.process_inline_formatting(text)
                p = Paragraph(text, styles['Heading2'])
                story.append(p)
                story.append(Spacer(1, 12))
            elif line.startswith('#'):
                text = line.replace('#', '').strip()
                text = self.process_inline_formatting(text)
                p = Paragraph(text, styles['Heading1'])
                story.append(p)
                story.append(Spacer(1, 12))
            
            # Handle bullet points
            elif line.startswith('- ') or line.startswith('* '):
                text = line[2:] if line[1] == ' ' else line[1:]
                text = self.process_inline_formatting(text)
                p = Paragraph(f'&bull; {text}', styles['Normal'])
                story.append(p)
            
            # Handle numbered lists
            elif re.match(r'^\d+\.\s', line):
                text = re.sub(r'^\d+\.\s', '', line)
                text = self.process_inline_formatting(text)
                p = Paragraph(text, styles['Normal'])
                story.append(p)
            
            # Regular paragraph
            else:
                text = self.process_inline_formatting(line)
                p = Paragraph(text, styles['Normal'])
                story.append(p)
                story.append(Spacer(1, 6))
            
            i += 1
    
    def process_inline_formatting(self, text):
        """Process inline formatting like bold, italic, color"""
        # Process [BOLD]...[/BOLD]
        bold_pattern = r'\[BOLD\](.*?)\[/BOLD\]'
        text = re.sub(bold_pattern, r'<b>\1</b>', text)
        
        # Process [ITALIC]...[/ITALIC]
        italic_pattern = r'\[ITALIC\](.*?)\[/ITALIC\]'
        text = re.sub(italic_pattern, r'<i>\1</i>', text)
        
        # Process [COLOR:RRGGBB]...[/COLOR]
        color_pattern = r'\[COLOR:([A-Fa-f0-9]{6})\](.*?)\[/COLOR\]'
        def color_replace(match):
            color = match.group(1)
            content = match.group(2)
            return f'<font color="#{color}">{content}</font>'
        text = re.sub(color_pattern, color_replace, text)
        
        return text
    
    def create_table_from_markdown(self, table_lines, styles):
        """Create table from markdown syntax"""
        lines = [line.strip() for line in table_lines if line.strip()]
        
        if not lines:
            return None
        
        # Parse headers
        headers = [cell.strip() for cell in lines[0].split('|') if cell.strip()]
        
        # Parse data rows (skip separator line)
        data_rows = []
        for line in lines[2:]:  # Skip header and separator
            cells = [cell.strip() for cell in line.split('|') if cell.strip()]
            if cells:
                data_rows.append(cells)
        
        if not headers:
            return None
        
        # Combine headers and data
        data = [headers] + data_rows
        
        # Create table
        table = Table(data)
        
        # Style the table
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), HexColor('EEEEEE')),
            ('TEXTCOLOR', (0, 0), (-1, 0), black),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 12),
            ('FONTSIZE', (0, 1), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('TOPPADDING', (0, 0), (-1, -1), 6),
            ('LEFTPADDING', (0, 0), (-1, -1), 6),
            ('RIGHTPADDING', (0, 0), (-1, -1), 6),
            ('BACKGROUND', (0, 1), (-1, -1), HexColor('FFFFFF')),
            ('GRID', (0, 0), (-1, -1), 0.5, HexColor('000000')),
        ]))
        
        return table


class PageBreak:
    def __init__(self):
        pass
    
    def wrap(self, availWidth, availHeight):
        return 0, 0
    
    def draw(self):
        from reportlab.lib.pagesizes import A4
        canvas.canv.showPage()
