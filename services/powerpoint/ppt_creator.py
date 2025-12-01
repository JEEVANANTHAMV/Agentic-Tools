from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from bs4 import BeautifulSoup
import re
import requests
from io import BytesIO
from datetime import datetime
from config import settings
import os

class PresentationCreator:
    def __init__(self):
        self.default_font_name = settings.DEFAULT_FONT_NAME
        self.default_font_size = settings.DEFAULT_FONT_SIZE
        self.slide_width = Inches(10)  # Standard 16:9 aspect ratio
        self.slide_height = Inches(5.625)
    
    def create_presentation(self, content: str, filename: str = None) -> BytesIO:
        """Create a PowerPoint presentation from HTML content and return as BytesIO"""
        # Create presentation
        prs = Presentation()
        
        # Set slide size to 16:9 aspect ratio
        prs.slide_width = self.slide_width
        prs.slide_height = self.slide_height
        
        # Parse HTML content
        soup = BeautifulSoup(content, 'html.parser')
        
        # Extract slides
        slides = soup.find_all('div', class_='slide')
        
        # If no slides found, treat the entire content as one slide
        if not slides:
            self.create_slide_from_content(prs, soup)
        else:
            for slide_soup in slides:
                self.create_slide_from_content(prs, slide_soup)
        
        # Save to BytesIO
        prs_stream = BytesIO()
        prs.save(prs_stream)
        prs_stream.seek(0)
        
        return prs_stream
    
    def create_slide_from_content(self, prs, slide_soup):
        """Create a slide from parsed HTML content"""
        # Add a slide
        slide_layout = prs.slide_layouts[1]  # Title and Content layout
        slide = prs.slides.add_slide(slide_layout)
        
        # Process content
        title = slide_soup.find(['h1', 'h2', 'h3', 'h4', 'h5', 'h6'])
        if title:
            title_text = title.get_text().strip()
            slide.shapes.title.text = title_text
            title_font = slide.shapes.title.text_frame.paragraphs[0].font
            title_font.name = self.default_font_name
            title_font.size = Pt(self.default_font_size + 4)
            title_font.bold = True
            title.decompose()  # Remove title from content
        
        # Get content placeholder
        content_placeholder = slide.placeholders[1] if len(slide.placeholders) > 1 else None
        
        if content_placeholder:
            # Clear existing content
            for paragraph in content_placeholder.text_frame.paragraphs:
                p = paragraph._p
                p.getparent().remove(p)
            
            # Process remaining content
            self.process_content(slide_soup, content_placeholder.text_frame)
        else:
            # If no placeholder, add content directly to slide
            self.process_content(slide_soup, None, slide)
    
    def process_content(self, soup, text_frame=None, slide=None):
        """Process HTML content and add to slide"""
        if not text_frame and not slide:
            return
        
        # Process each element
        for element in soup.children:
            if hasattr(element, 'name'):
                tag_name = element.name.lower()
                
                # Handle headings
                if tag_name in ['h1', 'h2', 'h3', 'h4', 'h5', 'h6']:
                    self.add_heading(element, text_frame, slide, tag_name)
                
                # Handle paragraphs
                elif tag_name == 'p':
                    self.add_paragraph(element, text_frame, slide)
                
                # Handle lists
                elif tag_name in ['ul', 'ol']:
                    self.add_list(element, text_frame, slide)
                
                # Handle images
                elif tag_name == 'img':
                    self.add_image(element, slide)
                
                # Handle tables
                elif tag_name == 'table':
                    self.add_table(element, slide)
                
                # Handle divs (recursive)
                elif tag_name == 'div':
                    self.process_content(element, text_frame, slide)
    
    def add_heading(self, element, text_frame, slide, tag_name):
        """Add a heading to the slide"""
        text = element.get_text().strip()
        if not text:
            return
        
        if text_frame:
            p = text_frame.add_paragraph()
        else:
            left = Inches(1)
            top = Inches(1.5)
            width = Inches(8)
            height = Inches(0.5)
            txBox = slide.shapes.add_textbox(left, top, width, height)
            p = txBox.text_frame.add_paragraph()
        
        p.text = text
        p.font.name = self.default_font_name
        
        # Set font size based on heading level
        if tag_name == 'h1':
            p.font.size = Pt(self.default_font_size + 6)
            p.font.bold = True
        elif tag_name == 'h2':
            p.font.size = Pt(self.default_font_size + 4)
            p.font.bold = True
        elif tag_name == 'h3':
            p.font.size = Pt(self.default_font_size + 2)
            p.font.bold = True
        else:
            p.font.size = Pt(self.default_font_size)
            p.font.bold = True
    
    def add_paragraph(self, element, text_frame, slide):
        """Add a paragraph to the slide"""
        text = element.get_text().strip()
        if not text:
            return
        
        if text_frame:
            p = text_frame.add_paragraph()
        else:
            left = Inches(1)
            top = Inches(1.5)
            width = Inches(8)
            height = Inches(0.5)
            txBox = slide.shapes.add_textbox(left, top, width, height)
            p = txBox.text_frame.add_paragraph()
        
        p.text = text
        p.font.name = self.default_font_name
        p.font.size = Pt(self.default_font_size)
        
        # Check for alignment
        if element.get('align') == 'center':
            p.alignment = PP_ALIGN.CENTER
        elif element.get('align') == 'right':
            p.alignment = PP_ALIGN.RIGHT
    
    def add_list(self, element, text_frame, slide):
        """Add a list to the slide"""
        items = element.find_all('li', recursive=False)
        if not items:
            return
        
        for item in items:
            text = item.get_text().strip()
            if not text:
                continue
            
            if text_frame:
                p = text_frame.add_paragraph()
            else:
                left = Inches(1.5) if element.name == 'ul' else Inches(1)
                top = Inches(1.5)
                width = Inches(7)
                height = Inches(0.5)
                txBox = slide.shapes.add_textbox(left, top, width, height)
                p = txBox.text_frame.add_paragraph()
            
            p.text = text
            p.font.name = self.default_font_name
            p.font.size = Pt(self.default_font_size)
            
            # Set list level
            if element.name == 'ul':
                p.level = 1
            elif element.name == 'ol':
                p.level = 0
    
    def add_image(self, element, slide):
        """Add an image to the slide"""
        src = element.get('src')
        if not src:
            return
        
        try:
            # If it's a URL, download the image
            if src.startswith('http'):
                response = requests.get(src)
                img_data = BytesIO(response.content)
            else:
                # If it's a local path, open the file
                img_data = src
            
            # Get image dimensions
            img_width = Inches(6)  # Default width
            img_height = Inches(4)  # Default height
            
            # Check for custom dimensions
            width_attr = element.get('width')
            height_attr = element.get('height')
            
            if width_attr:
                img_width = Inches(float(width_attr))
            if height_attr:
                img_height = Inches(float(height_attr))
            
            # Calculate position (centered)
            left = (self.slide_width - img_width) / 2
            top = (self.slide_height - img_height) / 2
            
            # Add image to slide
            slide.shapes.add_picture(img_data, left, top, width=img_width, height=img_height)
        except Exception as e:
            print(f"Error adding image: {e}")
    
    def add_table(self, element, slide):
        """Add a table to the slide"""
        rows = element.find_all('tr')
        if not rows:
            return
        
        # Determine number of columns
        cols = max(len(row.find_all(['td', 'th'])) for row in rows)
        
        # Calculate table dimensions and position
        left = Inches(1)
        top = Inches(1.5)
        width = Inches(8)
        height = Inches(4)
        
        # Add table to slide
        table = slide.shapes.add_table(rows=len(rows), cols=cols, left=left, top=top, width=width, height=height).table
        
        # Fill table with data
        for row_idx, row in enumerate(rows):
            cells = row.find_all(['td', 'th'])
            for col_idx, cell in enumerate(cells):
                if col_idx < cols:
                    text = cell.get_text().strip()
                    table.cell(row_idx, col_idx).text = text
                    
                    # Format header cells
                    if cell.name == 'th':
                        table.cell(row_idx, col_idx).fill.solid()
                        table.cell(row_idx, col_idx).fill.fore_color.rgb = RGBColor(79, 129, 189)  # Blue
                        table.cell(row_idx, col_idx).text_frame.paragraphs[0].font.bold = True
                        table.cell(row_idx, col_idx).text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)  # White
    
    def generate_filename(self, filename: str = None) -> str:
        """Generate a filename with timestamp if not provided"""
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = filename or f"presentation_{timestamp}"
        if not filename.endswith('.pptx'):
            filename += '.pptx'
        return filename
    
    def generate_object_name(self, filename: str) -> str:
        """Generate MinIO object name with date-based folder structure"""
        today = datetime.now()
        return f"{today.strftime('%Y')}/{today.strftime('%m')}/{today.strftime('%d')}/{filename}"