from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
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
        
        # Color constants for fallback
        self.black = RGBColor(0, 0, 0)
        self.white = RGBColor(255, 255, 255)
        self.blue = RGBColor(66, 135, 245)
    
    def create_presentation(self, content: str, filename: str = None) -> BytesIO:
        """Create a PowerPoint presentation from HTML content and return as BytesIO"""
        prs = Presentation()
        prs.slide_width = self.slide_width
        prs.slide_height = self.slide_height
        
        soup = BeautifulSoup(content, 'html.parser')
        slides = soup.find_all('div', class_='slide')
        
        if not slides:
            self.create_slide_from_content(prs, soup, 1)
        else:
            for idx, slide_soup in enumerate(slides, 1):
                self.create_slide_from_content(prs, slide_soup, idx)
        
        prs_stream = BytesIO()
        prs.save(prs_stream)
        prs_stream.seek(0)
        return prs_stream
    
    def apply_slide_background(self, slide, slide_soup=None):
        """Apply background to slide from style or default to white"""
        try:
            fill = slide.background.fill
            fill.solid()
            
            if slide_soup:
                styles = self.parse_style(slide_soup)
                bg_color = self.parse_color(styles.get('background-color'))
                if bg_color:
                    fill.fore_color.rgb = bg_color
                    return

            fill.fore_color.rgb = self.white
        except Exception as e:
            print(f"Error applying background: {e}")
    
    def parse_color(self, color_str):
        """Parse color from hex, rgb, or basic names"""
        if not color_str:
            return None
        
        color_str = color_str.strip().lower()
        
        named_colors = {
            'white': RGBColor(255, 255, 255),
            'black': RGBColor(0, 0, 0),
            'red': RGBColor(255, 0, 0),
            'green': RGBColor(0, 255, 0),
            'blue': RGBColor(0, 0, 255),
            'gray': RGBColor(128, 128, 128),
            'darkgray': RGBColor(64, 64, 64),
            'lightgray': RGBColor(211, 211, 211)
        }
        
        if color_str in named_colors:
            return named_colors[color_str]
        
        hex_match = re.match(r'#?([0-9a-f]{6})', color_str)
        if hex_match:
            hex_val = hex_match.group(1)
            return RGBColor(int(hex_val[0:2], 16), int(hex_val[2:4], 16), int(hex_val[4:6], 16))
        
        rgb_match = re.match(r'rgb\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)', color_str)
        if rgb_match:
            return RGBColor(int(rgb_match.group(1)), int(rgb_match.group(2)), int(rgb_match.group(3)))
        
        return None
    
    def parse_dimension(self, dim_str, reference_size):
        if not dim_str: return None
        dim_str = dim_str.strip().lower()
        try:
            if dim_str.endswith('%'):
                return int(reference_size * float(dim_str.replace('%', '')) / 100.0)
            elif dim_str.endswith('px'):
                return Inches(float(dim_str.replace('px', '')) / 96.0)
            else:
                return Inches(float(dim_str))
        except ValueError:
            return None

    def parse_style(self, element):
        style_str = element.get('style', '')
        styles = {}
        if style_str:
            for declaration in style_str.split(';'):
                if ':' in declaration:
                    key, value = declaration.split(':', 1)
                    styles[key.strip().lower()] = value.strip()
        return styles
    
    def create_slide_from_content(self, prs, slide_soup, slide_num):
        is_custom = any('position' in self.parse_style(d) for d in slide_soup.find_all(['div', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'p', 'ul', 'ol', 'img', 'table']))
        
        slide_layout = prs.slide_layouts[6] if is_custom else prs.slide_layouts[1]
        slide = prs.slides.add_slide(slide_layout)
        
        self.apply_slide_background(slide, slide_soup)
        
        title_tag = slide_soup.find(['h1', 'h2', 'h3', 'h4', 'h5', 'h6'])
        if title_tag:
            self.add_element_to_slide(title_tag, slide, is_custom, is_title=True)
            title_tag.decompose()
        
        if is_custom:
            self.process_content(slide_soup, None, slide)
        else:
            content_placeholder = slide.placeholders[1] if len(slide.placeholders) > 1 else None
            if content_placeholder and content_placeholder.text_frame:
                for paragraph in content_placeholder.text_frame.paragraphs:
                    p = paragraph._p
                    p.getparent().remove(p)
                self.process_content(slide_soup, content_placeholder.text_frame, slide)
            else:
                self.process_content(slide_soup, None, slide)

    def process_content(self, soup, text_frame=None, slide=None):
        for element in soup.children:
            if hasattr(element, 'name') and element.name:
                tag_name = element.name.lower()
                styles = self.parse_style(element)
                
                if tag_name == 'div':
                    if 'position' in styles or 'background-color' in styles:
                        new_shape = self.add_box(element, slide, styles)
                        if new_shape and hasattr(new_shape, 'text_frame'):
                            self.process_content(element, new_shape.text_frame, slide)
                    else:
                        self.process_content(element, text_frame, slide)
                elif tag_name in ['h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'p']:
                    self.add_element_to_slide(element, slide, text_frame is None, text_frame)
                elif tag_name in ['ul', 'ol']:
                    self.add_list(element, text_frame, slide)
                elif tag_name == 'img':
                    self.add_image(element, slide)
                elif tag_name == 'table':
                    self.add_table(element, slide)

    def add_element_to_slide(self, element, slide, create_new=True, text_frame=None, is_title=False):
        styles = self.parse_style(element)
        tag_name = element.name.lower()
        
        target_frame = None
        if create_new:
            left = self.parse_dimension(styles.get('left'), self.slide_width) or Inches(0.5)
            top = self.parse_dimension(styles.get('top'), self.slide_height) or (Inches(0.2) if is_title else Inches(1.5))
            width = self.parse_dimension(styles.get('width'), self.slide_width) or (self.slide_width - 2 * left)
            height = self.parse_dimension(styles.get('height'), self.slide_height) or Inches(0.8)
            txBox = slide.shapes.add_textbox(left, top, width, height)
            target_frame = txBox.text_frame
        elif is_title and slide.shapes.title:
            target_frame = slide.shapes.title.text_frame
        else:
            target_frame = text_frame
            
        if not target_frame:
            return

        p = target_frame.add_paragraph()
        self.add_rich_text(element, p)
        p.font.name = self.default_font_name
        
        # Sizing logic
        if is_title:
            p.font.size = Pt(36)
            p.font.bold = True
        else:
            level_map = {'h1': 32, 'h2': 28, 'h3': 24, 'h4': 20, 'p': 16}
            p.font.size = Pt(level_map.get(tag_name, 16))
            if tag_name.startswith('h'): p.font.bold = True

        if styles.get('text-align') == 'center':
            p.alignment = PP_ALIGN.CENTER
        elif styles.get('text-align') == 'right':
            p.alignment = PP_ALIGN.RIGHT
            
        color = self.parse_color(styles.get('color'))
        if color:
            p.font.color.rgb = color
        else:
            p.font.color.rgb = self.black

    def add_box(self, element, slide, styles):
        left = self.parse_dimension(styles.get('left'), self.slide_width) or Inches(1)
        top = self.parse_dimension(styles.get('top'), self.slide_height) or Inches(1)
        width = self.parse_dimension(styles.get('width'), self.slide_width) or Inches(4)
        height = self.parse_dimension(styles.get('height'), self.slide_height) or Inches(3)
        
        shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if 'border-radius' in styles else MSO_SHAPE.RECTANGLE
        shape = slide.shapes.add_shape(shape_type, left, top, width, height)
        
        bg_color = self.parse_color(styles.get('background-color'))
        if bg_color:
            shape.fill.solid()
            shape.fill.fore_color.rgb = bg_color
        else:
            shape.fill.background()
            
        border_color = self.parse_color(styles.get('border-color'))
        if border_color:
            shape.line.color.rgb = border_color
            shape.line.width = Pt(1)
        else:
            shape.line.fill.background()
            
        if styles.get('align-items') == 'center':
            shape.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
            
        return shape

    def add_rich_text(self, element, paragraph):
        if not list(element.children):
            run = paragraph.add_run()
            run.text = element.get_text().strip()
            return

        for child in element.children:
            if isinstance(child, str):
                if child.strip():
                    run = paragraph.add_run()
                    run.text = child
            else:
                run = paragraph.add_run()
                run.text = child.get_text()
                styles = self.parse_style(child)
                if child.name in ['strong', 'b']: run.font.bold = True
                if child.name in ['em', 'i']: run.font.italic = True
                color = self.parse_color(styles.get('color'))
                if color: run.font.color.rgb = color
                sz = styles.get('font-size')
                if sz:
                    try: run.font.size = Pt(int(float(sz.replace('px', ''))))
                    except: pass

    def add_list(self, element, text_frame, slide):
        items = element.find_all('li', recursive=False)
        styles = self.parse_style(element)
        target_frame = text_frame
        if not target_frame:
            left = self.parse_dimension(styles.get('left'), self.slide_width) or Inches(1)
            top = self.parse_dimension(styles.get('top'), self.slide_height) or Inches(1.5)
            width = self.parse_dimension(styles.get('width'), self.slide_width) or Inches(8)
            height = Inches(2)
            txBox = slide.shapes.add_textbox(left, top, width, height)
            target_frame = txBox.text_frame
        
        for item in items:
            p = target_frame.add_paragraph()
            self.add_rich_text(item, p)
            p.font.name = self.default_font_name
            p.font.size = Pt(18)
            p.level = 0 if element.name == 'ol' else 1
            color = self.parse_color(self.parse_style(item).get('color'))
            p.font.color.rgb = color if color else self.black

    def add_image(self, element, slide):
        src = element.get('src')
        if not src: return
        try:
            img_data = BytesIO(requests.get(src).content) if src.startswith('http') else src
            styles = self.parse_style(element)
            w = self.parse_dimension(element.get('width') or styles.get('width'), self.slide_width) or Inches(4)
            h = self.parse_dimension(element.get('height') or styles.get('height'), self.slide_height) or Inches(3)
            l = self.parse_dimension(styles.get('left'), self.slide_width) or (self.slide_width - w) / 2
            t = self.parse_dimension(styles.get('top'), self.slide_height) or (self.slide_height - h) / 2
            slide.shapes.add_picture(img_data, l, t, width=w, height=h)
        except: pass

    def add_table(self, element, slide):
        rows_tags = element.find_all('tr')
        if not rows_tags: return
        styles = self.parse_style(element)
        cols = max(len(row.find_all(['td', 'th'])) for row in rows_tags)
        l = self.parse_dimension(styles.get('left'), self.slide_width) or Inches(0.5)
        t = self.parse_dimension(styles.get('top'), self.slide_height) or Inches(1.2)
        w = self.parse_dimension(styles.get('width'), self.slide_width) or (self.slide_width - 2 * l)
        table = slide.shapes.add_table(len(rows_tags), cols, l, t, w, Inches(len(rows_tags) * 0.4)).table
        for r_idx, row in enumerate(rows_tags):
            for c_idx, cell in enumerate(row.find_all(['td', 'th'])):
                if c_idx < cols:
                    p = table.cell(r_idx, c_idx).text_frame.paragraphs[0]
                    self.add_rich_text(cell, p)
                    p.font.name = self.default_font_name
                    p.font.size = Pt(12)
                    if cell.name == 'th':
                        table.cell(r_idx, c_idx).fill.solid()
                        table.cell(r_idx, c_idx).fill.fore_color.rgb = self.blue
                        p.font.bold, p.font.color.rgb = True, self.white
                    else:
                        if r_idx % 2 == 0:
                            table.cell(r_idx, c_idx).fill.solid()
                            table.cell(r_idx, c_idx).fill.fore_color.rgb = RGBColor(245, 248, 255)
                        p.font.color.rgb = self.black

    def generate_filename(self, filename: str = None) -> str:
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = filename or f"presentation_{timestamp}"
        if not filename.endswith('.pptx'): filename += '.pptx'
        return filename