import json
import base64
from datetime import datetime
from io import BytesIO
from config import settings

class VisualizationCreator:
    def __init__(self):
        self.default_font_name = settings.DEFAULT_FONT_NAME
        self.default_font_size = settings.DEFAULT_FONT_SIZE
    
    def create_visualization(self, data: str, chart_type: str, filename: str = None) -> BytesIO:
        """Create a visualization from data and return as HTML BytesIO"""
        try:
            # Parse data
            chart_data = json.loads(data)
            
            # Generate HTML with Chart.js
            html = self.generate_chart_html(chart_data, chart_type)
            
            # Convert to BytesIO
            output = BytesIO(html.encode('utf-8'))
            output.seek(0)
            
            return output
            
        except json.JSONDecodeError as e:
            raise ValueError(f"Invalid JSON data: {str(e)}")
    
    def generate_chart_html(self, data: dict, chart_type: str) -> str:
        """Generate HTML with embedded Chart.js chart"""
        chart_type_map = {
            'bar': 'bar',
            'horizontal_bar': 'bar',
            'line': 'line',
            'area': 'line',
            'pie': 'pie',
            'doughnut': 'doughnut',
            'scatter': 'scatter',
            'bubble': 'bubble',
            'polar': 'polarArea',
            'radar': 'radar'
        }
        
        js_type = chart_type_map.get(chart_type, 'bar')
        
        # Extract labels and datasets
        labels = data.get('labels', [])
        datasets = data.get('datasets', [])
        
        # If data is in simple format, convert to datasets format
        if 'data' in data and not datasets:
            datasets = [{
                'label': 'Dataset',
                'data': data.get('data', [])
            }]
        
        # Generate dataset JSON
        datasets_json = json.dumps(datasets)
        labels_json = json.dumps(labels)
        
        html = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{chart_type.title()} Chart</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
        body {{
            font-family: {self.default_font_name}, Arial, sans-serif;
            max-width: 1000px;
            margin: 0 auto;
            padding: 20px;
            background-color: #f5f5f5;
        }}
        .chart-container {{
            background-color: white;
            border-radius: 8px;
            padding: 20px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }}
        h1 {{
            color: #333;
            text-align: center;
        }}
        canvas {{
            max-height: 500px;
        }}
    </style>
</head>
<body>
    <h1>{chart_type.title()} Chart</h1>
    <div class="chart-container">
        <canvas id="myChart"></canvas>
    </div>
    <script>
        const ctx = document.getElementById('myChart').getContext('2d');
        const chart = new Chart(ctx, {{
            type: '{js_type}',
            data: {{
                labels: {labels_json},
                datasets: {datasets_json}
            }},
            options: {{
                responsive: true,
                maintainAspectRatio: true,
                plugins: {{
                    legend: {{
                        position: 'top',
                    }},
                    title: {{
                        display: true,
                        text: '{chart_type.title()} Visualization'
                    }}
                }},
                scales: {{
                    y: {{
                        beginAtZero: true
                    }}
                }}
            }}
        }});
    </script>
</body>
</html>"""
        
        return html
    
    def create_static_image(self, data: dict, chart_type: str) -> BytesIO:
        """Create a static image of the chart (placeholder - would use matplotlib in production)"""
        # For now, return a simple SVG placeholder
        svg = """<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" width="800" height="600">
    <rect width="800" height="600" fill="#f5f5f5"/>
    <text x="400" y="300" text-anchor="middle" font-family="Arial" font-size="24">
        Chart visualization would appear here
    </text>
    <text x="400" y="340" text-anchor="middle" font-family="Arial" font-size="16" fill="#666">
        (Static image generation requires additional dependencies)
    </text>
</svg>"""
        
        output = BytesIO(svg.encode('utf-8'))
        output.seek(0)
        
        return output
    
    def generate_filename(self, filename: str = None) -> str:
        """Generate a filename with timestamp if not provided"""
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = filename or f"chart_{timestamp}"
        if not filename.endswith('.html'):
            filename += '.html'
        return filename
