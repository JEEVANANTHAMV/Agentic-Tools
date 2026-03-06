# Guidelines for Data Visualizer API

## Tool Name - data_visualizer

## Basic Structure

The data visualizer allows you to create various types of charts and visualizations from data. It supports multiple chart types, custom styling, and various output formats.

## Input Format

### Data Format

Provide your data as a JSON string or CSV string:

```json
{
  "labels": ["January", "February", "March", "April"],
  "datasets": [
    {
      "label": "Sales",
      "data": [120, 150, 180, 200]
    },
    {
      "label": "Expenses",
      "data": [80, 100, 120, 140]
    }
  ]
}
```

Or CSV format:

```
month,sales,expenses
January,120,80
February,150,100
March,180,120
April,200,140
```

## Chart Types

### Bar Chart

```
[CHART:bar]
```

Options:
- `bar` - Vertical bar chart
- `horizontal_bar` - Horizontal bar chart
- `grouped_bar` - Grouped bar chart
- `stacked_bar` - Stacked bar chart

### Line Chart

```
[CHART:line]
```

Options:
- `line` - Standard line chart
- `area` - Area chart
- `spline` - Smooth line chart
- `step` - Step line chart

### Pie Chart

```
[CHART:pie]
```

Options:
- `pie` - Standard pie chart
- `doughnut` - Doughnut chart
- `polar` - Polar area chart

### Scatter Plot

```
[CHART:scatter]
```

Options:
- `scatter` - Standard scatter plot
- `bubble` - Bubble chart

### Other Charts

```
[CHART:radar]
[CHART:polar]
[CHART:funnel]
[CHART:gauge]
[CHART:heatmap]
[CHART:treemap]
```

## Styling Options

### Colors

```
[COLORS:#FF6384,#36A2EB,#FFCE56,#4BC0C0]
```

Or use color names:
```
[COLORS:red,blue,yellow,green]
```

### Title

```
[TITLE:Sales Report]
```

### Axis Labels

```
[X_LABEL:Month]
[Y_LABEL:Amount]
```

### Legend

```
[LEGEND:top]
[LEGEND:bottom]
[LEGEND:left]
[LEGEND:right]
[LEGEND:none]
```

### Grid

```
[GRID:enabled]
[GRID:disabled]
```

## Content Format with Options

The `data` parameter contains the data, and options can be included:

```json
{
  "labels": ["Q1", "Q2", "Q3", "Q4"],
  "data": [100, 150, 200, 250]
}

[CHART:bar]
[TITLE:Quarterly Sales]
[X_LABEL:Quarter]
[Y_LABEL:Revenue]
```

## Combining Options

You can combine multiple options:

```json
{
  "labels": ["Product A", "Product B", "Product C"],
  "datasets": [
    {"label": "2023", "data": [100, 150, 200]},
    {"label": "2024", "data": [120, 180, 220]}
  ]
}

[CHART:grouped_bar]
[TITLE:Product Comparison]
[COLORS:#FF6384,#36A2EB]
[LEGEND:top]
```

## Best Practices

1. Choose appropriate chart type for your data
2. Use clear and descriptive labels
3. Keep color schemes consistent
4. Avoid cluttering with too many data series
5. Include legends for multi-series charts
6. Use appropriate scales for axes
7. Consider accessibility (color blindness)
8. Test visualizations with sample data

## Example Content

### Example 1: Bar Chart

```json
{
  "labels": ["January", "February", "March", "April", "May"],
  "datasets": [
    {
      "label": "Sales",
      "data": [120, 150, 180, 200, 220]
    }
  ]
}

[CHART:bar]
[TITLE:Monthly Sales]
[X_LABEL:Month]
[Y_LABEL:Revenue ($)]
```

### Example 2: Line Chart

```json
{
  "labels": ["Week 1", "Week 2", "Week 3", "Week 4"],
  "datasets": [
    {
      "label": "Page Views",
      "data": [1000, 1200, 1500, 1800]
    },
    {
      "label": "Unique Visitors",
      "data": [500, 600, 750, 900]
    }
  ]
}

[CHART:line]
[TITLE:Website Traffic]
[LEGEND:top]
```

### Example 3: Pie Chart

```json
{
  "labels": ["Electronics", "Clothing", "Food", "Other"],
  "datasets": [
    {
      "data": [35, 25, 20, 20]
    }
  ]
}

[CHART:pie]
[TITLE:Sales by Category]
[COLORS:#FF6384,#36A2EB,#FFCE56,#4BC0C0]
```

### Example 4: Scatter Plot

```json
{
  "datasets": [
    {
      "label": "Products",
      "data": [
        {"x": 10, "y": 20},
        {"x": 15, "y": 32},
        {"x": 20, "y": 45},
        {"x": 25, "y": 55}
      ]
    }
  ]
}

[CHART:scatter]
[TITLE:Price vs Sales]
[X_LABEL:Price]
[Y_LABEL:Units Sold]
```

### Example 5: Area Chart

```json
{
  "labels": ["2020", "2021", "2022", "2023"],
  "datasets": [
    {
      "label": "Revenue",
      "data": [500, 750, 1000, 1250]
    }
  ]
}

[CHART:area]
[TITLE:Revenue Growth]
[X_LABEL:Year]
[Y_LABEL:Revenue ($K)]
```

## API Call Format

To create a visualization, make a POST request to the endpoint with the following JSON structure:

```json
{
  "data": "{\"labels\":[\"A\",\"B\",\"C\"],\"data\":[10,20,30]}",
  "chart_type": "bar",
  "filename": "chart.png"
}
```

### Example cURL Request

```bash
curl -X 'POST' \
  'http://localhost:19801/api/v1/create-visualization' \
  -H 'accept: application/json' \
  -H 'Content-Type: application/json' \
  -d '{
  "data": "{\"labels\":[\"January\",\"February\",\"March\"],\"datasets\":[{\"label\":\"Sales\",\"data\":[120,150,180]}]}",
  "chart_type": "bar",
  "filename": "sales_chart.png"
}'
```

## Tool Call Integration

When integrating this tool into your application, use the following format:

```javascript
{
  "tool_name": "data_visualizer",
  "parameters": {
    "data": "[Your data in JSON or CSV format]",
    "chart_type": "bar",
    "filename": "output_filename.png"
  }
}
```

By following these guidelines, you can effectively create charts and visualizations using the data_visualizer tool.
