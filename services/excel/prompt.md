

# Guidelines for Generating Content for Excel File Generator API

## Tool Name - generate_presentation

## Basic Structure

The content should be formatted using a combination of Markdown-like syntax for tables and special tags for cell formatting. Each sheet in the Excel file starts with a sheet definition.

## Sheet Creation

To create a new sheet in the Excel file, start a line with `#` followed by the sheet name:

```
# Sales Report
```

This will create a new sheet named "Sales Report". All content following this line (until another sheet definition) will be added to this sheet.

## Table Creation

To create a table, use markdown-style table syntax:

```
|Header1|Header2|Header3|
|-------|-------|-------|
|Row1Col1|Row1Col2|Row1Col3|
|Row2Col1|Row2Col2|Row2Col3|
```

The first line defines the headers, the second line (with dashes) separates headers from data, and subsequent lines define the data rows.

## Cell Formatting

### Bold Text

Use `[BOLD]` and `[/BOLD]` tags to make text bold:

```
[BOLD]This text will be bold[/BOLD]
```

### Italic Text

Use `[ITALIC]` and `[/ITALIC]` tags to make text italic:

```
[ITALIC]This text will be italic[/ITALIC]
```

### Colored Text

Use `[COLOR:RRGGBB]` and `[/COLOR]` tags to change text color, where RRGGBB is the hex color code:

```
[COLOR:FF0000]This text will be red[/COLOR]
[COLOR:00FF00]This text will be green[/COLOR]
[COLOR:0000FF]This text will be blue[/COLOR]
```

### Text Alignment

Use `[ALIGN:left|center|right]` and `[/ALIGN]` tags to align text:

```
[ALIGN:left]This text will be left-aligned[/ALIGN]
[ALIGN:center]This text will be center-aligned[/ALIGN]
[ALIGN:right]This text will be right-aligned[/ALIGN]
```

### Cell Borders

Use `[BORDER]` and `[/BORDER]` tags to add borders to a cell:

```
[BORDER]This cell will have borders[/BORDER]
```

## Regular Text Content

For regular text content that doesn't require special formatting, simply add the text:

```
This is regular text content.
```

## Combining Formatting Options

You can combine multiple formatting options by nesting tags:

```
[BOLD][COLOR:FF0000]This text will be bold and red[/COLOR][/BOLD]
[ITALIC][ALIGN:center]This text will be italic and center-aligned[/ITALIC]
[BORDER][BOLD][COLOR:0000FF]This cell will have borders, bold text, and blue color[/COLOR][/BORDER]
```

## Multiple Sheets

To create multiple sheets, simply add multiple sheet definitions:

```
# Sales Report

|Product|Quantity|Price|Total|
|-------|-------|-----|-----|
|Product A|10|[COLOR:FF0000]$19.99[/COLOR]|$199.90|
|Product B|5|$29.99|$149.95|

# Customer Information

|Name|Email|Phone|
|-------|-------|-------|
|[BOLD]John Doe[/BOLD]|john@example.com|555-1234|
|Jane Smith|[ITALIC]jane@example.com[/ITALIC]|555-5678|
```

## Best Practices

1. Always start a table with a header row followed by a separator row with dashes.
2. Use consistent formatting for similar data across tables.
3. When using color codes, ensure they provide sufficient contrast for readability.
4. Use borders to highlight important information or to create visual separation.
5. Align numeric data to the right for better readability.
6. Use bold formatting for headers and important values.
7. Keep sheet names concise but descriptive.
8. Ensure all formatting tags are properly closed to avoid errors.
9. Test with a variety of content to ensure all formatting works as expected.
10. For complex spreadsheets, plan the structure before writing the content.

## Example Content

```
# Sales Report

|Product|Quantity|Price|Total|
|-------|-------|-----|-----|
|Product A|10|[COLOR:FF0000]$19.99[/COLOR]|$199.90|
|Product B|5|$29.99|$149.95|
|Product C|15|[BOLD]$9.99[/BOLD]|$149.85|

# Summary

[ALIGN:center]Total Products: 3[/ALIGN]
[ALIGN:right]Total Revenue: $499.70[/ALIGN]
[BORDER][BOLD][COLOR:0000FF]This is important information[/COLOR][/BORDER]

# Customer Information

|Name|Email|Phone|
|-------|-------|-------|
|[BOLD]John Doe[/BOLD]|john@example.com|555-1234|
|Jane Smith|[ITALIC]jane@example.com[/ITALIC]|555-5678|
```

## API Call Format

To generate an Excel file, make a POST request to the endpoint with the following JSON structure:

```json
{
  "content": "Your formatted content string here",
  "filename": "desired_filename.xlsx"
}
```

### Example cURL Request

```bash
curl -X 'POST' \
  'http://0.0.0.0:19801/api/v1/generate-excel' \
  -H 'accept: application/json' \
  -H 'Content-Type: application/json' \
  -d '{
  "content": "# Sales Report\n\n|Product|Quantity|Price|Total|\n|-------|-------|-----|-----|\n|Product A|10|[COLOR:FF0000]$19.99[/COLOR]|$199.90|\n|Product B|5|$29.99|$149.95|\n|Product C|15|[BOLD]$9.99[/BOLD]|$149.85|\n\n# Summary\n\n[ALIGN:center]Total Products: 3[/ALIGN]\n[ALIGN:right]Total Revenue: $499.70[/ALIGN]\n[BORDER]This is important information[/BORDER]",
  "filename": "sales_report.xlsx"
}'
```

## Tool Call Integration

When integrating this tool into your application, use the following format:

```javascript
{
  "tool_name": "generate_presentation",
  "parameters": {
    "content": "[Your formatted Excel content string]",
    "filename": "output_filename.xlsx"
  }
}
```

By following these guidelines, you can create well-formatted Excel files using the generate_presentation tool.