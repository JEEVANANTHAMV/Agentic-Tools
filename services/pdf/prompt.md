# Guidelines for Generating Content for PDF Converter API

## Tool Name - pdf_converter

## Basic Structure

The PDF converter allows you to create PDF documents from various content types including plain text, HTML, and structured data. The content is formatted using a combination of Markdown-like syntax and special tags for styling.

## Document Structure

### Page Setup

To configure page settings, use the following syntax at the beginning of your document:

```
[PAGE:orientation,size,margins]
```

- orientation: portrait or landscape
- size: A4, A3, Letter, Legal, etc.
- margins: top,bottom,left,right in inches (e.g., 1,1,1,1)

Example:
```
[PAGE:portrait,A4,1,1,1,1]
```

### Headers and Footers

```
[HEADER:alignment]Your header content[/HEADER]
[FOOTER:alignment]Your footer content[/FOOTER]
```

Alignment options: left, center, right

## Text Formatting

### Headings

```
# Heading Level 1
## Heading Level 2
### Heading Level 3
#### Heading Level 4
##### Heading Level 5
###### Heading Level 6
```

### Bold and Italic

```
[BOLD]This text is bold[/BOLD]
[ITALIC]This text is italic[/ITALIC]
[BOLD][ITALIC]This text is bold and italic[/ITALIC][/BOLD]
```

### Text Color

```
[COLOR:RRGGBB]This text has a custom color[/COLOR]
```

Common colors:
- Red: FF0000
- Green: 00FF00
- Blue: 0000FF
- Black: 000000
- White: FFFFFF

### Font Settings

```
[FONT:font_name,size]Text with custom font[/FONT]
```

Example:
```
[FONT:Arial,12]This text uses Arial 12pt[/FONT]
[FONT:Times New Roman,14]This text uses Times New Roman 14pt[/FONT]
```

## Content Elements

### Paragraphs

Separate paragraphs with blank lines:

```
This is the first paragraph.

This is the second paragraph.
```

### Lists

#### Unordered Lists

```
- Item 1
- Item 2
- Item 3
```

Or:

```
* Item 1
* Item 2
* Item 3
```

#### Ordered Lists

```
1. First item
2. Second item
3. Third item
```

### Tables

```
|Header 1|Header 2|Header 3|
|--------|--------|--------|
|Cell 1  |Cell 2  |Cell 3  |
|Cell 4  |Cell 5  |Cell 6  |
```

### Images

```
[IMAGE:image_path,width,height]
```

Example:
```
[IMAGE:logo.png,200,100]
```

### Links

```
[LINK:URL]Link text[/LINK]
```

Example:
```
[LINK:https://example.com]Visit Example.com[/LINK]
```

## Page Elements

### Page Break

```
[PAGEBREAK]
```

### Page Number

```
[PAGENUM]
```

### Date and Time

```
[DATE:format]
[TIME:format]
```

Format options: YYYY-MM-DD, MM/DD/YYYY, DD-Mon-YYYY, etc.

## Combining Formatting

You can combine multiple formatting options:

```
[PAGE:portrait,A4,1,1,1,1]

[HEADER:center]Document Title[/HEADER]
[FOOTER:right]Page [PAGENUM] | [DATE:MM/DD/YYYY][/FOOTER]

[FONT:Arial,16][BOLD]# Main Title[/BOLD][/FONT]

[FONT:Arial,12]This is a regular paragraph with [BOLD]bold text[/BOLD] and [ITALIC]italic text[/ITALIC].[/FONT]

[FONT:Arial,12]### Features[/FONT]

- [BOLD]Feature 1:[/BOLD] Description of feature one
- [BOLD]Feature 2:[/BOLD] Description of feature two
- [BOLD]Feature 3:[/BOLD] Description of feature three

[PAGEBREAK]

[FONT:Arial,12]### Data Table[/FONT]

|Name|Age|Department|
|--------|--------|--------|
|John Doe|30|Engineering|
|Jane Smith|28|Marketing|
|Bob Johnson|35|Sales|

[FOOTER:center]Confidential Document - [DATE:YYYY-MM-DD][/FOOTER]
```

## Best Practices

1. Always specify page settings at the beginning of the document
2. Use consistent font styles throughout the document
3. Keep margins readable (minimum 0.5 inches)
4. Use page breaks to separate major sections
5. Include headers and footers for professional documents
6. Test with various content to ensure proper rendering
7. Use standard fonts for better compatibility
8. Keep images optimized for PDF (appropriate resolution)

## Example Content

```
[PAGE:portrait,A4,1,1,1,1]

[HEADER:center]Monthly Report[/HEADER]
[FOOTER:right]Page [PAGENUM][/FOOTER]

[FONT:Arial,18][BOLD]# Sales Report - January 2024[/BOLD][/FONT]

[FONT:Arial,12]Prepared by: Sales Department[/FONT]
[FONT:Arial,12]Date: [DATE:MM/DD/YYYY][/FONT]

[FONT:Arial,14][BOLD]## Executive Summary[/BOLD][/FONT]

[FONT:Arial,11]This report provides an overview of sales performance for January 2024. Key highlights include:[/FONT]

[FONT:Arial,11]- [BOLD]Total Revenue:[/BOLD] $1,250,000[/FONT]
[FONT:Arial,11]- [BOLD]Growth Rate:[/BOLD] [COLOR:00FF00]15%[/COLOR] increase from previous month[/FONT]
[FONT:Arial,11]- [BOLD]Top Product:[/BOLD] Product X with $350,000 in sales[/FONT]

[PAGEBREAK]

[FONT:Arial,14][BOLD]## Sales by Region[/BOLD][/FONT]

|Region|Revenue|Growth|Target|
|--------|--------|--------|--------|
|North|$450,000|12%|$400,000|
|South|$380,000|18%|$350,000|
|East|$270,000|8%|$300,000|
|West|$150,000|22%|$200,000|

[FONT:Arial,11][ITALIC]Note: All figures are in USD.[/ITALIC][/FONT]

[FOOTER:center]Confidential - For Internal Use Only[/FOOTER]
```

## API Call Format

To generate a PDF file, make a POST request to the endpoint with the following JSON structure:

```json
{
  "content": "Your formatted content string here",
  "filename": "desired_filename.pdf"
}
```

### Example cURL Request

```bash
curl -X 'POST' \
  'http://localhost:19801/api/v1/convert-to-pdf' \
  -H 'accept: application/json' \
  -H 'Content-Type: application/json' \
  -d '{
  "content": "[PAGE:portrait,A4,1,1,1,1]\n\n[FONT:Arial,16][BOLD]# Document Title[/BOLD][/FONT]\n\n[FONT:Arial,12]This is a sample PDF document.[/FONT]",
  "filename": "document.pdf"
}'
```

## Tool Call Integration

When integrating this tool into your application, use the following format:

```javascript
{
  "tool_name": "pdf_converter",
  "parameters": {
    "content": "[Your formatted PDF content string]",
    "filename": "output_filename.pdf"
  }
}
```

By following these guidelines, you can create well-formatted PDF documents using the pdf_converter tool.
