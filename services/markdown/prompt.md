# Guidelines for Markdown Converter API

## Tool Name - markdown_converter

## Basic Structure

The Markdown converter allows you to convert Markdown content to various output formats including HTML, PDF, and DOCX. It supports custom styling, embedded resources, and various Markdown extensions.

## Input Format

### Markdown Content

Provide your Markdown content as a string:

```markdown
# Document Title

This is a paragraph with **bold** and *italic* text.

## Features

- Feature 1
- Feature 2
- Feature 3

### Code Example

```python
def hello():
    print("Hello, World!")
```

[Link Text](https://example.com)
```

## Output Formats

### HTML

Convert Markdown to HTML:

```
[OUTPUT:html]
```

Options:
- `html` - Standard HTML output
- `html5` - HTML5 compliant output
- `xhtml` - XHTML output

### PDF

Convert Markdown to PDF:

```
[OUTPUT:pdf]
```

Options:
- `pdf` - Standard PDF output
- `pdf:a4` - A4 sized PDF
- `pdf:letter` - Letter sized PDF

### DOCX

Convert Markdown to Word document:

```
[OUTPUT:docx]
```

Options:
- `docx` - Standard Word document
- `docx:template` - Use custom template

## Styling Options

### Custom CSS (for HTML output)

```
[STYLE:css_file.css]
```

Or inline styles:

```
[STYLE:inline]
h1 { color: #333; font-size: 2em; }
p { line-height: 1.6; }
[/STYLE]
```

### Theme Selection

```
[THEME:light]
[THEME:dark]
[THEME:monokai]
```

Available themes: light, dark, monokai, solarized, github

## Markdown Extensions

### Tables

```markdown
| Header 1 | Header 2 | Header 3 |
|----------|----------|----------|
| Cell 1   | Cell 2   | Cell 3   |
| Cell 4   | Cell 5   | Cell 6   |
```

### Task Lists

```markdown
- [x] Completed task
- [ ] Pending task
- [ ] Another pending task
```

### Checkboxes

```markdown
- [x] Item 1
- [ ] Item 2
```

### Footnotes

```markdown
Here is a footnote reference.[^1]

[^1]: Here is the footnote.
```

### Code Blocks with Syntax Highlighting

```markdown
```python
def greet(name):
    return f"Hello, {name}!"
```
```

### Blockquotes

```markdown
> This is a blockquote.
> It can span multiple lines.
```

### Horizontal Rules

```markdown
---

***

___
```

### Images

```markdown
![Alt text](image.jpg "Title")
```

### Links

```markdown
[Link text](https://example.com)

[Reference][ref]

[ref]: https://example.com
```

## Content Format with Options

The `content` parameter can include both the Markdown and output options:

```markdown
# My Document

This is the content.

[OUTPUT:html]
[THEME:github]
```

## Combining Options

You can combine multiple options:

```markdown
# Report

## Summary

This is a summary.

[OUTPUT:pdf]
[THEME:light]
[STYLE:custom.css]
```

## Best Practices

1. Use proper Markdown syntax
2. Include alt text for images
3. Use meaningful link text
4. Keep code blocks properly formatted
5. Test output in target format
6. Use consistent heading levels
7. Validate links before conversion
8. Consider accessibility in styling

## Example Content

### Example 1: Convert to HTML

```markdown
# Welcome

This is a **welcome** message with *formatting*.

## Features

- Fast conversion
- Multiple formats
- Custom styling

### Code Example

```javascript
console.log("Hello, World!");
```

[Learn More](https://example.com)
```

With options:
```
[OUTPUT:html]
[THEME:github]
```

### Example 2: Convert to PDF

```markdown
# Monthly Report

## Executive Summary

This report covers the monthly performance.

### Key Metrics

| Metric | Value | Change |
|--------|-------|--------|
| Revenue | $100K | +10% |
| Users | 5000 | +5% |

> Note: All figures are preliminary.
```

With options:
```
[OUTPUT:pdf]
[THEME:light]
```

### Example 3: Convert to DOCX

```markdown
# Project Documentation

## Overview

This document describes the project.

## Installation

1. Clone the repository
2. Install dependencies
3. Run the application

## Usage

```bash
npm start
```
```

With options:
```
[OUTPUT:docx]
```

## API Call Format

To convert Markdown, make a POST request to the endpoint with the following JSON structure:

```json
{
  "content": "Your Markdown content string",
  "output_format": "html",
  "filename": "output.html"
}
```

### Example cURL Request

```bash
curl -X 'POST' \
  'http://101.53.140.44:8002/api/v1/convert-markdown' \
  -H 'accept: application/json' \
  -H 'Content-Type: application/json' \
  -d '{
  "content": "# Welcome\n\nThis is a **welcome** message.\n\n## Features\n\n- Fast conversion\n- Multiple formats",
  "output_format": "html",
  "filename": "welcome.html"
}'
```

## Tool Call Integration

When integrating this tool into your application, use the following format:

```javascript
{
  "tool_name": "markdown_converter",
  "parameters": {
    "content": "[Your Markdown content string]",
    "output_format": "html",
    "filename": "output_filename.html"
  }
}
```

By following these guidelines, you can effectively convert Markdown to various formats using the markdown_converter tool.
