# Guidelines for Generating Content for PowerPoint Presentation API

## Tool Name - presentation_creator

## Basic Structure

The presentation creator allows you to create PowerPoint presentations with multiple slides, layouts, and content. The content is formatted using a combination of Markdown-like syntax and special tags for slide formatting.

## Slide Creation

### New Slide

To create a new slide, use the slide delimiter:

```
---
```

Each section between delimiters represents a new slide.

### Slide Layouts

Specify the layout at the beginning of a slide:

```
[LAYOUT:title]
[LAYOUT:content]
[LAYOUT:two_column]
[LAYOUT:comparison]
[LAYOUT:section_header]
[LAYOUT:blank]
```

## Content Elements

### Title

```
# Slide Title
```

### Subtitle

```
## Subtitle
```

### Text Content

Regular text content:

```
This is regular text content on the slide.
```

### Bullet Points

```
- First bullet point
- Second bullet point
- Third bullet point
```

Or:

```
* First item
* Second item
* Third item
```

### Numbered Lists

```
1. First item
2. Second item
3. Third item
```

### Text Formatting

#### Bold

```
[BOLD]This text is bold[/BOLD]
```

#### Italic

```
[ITALIC]This text is italic[/ITALIC]
```

#### Color

```
[COLOR:RRGGBB]This text has a custom color[/COLOR]
```

Examples:
- Red: `[COLOR:FF0000]Red text[/COLOR]`
- Blue: `[COLOR:0000FF]Blue text[/COLOR]`
- Green: `[COLOR:00FF00]Green text[/COLOR]`

#### Font Size

```
[SIZE:24]Large text[/SIZE]
[SIZE:18]Medium text[/SIZE]
[SIZE:12]Small text[/SIZE]
```

## Tables

```
|Header 1|Header 2|Header 3|
|--------|--------|--------|
|Cell 1  |Cell 2  |Cell 3  |
|Cell 4  |Cell 5  |Cell 6  |
```

## Images

```
[IMAGE:image_path]
```

Or with sizing:

```
[IMAGE:image_path,width,height]
```

Example:
```
[IMAGE:logo.png,200,100]
```

## Shapes

### Rectangle

```
[SHAPE:rectangle,width,height,color]
```

Example:
```
[SHAPE:rectangle,300,150,FF6384]
```

### Circle

```
[SHAPE:circle,diameter,color]
```

Example:
```
[SHAPE:circle,100,36A2EB]
```

## Two Column Layout

```
[LAYOUT:two_column]

# Title

[LEFT]
Left column content
- Point 1
- Point 2
[/LEFT]

[RIGHT]
Right column content
- Point 1
- Point 2
[/RIGHT]
```

## Section Headers

```
[LAYOUT:section_header]

# Section Title

## Section Subtitle
```

## Slide Background

```
[BACKGROUND:color]
[BACKGROUND:image_path]
[BACKGROUND:gradient]
```

Examples:
```
[BACKGROUND:FFFFFF]  # White background
[BACKGROUND:000000]  # Black background
[BACKGROUND:logo_bg.png]  # Image background
```

## Animations

```
[ANIMATION:fade]
[ANIMATION:slide_in]
[ANIMATION:zoom]
[ANIMATION:bounce]
```

Apply to elements:
```
[ANIMATION:fade]- This bullet will fade in[/ANIMATION]
```

## Transitions

```
[TRANSITION:fade]
[TRANSITION:slide]
[TRANSITION:dissolve]
[TRANSITION:push]
```

## Combining Elements

You can combine multiple elements on a slide:

```
[LAYOUT:content]

# Quarterly Results

## Q4 2024 Performance

[BOLD]Key Highlights:[/BOLD]

- [COLOR:00FF00]Revenue increased by 25%[/COLOR]
- [COLOR:0000FF]New customers: 1,500[/COLOR]
- [COLOR:FF0000]Expenses reduced by 10%[/COLOR]

|Metric|Q3|Q4|Change|
|------|------|------|------|
|Revenue|$800K|$1M|+25%|
|Customers|5000|6500|+30%|
|Profit|$200K|$300K|+50%|

[IMAGE:chart.png,400,250]
```

## Best Practices

1. Keep slides concise and focused
2. Use consistent formatting throughout
3. Limit text per slide (6x6 rule: 6 bullets, 6 words each)
4. Use high-quality images
5. Choose readable font sizes (minimum 18pt for body text)
6. Use contrasting colors for readability
7. Include clear titles on all slides
8. Use animations sparingly

## Example Presentation

```
[LAYOUT:title]
[BACKGROUND:gradient]

# Annual Report 2024

## Presented by Sales Team

[IMAGE:logo.png,150,150]

[TRANSITION:fade]

---

[LAYOUT:section_header]

# Executive Summary

## Key Highlights from 2024

[TRANSITION:slide]

---

[LAYOUT:content]

# Performance Overview

[BOLD]Year-over-Year Growth:[/BOLD]

- [COLOR:00FF00]Revenue: +25%[/COLOR]
- [COLOR:00FF00]Customers: +30%[/COLOR]
- [COLOR:00FF00]Market Share: +15%[/COLOR]

[SIZE:14]Our strategic initiatives have delivered exceptional results across all key metrics.[/SIZE]

[TRANSITION:dissolve]

---

[LAYOUT:two_column]

# Regional Performance

[LEFT]
## North America

- Revenue: $5M
- Growth: +20%
- Customers: 3000
[/LEFT]

[RIGHT]
## International

- Revenue: $3M
- Growth: +35%
- Customers: 2000
[/RIGHT]

[TRANSITION:slide]

---

[LAYOUT:content]

# Financial Summary

|Quarter|Revenue|Expenses|Profit|
|-------|--------|--------|------|
|Q1|$2.0M|$1.5M|$0.5M|
|Q2|$2.5M|$1.6M|$0.9M|
|Q3|$2.8M|$1.7M|$1.1M|
|Q4|$3.2M|$1.8M|$1.4M|

[IMAGE:revenue_chart.png,500,300]

[TRANSITION:fade]

---

[LAYOUT:title]

# Thank You

## Questions & Answers

[IMAGE:contact_info.png,300,100]
```

## API Call Format

To generate a presentation, make a POST request to the endpoint with the following JSON structure:

```json
{
  "content": "Your formatted presentation content string",
  "filename": "presentation.pptx"
}
```

### Example cURL Request

```bash
curl -X 'POST' \
  'http://localhost:19801/api/v1/generate-presentation' \
  -H 'accept: application/json' \
  -H 'Content-Type: application/json' \
  -d '{
  "content": "[LAYOUT:title]\n\n# Welcome\n\n## Presentation Title\n\n---\n\n[LAYOUT:content]\n\n# Agenda\n\n- Topic 1\n- Topic 2\n- Topic 3",
  "filename": "presentation.pptx"
}'
```

## Tool Call Integration

When integrating this tool into your application, use the following format:

```javascript
{
  "tool_name": "presentation_creator",
  "parameters": {
    "content": "[Your formatted presentation content string]",
    "filename": "output_filename.pptx"
  }
}
```

By following these guidelines, you can create well-formatted PowerPoint presentations using the presentation_creator tool.
