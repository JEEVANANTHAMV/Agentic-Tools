Guidelines for Generating Content for Word Document Generator API

Tool Name - document_writer

Basic Structure

The content should be formatted using a combination of Markdown-like syntax and special tags for font formatting. Each paragraph can have its own font settings specified at the beginning.

Font Formatting

Paragraph-Level Font Settings

[FONT:FontName,FontSize]Your paragraph content here

- FontName: Any valid font name (e.g., Arial, Calibri, Times New Roman)

- FontSize: Integer value in points (e.g., 11, 12, 14)

- This setting applies to the entire paragraph unless overridden by inline formatting

Inline Font Changes

[FONT:FontName,FontSize]Text with custom font[/FONT]

- Use this to change font within a paragraph

- The closing tag [/FONT] is required

Inline Size Changes

[SIZE:FontSize]Text with custom size[/SIZE]

- Use this to change only the font size within a paragraph

- The closing tag [/SIZE] is required

Text Formatting

Bold Text (will 2 stars fist and last)

**This text will be bold* *

Italic Text

*This text will be italic*

Bold and Italic Text

***This text will be both bold and italic***

Headings

### Heading 1

# This is a level 1 heading

Heading 2

## This is a level 2 heading

Heading 3

### This is a level 3 heading

Note: Headings can also have paragraph-level font settings:

[FONT:Arial,16]# Document Title

Lists

Bullet Points

- First bullet point

- Second bullet point

- Third bullet point

or

* First bullet point

* Second bullet point

* Third bullet point

### Numbered Lists

1. First item

2. Second item

3. Third item

Note: List items can have paragraph-level font settings:

[FONT:Arial,12]- Bullet point with custom font

Tables

Use Markdown-style table syntax:

|Header1|Header2|Header3|

|-------|-------|-------|

|Row1Col1|Row1Col2|Row1Col3|

|Row2Col1|Row2Col2|Row2Col3|

Tables use the font settings from the paragraph that precedes them.

Empty Paragraphs

[FONT:Calibri,12]

Just an empty line with font settings will create an empty paragraph with those settings.

Combining Formatting

You can combine multiple formatting options:

[FONT:Arial,12]This paragraph has bold text and italic text.

[FONT:Calibri,11]You can use [FONT:Courier New,10]different fonts[/FONT] within the same paragraph.

Best Practices

1. Always specify font settings at the beginning of each paragraph for consistent formatting

2. Use standard fonts (Arial, Calibri, Times New Roman, Verdana) for better compatibility

3. Keep font sizes between 8pt and 18pt for readability

4. When mixing formatting, be careful with nested tags - ensure all closing tags match

5. For complex documents, plan the structure before writing the content

6. Test with a variety of content to ensure all formatting works as expected

Example Content

[FONT:Arial,16]# Document Title

[FONT:Calibri,12]This is a regular paragraph with default font settings.

[FONT:Times New Roman,11]This paragraph uses Times New Roman font.

[FONT:Arial,14]## Section Header

[FONT:Calibri,12]Here's some bold text and italic text in the same paragraph.

[FONT:Verdana,12]You can also have ***bold and italic text*** together.

[FONT:Calibri,12]### Subsection

[FONT:Arial,11]This is a list:

- First item with [FONT:Courier New,10]different font[/FONT]

- Second item

- Third item with bold text

[FONT:Calibri,12]And a numbered list:

1. First numbered item

2. Second item with [FONT:Georgia,14]larger font[/FONT]

3. Third item

[FONT:Arial,12]### Table Example

[FONT:Calibri,11]|Name|Age|Occupation|

|-------|-------|-------|

|John Doe|30|Software Engineer|

|Jane Smith|28|Designer|

|Bob Johnson|45|Manager|

[FONT:Times New Roman,12]This paragraph has some inline formatting: [FONT:Arial,10]Arial 10pt[/FONT], [FONT:Verdana,14]Verdana 14pt[/FONT], and back to [FONT:Times New Roman,12]Times New Roman 12pt[/FONT].

[FONT:Calibri,12]You can also use size tags: [SIZE:10]Small text[/SIZE], [SIZE:14]Medium text[/SIZE], and [SIZE:18]Large text[/SIZE].