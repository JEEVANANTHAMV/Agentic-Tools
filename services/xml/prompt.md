# Guidelines for XML Parser API

## Tool Name - xml_parser

## Basic Structure

The XML parser allows you to parse, validate, transform, and manipulate XML documents. It supports XSD validation, XSLT transformations, and various XML operations.

## Input Format

### XML Content

Provide your XML data as a properly formatted string:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<company>
  <employee id="001">
    <name>John Doe</name>
    <department>Engineering</department>
    <salary>75000</salary>
  </employee>
  <employee id="002">
    <name>Jane Smith</name>
    <department>Marketing</department>
    <salary>65000</salary>
  </employee>
</company>
```

## XML Formatting

### Pretty Print

Format XML with indentation for readability:

```
[FORMAT:pretty,indent:2]
```

### Minify

Remove all whitespace for compact XML:

```
[FORMAT:minify]
```

## XSD Validation

### Schema Definition

Include an XSD schema to validate the XML:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<xs:schema xmlns:xs="http://www.w3.org/2001/XMLSchema">
  <xs:element name="company">
    <xs:complexType>
      <xs:sequence>
        <xs:element name="employee" maxOccurs="unbounded">
          <xs:complexType>
            <xs:sequence>
              <xs:element name="name" type="xs:string"/>
              <xs:element name="department" type="xs:string"/>
              <xs:element name="salary" type="xs:integer"/>
            </xs:sequence>
            <xs:attribute name="id" type="xs:string" use="required"/>
          </xs:complexType>
        </xs:element>
      </xs:sequence>
    </xs:complexType>
  </xs:element>
</xs:schema>
```

### Validation in Content

```
[VALIDATE:XSD:schema_file.xsd]
```

## XSLT Transformations

### Basic Transformation

Apply an XSLT stylesheet to transform the XML:

```
[TRANSFORM:XSLT:stylesheet.xsl]
```

### Inline XSLT

Include XSLT directly in the transform parameter:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
  <xsl:template match="/">
    <html>
      <body>
        <h2>Employee List</h2>
        <table border="1">
          <tr>
            <th>Name</th>
            <th>Department</th>
            <th>Salary</th>
          </tr>
          <xsl:for-each select="company/employee">
            <tr>
              <td><xsl:value-of select="name"/></td>
              <td><xsl:value-of select="department"/></td>
              <td><xsl:value-of select="salary"/></td>
            </tr>
          </xsl:for-each>
        </table>
      </body>
    </html>
  </xsl:template>
</xsl:stylesheet>
```

## XPath Queries

### Extract Values

Use XPath to extract specific values:

```
[EXTRACT://employee/name]
[EXTRACT://employee[@id='001']/salary]
```

### Filter Elements

Filter elements based on conditions:

```
[FILTER://employee:department:Engineering]
[FILTER://employee:salary>70000]
```

## Content Format with Operations

The `content` parameter can include both the XML data and inline operation instructions:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<root>
  <item>Value</item>
</root>

[FORMAT:pretty,indent:2]
```

## Combining Operations

You can combine multiple operations:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<company>
  <employee id="001">
    <name>John Doe</name>
    <department>Engineering</department>
  </employee>
</company>

[FORMAT:pretty,indent:2]
[VALIDATE:XSD:schema.xsd]
```

## Best Practices

1. Always validate XML syntax before processing
2. Use XSD schemas for data validation
3. Keep XSLT transformations modular
4. Test XPath expressions on sample data
5. Handle namespaces correctly
6. Use appropriate character encoding (UTF-8 recommended)
7. Document complex transformations
8. Consider security when parsing external XML

## Example Content

### Example 1: Pretty Print XML

```xml
<?xml version="1.0"?><company><employee><name>John Doe</name></employee></company>

[FORMAT:pretty,indent:2]
```

Output:
```xml
<?xml version="1.0"?>
<company>
  <employee>
    <name>John Doe</name>
  </employee>
</company>
```

### Example 2: Transform to HTML

XML:
```xml
<?xml version="1.0" encoding="UTF-8"?>
<products>
  <product id="1">
    <name>Widget A</name>
    <price>19.99</price>
  </product>
  <product id="2">
    <name>Widget B</name>
    <price>29.99</price>
  </product>
</products>
```

XSLT:
```xml
<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
  <xsl:template match="/">
    <html>
      <body>
        <h2>Products</h2>
        <ul>
          <xsl:for-each select="products/product">
            <li>
              <xsl:value-of select="name"/> - $<xsl:value-of select="price"/>
            </li>
          </xsl:for-each>
        </ul>
      </body>
    </html>
  </xsl:template>
</xsl:stylesheet>
```

### Example 3: Extract and Filter

```xml
<?xml version="1.0" encoding="UTF-8"?>
<employees>
  <employee department="Engineering">
    <name>John Doe</name>
    <salary>75000</salary>
  </employee>
  <employee department="Marketing">
    <name>Jane Smith</name>
    <salary>65000</salary>
  </employee>
</employees>

[EXTRACT://employee[@department='Engineering']/name]
```

## API Call Format

To parse XML data, make a POST request to the endpoint with the following JSON structure:

```json
{
  "content": "Your XML content string",
  "transform": "XSLT transformation string or file reference",
  "filename": "output.xml"
}
```

### Example cURL Request

```bash
curl -X 'POST' \
  'http://101.53.140.44:8002/api/v1/parse-xml' \
  -H 'accept: application/json' \
  -H 'Content-Type: application/json' \
  -d '{
  "content": "<?xml version=\"1.0\"?><company><employee><name>John Doe</name></employee></company>",
  "transform": "<?xml version=\"1.0\"?><xsl:stylesheet version=\"1.0\" xmlns:xsl=\"http://www.w3.org/1999/XSL/Transform\"><xsl:template match=\"/\"><html><body><xsl:value-of select=\"/company/employee/name\"/></body></html></xsl:template></xsl:stylesheet>",
  "filename": "transformed_output.xml"
}'
```

## Tool Call Integration

When integrating this tool into your application, use the following format:

```javascript
{
  "tool_name": "xml_parser",
  "parameters": {
    "content": "[Your XML content string]",
    "transform": "[Optional XSLT transformation]",
    "filename": "output_filename.xml"
  }
}
```

By following these guidelines, you can effectively parse, validate, and transform XML documents using the xml_parser tool.
