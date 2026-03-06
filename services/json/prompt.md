# Guidelines for JSON Formatter API

## Tool Name - json_formatter

## Basic Structure

The JSON formatter allows you to format, validate, transform, and manipulate JSON data. It supports schema validation, pretty printing, minification, and various transformations.

## Input Format

### JSON Content

Provide your JSON data as a properly formatted string:

```json
{
  "name": "John Doe",
  "age": 30,
  "email": "john@example.com",
  "address": {
    "street": "123 Main St",
    "city": "New York",
    "zip": "10001"
  },
  "hobbies": ["reading", "gaming", "traveling"]
}
```

## Formatting Options

### Pretty Print

Format JSON with indentation for readability:

```
[FORMAT:pretty,indent:2]
```

Indent options: 2, 4, or any positive integer

### Minify

Remove all whitespace for compact JSON:

```
[FORMAT:minify]
```

### Sort Keys

Sort object keys alphabetically:

```
[FORMAT:sort_keys]
```

## Schema Validation

### JSON Schema

Include a JSON schema to validate the data:

```json
{
  "type": "object",
  "properties": {
    "name": {"type": "string"},
    "age": {"type": "integer", "minimum": 0},
    "email": {"type": "string", "format": "email"}
  },
  "required": ["name", "email"]
}
```

### Validation in Content

```
[VALIDATE:schema]
```

## Transformations

### Extract Path

Extract specific values using JSONPath:

```
[EXTRACT:$.address.city]
[EXTRACT:$.hobbies[*]]
```

### Filter Objects

Filter array elements based on conditions:

```
[FILTER_ARRAY:$.items:price>100]
```

### Map Transform

Transform array elements:

```
[MAP:$.items:uppercase:name]
```

### Rename Keys

Rename object keys:

```
[RENAME:old_key:new_key]
```

### Add Field

Add new fields to objects:

```
[ADD_FIELD:timestamp:CURRENT_DATE]
[ADD_FIELD:status:active]
```

### Remove Field

Remove fields from objects:

```
[REMOVE_FIELD:temp_field]
```

## Content Format with Operations

The `content` parameter can include both the JSON data and inline operation instructions:

```json
{
  "name": "John Doe",
  "age": 30
}

[FORMAT:pretty,indent:4]
[VALIDATE:schema]
```

## Schema Format

The `schema` parameter accepts a JSON Schema object:

```json
{
  "type": "object",
  "properties": {
    "id": {"type": "integer"},
    "name": {"type": "string", "minLength": 1},
    "email": {"type": "string", "format": "email"},
    "age": {"type": "integer", "minimum": 0, "maximum": 150},
    "active": {"type": "boolean"},
    "metadata": {
      "type": "object",
      "properties": {
        "created_at": {"type": "string", "format": "date-time"},
        "tags": {
          "type": "array",
          "items": {"type": "string"}
        }
      }
    }
  },
  "required": ["id", "name", "email"]
}
```

## Combining Operations

You can combine multiple operations:

```json
{
  "name": "john doe",
  "age": 30,
  "email": "JOHN@EXAMPLE.COM"
}

[FORMAT:pretty,indent:2]
[TRANSFORM:name:capitalize]
[TRANSFORM:email:lowercase]
[ADD_FIELD:processed:CURRENT_DATE]
```

## Best Practices

1. Always validate JSON syntax before processing
2. Use JSON Schema for data validation
3. Keep schemas modular and reusable
4. Test transformations on sample data
5. Handle nested structures carefully
6. Use appropriate data types in schemas
7. Document complex transformations
8. Consider performance for large JSON files

## Example Content

### Example 1: Pretty Print with Validation

```json
{"name":"John Doe","age":30,"email":"john@example.com"}

[FORMAT:pretty,indent:2]
```

With schema:
```json
{
  "type": "object",
  "properties": {
    "name": {"type": "string"},
    "age": {"type": "integer"},
    "email": {"type": "string", "format": "email"}
  },
  "required": ["name", "email"]
}
```

### Example 2: Transform and Format

```json
{
  "users": [
    {"name": "john doe", "email": "JOHN@EXAMPLE.COM"},
    {"name": "jane smith", "email": "JANE@EXAMPLE.COM"}
  ]
}

[TRANSFORM:users[*].name:capitalize]
[TRANSFORM:users[*].email:lowercase]
[FORMAT:pretty,indent:2]
```

### Example 3: Extract and Filter

```json
{
  "products": [
    {"name": "Widget A", "price": 19.99, "category": "electronics"},
    {"name": "Widget B", "price": 29.99, "category": "home"},
    {"name": "Widget C", "price": 9.99, "category": "electronics"}
  ]
}

[FILTER_ARRAY:products:category:electronics]
[EXTRACT:products[*].name]
[FORMAT:pretty,indent:2]
```

## API Call Format

To format JSON data, make a POST request to the endpoint with the following JSON structure:

```json
{
  "content": "Your JSON content string",
  "schema": {
    "type": "object",
    "properties": {
      "name": {"type": "string"},
      "age": {"type": "integer"}
    }
  },
  "filename": "formatted_data.json"
}
```

### Example cURL Request

```bash
curl -X 'POST' \
  'http://localhost:19801/api/v1/format-json' \
  -H 'accept: application/json' \
  -H 'Content-Type: application/json' \
  -d '{
  "content": "{\"name\":\"John Doe\",\"age\":30,\"email\":\"john@example.com\"}",
  "schema": {
    "type": "object",
    "properties": {
      "name": {"type": "string"},
      "age": {"type": "integer"},
      "email": {"type": "string", "format": "email"}
    },
    "required": ["name", "email"]
  },
  "filename": "user_data.json"
}'
```

## Tool Call Integration

When integrating this tool into your application, use the following format:

```javascript
{
  "tool_name": "json_formatter",
  "parameters": {
    "content": "[Your JSON content string]",
    "schema": {
      "type": "object",
      "properties": {
        "name": {"type": "string"}
      }
    },
    "filename": "output_filename.json"
  }
}
```

By following these guidelines, you can effectively format, validate, and transform JSON data using the json_formatter tool.
