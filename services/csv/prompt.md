# Guidelines for Processing CSV Data API

## Tool Name - csv_processor

## Basic Structure

The CSV processor allows you to process, transform, filter, and format CSV data. You can perform various operations on CSV content including filtering rows, sorting columns, transforming data, and more.

## Input Format

### CSV Content

Provide your CSV data as a string with proper formatting:

```
name,age,department,salary
John Doe,30,Engineering,75000
Jane Smith,28,Marketing,65000
Bob Johnson,35,Sales,70000
```

## Operations

The `operations` parameter accepts an array of operation objects. Each operation specifies the type of transformation to apply.

### Filter Operation

Filter rows based on conditions:

```json
{
  "operation": "filter",
  "column": "age",
  "condition": "greater_than",
  "value": 30
}
```

Available conditions:
- `equals` - Exact match
- `not_equals` - Not equal to
- `greater_than` - Greater than value
- `less_than` - Less than value
- `greater_than_or_equal` - Greater than or equal
- `less_than_or_equal` - Less than or equal
- `contains` - Contains substring
- `starts_with` - Starts with substring
- `ends_with` - Ends with substring

### Sort Operation

Sort rows by column values:

```json
{
  "operation": "sort",
  "column": "salary",
  "order": "descending"
}
```

Order options: `ascending`, `descending`

### Select Operation

Select specific columns to include in output:

```json
{
  "operation": "select",
  "columns": ["name", "department"]
}
```

### Transform Operation

Transform column values:

```json
{
  "operation": "transform",
  "column": "name",
  "function": "uppercase"
}
```

Available functions:
- `uppercase` - Convert to uppercase
- `lowercase` - Convert to lowercase
- `trim` - Remove whitespace
- `round` - Round numeric values
- `format_currency` - Format as currency
- `format_percentage` - Format as percentage

### Add Column Operation

Add a new calculated column:

```json
{
  "operation": "add_column",
  "name": "bonus",
  "formula": "salary * 0.1"
}
```

### Remove Column Operation

Remove unwanted columns:

```json
{
  "operation": "remove_column",
  "columns": ["age"]
}
```

### Aggregate Operation

Perform aggregations on data:

```json
{
  "operation": "aggregate",
  "group_by": "department",
  "aggregations": [
    {"column": "salary", "function": "sum"},
    {"column": "salary", "function": "average"},
    {"column": "name", "function": "count"}
  ]
}
```

Available aggregation functions: `sum`, `average`, `count`, `min`, `max`

## Content Format with Operations

The `content` parameter can include both the CSV data and inline operation instructions:

```
# CSV Data
name,age,department,salary
John Doe,30,Engineering,75000
Jane Smith,28,Marketing,65000
Bob Johnson,35,Sales,70000

# Operations
[FILTER:age>30]
[SORT:salary:descending]
[SELECT:name,department,salary]
```

## Inline Operations Syntax

### Filter

```
[FILTER:column:condition:value]
```

Examples:
```
[FILTER:age:greater_than:30]
[FILTER:department:equals:Engineering]
[FILTER:name:contains:John]
```

### Sort

```
[SORT:column:order]
```

Examples:
```
[SORT:salary:descending]
[SORT:name:ascending]
```

### Select

```
[SELECT:column1,column2,column3]
```

Example:
```
[SELECT:name,department]
```

### Transform

```
[TRANSFORM:column:function]
```

Examples:
```
[TRANSFORM:name:uppercase]
[TRANSFORM:salary:format_currency]
```

## Combining Operations

You can combine multiple operations in sequence:

```
# CSV Data
name,age,department,salary
John Doe,30,Engineering,75000
Jane Smith,28,Marketing,65000
Bob Johnson,35,Sales,70000
Alice Brown,32,Engineering,80000

# Operations
[FILTER:department:equals:Engineering]
[SORT:salary:descending]
[TRANSFORM:salary:format_currency]
[SELECT:name,salary]
```

## Best Practices

1. Always validate CSV data before processing
2. Use clear column names without special characters
3. Handle missing values appropriately
4. Test operations on sample data first
5. Chain operations in logical order (filter before aggregate)
6. Use appropriate data types for numeric operations
7. Consider performance for large datasets
8. Document complex transformation pipelines

## Example Content

### Example 1: Filter and Sort

```
name,age,department,salary
John Doe,30,Engineering,75000
Jane Smith,28,Marketing,65000
Bob Johnson,35,Sales,70000
Alice Brown,32,Engineering,80000

[FILTER:salary:greater_than:70000]
[SORT:salary:descending]
```

### Example 2: Transform and Format

```
product,quantity,price
Widget A,100,19.99
Widget B,50,29.99
Widget C,75,9.99

[TRANSFORM:price:format_currency]
[ADD_COLUMN:name:total,formula:quantity * price]
[SORT:total:descending]
```

### Example 3: Aggregate by Group

```
date,department,revenue
2024-01-01,Sales,5000
2024-01-01,Marketing,3000
2024-01-02,Sales,6000
2024-01-02,Marketing,4000

[AGGREGATE:department:sum:revenue]
```

## API Call Format

To process CSV data, make a POST request to the endpoint with the following JSON structure:

```json
{
  "content": "Your CSV content with optional inline operations",
  "operations": [
    {"operation": "filter", "column": "age", "condition": "greater_than", "value": 30},
    {"operation": "sort", "column": "salary", "order": "descending"}
  ],
  "filename": "processed_data.csv"
}
```

### Example cURL Request

```bash
curl -X 'POST' \
  'http://localhost:19801/api/v1/process-csv' \
  -H 'accept: application/json' \
  -H 'Content-Type: application/json' \
  -d '{
  "content": "name,age,department,salary\nJohn Doe,30,Engineering,75000\nJane Smith,28,Marketing,65000\nBob Johnson,35,Sales,70000",
  "operations": [
    {"operation": "filter", "column": "age", "condition": "greater_than", "value": 30},
    {"operation": "sort", "column": "salary", "order": "descending"}
  ],
  "filename": "filtered_employees.csv"
}'
```

## Tool Call Integration

When integrating this tool into your application, use the following format:

```javascript
{
  "tool_name": "csv_processor",
  "parameters": {
    "content": "[Your CSV content string]",
    "operations": [
      {"operation": "filter", "column": "age", "condition": "greater_than", "value": 30}
    ],
    "filename": "output_filename.csv"
  }
}
```

By following these guidelines, you can effectively process and transform CSV data using the csv_processor tool.
