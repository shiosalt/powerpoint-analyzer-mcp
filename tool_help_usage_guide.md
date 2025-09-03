# PowerPoint Analyzer MCP - Tool Help Usage Guide

## Overview
The PowerPoint Analyzer MCP includes a comprehensive tool help system that provides detailed documentation for all available tools.

## Available Functions

### 1. get_tool_help(tool_name: str) -> str
Returns formatted help text for a specific tool.

**Example:**
```python
help_text = get_tool_help("query_slides")
print(help_text)
```

### 2. get_tool_examples(tool_name: str) -> List[Dict]
Returns usage examples for a specific tool.

**Example:**
```python
examples = get_tool_examples("query_slides")
for example in examples:
    print(f"Example: {example['name']}")
    print(f"Criteria: {example['search_criteria']}")
```

### 3. get_parameter_help(tool_name: str, parameter_name: str) -> Dict
Returns detailed help for a specific parameter.

**Example:**
```python
param_help = get_parameter_help("query_slides", "search_criteria")
schema = param_help.get('schema', {})
```

## Using the MCP Tool

### Via MCP Client
```json
{
  "method": "tools/call",
  "params": {
    "name": "tool_help",
    "arguments": {
      "tool_name": "query_slides"
    }
  }
}
```

### Response Format
The tool_help MCP tool returns formatted markdown documentation including:
- Tool description
- Parameter specifications with types and requirements
- Detailed schema for complex parameters
- Usage examples with real scenarios
- Important notes and best practices

## Query Slides Tool Documentation

The query_slides tool supports flexible filtering with the following structure:

### search_criteria Schema
```json
{
  "title": {
    "contains": "string",
    "starts_with": "string", 
    "ends_with": "string",
    "regex": "string",
    "one_of": ["pattern1", "pattern2"]
  },
  "content": {
    "contains_text": "string",
    "has_tables": true/false,
    "has_charts": true/false,
    "has_images": true/false,
    "object_count": {"min": number, "max": number}
  },
  "layout": {
    "type": "layout_type",
    "name": "layout_name"
  },
  "slide_numbers": [1, 2, 3],
  "section": "section_name"
}
```

### return_fields Options
- slide_number (always included)
- title
- subtitle  
- layout
- object_counts
- preview_text
- table_info
- full_content

## Best Practices

1. **Use specific filters**: Combine multiple criteria for precise results
2. **Limit return fields**: Only request needed fields for better performance
3. **Set appropriate limits**: Use the limit parameter to control result size
4. **Test with examples**: Use provided examples as starting points
5. **Check help regularly**: Tool capabilities may expand over time

## Troubleshooting

- **No results**: Check filter criteria are not too restrictive
- **Too many results**: Add more specific filters or reduce limit
- **Invalid parameters**: Use get_parameter_help() for parameter details
- **Schema questions**: Refer to the detailed schema documentation

Generated on: 1756835114.0
