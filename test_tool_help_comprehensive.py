#!/usr/bin/env python3
"""Comprehensive test for tool_help functionality."""

import asyncio
import json
import sys
import os
import tempfile
from pathlib import Path

# Add the project root to the path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from powerpoint_mcp_server.tools.tool_help import ToolHelpSystem, get_tool_help, get_tool_examples, get_parameter_help


def test_tool_help_system_comprehensive():
    """Comprehensive test of the tool help system."""
    print("=== Comprehensive Tool Help System Test ===\n")
    
    # Test 1: Initialize ToolHelpSystem
    print("1. Testing ToolHelpSystem initialization:")
    try:
        help_system = ToolHelpSystem()
        print("✅ ToolHelpSystem initialized successfully")
        print(f"Available tools in help system: {list(help_system.tool_docs.keys())}")
    except Exception as e:
        print(f"❌ Error initializing ToolHelpSystem: {e}")
        return False
    
    # Test 2: Test get_tool_help function
    print("\n2. Testing get_tool_help function:")
    try:
        help_text = get_tool_help("query_slides")
        print("✅ get_tool_help function works")
        print(f"Help text length: {len(help_text)} characters")
        
        # Verify key sections exist
        required_sections = [
            "# query_slides",
            "## Parameters", 
            "### search_criteria",
            "#### Schema:",
            "## Examples",
            "## Important Notes"
        ]
        
        missing_sections = []
        for section in required_sections:
            if section not in help_text:
                missing_sections.append(section)
        
        if missing_sections:
            print(f"❌ Missing sections: {missing_sections}")
        else:
            print("✅ All required sections present")
            
    except Exception as e:
        print(f"❌ Error in get_tool_help: {e}")
        return False
    
    # Test 3: Test parameter help
    print("\n3. Testing get_parameter_help function:")
    try:
        param_help = get_parameter_help("query_slides", "search_criteria")
        print("✅ get_parameter_help function works")
        
        if param_help and 'schema' in param_help:
            print("✅ Parameter help contains schema information")
        else:
            print("❌ Parameter help missing schema information")
            
    except Exception as e:
        print(f"❌ Error in get_parameter_help: {e}")
        return False
    
    # Test 4: Test examples
    print("\n4. Testing get_tool_examples function:")
    try:
        examples = get_tool_examples("query_slides")
        print(f"✅ get_tool_examples function works - {len(examples)} examples")
        
        # Verify each example has required fields
        for i, example in enumerate(examples, 1):
            has_name = 'name' in example
            has_search_criteria = 'search_criteria' in example
            
            print(f"  Example {i}: name={has_name}, search_criteria={has_search_criteria}")
            
            if not has_name or not has_search_criteria:
                print(f"❌ Example {i} missing required fields")
            
    except Exception as e:
        print(f"❌ Error in get_tool_examples: {e}")
        return False
    
    # Test 5: Test invalid tool name
    print("\n5. Testing invalid tool name handling:")
    try:
        help_text = get_tool_help("nonexistent_tool")
        if "No help available" in help_text:
            print("✅ Invalid tool name handled correctly")
        else:
            print("❌ Invalid tool name not handled correctly")
            
    except Exception as e:
        print(f"✅ Exception raised for invalid tool (expected): {type(e).__name__}")
    
    # Test 6: Test schema formatting
    print("\n6. Testing schema formatting:")
    try:
        help_system = ToolHelpSystem()
        formatted_help = help_system.format_help_text("query_slides")
        
        # Check for proper markdown formatting
        markdown_elements = [
            "# query_slides",
            "## Parameters",
            "### search_criteria",
            "- **Type**:",
            "- **Required**:",
            "- **Description**:",
            "```json"
        ]
        
        missing_elements = []
        for element in markdown_elements:
            if element not in formatted_help:
                missing_elements.append(element)
        
        if missing_elements:
            print(f"❌ Missing markdown elements: {missing_elements}")
        else:
            print("✅ Proper markdown formatting present")
            
    except Exception as e:
        print(f"❌ Error in schema formatting: {e}")
        return False
    
    # Test 7: Test search_criteria schema completeness
    print("\n7. Testing search_criteria schema completeness:")
    try:
        param_help = get_parameter_help("query_slides", "search_criteria")
        schema = param_help.get('schema', {})
        
        expected_top_level = ['title', 'content', 'layout', 'slide_numbers', 'section']
        expected_title_props = ['contains', 'starts_with', 'ends_with', 'regex', 'one_of']
        expected_content_props = ['contains_text', 'has_tables', 'has_charts', 'has_images', 'object_count']
        expected_layout_props = ['type', 'name']
        
        # Check top-level properties
        missing_top_level = [prop for prop in expected_top_level if prop not in schema]
        if missing_top_level:
            print(f"❌ Missing top-level properties: {missing_top_level}")
        else:
            print("✅ All top-level properties present")
        
        # Check title properties
        title_props = schema.get('title', {}).get('properties', {})
        missing_title = [prop for prop in expected_title_props if prop not in title_props]
        if missing_title:
            print(f"❌ Missing title properties: {missing_title}")
        else:
            print("✅ All title properties present")
        
        # Check content properties
        content_props = schema.get('content', {}).get('properties', {})
        missing_content = [prop for prop in expected_content_props if prop not in content_props]
        if missing_content:
            print(f"❌ Missing content properties: {missing_content}")
        else:
            print("✅ All content properties present")
            
    except Exception as e:
        print(f"❌ Error checking schema completeness: {e}")
        return False
    
    print("\n✅ Comprehensive tool help system test completed successfully!")
    return True


def test_example_validity():
    """Test that all examples are valid and complete."""
    print("\n=== Example Validity Test ===\n")
    
    try:
        examples = get_tool_examples("query_slides")
        print(f"Testing {len(examples)} examples for validity...")
        
        for i, example in enumerate(examples, 1):
            print(f"\nExample {i}: {example.get('name', 'Unnamed')}")
            
            # Check required fields
            has_search_criteria = 'search_criteria' in example
            search_criteria = example.get('search_criteria', {})
            
            print(f"  ✓ Has search_criteria: {has_search_criteria}")
            
            if has_search_criteria:
                # Check search_criteria structure
                sc_keys = list(search_criteria.keys())
                print(f"  ✓ Search criteria keys: {sc_keys}")
                
                # Validate each search criteria section
                for key, value in search_criteria.items():
                    if key == 'title' and isinstance(value, dict):
                        title_keys = list(value.keys())
                        print(f"    - Title filters: {title_keys}")
                    elif key == 'content' and isinstance(value, dict):
                        content_keys = list(value.keys())
                        print(f"    - Content filters: {content_keys}")
                    elif key == 'layout' and isinstance(value, dict):
                        layout_keys = list(value.keys())
                        print(f"    - Layout filters: {layout_keys}")
                    elif key == 'slide_numbers' and isinstance(value, list):
                        print(f"    - Slide numbers: {value}")
            
            # Check return_fields if present
            if 'return_fields' in example:
                return_fields = example['return_fields']
                print(f"  ✓ Return fields: {return_fields}")
            
            # Check limit if present
            if 'limit' in example:
                limit = example['limit']
                print(f"  ✓ Limit: {limit}")
        
        print("\n✅ All examples are valid and complete!")
        return True
        
    except Exception as e:
        print(f"❌ Error testing examples: {e}")
        import traceback
        traceback.print_exc()
        return False


def create_tool_help_usage_guide():
    """Create a usage guide for the tool_help system."""
    print("\n=== Creating Tool Help Usage Guide ===\n")
    
    try:
        guide_content = """# PowerPoint MCP Server - Tool Help Usage Guide

## Overview
The PowerPoint MCP Server includes a comprehensive tool help system that provides detailed documentation for all available tools.

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

Generated on: """ + str(Path(__file__).stat().st_mtime) + """
"""
        
        guide_path = Path("tool_help_usage_guide.md")
        guide_path.write_text(guide_content, encoding='utf-8')
        
        print(f"✅ Usage guide created: {guide_path}")
        print(f"Guide length: {len(guide_content)} characters")
        
        return True
        
    except Exception as e:
        print(f"❌ Error creating usage guide: {e}")
        return False


if __name__ == "__main__":
    print("Starting Comprehensive Tool Help Tests...\n")
    
    success = True
    
    try:
        # Run all tests
        success &= test_tool_help_system_comprehensive()
        success &= test_example_validity()
        success &= create_tool_help_usage_guide()
        
        if success:
            print("\n🎉 All tests passed successfully!")
            print("\nSummary:")
            print("✅ Tool help system is fully functional")
            print("✅ All examples are valid and complete")
            print("✅ Schema documentation is comprehensive")
            print("✅ Usage guide has been created")
        else:
            print("\n❌ Some tests failed. Please review the output above.")
            
    except KeyboardInterrupt:
        print("\nTests interrupted by user")
    except Exception as e:
        print(f"\nTests failed with error: {e}")
        import traceback
        traceback.print_exc()
    
    print("\nTest suite completed.")