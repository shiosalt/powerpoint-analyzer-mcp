#!/usr/bin/env python3
"""Test script for tool_help functionality."""

import asyncio
import json
import sys
import os

# Add the project root to the path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from powerpoint_mcp_server.server import PowerPointMCPServer
from powerpoint_mcp_server.tools.tool_help import get_tool_help, get_tool_examples


async def test_tool_help_system():
    """Test the tool help system comprehensively."""
    print("=== Tool Help System Test ===\n")
    
    # Test 1: Direct tool help function
    print("1. Testing direct tool help function:")
    try:
        help_text = get_tool_help("query_slides")
        print(f"✅ Successfully got help for query_slides")
        print(f"Help text length: {len(help_text)} characters")
        print(f"First 200 characters: {help_text[:200]}...")
    except Exception as e:
        print(f"❌ Error getting help: {e}")
    
    print("\n" + "="*50 + "\n")
    
    # Test 2: Server tool list
    print("2. Testing server tool list:")
    try:
        server = PowerPointMCPServer()
        tools_list = await server._get_tools_list()
        
        print(f"✅ Server initialized successfully")
        print(f"Total tools available: {len(tools_list)}")
        
        tool_names = [tool['name'] for tool in tools_list]
        print("Available tools:")
        for i, name in enumerate(tool_names, 1):
            print(f"  {i}. {name}")
        
        # Check if tool_help is in the list
        if 'tool_help' in tool_names:
            print("✅ tool_help is in the tools list")
        else:
            print("❌ tool_help is NOT in the tools list")
            
    except Exception as e:
        print(f"❌ Error testing server: {e}")
        import traceback
        traceback.print_exc()
    
    print("\n" + "="*50 + "\n")
    
    # Test 3: Tool help via server call
    print("3. Testing tool_help via server call:")
    try:
        server = PowerPointMCPServer()
        
        # Test tool_help call
        arguments = {"tool_name": "query_slides"}
        result = await server._tool_help(arguments)
        
        print("✅ Successfully called _tool_help method")
        print(f"Result type: {type(result)}")
        
        if hasattr(result, 'content') and result.content:
            content = result.content[0].text
            print(f"Content length: {len(content)} characters")
            print(f"First 200 characters: {content[:200]}...")
        
    except Exception as e:
        print(f"❌ Error calling tool_help: {e}")
        import traceback
        traceback.print_exc()
    
    print("\n" + "="*50 + "\n")
    
    # Test 4: Test examples
    print("4. Testing tool examples:")
    try:
        examples = get_tool_examples("query_slides")
        print(f"✅ Successfully got {len(examples)} examples for query_slides")
        
        for i, example in enumerate(examples, 1):
            print(f"  Example {i}: {example.get('name', 'Unnamed')}")
            
    except Exception as e:
        print(f"❌ Error getting examples: {e}")
    
    print("\n" + "="*50 + "\n")
    
    # Test 5: Test invalid tool name
    print("5. Testing invalid tool name:")
    try:
        help_text = get_tool_help("nonexistent_tool")
        if "No help available" in help_text:
            print("✅ Correctly handled invalid tool name")
        else:
            print("❌ Did not handle invalid tool name correctly")
    except Exception as e:
        print(f"✅ Correctly raised exception for invalid tool: {e}")


async def test_query_slides_help_details():
    """Test the detailed help for query_slides specifically."""
    print("=== Query Slides Help Details Test ===\n")
    
    help_text = get_tool_help("query_slides")
    
    # Check for key sections
    required_sections = [
        "# query_slides",
        "## Parameters", 
        "### search_criteria",
        "#### Schema:",
        "## Examples",
        "## Important Notes"
    ]
    
    print("Checking for required sections:")
    for section in required_sections:
        if section in help_text:
            print(f"✅ Found: {section}")
        else:
            print(f"❌ Missing: {section}")
    
    # Check for specific search_criteria properties
    search_criteria_props = [
        "title", "content", "layout", "slide_numbers", "section",
        "contains", "starts_with", "ends_with", "regex", "one_of",
        "contains_text", "has_tables", "has_charts", "has_images", "object_count"
    ]
    
    print(f"\nChecking for search_criteria properties:")
    for prop in search_criteria_props:
        if prop in help_text:
            print(f"✅ Found: {prop}")
        else:
            print(f"❌ Missing: {prop}")
    
    # Check for examples
    examples = get_tool_examples("query_slides")
    print(f"\nExamples available: {len(examples)}")
    
    for i, example in enumerate(examples, 1):
        name = example.get('name', 'Unnamed')
        has_search_criteria = 'search_criteria' in example
        has_return_fields = 'return_fields' in example
        print(f"  {i}. {name} - search_criteria: {has_search_criteria}, return_fields: {has_return_fields}")


if __name__ == "__main__":
    print("Starting Tool Help System Tests...\n")
    
    try:
        asyncio.run(test_tool_help_system())
        print("\n" + "="*70 + "\n")
        asyncio.run(test_query_slides_help_details())
        
    except KeyboardInterrupt:
        print("\nTest interrupted by user")
    except Exception as e:
        print(f"\nTest failed with error: {e}")
        import traceback
        traceback.print_exc()
    
    print("\nTest completed.")