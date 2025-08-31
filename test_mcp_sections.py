#!/usr/bin/env python3
"""
Test script for MCP server section extraction.
"""

import sys
import os
sys.path.insert(0, os.path.abspath('.'))

from powerpoint_mcp_server.server import PowerPointMCPServer
from powerpoint_mcp_server.utils.zip_extractor import ZipExtractor
import asyncio
import logging

# Enable debug logging
logging.basicConfig(level=logging.DEBUG)

async def test_mcp_section_extraction():
    """Test MCP server section extraction."""
    pptx_path = r"J:\MCP\powerpoint-analyzer\tests\test_files\test_complex.pptx"
    
    if not os.path.exists(pptx_path):
        print(f"File not found: {pptx_path}")
        return
    
    print(f"Testing MCP server section extraction from: {pptx_path}")
    print("=" * 60)
    
    try:
        # Test ZipExtractor directly first
        print("1. Testing ZipExtractor directly...")
        with ZipExtractor(pptx_path) as extractor:
            presentation_xml = extractor.read_xml_content('ppt/presentation.xml')
            if presentation_xml:
                print(f"   ✓ Successfully read presentation.xml ({len(presentation_xml)} chars)")
                
                # Check if section information is in the XML
                if 'section' in presentation_xml.lower():
                    print("   ✓ 'section' found in presentation.xml")
                else:
                    print("   ✗ 'section' not found in presentation.xml")
            else:
                print("   ✗ Failed to read presentation.xml")
        
        print("\n2. Testing server _process_powerpoint_file method...")
        server = PowerPointMCPServer()
        result = await server._process_powerpoint_file(pptx_path)
        
        print(f"   Result keys: {list(result.keys())}")
        
        if 'sections' in result:
            sections = result['sections']
            print(f"   ✓ Found {len(sections)} sections:")
            for i, section in enumerate(sections, 1):
                print(f"     Section {i}: {section.get('name', 'N/A')} (ID: {section.get('id', 'N/A')})")
        else:
            print("   ✗ No 'sections' key in result")
        
        print("\n3. Testing MCP tool call...")
        # Simulate MCP tool call
        arguments = {"file_path": pptx_path}
        tool_result = await server._extract_powerpoint_content(arguments)
        
        # Parse the JSON result
        import json
        result_text = tool_result.content[0].text
        parsed_result = json.loads(result_text)
        
        if 'sections' in parsed_result:
            sections = parsed_result['sections']
            print(f"   ✓ MCP tool returned {len(sections)} sections:")
            for i, section in enumerate(sections, 1):
                print(f"     Section {i}: {section.get('name', 'N/A')} (ID: {section.get('id', 'N/A')})")
        else:
            print("   ✗ No 'sections' key in MCP tool result")
            
    except Exception as e:
        print(f"Error during testing: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    asyncio.run(test_mcp_section_extraction())