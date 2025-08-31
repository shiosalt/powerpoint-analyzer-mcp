#!/usr/bin/env python3
"""
Test script for server-level section extraction.
"""

import sys
import os
sys.path.insert(0, os.path.abspath('.'))

from powerpoint_mcp_server.server import PowerPointMCPServer
import asyncio
import logging

# Enable debug logging
logging.basicConfig(level=logging.DEBUG)

async def test_server_section_extraction():
    """Test section extraction through the server."""
    pptx_path = r"J:\MCP\powerpoint-analyzer\tests\test_files\test_complex.pptx"
    
    if not os.path.exists(pptx_path):
        print(f"File not found: {pptx_path}")
        return
    
    print(f"Testing server section extraction from: {pptx_path}")
    print("=" * 60)
    
    try:
        # Initialize server
        server = PowerPointMCPServer()
        
        # Test the internal method directly
        result = await server._process_powerpoint_file(pptx_path)
        
        print(f"Server result keys: {list(result.keys())}")
        
        if 'sections' in result:
            sections = result['sections']
            print(f"\nExtracted {len(sections)} sections:")
            for i, section in enumerate(sections, 1):
                print(f"  Section {i}:")
                print(f"    Name: {section.get('name', 'N/A')}")
                print(f"    ID: {section.get('id', 'N/A')}")
                print(f"    Slide count: {section.get('slide_count', 0)}")
                print(f"    Slide IDs: {section.get('slide_ids', [])}")
        else:
            print("No 'sections' key in result")
            
        # Also test metadata
        if 'metadata' in result:
            metadata = result['metadata']
            print(f"\nMetadata keys: {list(metadata.keys())}")
        
    except Exception as e:
        print(f"Error during testing: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    asyncio.run(test_server_section_extraction())