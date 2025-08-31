#!/usr/bin/env python3
"""
Test script for section extraction functionality.
"""

import sys
import os
sys.path.insert(0, os.path.abspath('.'))

from powerpoint_mcp_server.core.content_extractor import ContentExtractor
from powerpoint_mcp_server.utils.zip_extractor import ZipExtractor
import logging

# Enable debug logging
logging.basicConfig(level=logging.DEBUG)

def test_section_extraction():
    """Test section extraction functionality."""
    pptx_path = r"tests\test_files\test_complex.pptx"
    
    if not os.path.exists(pptx_path):
        print(f"File not found: {pptx_path}")
        return
    
    print(f"Testing section extraction from: {pptx_path}")
    print("=" * 60)
    
    try:
        # Initialize components
        content_extractor = ContentExtractor()
        
        # Extract presentation.xml directly using zipfile
        import zipfile
        with zipfile.ZipFile(pptx_path, 'r') as zip_file:
            presentation_xml = zip_file.read('ppt/presentation.xml').decode('utf-8')
            
            if presentation_xml:
                print("Successfully read presentation.xml")
                print(f"XML length: {len(presentation_xml)} characters")
                
                # Test section extraction
                sections = content_extractor.extract_section_information(presentation_xml)
                
                print(f"\nExtracted {len(sections)} sections:")
                for i, section in enumerate(sections, 1):
                    print(f"  Section {i}:")
                    print(f"    Name: {section.get('name', 'N/A')}")
                    print(f"    ID: {section.get('id', 'N/A')}")
                    print(f"    Slide count: {section.get('slide_count', 0)}")
                    print(f"    Slide IDs: {section.get('slide_ids', [])}")
                
                if not sections:
                    print("No sections found. Debugging XML content...")
                    
                    # Show first 2000 characters of XML
                    print("\nFirst 2000 characters of presentation.xml:")
                    print("-" * 40)
                    print(presentation_xml[:2000])
                    
                    # Look for section-related content
                    if 'section' in presentation_xml.lower():
                        print("\n'section' found in XML content")
                        lines = presentation_xml.split('\n')
                        for i, line in enumerate(lines):
                            if 'section' in line.lower():
                                print(f"Line {i+1}: {line.strip()}")
                    else:
                        print("\n'section' not found in XML content")
            else:
                print("Failed to read presentation.xml")
                
    except Exception as e:
        print(f"Error during testing: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_section_extraction()