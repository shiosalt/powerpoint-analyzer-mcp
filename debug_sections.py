#!/usr/bin/env python3
"""
Debug script for PowerPoint section information extraction.
"""

import zipfile
import xml.etree.ElementTree as ET
import sys
import os

def debug_powerpoint_sections(pptx_path):
    """Debug PowerPoint section information."""
    if not os.path.exists(pptx_path):
        print(f"File not found: {pptx_path}")
        return
    
    print(f"Debugging PowerPoint sections in: {pptx_path}")
    print("=" * 60)
    
    try:
        with zipfile.ZipFile(pptx_path, 'r') as zip_file:
            # Read presentation.xml
            presentation_xml = zip_file.read('ppt/presentation.xml').decode('utf-8')
            
            # Parse XML
            root = ET.fromstring(presentation_xml)
            
            # Define namespaces
            namespaces = {
                'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
                'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
                'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
            }
            
            print("1. Looking for section list (p:sectionLst)...")
            section_list = root.find('.//p:sectionLst', namespaces)
            if section_list is not None:
                print("   ✓ Found p:sectionLst element")
                
                sections = section_list.findall('.//p:section', namespaces)
                print(f"   Found {len(sections)} sections:")
                
                for i, section in enumerate(sections, 1):
                    name = section.get('name', 'Unnamed Section')
                    section_id = section.get('id', 'No ID')
                    print(f"     Section {i}: name='{name}', id='{section_id}'")
                    
                    # Look for slide references in section
                    slide_refs = section.findall('.//p:sldId', namespaces)
                    if slide_refs:
                        print(f"       Slides in section: {len(slide_refs)}")
                        for slide_ref in slide_refs:
                            slide_id = slide_ref.get('id', 'No ID')
                            r_id = slide_ref.get('r:id', 'No r:id')
                            print(f"         Slide: id='{slide_id}', r:id='{r_id}'")
                    else:
                        print("       No slide references found in section")
            else:
                print("   ✗ No p:sectionLst element found")
            
            print("\n2. Checking presentation structure...")
            
            # Check slide master list
            slide_master_list = root.find('.//p:sldMasterIdLst', namespaces)
            if slide_master_list is not None:
                masters = slide_master_list.findall('.//p:sldMasterId', namespaces)
                print(f"   Found {len(masters)} slide masters")
            
            # Check slide list
            slide_list = root.find('.//p:sldIdLst', namespaces)
            if slide_list is not None:
                slides = slide_list.findall('.//p:sldId', namespaces)
                print(f"   Found {len(slides)} slides:")
                for i, slide in enumerate(slides, 1):
                    slide_id = slide.get('id', 'No ID')
                    r_id = slide.get('r:id', 'No r:id')
                    print(f"     Slide {i}: id='{slide_id}', r:id='{r_id}'")
            
            print("\n3. Raw XML structure around sections...")
            # Look for any element containing 'section' in its name
            for elem in root.iter():
                if 'section' in elem.tag.lower():
                    print(f"   Found element: {elem.tag}")
                    print(f"     Attributes: {elem.attrib}")
                    if elem.text and elem.text.strip():
                        print(f"     Text: {elem.text.strip()}")
            
            print("\n4. Full presentation.xml content (first 3000 chars):")
            print("-" * 40)
            print(presentation_xml[:3000])
            if len(presentation_xml) > 3000:
                print("... (truncated)")
            
    except Exception as e:
        print(f"Error debugging sections: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python debug_sections.py <path_to_pptx_file>")
        sys.exit(1)
    
    pptx_path = sys.argv[1]
    debug_powerpoint_sections(pptx_path)