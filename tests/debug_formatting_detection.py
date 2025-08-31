"""
Debug script for formatting detection issues.

This script provides detailed debugging information about why certain
formatting attributes (bold, italic, underline, strikethrough) are not
being detected in PowerPoint files.
"""

import sys
import os
import logging
import json
from pathlib import Path
import xml.etree.ElementTree as ET

# Add parent directory to path
sys.path.append(os.path.join(os.path.dirname(__file__), '..'))

from powerpoint_mcp_server.core.content_extractor import ContentExtractor
from powerpoint_mcp_server.core.xml_parser import XMLParser

# Configure detailed logging
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


class FormattingDebugger:
    """Debug class for investigating formatting detection issues."""
    
    def __init__(self):
        self.content_extractor = ContentExtractor()
        self.xml_parser = XMLParser()
    
    def debug_file_structure(self, file_path: str):
        """Debug the internal structure of a PowerPoint file."""
        print(f"\n=== DEBUGGING FILE STRUCTURE: {file_path} ===")
        
        try:
            import zipfile
            zip_data = zipfile.ZipFile(file_path, 'r')
            if not zip_data:
                print("ERROR: Failed to load file")
                return
            
            print(f"File loaded successfully. Archive contains {len(zip_data.namelist())} files:")
            
            # List all files in the archive
            for name in sorted(zip_data.namelist()):
                if name.startswith('ppt/slides/slide'):
                    print(f"  📄 {name}")
                elif name.startswith('ppt/'):
                    print(f"  📁 {name}")
            
            # Examine slide XML content
            slide_files = [name for name in zip_data.namelist() if name.startswith('ppt/slides/slide') and name.endswith('.xml')]
            
            for slide_file in slide_files[:3]:  # Examine first 3 slides
                print(f"\n--- EXAMINING {slide_file} ---")
                self.debug_slide_xml(zip_data, slide_file)
                
        except Exception as e:
            print(f"ERROR debugging file structure: {e}")
            logger.exception("File structure debug failed")
    
    def debug_slide_xml(self, zip_data, slide_file: str):
        """Debug the XML content of a specific slide."""
        try:
            xml_content = zip_data.read(slide_file).decode('utf-8')
            root = ET.fromstring(xml_content)
            
            print(f"  Root element: {root.tag}")
            print(f"  Namespaces: {root.attrib}")
            
            # Find all text runs with formatting
            namespaces = {
                'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
                'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
                'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
            }
            
            # Find all text runs
            runs = root.findall('.//a:r', namespaces)
            print(f"  Found {len(runs)} text runs")
            
            for i, run in enumerate(runs[:5]):  # Examine first 5 runs
                print(f"\n    --- RUN {i+1} ---")
                self.debug_text_run(run, namespaces)
                
        except Exception as e:
            print(f"  ERROR debugging slide XML: {e}")
            logger.exception(f"Slide XML debug failed for {slide_file}")
    
    def debug_text_run(self, run: ET.Element, namespaces: dict):
        """Debug a specific text run element."""
        try:
            # Get text content
            text_elem = run.find('.//a:t', namespaces)
            text_content = text_elem.text if text_elem is not None and text_elem.text else "[NO TEXT]"
            print(f"      Text: '{text_content}'")
            
            # Get run properties
            r_pr = run.find('.//a:rPr', namespaces)
            if r_pr is None:
                print("      No run properties found")
                return
            
            print(f"      Run properties XML:")
            print(f"        {ET.tostring(r_pr, encoding='unicode')}")
            
            # Check for specific formatting elements
            formatting_checks = [
                ('Bold', './/a:b'),
                ('Italic', './/a:i'),
                ('Underline', './/a:u'),
                ('Strikethrough', './/a:strike'),
                ('Highlight', './/a:highlight')
            ]
            
            for fmt_name, xpath in formatting_checks:
                elem = r_pr.find(xpath, namespaces)
                if elem is not None:
                    val = elem.get('val', 'DEFAULT')
                    print(f"      {fmt_name}: Found with val='{val}'")
                else:
                    print(f"      {fmt_name}: Not found")
            
            # Check for font size
            sz_elem = r_pr.find('.//a:sz', namespaces)
            if sz_elem is not None:
                size_val = sz_elem.get('val', 'NO_VAL')
                print(f"      Font size: {size_val} (hundredths of points)")
            
            # Check for font color
            solid_fill = r_pr.find('.//a:solidFill', namespaces)
            if solid_fill is not None:
                srgb_clr = solid_fill.find('.//a:srgbClr', namespaces)
                scheme_clr = solid_fill.find('.//a:schemeClr', namespaces)
                if srgb_clr is not None:
                    color_val = srgb_clr.get('val', 'NO_VAL')
                    print(f"      Font color (RGB): #{color_val}")
                elif scheme_clr is not None:
                    color_val = scheme_clr.get('val', 'NO_VAL')
                    print(f"      Font color (Scheme): {color_val}")
            
        except Exception as e:
            print(f"      ERROR debugging text run: {e}")
            logger.exception("Text run debug failed")
    
    def debug_content_extraction(self, file_path: str):
        """Debug the content extraction process."""
        print(f"\n=== DEBUGGING CONTENT EXTRACTION: {file_path} ===")
        
        try:
            import zipfile
            with zipfile.ZipFile(file_path, 'r') as zip_data:
                # Get slide files
                slide_files = [name for name in zip_data.namelist() 
                              if name.startswith('ppt/slides/slide') and name.endswith('.xml')]
                
                slides_data = []
                for i, slide_file in enumerate(sorted(slide_files), 1):
                    slide_xml = zip_data.read(slide_file).decode('utf-8')
                    slide_info = self.content_extractor.extract_slide_content(slide_xml, i)
                    slides_data.append({
                        'slide_number': i,
                        'title': slide_info.title,
                        'text_elements': slide_info.text_elements
                    })
            
            print(f"Extracted {len(slides_data)} slides")
            
            for slide in slides_data:
                slide_num = slide.get('slide_number', 0)
                title = slide.get('title', 'No title')
                text_elements = slide.get('text_elements', [])
                
                print(f"\n--- SLIDE {slide_num}: {title} ---")
                print(f"  Text elements: {len(text_elements)}")
                
                for i, text_elem in enumerate(text_elements):
                    content = text_elem.get('content_plain', '')[:50]
                    print(f"    Element {i+1}: '{content}...'")
                    print(f"      Bold: {text_elem.get('bolded', 0)}")
                    print(f"      Italic: {text_elem.get('italic', 0)}")
                    print(f"      Underlined: {text_elem.get('underlined', 0)}")
                    print(f"      Highlighted: {text_elem.get('highlighted', 0)}")
                    print(f"      Strikethrough: {text_elem.get('strikethrough', 0)}")
                    print(f"      Hyperlinks: {len(text_elem.get('hyperlinks', []))}")
                    print(f"      Font sizes: {text_elem.get('font_sizes', [])}")
                    print(f"      Font colors: {text_elem.get('font_colors', [])}")
                
        except Exception as e:
            print(f"ERROR debugging content extraction: {e}")
            logger.exception("Content extraction debug failed")
    
    def run_comprehensive_debug(self, file_path: str):
        """Run comprehensive debugging on a PowerPoint file."""
        print(f"🔍 COMPREHENSIVE FORMATTING DEBUG")
        print(f"File: {file_path}")
        print("=" * 80)
        
        if not Path(file_path).exists():
            print(f"❌ ERROR: File not found: {file_path}")
            return
        
        # Debug file structure and XML content
        self.debug_file_structure(file_path)
        
        # Debug content extraction
        self.debug_content_extraction(file_path)
        
        print("\n" + "=" * 80)
        print("🔍 DEBUG COMPLETE")


def main():
    """Main function to run formatting debugging."""
    test_file = Path(__file__).parent / "test_files" / "test_complex.pptx"
    
    if len(sys.argv) > 1:
        test_file = Path(sys.argv[1])
    
    debugger = FormattingDebugger()
    debugger.run_comprehensive_debug(str(test_file))


if __name__ == "__main__":
    main()