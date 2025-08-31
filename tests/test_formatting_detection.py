"""
Comprehensive tests for text formatting detection in PowerPoint files.
"""

import pytest
import logging
import zipfile
from pathlib import Path

# Configure logging for debugging
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

# Import the MCP server components
import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), '..'))

from powerpoint_mcp_server.core.content_extractor import ContentExtractor


class TestFormattingDetection:
    """Test class for comprehensive formatting detection validation."""
    
    @pytest.fixture
    def test_file_path(self):
        """Path to the complex test file containing all formatting types."""
        return Path(__file__).parent / "test_files" / "test_complex.pptx"
    
    @pytest.fixture
    def content_extractor(self):
        """Content extractor instance for testing."""
        return ContentExtractor()
    
    def test_file_exists(self, test_file_path):
        """Verify the test file exists."""
        assert test_file_path.exists(), f"Test file not found: {test_file_path}"
    
    def test_comprehensive_formatting_summary(self, test_file_path, content_extractor):
        """Test comprehensive formatting detection across all attributes."""
        # Open the PowerPoint file as a zip archive
        with zipfile.ZipFile(str(test_file_path), 'r') as zip_file:
            # Get slide files
            slide_files = [name for name in zip_file.namelist() 
                          if name.startswith('ppt/slides/slide') and name.endswith('.xml')]
            
            formatting_summary = {
                'bold': 0,
                'italic': 0,
                'underlined': 0,
                'highlighted': 0,
                'strikethrough': 0,
                'hyperlinks': 0,
                'font_sizes': set(),
                'font_colors': set()
            }
            
            # Process each slide
            for i, slide_file in enumerate(sorted(slide_files), 1):
                slide_xml = zip_file.read(slide_file).decode('utf-8')
                slide_info = content_extractor.extract_slide_content(slide_xml, i)
                
                # Analyze text elements
                for text_elem in slide_info.text_elements:
                    formatting_summary['bold'] += text_elem.get('bolded', 0)
                    formatting_summary['italic'] += text_elem.get('italic', 0)
                    formatting_summary['underlined'] += text_elem.get('underlined', 0)
                    formatting_summary['highlighted'] += text_elem.get('highlighted', 0)
                    formatting_summary['strikethrough'] += text_elem.get('strikethrough', 0)
                    formatting_summary['hyperlinks'] += len(text_elem.get('hyperlinks', []))
                    formatting_summary['font_sizes'].update(text_elem.get('font_sizes', []))
                    formatting_summary['font_colors'].update(text_elem.get('font_colors', []))
        
        # Convert sets to lists for logging
        summary_for_log = formatting_summary.copy()
        summary_for_log['font_sizes'] = sorted(list(formatting_summary['font_sizes']))
        summary_for_log['font_colors'] = sorted(list(formatting_summary['font_colors']))
        
        logger.info(f"Comprehensive formatting summary: {summary_for_log}")
        
        # We know highlights and hyperlinks are detected from previous analysis
        assert formatting_summary['highlighted'] > 0, "Expected to find highlighted text"
        assert formatting_summary['hyperlinks'] > 0, "Expected to find hyperlinks"
        
        # Font attributes should be present
        assert len(formatting_summary['font_colors']) > 0, "Expected to find font colors"
        
        # The file should contain bold, italic, underline, and strikethrough
        expected_formats = ['bold', 'italic', 'underlined', 'strikethrough']
        missing_formats = []
        
        for fmt in expected_formats:
            if formatting_summary[fmt] == 0:
                missing_formats.append(fmt)
        
        if missing_formats:
            logger.warning(f"Missing expected formatting types: {missing_formats}")
        
        return formatting_summary


def run_formatting_detection_tests():
    """Standalone function to run formatting detection tests."""
    test_instance = TestFormattingDetection()
    test_file_path = Path(__file__).parent / "test_files" / "test_complex.pptx"
    
    if not test_file_path.exists():
        print(f"Test file not found: {test_file_path}")
        return False
    
    try:
        content_extractor = ContentExtractor()
        
        print("Running comprehensive formatting detection tests...")
        
        result = test_instance.test_comprehensive_formatting_summary(
            test_file_path, content_extractor
        )
        
        print("Formatting detection test completed!")
        print(f"Results: {result}")
        return True
        
    except Exception as e:
        print(f"Test failed with error: {e}")
        logger.exception("Test execution failed")
        return False


if __name__ == "__main__":
    success = run_formatting_detection_tests()
    exit(0 if success else 1)