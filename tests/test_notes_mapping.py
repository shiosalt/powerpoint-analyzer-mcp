#!/usr/bin/env python3
"""
Test script to debug notes-slide mapping issue.
"""

import logging
import sys
import os

# Add the project root to the path
sys.path.insert(0, os.path.abspath('.'))

from powerpoint_mcp_server.core.content_extractor import ContentExtractor
from powerpoint_mcp_server.utils.zip_extractor import ZipExtractor

# Set up logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def test_notes_mapping():
    """Test the notes-slide mapping functionality."""
    test_file = r"J:\MCP\powerpoint-analyzer\tests\test_files\test_complex.pptx"
    
    if not os.path.exists(test_file):
        logger.error(f"Test file not found: {test_file}")
        return
    
    logger.info(f"Testing notes mapping with file: {test_file}")
    
    extractor = ContentExtractor()
    zip_extractor = ZipExtractor(test_file)
    
    with zip_extractor.extract_archive():
        # List all files in the zip for debugging
        logger.info("All files in the PowerPoint archive:")
        all_files = zip_extractor.list_archive_contents()
        for filename in all_files:
            if 'notes' in filename.lower():
                logger.info(f"  Notes-related file: {filename}")
        
        # Test the notes mapping
        notes_mapping = extractor._build_notes_slide_mapping(zip_extractor)
        logger.info(f"Notes mapping result: {notes_mapping}")
        
        # Test extracting notes content
        logger.info("Testing notes extraction:")
        notes = extractor.extract_notes(zip_extractor)
        for note in notes:
            logger.info(f"  Note for slide {note['slide_number']}: {note['content'][:100]}...")

if __name__ == "__main__":
    test_notes_mapping()