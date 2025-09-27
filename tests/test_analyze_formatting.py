#!/usr/bin/env python3
"""
Simple test script to verify analyze_text_formatting tool functionality.
"""

import asyncio
import json
import logging
import sys
from pathlib import Path

# Add the project root to the path
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

from powerpoint_mcp_server.server import PowerPointMCPServer

# Configure logging
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)

logger = logging.getLogger(__name__)

async def test_analyze_text_formatting():
    """Test the analyze_text_formatting functionality."""
    
    # Test file path
    test_file = "tests/test_files/test_complex.pptx"
    
    # Check if test file exists
    if not Path(test_file).exists():
        logger.error(f"Test file not found: {test_file}")
        return False
    
    try:
        # Create server instance
        server = PowerPointMCPServer()
        
        # Test analyze_text_formatting
        logger.info("Testing analyze_text_formatting...")
        
        arguments = {
            "file_path": test_file,
            "slide_numbers": [],  # Analyze all slides
            "formatting_filter": {
                "include_bold": True,
                "include_formatting_details": True
            },
            "grouping": "by_formatting_type"
        }
        
        result = await server._analyze_text_formatting(arguments)
        
        # Extract text content from CallToolResult
        content_text = ""
        if result.content:
            for content_item in result.content:
                if hasattr(content_item, 'text'):
                    content_text += content_item.text
        
        # Parse the JSON response
        response_data = json.loads(content_text)
        
        logger.info("=== ANALYZE TEXT FORMATTING RESULT ===")
        logger.info(f"Response keys: {list(response_data.keys())}")
        
        # Check for formatting_summary
        if "formatting_summary" in response_data:
            formatting_summary = response_data["formatting_summary"]
            logger.info(f"Formatting summary: {json.dumps(formatting_summary, indent=2)}")
            
            # Check formatting_counts
            if "formatting_counts" in formatting_summary:
                formatting_counts = formatting_summary["formatting_counts"]
                logger.info(f"Formatting counts: {json.dumps(formatting_counts, indent=2)}")
                
                # Calculate total formatting
                total_formatting = sum([
                    formatting_counts.get("bold", 0),
                    formatting_counts.get("italic", 0),
                    formatting_counts.get("underline", 0),
                    formatting_counts.get("highlight", 0),
                    formatting_counts.get("strikethrough", 0),
                    formatting_counts.get("colored_text", 0),
                    formatting_counts.get("hyperlinks", 0)
                ])
                
                logger.info(f"Total formatting found: {total_formatting}")
                
                if total_formatting > 0:
                    logger.info("✅ SUCCESS: Formatting counts are non-zero")
                    return True
                else:
                    logger.warning("⚠️  WARNING: No formatting detected - this might indicate a bug")
                    return False
            else:
                logger.error("❌ ERROR: formatting_counts not found in formatting_summary")
                return False
        else:
            logger.error("❌ ERROR: formatting_summary not found in response")
            logger.info(f"Full response: {json.dumps(response_data, indent=2)}")
            return False
            
    except Exception as e:
        logger.error(f"❌ ERROR: Exception occurred: {e}")
        import traceback
        logger.error(f"Traceback: {traceback.format_exc()}")
        return False

async def test_extract_formatted_text():
    """Test the extract_formatted_text functionality."""
    
    # Test file path
    test_file = "tests/test_files/test_complex.pptx"
    
    try:
        # Create server instance
        server = PowerPointMCPServer()
        
        # Test extract_formatted_text for different types
        formatting_types = ["bold", "italic", "hyperlinks"]
        
        for formatting_type in formatting_types:
            logger.info(f"Testing extract_formatted_text for {formatting_type}...")
            
            arguments = {
                "file_path": test_file,
                "formatting_type": formatting_type,
                "slide_numbers": []  # All slides
            }
            
            result = await server._extract_text_formatting(arguments)
            
            # Extract text content from CallToolResult
            content_text = ""
            if result.content:
                for content_item in result.content:
                    if hasattr(content_item, 'text'):
                        content_text += content_item.text
            
            # Parse the JSON response
            response_data = json.loads(content_text)
            
            logger.info(f"=== EXTRACT TEXT FORMATTING RESULT ({formatting_type}) ===")
            
            # Check basic structure
            if "formatting_type" in response_data and "summary" in response_data:
                summary = response_data["summary"]
                logger.info(f"Summary: {json.dumps(summary, indent=2)}")
                
                total_segments = summary.get("total_formatted_segments", 0)
                logger.info(f"Total {formatting_type} segments found: {total_segments}")
                
                if total_segments > 0:
                    logger.info(f"✅ SUCCESS: {formatting_type} formatting detected")
                else:
                    logger.info(f"ℹ️  INFO: No {formatting_type} formatting found")
            else:
                logger.error(f"❌ ERROR: Invalid response structure for {formatting_type}")
                logger.info(f"Response keys: {list(response_data.keys())}")
        
        return True
        
    except Exception as e:
        logger.error(f"❌ ERROR: Exception in extract_formatted_text test: {e}")
        import traceback
        logger.error(f"Traceback: {traceback.format_exc()}")
        return False

async def main():
    """Run all tests."""
    logger.info("Starting PowerPoint MCP formatting tests...")
    
    # Test analyze_text_formatting
    analyze_success = await test_analyze_text_formatting()
    
    # Test extract_formatted_text
    extract_success = await test_extract_formatted_text()
    
    if analyze_success and extract_success:
        logger.info("🎉 All tests completed successfully!")
        return True
    else:
        logger.error("❌ Some tests failed")
        return False

if __name__ == "__main__":
    success = asyncio.run(main())
    sys.exit(0 if success else 1)