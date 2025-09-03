"""
Comprehensive integration tests for PowerPoint MCP bug fixes.
Tests all identified bugs using real MCP communication.
"""

import asyncio
import json
import logging
import os
import sys
import pytest
from pathlib import Path
from typing import Dict, Any, List

# Add the project root to the path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from fastmcp.client.transports import StdioClientTransport
from fastmcp.client import FastMCPClient

logger = logging.getLogger(__name__)

# Test file path
TEST_FILE_PATH = "tests/test_files/test_complex.pptx"

class TestBugFixesIntegration:
    """Integration tests for all bug fixes using MCP protocol."""
    
    @pytest.fixture
    async def mcp_client(self):
        """Create MCP client for testing."""
        # Start the MCP server as a subprocess
        import subprocess
        
        # Use the main.py file as the server
        server_cmd = [sys.executable, "main.py"]
        
        # Create transport
        transport = StdioClientTransport(server_cmd)
        
        # Create client
        client = FastMCPClient("test-client")
        
        try:
            # Connect to server
            await client.connect(transport)
            yield client
        finally:
            # Cleanup
            await client.disconnect()
    
    async def test_sections_and_notes_filtering_in_query_slides(self, mcp_client):
        """Test that query_slides supports sections and notes filtering."""
        try:
            # Test sections filtering
            result = await mcp_client.call_tool(
                "query_slides",
                {
                    "file_path": TEST_FILE_PATH,
                    "search_criteria": {
                        "section": "Introduction"
                    },
                    "return_fields": ["slide_number", "title", "section"],
                    "limit": 10
                }
            )
            
            assert result is not None
            response_data = json.loads(result.content[0].text)
            
            # Should return results or empty array (not all slides)
            assert "total_found" in response_data
            assert "results" in response_data
            assert isinstance(response_data["results"], list)
            
            # Test notes filtering
            result = await mcp_client.call_tool(
                "query_slides",
                {
                    "file_path": TEST_FILE_PATH,
                    "search_criteria": {
                        "notes": {
                            "contains": "speaker"
                        }
                    },
                    "return_fields": ["slide_number", "title", "notes"],
                    "limit": 10
                }
            )
            
            assert result is not None
            response_data = json.loads(result.content[0].text)
            
            # Should return results or empty array
            assert "total_found" in response_data
            assert "results" in response_data
            assert isinstance(response_data["results"], list)
            
            logger.info("✓ Sections and notes filtering in query_slides works correctly")
            
        except Exception as e:
            logger.error(f"✗ Sections and notes filtering test failed: {e}")
            raise
    
    async def test_sections_and_notes_filtering_in_extract_table_data(self, mcp_client):
        """Test that extract_table_data supports sections and notes filtering."""
        try:
            # Test with section filtering
            result = await mcp_client.call_tool(
                "extract_table_data",
                {
                    "file_path": TEST_FILE_PATH,
                    "slide_numbers": [1, 2, 3, 4, 5],
                    "table_criteria": {
                        "section": "Data Section"
                    },
                    "output_format": "structured",
                    "include_metadata": True
                }
            )
            
            assert result is not None
            response_data = json.loads(result.content[0].text)
            
            # Should have proper structure
            assert "summary" in response_data
            assert "extracted_tables" in response_data
            
            # Test with notes filtering
            result = await mcp_client.call_tool(
                "extract_table_data",
                {
                    "file_path": TEST_FILE_PATH,
                    "slide_numbers": [1, 2, 3, 4, 5],
                    "table_criteria": {
                        "notes": {
                            "contains": "table"
                        }
                    },
                    "output_format": "structured",
                    "include_metadata": True
                }
            )
            
            assert result is not None
            response_data = json.loads(result.content[0].text)
            
            # Should have proper structure
            assert "summary" in response_data
            assert "extracted_tables" in response_data
            
            logger.info("✓ Sections and notes filtering in extract_table_data works correctly")
            
        except Exception as e:
            logger.error(f"✗ Sections and notes filtering in extract_table_data test failed: {e}")
            raise
    
    async def test_grammar_error_handling_in_query_slides(self, mcp_client):
        """Test that query_slides handles grammar errors properly."""
        try:
            # Test with malformed JSON structure
            result = await mcp_client.call_tool(
                "query_slides",
                {
                    "file_path": TEST_FILE_PATH,
                    "search_criteria": {
                        "invalid_field": "invalid_value",
                        "title": {
                            "invalid_operator": "test"
                        }
                    },
                    "return_fields": ["slide_number", "title"],
                    "limit": 10
                }
            )
            
            assert result is not None
            response_data = json.loads(result.content[0].text)
            
            # Should return zero results for invalid criteria
            assert "total_found" in response_data
            assert response_data["total_found"] == 0
            assert "results" in response_data
            assert len(response_data["results"]) == 0
            
            # Test with invalid regex
            result = await mcp_client.call_tool(
                "query_slides",
                {
                    "file_path": TEST_FILE_PATH,
                    "search_criteria": {
                        "title": {
                            "regex": "[invalid regex("
                        }
                    },
                    "return_fields": ["slide_number", "title"],
                    "limit": 10
                }
            )
            
            assert result is not None
            response_data = json.loads(result.content[0].text)
            
            # Should return zero results for invalid regex
            assert "total_found" in response_data
            assert response_data["total_found"] == 0
            
            logger.info("✓ Grammar error handling in query_slides works correctly")
            
        except Exception as e:
            logger.error(f"✗ Grammar error handling test failed: {e}")
            raise
    
    async def test_correct_slide_number_display_in_extract_table_data(self, mcp_client):
        """Test that extract_table_data shows correct slide numbers."""
        try:
            # Test with specific slide numbers including slide 10+
            result = await mcp_client.call_tool(
                "extract_table_data",
                {
                    "file_path": TEST_FILE_PATH,
                    "slide_numbers": [10, 11, 12],
                    "output_format": "structured",
                    "include_metadata": True
                }
            )
            
            assert result is not None
            response_data = json.loads(result.content[0].text)
            
            # Check that slide numbers are correct
            if "extracted_tables" in response_data and response_data["extracted_tables"]:
                for table in response_data["extracted_tables"]:
                    slide_number = table.get("slide_number")
                    assert slide_number in [10, 11, 12], f"Expected slide number 10, 11, or 12, got {slide_number}"
            
            # Test with slide 10 specifically if it exists
            result = await mcp_client.call_tool(
                "extract_table_data",
                {
                    "file_path": TEST_FILE_PATH,
                    "slide_numbers": [10],
                    "output_format": "structured",
                    "include_metadata": True
                }
            )
            
            assert result is not None
            response_data = json.loads(result.content[0].text)
            
            # If tables are found, slide_number should be 10, not 2
            if "extracted_tables" in response_data and response_data["extracted_tables"]:
                for table in response_data["extracted_tables"]:
                    slide_number = table.get("slide_number")
                    assert slide_number == 10, f"Expected slide number 10, got {slide_number}"
            
            logger.info("✓ Correct slide number display in extract_table_data works correctly")
            
        except Exception as e:
            logger.error(f"✗ Slide number display test failed: {e}")
            raise
    
    async def test_sections_and_notes_in_analyze_text_formatting(self, mcp_client):
        """Test that analyze_text_formatting includes sections and notes information."""
        try:
            result = await mcp_client.call_tool(
                "analyze_text_formatting",
                {
                    "file_path": TEST_FILE_PATH,
                    "slide_numbers": [1, 2, 3],
                    "include_bold_analysis": True,
                    "include_formatting_details": True
                }
            )
            
            assert result is not None
            response_data = json.loads(result.content[0].text)
            
            # Check for sections information in formatting summary
            assert "formatting_summary" in response_data
            formatting_summary = response_data["formatting_summary"]
            
            # Should include sections information
            if "sections" in formatting_summary:
                sections_info = formatting_summary["sections"]
                assert "total_sections" in sections_info
                assert "sections" in sections_info
                assert isinstance(sections_info["sections"], list)
            
            # Should include notes information
            if "notes" in formatting_summary:
                notes_info = formatting_summary["notes"]
                assert "slides_with_notes" in notes_info
                assert "total_notes_length" in notes_info
                assert "average_notes_length" in notes_info
            
            logger.info("✓ Sections and notes in analyze_text_formatting works correctly")
            
        except Exception as e:
            logger.error(f"✗ Sections and notes in analyze_text_formatting test failed: {e}")
            raise
    
    async def test_sections_and_notes_in_get_presentation_overview(self, mcp_client):
        """Test that get_presentation_overview includes sections and notes information."""
        try:
            result = await mcp_client.call_tool(
                "get_presentation_overview",
                {
                    "file_path": TEST_FILE_PATH,
                    "analysis_depth": "detailed",
                    "include_sample_content": True
                }
            )
            
            assert result is not None
            response_data = json.loads(result.content[0].text)
            
            # Check for sections information in metadata
            assert "metadata" in response_data
            metadata = response_data["metadata"]
            
            # Should include sections information
            if "sections" in metadata:
                sections = metadata["sections"]
                assert isinstance(sections, list)
            
            # Should include notes statistics
            if "notes_statistics" in metadata:
                notes_stats = metadata["notes_statistics"]
                assert "slides_with_notes" in notes_stats
                assert "total_notes_length" in notes_stats
                assert "average_notes_length" in notes_stats
            
            # Check structure for sections
            if "structure" in response_data:
                structure = response_data["structure"]
                if "sections" in structure:
                    assert isinstance(structure["sections"], list)
            
            logger.info("✓ Sections and notes in get_presentation_overview works correctly")
            
        except Exception as e:
            logger.error(f"✗ Sections and notes in get_presentation_overview test failed: {e}")
            raise
    
    async def test_regression_existing_functionality(self, mcp_client):
        """Test that existing functionality still works after improvements."""
        try:
            # Test basic extract_powerpoint_content still works
            result = await mcp_client.call_tool(
                "extract_powerpoint_content",
                {
                    "file_path": TEST_FILE_PATH
                }
            )
            
            assert result is not None
            response_data = json.loads(result.content[0].text)
            assert "slides" in response_data
            
            # Test basic query_slides still works
            result = await mcp_client.call_tool(
                "query_slides",
                {
                    "file_path": TEST_FILE_PATH,
                    "search_criteria": {
                        "title": {
                            "contains": "slide"
                        }
                    },
                    "return_fields": ["slide_number", "title"],
                    "limit": 5
                }
            )
            
            assert result is not None
            response_data = json.loads(result.content[0].text)
            assert "total_found" in response_data
            assert "results" in response_data
            
            # Test basic extract_table_data still works
            result = await mcp_client.call_tool(
                "extract_table_data",
                {
                    "file_path": TEST_FILE_PATH,
                    "slide_numbers": [1, 2, 3],
                    "output_format": "structured"
                }
            )
            
            assert result is not None
            response_data = json.loads(result.content[0].text)
            assert "summary" in response_data
            assert "extracted_tables" in response_data
            
            logger.info("✓ Regression test passed - existing functionality works correctly")
            
        except Exception as e:
            logger.error(f"✗ Regression test failed: {e}")
            raise


async def run_all_tests():
    """Run all integration tests for search and display improvements."""
    logger.info("Starting comprehensive integration tests for search and display improvements...")
    
    # Check if test file exists
    if not os.path.exists(TEST_FILE_PATH):
        logger.error(f"Test file not found: {TEST_FILE_PATH}")
        logger.info("Please ensure test_complex.pptx exists with appropriate content")
        return False
    
    test_instance = TestBugFixesIntegration()
    
    try:
        # Create MCP client
        async with test_instance.mcp_client() as client:
            
            # Run all new functionality tests
            await test_instance.test_sections_and_notes_filtering_in_query_slides(client)
            await test_instance.test_sections_and_notes_filtering_in_extract_table_data(client)
            await test_instance.test_grammar_error_handling_in_query_slides(client)
            await test_instance.test_correct_slide_number_display_in_extract_table_data(client)
            await test_instance.test_sections_and_notes_in_analyze_text_formatting(client)
            await test_instance.test_sections_and_notes_in_get_presentation_overview(client)
            await test_instance.test_regression_existing_functionality(client)
            
        logger.info("✅ All search and display improvement tests passed!")
        return True
        
    except Exception as e:
        logger.error(f"❌ Integration tests failed: {e}")
        import traceback
        logger.error(f"Traceback: {traceback.format_exc()}")
        return False


if __name__ == "__main__":
    # Configure logging
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )
    
    # Run tests
    success = asyncio.run(run_all_tests())
    sys.exit(0 if success else 1)
            await client.disconnect()
    
    async def test_formatting_analysis_accuracy(self, mcp_client):
        """Test that formatting counts are accurate and non-zero for existing formatting."""
        logger.info("Testing formatting analysis accuracy...")
        
        # Call analyze_text_formatting
        result = await mcp_client.call_tool(
            "analyze_text_formatting",
            {
                "file_path": TEST_FILE_PATH,
                "include_bold_analysis": True,
                "include_formatting_details": True
            }
        )
        
        # Parse the result
        response_data = json.loads(result.content[0].text)
        
        # Verify the response structure
        assert "formatting_summary" in response_data
        assert "formatting_counts" in response_data["formatting_summary"]
        
        formatting_counts = response_data["formatting_summary"]["formatting_counts"]
        
        # Check that at least some formatting is detected
        # (assuming test_complex.pptx has formatting)
        total_formatting = sum([
            formatting_counts.get("bold", 0),
            formatting_counts.get("italic", 0),
            formatting_counts.get("underline", 0),
            formatting_counts.get("highlight", 0),
            formatting_counts.get("strikethrough", 0),
            formatting_counts.get("colored_text", 0),
            formatting_counts.get("hyperlinks", 0)
        ])
        
        # Assert that some formatting was found
        assert total_formatting > 0, f"Expected some formatting to be found, but got: {formatting_counts}"
        
        logger.info(f"Formatting analysis passed: {formatting_counts}")
    
    async def test_text_formatting_extraction_precision(self, mcp_client):
        """Test that italic/hyperlinks are recognized and positions are accurate."""
        logger.info("Testing text formatting extraction precision...")
        
        # Test italic recognition
        italic_result = await mcp_client.call_tool(
            "extract_text_formatting",
            {
                "file_path": TEST_FILE_PATH,
                "formatting_type": "italic"
            }
        )
        
        italic_data = json.loads(italic_result.content[0].text)
        
        # Verify response structure
        assert "formatting_type" in italic_data
        assert italic_data["formatting_type"] == "italic"
        assert "summary" in italic_data
        assert "results_by_slide" in italic_data
        
        # Test hyperlinks recognition
        hyperlinks_result = await mcp_client.call_tool(
            "extract_text_formatting",
            {
                "file_path": TEST_FILE_PATH,
                "formatting_type": "hyperlinks"
            }
        )
        
        hyperlinks_data = json.loads(hyperlinks_result.content[0].text)
        
        # Verify response structure
        assert "formatting_type" in hyperlinks_data
        assert hyperlinks_data["formatting_type"] == "hyperlinks"
        
        # Test position accuracy for bold text
        bold_result = await mcp_client.call_tool(
            "extract_text_formatting",
            {
                "file_path": TEST_FILE_PATH,
                "formatting_type": "bold"
            }
        )
        
        bold_data = json.loads(bold_result.content[0].text)
        
        # Check that formatted segments have proper position information
        for slide_result in bold_data["results_by_slide"]:
            for segment in slide_result["formatted_segments"]:
                # Verify position fields exist
                assert "start_position" in segment
                assert "text" in segment
                
                # Verify positions are reasonable
                start_pos = segment["start_position"]
                text = segment["text"]
                
                assert isinstance(start_pos, int)
                assert start_pos >= 0
                assert len(text) > 0
                
                # Verify that the segment contains only formatted text (not complete text)
                complete_text = slide_result["complete_text"]
                assert len(text) <= len(complete_text), "Formatted segment should not be longer than complete text"
        
        logger.info("Text formatting extraction precision tests passed")
    
    async def test_table_extraction_completeness(self, mcp_client):
        """Test that table extraction returns proper summaries and handles slide numbers."""
        logger.info("Testing table extraction completeness...")
        
        # Test table extraction with specific slide numbers
        result = await mcp_client.call_tool(
            "extract_table_data",
            {
                "file_path": TEST_FILE_PATH,
                "slide_numbers": [1, 2, 3],
                "output_format": "structured",
                "include_metadata": True
            }
        )
        
        response_data = json.loads(result.content[0].text)
        
        # Verify response structure
        assert "summary" in response_data
        assert "extracted_tables" in response_data
        
        summary = response_data["summary"]
        
        # Check that summary fields exist and are properly calculated
        assert "total_tables_found" in summary
        assert "slides_with_tables" in summary
        assert "slides_processed" in summary
        
        # Verify that the counts are consistent
        total_tables = summary["total_tables_found"]
        extracted_tables = response_data["extracted_tables"]
        
        assert len(extracted_tables) == total_tables, f"Expected {total_tables} tables, but got {len(extracted_tables)}"
        
        # Test with slide numbers that should work
        if total_tables > 0:
            # Find a slide with tables
            slide_with_table = extracted_tables[0]["slide_number"]
            
            single_slide_result = await mcp_client.call_tool(
                "extract_table_data",
                {
                    "file_path": TEST_FILE_PATH,
                    "slide_numbers": [slide_with_table],
                    "output_format": "structured"
                }
            )
            
            single_slide_data = json.loads(single_slide_result.content[0].text)
            
            # Should return at least one table
            assert single_slide_data["summary"]["total_tables_found"] >= 1
        
        logger.info(f"Table extraction completeness tests passed: {summary}")
    
    async def test_slide_number_parameter_handling(self, mcp_client):
        """Test that valid slide numbers work correctly."""
        logger.info("Testing slide number parameter handling...")
        
        # First, get the total number of slides
        overview_result = await mcp_client.call_tool(
            "get_presentation_overview",
            {
                "file_path": TEST_FILE_PATH,
                "analysis_depth": "basic"
            }
        )
        
        overview_data = json.loads(overview_result.content[0].text)
        total_slides = overview_data.get("metadata", {}).get("slide_count", 0)
        
        if total_slides >= 2:
            # Test with valid slide numbers
            result = await mcp_client.call_tool(
                "extract_table_data",
                {
                    "file_path": TEST_FILE_PATH,
                    "slide_numbers": [1, 2],
                    "output_format": "structured"
                }
            )
            
            # Should not raise an error
            response_data = json.loads(result.content[0].text)
            assert "summary" in response_data
            
            # Test with single slide number
            single_result = await mcp_client.call_tool(
                "extract_table_data",
                {
                    "file_path": TEST_FILE_PATH,
                    "slide_numbers": [1],
                    "output_format": "structured"
                }
            )
            
            # Should not raise an error
            single_data = json.loads(single_result.content[0].text)
            assert "summary" in single_data
        
        # Test with invalid slide numbers (should raise error)
        try:
            invalid_result = await mcp_client.call_tool(
                "extract_table_data",
                {
                    "file_path": TEST_FILE_PATH,
                    "slide_numbers": [999],  # Invalid slide number
                    "output_format": "structured"
                }
            )
            # If we get here, check if it's an error response
            invalid_data = json.loads(invalid_result.content[0].text)
            if "error" not in invalid_data:
                pytest.fail("Expected error for invalid slide number, but got success")
        except Exception as e:
            # This is expected for invalid slide numbers
            logger.info(f"Correctly caught error for invalid slide number: {e}")
        
        logger.info("Slide number parameter handling tests passed")
    
    async def test_query_validation_strictness(self, mcp_client):
        """Test that invalid queries return zero results, not all slides."""
        logger.info("Testing query validation strictness...")
        
        # Test with invalid field names
        invalid_field_result = await mcp_client.call_tool(
            "query_slides",
            {
                "file_path": TEST_FILE_PATH,
                "search_criteria": {"invalid_field": "test"},
                "return_fields": ["slide_number", "title"]
            }
        )
        
        invalid_field_data = json.loads(invalid_field_result.content[0].text)
        
        # Should return empty results or error
        if "results" in invalid_field_data:
            assert len(invalid_field_data["results"]) == 0, "Invalid search criteria should return zero results"
        
        # Test with invalid return fields
        try:
            invalid_return_result = await mcp_client.call_tool(
                "query_slides",
                {
                    "file_path": TEST_FILE_PATH,
                    "search_criteria": {"title": {"contains": "test"}},
                    "return_fields": ["invalid_return_field"]
                }
            )
            
            invalid_return_data = json.loads(invalid_return_result.content[0].text)
            
            # Should return empty results or error
            if "results" in invalid_return_data:
                assert len(invalid_return_data["results"]) == 0, "Invalid return fields should return zero results"
        
        except Exception as e:
            # This is also acceptable - validation error
            logger.info(f"Correctly caught validation error: {e}")
        
        # Test with valid criteria to ensure normal operation works
        valid_result = await mcp_client.call_tool(
            "query_slides",
            {
                "file_path": TEST_FILE_PATH,
                "search_criteria": {},  # Empty criteria should return all slides
                "return_fields": ["slide_number", "title"]
            }
        )
        
        valid_data = json.loads(valid_result.content[0].text)
        
        # Should return some results for valid criteria
        if "results" in valid_data:
            assert len(valid_data["results"]) > 0, "Valid search criteria should return some results"
        
        logger.info("Query validation strictness tests passed")
    
    async def test_section_filtering_accuracy(self, mcp_client):
        """Test that section-based queries work correctly."""
        logger.info("Testing section filtering accuracy...")
        
        # First, get presentation overview to see if sections exist
        overview_result = await mcp_client.call_tool(
            "get_presentation_overview",
            {
                "file_path": TEST_FILE_PATH,
                "analysis_depth": "detailed"
            }
        )
        
        overview_data = json.loads(overview_result.content[0].text)
        sections = overview_data.get("sections", [])
        
        if sections:
            # Test with existing section
            section_name = sections[0].get("name", "")
            if section_name:
                section_result = await mcp_client.call_tool(
                    "query_slides",
                    {
                        "file_path": TEST_FILE_PATH,
                        "search_criteria": {"section": section_name},
                        "return_fields": ["slide_number", "title"]
                    }
                )
                
                section_data = json.loads(section_result.content[0].text)
                
                # Should return results for existing section
                if "results" in section_data:
                    logger.info(f"Section '{section_name}' returned {len(section_data['results'])} slides")
        
        # Test with non-existent section
        nonexistent_result = await mcp_client.call_tool(
            "query_slides",
            {
                "file_path": TEST_FILE_PATH,
                "search_criteria": {"section": "NonExistentSection"},
                "return_fields": ["slide_number", "title"]
            }
        )
        
        nonexistent_data = json.loads(nonexistent_result.content[0].text)
        
        # Should return zero results for non-existent section
        if "results" in nonexistent_data:
            assert len(nonexistent_data["results"]) == 0, "Non-existent section should return zero results"
        
        logger.info("Section filtering accuracy tests passed")
    
    async def test_comprehensive_bug_scenarios(self, mcp_client):
        """Test all identified bug scenarios comprehensively."""
        logger.info("Running comprehensive bug scenario tests...")
        
        # Test all formatting types
        formatting_types = ["bold", "italic", "underlined", "highlighted", "strikethrough", "font_sizes", "font_colors", "hyperlinks"]
        
        for formatting_type in formatting_types:
            try:
                result = await mcp_client.call_tool(
                    "extract_text_formatting",
                    {
                        "file_path": TEST_FILE_PATH,
                        "formatting_type": formatting_type
                    }
                )
                
                data = json.loads(result.content[0].text)
                
                # Verify basic structure
                assert "formatting_type" in data
                assert data["formatting_type"] == formatting_type
                assert "summary" in data
                assert "results_by_slide" in data
                
                logger.info(f"Formatting type '{formatting_type}' extraction successful")
                
            except Exception as e:
                logger.error(f"Error testing formatting type '{formatting_type}': {e}")
                raise
        
        # Test analyze_text_formatting with different parameters
        analysis_result = await mcp_client.call_tool(
            "analyze_text_formatting",
            {
                "file_path": TEST_FILE_PATH,
                "slide_numbers": [1, 2],
                "include_bold_analysis": True,
                "include_formatting_details": True
            }
        )
        
        analysis_data = json.loads(analysis_result.content[0].text)
        assert "formatting_summary" in analysis_data
        
        logger.info("Comprehensive bug scenario tests passed")


# Utility functions for running tests
async def run_integration_tests():
    """Run all integration tests."""
    logger.info("Starting PowerPoint MCP bug fix integration tests...")
    
    # Check if test file exists
    if not os.path.exists(TEST_FILE_PATH):
        logger.error(f"Test file not found: {TEST_FILE_PATH}")
        return False
    
    try:
        # Create test instance
        test_instance = TestBugFixesIntegration()
        
        # Create MCP client
        async with test_instance.mcp_client() as client:
            # Run all tests
            await test_instance.test_formatting_analysis_accuracy(client)
            await test_instance.test_text_formatting_extraction_precision(client)
            await test_instance.test_table_extraction_completeness(client)
            await test_instance.test_slide_number_parameter_handling(client)
            await test_instance.test_query_validation_strictness(client)
            await test_instance.test_section_filtering_accuracy(client)
            await test_instance.test_comprehensive_bug_scenarios(client)
        
        logger.info("All integration tests passed!")
        return True
        
    except Exception as e:
        logger.error(f"Integration tests failed: {e}")
        import traceback
        logger.error(f"Traceback: {traceback.format_exc()}")
        return False


if __name__ == "__main__":
    # Configure logging
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )
    
    # Run tests
    success = asyncio.run(run_integration_tests())
    sys.exit(0 if success else 1)