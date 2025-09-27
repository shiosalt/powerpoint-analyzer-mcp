"""
Comprehensive test suite for all MCP tools.
Tests each tool with all valid parameters and error conditions.
"""

import pytest
import asyncio
import json
import logging
from pathlib import Path
from typing import Dict, List, Any
from unittest.mock import AsyncMock, MagicMock

from tests.mcp_integration_test_framework import MCPIntegrationTestSuite, MCPTestClient, MCPTestResult

logger = logging.getLogger(__name__)


class TestComprehensiveMCPTools:
    """Comprehensive test suite for all MCP tools."""

    def setup_method(self):
        """Set up test fixtures."""
        self.test_files_dir = Path("tests/test_files")
        self.test_file = self.test_files_dir / "test_formatting_comprehensive.pptx"
        self.edge_case_file = self.test_files_dir / "test_edge_cases.pptx"

        # Ensure test files exist
        if not self.test_file.exists():
            pytest.skip("Test file not found. Run test_data_generator.py first.")

    # Removed tests for deleted tools: extract_powerpoint_content, get_powerpoint_attributes, get_slide_info

    @pytest.mark.asyncio
    async def test_extract_formatted_text_tool_all_types(self):
        """Test extract_formatted_text with all supported formatting types."""
        test_suite = MCPIntegrationTestSuite()

        try:
            assert await test_suite.setup()

            formatting_types = [
                "bold", "italic", "underlined", "highlighted",
                "strikethrough", "hyperlinks", "font_sizes", "font_colors"
            ]

            for formatting_type in formatting_types:
                success, response_data, response_time = await test_suite.client.call_tool(
                    "extract_formatted_text",
                    {
                        "file_path": str(self.test_file),
                        "formatting_type": formatting_type
                    }
                )

                assert success, f"Failed to extract {formatting_type} formatting: {response_data}"
                assert "formatting_type" in response_data
                assert response_data["formatting_type"] == formatting_type
                assert "summary" in response_data
                assert "results_by_slide" in response_data

                # Verify summary structure
                summary = response_data["summary"]
                assert "total_slides_analyzed" in summary
                assert "slides_with_formatting" in summary
                assert "total_formatted_segments" in summary

                # Verify results structure
                for slide_result in response_data["results_by_slide"]:
                    assert "slide_number" in slide_result
                    assert "complete_text" in slide_result
                    assert "format" in slide_result
                    assert "formatted_segments" in slide_result
                    assert slide_result["format"] == formatting_type

                    # Verify segment structure
                    for segment in slide_result["formatted_segments"]:
                        assert "text" in segment
                        assert "start_position" in segment
                        assert isinstance(segment["start_position"], int)

        finally:
            await test_suite.teardown()

    @pytest.mark.asyncio
    async def test_extract_formatted_text_with_slide_filter(self):
        """Test extract_formatted_text with slide number filtering."""
        test_suite = MCPIntegrationTestSuite()

        try:
            assert await test_suite.setup()

            # Test with specific slides
            success, response_data, response_time = await test_suite.client.call_tool(
                "extract_formatted_text",
                {
                    "file_path": str(self.test_file),
                    "formatting_type": "bold",
                    "slide_numbers": [1, 4]
                }
            )

            assert success, f"Failed to extract bold formatting with slide filter: {response_data}"

            # Verify only requested slides are included
            slide_numbers = [slide["slide_number"] for slide in response_data["results_by_slide"]]
            for slide_num in slide_numbers:
                assert slide_num in [1, 4], f"Unexpected slide number {slide_num} in filtered results"

        finally:
            await test_suite.teardown()


    @pytest.mark.asyncio
    async def test_query_slides_tool(self):
        """Test query_slides with various search criteria."""
        test_suite = MCPIntegrationTestSuite()

        try:
            assert await test_suite.setup()

            # Test query by slide numbers
            success, response_data, response_time = await test_suite.client.call_tool(
                "query_slides",
                {
                    "file_path": str(self.test_file),
                    "search_criteria": {"slide_numbers": [1, 2, 3]}
                }
            )

            assert success, f"Failed to query slides: {response_data}"

            # Verify response structure
            if "slides" in response_data:
                for slide in response_data["slides"]:
                    assert slide["slide_number"] in [1, 2, 3]

        finally:
            await test_suite.teardown()

    @pytest.mark.asyncio
    async def test_extract_table_data_tool(self):
        """Test extract_table_data tool."""
        test_suite = MCPIntegrationTestSuite()

        try:
            assert await test_suite.setup()

            success, response_data, response_time = await test_suite.client.call_tool(
                "extract_table_data",
                {
                    "file_path": str(self.test_file),
                    "slide_numbers": [1, 2, 3, 4, 5, 6, 7, 8]
                }
            )

            # This tool might not find tables in our test file, but should not error
            assert success, f"extract_table_data failed: {response_data}"

        finally:
            await test_suite.teardown()

    # Removed tests for deleted tools: get_presentation_overview, analyze_text_formatting

    @pytest.mark.asyncio
    async def test_error_conditions(self):
        """Test error handling for invalid parameters."""
        test_suite = MCPIntegrationTestSuite()

        try:
            assert await test_suite.setup()

            # Test invalid formatting type
            success, response_data, response_time = await test_suite.client.call_tool(
                "extract_formatted_text",
                {
                    "file_path": str(self.test_file),
                    "formatting_type": "invalid_type"
                }
            )

            # Should return error message, not crash
            assert "error" in response_data or "Invalid formatting_type" in str(response_data)

            # Test nonexistent file
            success, response_data, response_time = await test_suite.client.call_tool(
                "extract_formatted_text",
                {
                    "file_path": "nonexistent_file.pptx",
                    "formatting_type": "bold"
                }
            )

            # Should return error
            assert not success or "error" in response_data

        finally:
            await test_suite.teardown()

    @pytest.mark.asyncio
    async def test_performance_requirements(self):
        """Test that tools meet performance requirements."""
        test_suite = MCPIntegrationTestSuite()

        try:
            assert await test_suite.setup()

            # Test formatting extraction performance
            success, response_data, response_time = await test_suite.client.call_tool(
                "extract_formatted_text",
                {
                    "file_path": str(self.test_file),
                    "formatting_type": "bold"
                }
            )

            assert success
            assert response_time < 5.0, f"Formatting extraction took too long: {response_time}s"

        finally:
            await test_suite.teardown()

    def test_expected_results_validation(self):
        """Validate test results against expected outcomes."""
        # Load expected results
        expected_file = self.test_files_dir / "test_formatting_comprehensive.json"
        if not expected_file.exists():
            pytest.skip("Expected results file not found")

        with open(expected_file, 'r', encoding='utf-8') as f:
            expected_results = json.load(f)

        # Verify expected results structure
        assert "slide_1_bold" in expected_results
        assert "slide_2_italic" in expected_results
        assert "slide_3_underlined" in expected_results

        # Verify bold segments structure
        bold_segments = expected_results["slide_1_bold"]["bold_segments"]
        assert len(bold_segments) > 0

        for segment in bold_segments:
            assert "text" in segment
            assert "start_position" in segment
            assert isinstance(segment["start_position"], int)

    @pytest.mark.asyncio
    async def test_example_usage_query_slides(self):
        """Test all Example Usage cases for query_slides tool."""
        test_suite = MCPIntegrationTestSuite()

        try:
            assert await test_suite.setup()

            # Example 1: Find slides with "Sales" in the title
            # Note: Using actual file, so searching for common words like "Test" or "Slide"
            success, response_data, response_time = await test_suite.client.call_tool(
                "query_slides",
                {
                    "file_path": str(self.test_file),
                    "search_criteria": {"title": {"contains": "Test"}}
                }
            )
            assert success, f"Example 1 failed: {response_data}"
            logger.info(f"Example 1 - Title contains 'Test': {len(response_data.get('results', []))} results")

            # Example 2: Find slides with tables and specific layout
            success, response_data, response_time = await test_suite.client.call_tool(
                "query_slides",
                {
                    "file_path": str(self.test_file),
                    "search_criteria": {
                        "content": {"has_tables": True},
                        "layout": {"type": "title_content"}
                    }
                }
            )
            assert success, f"Example 2 failed: {response_data}"
            logger.info(f"Example 2 - Tables + layout: {len(response_data.get('results', []))} results")

            # Example 3: Find specific slides with custom return fields
            success, response_data, response_time = await test_suite.client.call_tool(
                "query_slides",
                {
                    "file_path": str(self.test_file),
                    "search_criteria": {"slide_numbers": [1, 3, 5]},
                    "return_fields": ["slide_number", "title", "preview_text"]
                }
            )
            assert success, f"Example 3 failed: {response_data}"
            results = response_data.get('results', [])
            logger.info(f"Example 3 - Specific slides: {len(results)} results")
            # Verify return fields are present
            if results:
                for result in results:
                    assert "slide_number" in result
                    assert "title" in result
                    assert "preview_text" in result

            # Example 4: Complex search with multiple criteria
            success, response_data, response_time = await test_suite.client.call_tool(
                "query_slides",
                {
                    "file_path": str(self.test_file),
                    "search_criteria": {
                        "title": {"regex": r".*[Tt]est.*"},  # More flexible regex
                        "content": {"has_tables": True, "has_images": False}
                    },
                    "limit": 10
                }
            )
            assert success, f"Example 4 failed: {response_data}"
            logger.info(f"Example 4 - Complex search: {len(response_data.get('results', []))} results")

        finally:
            await test_suite.teardown()

    @pytest.mark.asyncio
    async def test_example_usage_extract_table_data(self):
        """Test all Example Usage cases for extract_table_data tool."""
        test_suite = MCPIntegrationTestSuite()

        try:
            assert await test_suite.setup()

            # Example 1: Basic table extraction from all slides
            success, response_data, response_time = await test_suite.client.call_tool(
                "extract_table_data",
                {
                    "file_path": str(self.test_file)
                }
            )
            assert success, f"Example 1 failed: {response_data}"
            logger.info(f"Example 1 - Basic extraction: {response_data.get('summary', {}).get('total_tables_found', 0)} tables found")

            # Example 2: Extract tables from specific slides
            success, response_data, response_time = await test_suite.client.call_tool(
                "extract_table_data",
                {
                    "file_path": str(self.test_file),
                    "slide_numbers": [1, 2]
                }
            )
            assert success, f"Example 2 failed: {response_data}"
            logger.info(f"Example 2 - Specific slides: {response_data.get('summary', {}).get('total_tables_found', 0)} tables found")

            # Example 3: Extract tables with specific criteria
            success, response_data, response_time = await test_suite.client.call_tool(
                "extract_table_data",
                {
                    "file_path": str(self.test_file),
                    "slide_numbers": [1, 2],
                    "table_criteria": {"min_rows": 2, "header_contains": ["Name"]}
                }
            )
            assert success, f"Example 3 failed: {response_data}"
            logger.info(f"Example 3 - With criteria: {response_data.get('summary', {}).get('total_tables_found', 0)} tables found")

            # Example 4: Extract specific columns with formatting from all slides
            success, response_data, response_time = await test_suite.client.call_tool(
                "extract_table_data",
                {
                    "file_path": str(self.test_file),
                    "column_selection": {"specific_columns": ["Name", "Age"]},
                    "formatting_detection": {"detect_bold": True, "detect_colors": True}
                }
            )
            assert success, f"Example 4 failed: {response_data}"
            logger.info(f"Example 4 - Column selection + formatting: {response_data.get('summary', {}).get('total_tables_found', 0)} tables found")

        finally:
            await test_suite.teardown()

    @pytest.mark.asyncio
    async def test_example_usage_extract_formatted_text(self):
        """Test all Example Usage cases for extract_formatted_text tool."""
        test_suite = MCPIntegrationTestSuite()

        try:
            assert await test_suite.setup()

            # Example 1: Returns all bold text from all slides
            success, response_data, response_time = await test_suite.client.call_tool(
                "extract_formatted_text",
                {
                    "file_path": str(self.test_file),
                    "formatting_type": "bold"
                }
            )
            assert success, f"Example 1 failed: {response_data}"
            summary = response_data.get('summary', {})
            logger.info(f"Example 1 - Bold text: {summary.get('total_formatted_segments', 0)} segments found")

            # Example 2: Returns hyperlinks from slides 1 and 2 only
            success, response_data, response_time = await test_suite.client.call_tool(
                "extract_formatted_text",
                {
                    "file_path": str(self.test_file),
                    "formatting_type": "hyperlinks",
                    "slide_numbers": [1, 2]
                }
            )
            assert success, f"Example 2 failed: {response_data}"
            summary = response_data.get('summary', {})
            logger.info(f"Example 2 - Hyperlinks from slides 1,2: {summary.get('total_formatted_segments', 0)} segments found")

            # Additional test: Test other formatting types mentioned in the documentation
            formatting_types = ["italic", "underlined", "highlighted", "font_sizes", "font_colors"]
            for formatting_type in formatting_types:
                success, response_data, response_time = await test_suite.client.call_tool(
                    "extract_formatted_text",
                    {
                        "file_path": str(self.test_file),
                        "formatting_type": formatting_type
                    }
                )
                assert success, f"Additional test for {formatting_type} failed: {response_data}"
                summary = response_data.get('summary', {})
                logger.info(f"Additional test - {formatting_type}: {summary.get('total_formatted_segments', 0)} segments found")

        finally:
            await test_suite.teardown()


class TestMCPToolCoverage:
    """Test coverage analysis for MCP tools."""

    @pytest.mark.asyncio
    async def test_all_tools_coverage(self):
        """Verify all tools are tested and achieve good coverage."""
        test_suite = MCPIntegrationTestSuite()

        try:
            assert await test_suite.setup()

            # Get all available tools
            available_tools = await test_suite.client.list_available_tools()

            # Expected tools based on our implementation
            expected_tools = [
                "extract_formatted_text",
                "query_slides",
                "extract_table_data"
            ]

            # Verify all expected tools are available
            for tool in expected_tools:
                assert tool in available_tools, f"Expected tool {tool} not found in available tools: {available_tools}"

            # Run comprehensive tests
            await test_suite.test_all_tools()
            await test_suite.test_error_conditions()

            # Generate coverage report
            report = test_suite.generate_test_report()

            # Verify coverage requirements
            assert report["summary"]["success_rate"] >= 80.0, f"Success rate too low: {report['summary']['success_rate']}%"
            assert report["summary"]["total_tests"] >= len(expected_tools), "Not enough tests run"

            # Verify each tool was tested
            for tool in expected_tools:
                assert tool in report["tools_coverage"], f"Tool {tool} not covered in tests"
                assert report["tools_coverage"][tool]["total"] > 0, f"No tests run for tool {tool}"

        finally:
            await test_suite.teardown()


def run_comprehensive_test_suite():
    """Run the comprehensive test suite and generate report."""
    # Run pytest with coverage
    pytest.main([
        __file__,
        "-v",
        "--tb=short",
        "--durations=10"
    ])


if __name__ == "__main__":
    run_comprehensive_test_suite()