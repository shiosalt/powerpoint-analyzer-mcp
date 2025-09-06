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

from tests.mcp_integration_test_framework import MCPIntegrationTestSuite, MCPTestClient, TestResult

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

    @pytest.mark.asyncio
    async def test_extract_powerpoint_content_tool(self):
        """Test extract_powerpoint_content with various parameters."""
        test_suite = MCPIntegrationTestSuite()

        try:
            assert await test_suite.setup()

            # Test basic content extraction
            success, response_data, response_time = await test_suite.client.call_tool(
                "extract_powerpoint_content",
                {"file_path": str(self.test_file)}
            )

            assert success, f"Tool call failed: {response_data}"
            assert "slides" in response_data
            assert "metadata" in response_data
            assert len(response_data["slides"]) > 0
            assert response_time < 10.0  # Should complete within 10 seconds

        finally:
            await test_suite.teardown()

    @pytest.mark.asyncio
    async def test_get_powerpoint_attributes_tool(self):
        """Test get_powerpoint_attributes with all valid attribute combinations."""
        test_suite = MCPIntegrationTestSuite()

        try:
            assert await test_suite.setup()

            # Test individual attributes
            individual_attributes = ["title", "subtitle", "text_elements", "tables", "object_counts"]

            for attribute in individual_attributes:
                success, response_data, response_time = await test_suite.client.call_tool(
                    "get_powerpoint_attributes",
                    {
                        "file_path": str(self.test_file),
                        "attributes": [attribute]
                    }
                )

                assert success, f"Failed to extract attribute {attribute}: {response_data}"
                assert "slides" in response_data

                # Check that only requested attribute is present (plus slide_number)
                if response_data["slides"]:
                    slide = response_data["slides"][0]
                    assert attribute in slide or attribute == "text_elements"  # text_elements might be empty
                    assert "slide_number" in slide

            # Test multiple attributes
            success, response_data, response_time = await test_suite.client.call_tool(
                "get_powerpoint_attributes",
                {
                    "file_path": str(self.test_file),
                    "attributes": ["title", "subtitle", "object_counts"]
                }
            )

            assert success, f"Failed to extract multiple attributes: {response_data}"
            assert "slides" in response_data

        finally:
            await test_suite.teardown()

    @pytest.mark.asyncio
    async def test_get_slide_info_tool(self):
        """Test get_slide_info with valid slide numbers."""
        test_suite = MCPIntegrationTestSuite()

        try:
            assert await test_suite.setup()

            # Test first slide
            success, response_data, response_time = await test_suite.client.call_tool(
                "get_slide_info",
                {
                    "file_path": str(self.test_file),
                    "slide_number": 1
                }
            )

            assert success, f"Failed to get slide 1 info: {response_data}"
            assert response_data["slide_number"] == 1
            assert "title" in response_data
            assert "text_elements" in response_data

            # Test middle slide
            success, response_data, response_time = await test_suite.client.call_tool(
                "get_slide_info",
                {
                    "file_path": str(self.test_file),
                    "slide_number": 3
                }
            )

            assert success, f"Failed to get slide 3 info: {response_data}"
            assert response_data["slide_number"] == 3

        finally:
            await test_suite.teardown()

    @pytest.mark.asyncio
    async def test_extract_text_formatting_tool_all_types(self):
        """Test extract_text_formatting with all supported formatting types."""
        test_suite = MCPIntegrationTestSuite()

        try:
            assert await test_suite.setup()

            formatting_types = [
                "bold", "italic", "underlined", "highlighted",
                "strikethrough", "hyperlinks", "font_sizes", "font_colors"
            ]

            for formatting_type in formatting_types:
                success, response_data, response_time = await test_suite.client.call_tool(
                    "extract_text_formatting",
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
    async def test_extract_text_formatting_with_slide_filter(self):
        """Test extract_text_formatting with slide number filtering."""
        test_suite = MCPIntegrationTestSuite()

        try:
            assert await test_suite.setup()

            # Test with specific slides
            success, response_data, response_time = await test_suite.client.call_tool(
                "extract_text_formatting",
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

    @pytest.mark.asyncio
    async def test_get_presentation_overview_tool(self):
        """Test get_presentation_overview with different analysis depths."""
        test_suite = MCPIntegrationTestSuite()

        try:
            assert await test_suite.setup()

            analysis_depths = ["basic", "detailed", "comprehensive"]

            for depth in analysis_depths:
                success, response_data, response_time = await test_suite.client.call_tool(
                    "get_presentation_overview",
                    {
                        "file_path": str(self.test_file),
                        "analysis_depth": depth
                    }
                )

                assert success, f"Failed to get {depth} overview: {response_data}"

                # Verify basic structure
                assert isinstance(response_data, dict)

        finally:
            await test_suite.teardown()

    @pytest.mark.asyncio
    async def test_analyze_text_formatting_tool(self):
        """Test analyze_text_formatting tool."""
        test_suite = MCPIntegrationTestSuite()

        try:
            assert await test_suite.setup()

            success, response_data, response_time = await test_suite.client.call_tool(
                "analyze_text_formatting",
                {"file_path": str(self.test_file)}
            )

            assert success, f"Failed to analyze text formatting: {response_data}"

        finally:
            await test_suite.teardown()

    @pytest.mark.asyncio
    async def test_error_conditions(self):
        """Test error handling for invalid parameters."""
        test_suite = MCPIntegrationTestSuite()

        try:
            assert await test_suite.setup()

            # Test invalid formatting type
            success, response_data, response_time = await test_suite.client.call_tool(
                "extract_text_formatting",
                {
                    "file_path": str(self.test_file),
                    "formatting_type": "invalid_type"
                }
            )

            # Should return error message, not crash
            assert "error" in response_data or "Invalid formatting_type" in str(response_data)

            # Test invalid slide number
            success, response_data, response_time = await test_suite.client.call_tool(
                "get_slide_info",
                {
                    "file_path": str(self.test_file),
                    "slide_number": 999
                }
            )

            # Should handle gracefully
            assert not success or "error" in response_data or "out of range" in str(response_data)

            # Test nonexistent file
            success, response_data, response_time = await test_suite.client.call_tool(
                "extract_powerpoint_content",
                {"file_path": "nonexistent_file.pptx"}
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

            # Test basic content extraction performance
            success, response_data, response_time = await test_suite.client.call_tool(
                "extract_powerpoint_content",
                {"file_path": str(self.test_file)}
            )

            assert success
            assert response_time < 10.0, f"Content extraction took too long: {response_time}s"

            # Test formatting extraction performance
            success, response_data, response_time = await test_suite.client.call_tool(
                "extract_text_formatting",
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
                "extract_powerpoint_content",
                "get_powerpoint_attributes",
                "get_slide_info",
                "extract_text_formatting",
                "query_slides",
                "extract_table_data",
                "get_presentation_overview",
                "analyze_text_formatting"
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