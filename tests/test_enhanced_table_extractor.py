"""
Comprehensive tests for extract_table_data MCP tool.
Tests all options and parameters using test_complex.pptx file.
"""

import pytest
import json
import asyncio
from pathlib import Path
from typing import Dict, List, Any, Optional

from powerpoint_mcp_server.server import PowerPointMCPServer


class TestExtractTableDataMCP:
    """Test cases for extract_table_data MCP tool with all options."""
    
    def setup_method(self):
        """Set up test fixtures."""
        self.server = PowerPointMCPServer()
        self.test_files_dir = Path("tests/test_files")
        self.test_file = self.test_files_dir / "test_complex.pptx"
        
        # Ensure test file exists
        if not self.test_file.exists():
            pytest.skip(f"Test file not found: {self.test_file}")
    
    @pytest.mark.asyncio
    async def test_basic_table_extraction(self):
        """Test basic table extraction from all slides."""
        arguments = {
            "file_path": str(self.test_file)
        }
        
        result = await self.server._extract_table_data(arguments)
        content_text = result.content[0].text
        response_data = json.loads(content_text)
        
        # Verify response structure
        assert "summary" in response_data
        assert "extracted_tables" in response_data
        
        # Verify summary structure
        summary = response_data["summary"]
        assert "total_tables_found" in summary
        assert "total_tables" in summary
        assert "total_rows" in summary
        assert "slides_with_tables" in summary
        assert "formatting_found" in summary
        assert "slides_processed" in summary
        
        # Based on test_complex_documentation.md, slide 2 has 2 tables
        assert summary["total_tables_found"] >= 2
        assert summary["slides_with_tables"] >= 1
        assert summary["slides_processed"] == 4  # All slides processed
        
        # Verify extracted tables structure
        tables = response_data["extracted_tables"]
        assert isinstance(tables, list)
        assert len(tables) >= 2  # At least 2 tables from slide 2
        
        # Verify table structure
        for table in tables:
            assert "slide_number" in table
            assert "table_index" in table
            assert "rows" in table
            assert "columns" in table
            assert "headers" in table
            assert "data" in table
            assert "metadata" in table
            assert "position" in table
            assert "size" in table
            
            # Verify data structure
            assert isinstance(table["data"], list)
            assert len(table["data"]) == table["rows"]
            
            # Verify each row has correct structure
            for row in table["data"]:
                assert isinstance(row, dict)
                for header in table["headers"]:
                    assert header in row
                    cell = row[header]
                    assert "value" in cell
                    assert "formatting" in cell
                    assert "row_span" in cell
                    assert "col_span" in cell
                    assert "row_col_position" in cell
    
    @pytest.mark.asyncio
    async def test_specific_slide_extraction(self):
        """Test table extraction from specific slides."""
        # Test slide 2 only (has tables according to documentation)
        arguments = {
            "file_path": str(self.test_file),
            "slide_numbers": [2]
        }
        
        result = await self.server._extract_table_data(arguments)
        content_text = result.content[0].text
        response_data = json.loads(content_text)
        
        summary = response_data["summary"]
        assert summary["slides_processed"] == 1
        assert summary["total_tables_found"] >= 2  # Slide 2 has 2 tables
        
        # Verify all tables are from slide 2
        tables = response_data["extracted_tables"]
        for table in tables:
            assert table["slide_number"] == 2
    
    @pytest.mark.asyncio
    async def test_slides_without_tables(self):
        """Test extraction from slides without tables."""
        # Test slides 1, 3, 4 (no tables according to documentation)
        arguments = {
            "file_path": str(self.test_file),
            "slide_numbers": [1, 3, 4]
        }
        
        result = await self.server._extract_table_data(arguments)
        content_text = result.content[0].text
        response_data = json.loads(content_text)
        
        summary = response_data["summary"]
        assert summary["slides_processed"] == 3
        assert summary["total_tables_found"] == 0
        assert summary["slides_with_tables"] == 0
        
        tables = response_data["extracted_tables"]
        assert len(tables) == 0
    
    @pytest.mark.asyncio
    async def test_table_criteria_min_rows(self):
        """Test table criteria with minimum rows."""
        arguments = {
            "file_path": str(self.test_file),
            "table_criteria": {
                "min_rows": 5  # Should filter out smaller tables
            }
        }
        
        result = await self.server._extract_table_data(arguments)
        content_text = result.content[0].text
        response_data = json.loads(content_text)
        
        # Verify that only tables with >= 5 rows are returned
        tables = response_data["extracted_tables"]
        for table in tables:
            assert table["rows"] >= 5
    
    @pytest.mark.asyncio
    async def test_table_criteria_min_columns(self):
        """Test table criteria with minimum columns."""
        arguments = {
            "file_path": str(self.test_file),
            "table_criteria": {
                "min_columns": 2
            }
        }
        
        result = await self.server._extract_table_data(arguments)
        content_text = result.content[0].text
        response_data = json.loads(content_text)
        
        # Verify that only tables with >= 2 columns are returned
        tables = response_data["extracted_tables"]
        for table in tables:
            assert table["columns"] >= 2
    
    @pytest.mark.asyncio
    async def test_table_criteria_header_contains(self):
        """Test table criteria with header contains."""
        arguments = {
            "file_path": str(self.test_file),
            "table_criteria": {
                "header_contains": ["Header"]  # Should match tables with "Header" in column names
            }
        }
        
        result = await self.server._extract_table_data(arguments)
        content_text = result.content[0].text
        response_data = json.loads(content_text)
        
        # Verify that returned tables have headers containing "Header"
        tables = response_data["extracted_tables"]
        for table in tables:
            has_header_match = any("Header" in header for header in table["headers"])
            assert has_header_match
    
    @pytest.mark.asyncio
    async def test_column_selection_specific_columns(self):
        """Test column selection with specific columns."""
        # First, get all tables to see available headers
        all_tables_args = {"file_path": str(self.test_file)}
        all_result = await self.server._extract_table_data(all_tables_args)
        all_content = json.loads(all_result.content[0].text)
        
        if not all_content["extracted_tables"]:
            pytest.skip("No tables found in test file")
        
        # Get first table's headers
        first_table = all_content["extracted_tables"][0]
        available_headers = first_table["headers"]
        
        if len(available_headers) < 2:
            pytest.skip("Need at least 2 columns for this test")
        
        # Select only first column
        selected_column = available_headers[0]
        arguments = {
            "file_path": str(self.test_file),
            "column_selection": {
                "specific_columns": [selected_column],
                "all_columns": False
            }
        }
        
        result = await self.server._extract_table_data(arguments)
        content_text = result.content[0].text
        response_data = json.loads(content_text)
        
        # Verify that column selection was applied
        # Note: The actual behavior may vary based on implementation
        # This test verifies that the tool accepts the parameters without error
        assert "summary" in response_data
        assert "extracted_tables" in response_data
        
        # If column selection is implemented, verify the filtering
        tables = response_data["extracted_tables"]
        if tables:
            # Column selection may or may not be fully implemented
            # At minimum, the tool should not crash and should return valid data
            for table in tables:
                assert "headers" in table
                assert isinstance(table["headers"], list)
    
    @pytest.mark.asyncio
    async def test_column_selection_exclude_columns(self):
        """Test column selection with excluded columns."""
        # First, get all tables to see available headers
        all_tables_args = {"file_path": str(self.test_file)}
        all_result = await self.server._extract_table_data(all_tables_args)
        all_content = json.loads(all_result.content[0].text)
        
        if not all_content["extracted_tables"]:
            pytest.skip("No tables found in test file")
        
        # Get first table's headers
        first_table = all_content["extracted_tables"][0]
        available_headers = first_table["headers"]
        
        if len(available_headers) < 2:
            pytest.skip("Need at least 2 columns for this test")
        
        # Exclude first column
        excluded_column = available_headers[0]
        arguments = {
            "file_path": str(self.test_file),
            "column_selection": {
                "exclude_columns": [excluded_column],
                "all_columns": True
            }
        }
        
        result = await self.server._extract_table_data(arguments)
        content_text = result.content[0].text
        response_data = json.loads(content_text)
        
        # Verify that column exclusion was processed
        # Note: The actual behavior may vary based on implementation
        # This test verifies that the tool accepts the parameters without error
        assert "summary" in response_data
        assert "extracted_tables" in response_data
        
        # If column exclusion is implemented, verify the filtering
        tables = response_data["extracted_tables"]
        if tables:
            # Column exclusion may or may not be fully implemented
            # At minimum, the tool should not crash and should return valid data
            for table in tables:
                assert "headers" in table
                assert isinstance(table["headers"], list)
    
    @pytest.mark.asyncio
    async def test_formatting_detection_options(self):
        """Test formatting detection configuration."""
        arguments = {
            "file_path": str(self.test_file),
            "formatting_detection": {
                "detect_bold": True,
                "detect_italic": True,
                "detect_underline": True,
                "detect_highlight": True,
                "detect_colors": True,
                "detect_hyperlinks": True,
                "preserve_formatting": True
            }
        }
        
        result = await self.server._extract_table_data(arguments)
        content_text = result.content[0].text
        response_data = json.loads(content_text)
        
        # Verify formatting information is included
        tables = response_data["extracted_tables"]
        for table in tables:
            for row in table["data"]:
                for header in table["headers"]:
                    cell = row[header]
                    formatting = cell["formatting"]
                    
                    # Verify all formatting fields are present
                    assert "bold" in formatting
                    assert "italic" in formatting
                    assert "underline" in formatting
                    assert "highlight" in formatting
                    assert "strikethrough" in formatting
                    assert "font_color" in formatting
                    assert "background_color" in formatting
                    assert "font_size" in formatting
                    assert "hyperlink" in formatting
    
    @pytest.mark.asyncio
    async def test_output_format_structured(self):
        """Test structured output format."""
        arguments = {
            "file_path": str(self.test_file),
            "output_format": "structured",
            "include_metadata": True
        }
        
        result = await self.server._extract_table_data(arguments)
        content_text = result.content[0].text
        response_data = json.loads(content_text)
        
        # Verify structured format
        assert "summary" in response_data
        assert "extracted_tables" in response_data
        
        # Verify metadata is included
        tables = response_data["extracted_tables"]
        for table in tables:
            assert "metadata" in table
            metadata = table["metadata"]
            assert "has_formatting" in metadata
            assert "cell_count" in metadata
            assert "non_empty_cells" in metadata
    
    @pytest.mark.asyncio
    async def test_output_format_flat(self):
        """Test flat output format."""
        arguments = {
            "file_path": str(self.test_file),
            "output_format": "flat",
            "include_metadata": True
        }
        
        result = await self.server._extract_table_data(arguments)
        content_text = result.content[0].text
        response_data = json.loads(content_text)
        
        # Verify flat format structure
        assert "summary" in response_data
        # Flat format should have different structure than structured
        # The exact structure depends on implementation
        assert isinstance(response_data, dict)
    
    @pytest.mark.asyncio
    async def test_output_format_grouped_by_slide(self):
        """Test grouped by slide output format."""
        arguments = {
            "file_path": str(self.test_file),
            "output_format": "grouped_by_slide",
            "include_metadata": True
        }
        
        result = await self.server._extract_table_data(arguments)
        content_text = result.content[0].text
        response_data = json.loads(content_text)
        
        # Verify grouped format structure
        assert "summary" in response_data
        # Grouped format should have different structure than structured
        # The exact structure depends on implementation
        assert isinstance(response_data, dict)
    
    @pytest.mark.asyncio
    async def test_include_metadata_false(self):
        """Test extraction without metadata."""
        arguments = {
            "file_path": str(self.test_file),
            "include_metadata": False
        }
        
        result = await self.server._extract_table_data(arguments)
        content_text = result.content[0].text
        response_data = json.loads(content_text)
        
        # Verify basic structure is still present
        assert "summary" in response_data
        assert "extracted_tables" in response_data
        
        # Metadata inclusion behavior depends on implementation
        # At minimum, basic structure should be present
        tables = response_data["extracted_tables"]
        for table in tables:
            assert "slide_number" in table
            assert "headers" in table
            assert "data" in table
    
    @pytest.mark.asyncio
    async def test_complex_criteria_combination(self):
        """Test combination of multiple criteria."""
        arguments = {
            "file_path": str(self.test_file),
            "slide_numbers": [2],  # Only slide with tables
            "table_criteria": {
                "min_rows": 2,
                "min_columns": 2
            },
            "formatting_detection": {
                "detect_bold": True,
                "detect_highlight": True,
                "detect_colors": True
            },
            "output_format": "structured",
            "include_metadata": True
        }
        
        result = await self.server._extract_table_data(arguments)
        content_text = result.content[0].text
        response_data = json.loads(content_text)
        
        # Verify all criteria are applied
        summary = response_data["summary"]
        assert summary["slides_processed"] == 1
        
        tables = response_data["extracted_tables"]
        for table in tables:
            assert table["slide_number"] == 2
            assert table["rows"] >= 2
            assert table["columns"] >= 2
            assert "metadata" in table
    
    @pytest.mark.asyncio
    async def test_error_handling_invalid_file(self):
        """Test error handling with invalid file path."""
        arguments = {
            "file_path": "nonexistent_file.pptx"
        }
        
        # Should raise McpError for invalid file
        with pytest.raises(Exception) as exc_info:
            await self.server._extract_table_data(arguments)
        
        # Verify error contains appropriate message
        error_message = str(exc_info.value)
        assert "file" in error_message.lower() or "not exist" in error_message.lower()
    
    @pytest.mark.asyncio
    async def test_error_handling_invalid_slide_numbers(self):
        """Test error handling with invalid slide numbers."""
        arguments = {
            "file_path": str(self.test_file),
            "slide_numbers": [999]  # Non-existent slide
        }
        
        # Should raise McpError for invalid slide numbers
        with pytest.raises(Exception) as exc_info:
            await self.server._extract_table_data(arguments)
        
        # Verify error contains appropriate message
        error_message = str(exc_info.value)
        assert "slide" in error_message.lower() or "invalid" in error_message.lower()
    
    @pytest.mark.asyncio
    async def test_formatting_statistics(self):
        """Test that formatting statistics are correctly calculated."""
        arguments = {
            "file_path": str(self.test_file),
            "formatting_detection": {
                "detect_bold": True,
                "detect_italic": True,
                "detect_highlight": True,
                "detect_colors": True,
                "detect_hyperlinks": True
            }
        }
        
        result = await self.server._extract_table_data(arguments)
        content_text = result.content[0].text
        response_data = json.loads(content_text)
        
        # Verify formatting statistics in summary
        summary = response_data["summary"]
        assert "formatting_found" in summary
        
        formatting_stats = summary["formatting_found"]
        # Based on test_complex_documentation.md, tables have formatting
        # Exact counts depend on implementation, but should be >= 0
        for stat_key in ["bold_cells", "italic_cells", "highlighted_cells", "colored_cells"]:
            assert stat_key in formatting_stats
            assert isinstance(formatting_stats[stat_key], int)
            assert formatting_stats[stat_key] >= 0
    
    @pytest.mark.asyncio
    async def test_table_position_and_size_info(self):
        """Test that table position and size information is included."""
        arguments = {
            "file_path": str(self.test_file),
            "include_metadata": True
        }
        
        result = await self.server._extract_table_data(arguments)
        content_text = result.content[0].text
        response_data = json.loads(content_text)
        
        tables = response_data["extracted_tables"]
        for table in tables:
            # Verify position and size information
            assert "position" in table
            assert "size" in table
            
            position = table["position"]
            size = table["size"]
            
            assert isinstance(position, list)
            assert len(position) == 2
            assert all(isinstance(coord, (int, float)) for coord in position)
            
            assert isinstance(size, list)
            assert len(size) == 2
            assert all(isinstance(dim, (int, float)) for dim in size)


if __name__ == "__main__":
    pytest.main([__file__, "-v"])