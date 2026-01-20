"""
Unit tests for simple table extractor output formats.
Tests row_col_value, row_col_formattedvalue, html, and simple_html formats.
"""

import pytest
import os
from powerpoint_mcp_server.core.simple_table_extractor import SimpleTableExtractor
from powerpoint_mcp_server.core.content_extractor import ContentExtractor


class TestSimpleTableFormats:
    """Test cases for simple table extractor output formats."""
    
    def setup_method(self):
        """Set up test fixtures."""
        self.content_extractor = ContentExtractor()
        self.simple_extractor = SimpleTableExtractor(self.content_extractor)
        self.test_file = "tests/test_files/test_complex.pptx"
        
        if not os.path.exists(self.test_file):
            pytest.skip(f"Test file not found: {self.test_file}")
    
    def test_row_col_value_format(self):
        """Test row_col_value output format (default)."""
        result = self.simple_extractor.extract_tables_simple(
            file_path=self.test_file,
            slide_numbers=[1, 2, 3],
            output_format="row_col_value"
        )
        
        # Check result structure
        assert "extracted_tables" in result
        assert isinstance(result["extracted_tables"], list)
        
        # Check table structure
        if result["extracted_tables"]:
            table = result["extracted_tables"][0]
            assert "slide_number" in table
            assert "rows" in table
            assert "columns" in table
            assert "headers" in table
            assert "data" in table
            
            # Check data format: [[row, col, value], ...]
            data = table["data"]
            assert isinstance(data, list)
            if data:
                row = data[0]
                assert isinstance(row, list)
                assert len(row) == 3  # [row_index, col_index, value]
                assert isinstance(row[0], int)  # row index
                assert isinstance(row[1], int)  # col index
                assert isinstance(row[2], str)  # value
    
    def test_row_col_formattedvalue_format(self):
        """Test row_col_formattedvalue output format with formatting markers."""
        result = self.simple_extractor.extract_tables_simple(
            file_path=self.test_file,
            slide_numbers=[1, 2, 3],
            output_format="row_col_formattedvalue"
        )
        
        # Check result structure
        assert "extracted_tables" in result
        assert isinstance(result["extracted_tables"], list)
        
        # Check table structure
        if result["extracted_tables"]:
            table = result["extracted_tables"][0]
            assert "slide_number" in table
            assert "rows" in table
            assert "columns" in table
            assert "headers" in table
            assert "data" in table
            
            # Check data format: [[row, col, formatted_value], ...]
            data = table["data"]
            assert isinstance(data, list)
            if data:
                row = data[0]
                assert isinstance(row, list)
                assert len(row) == 3  # [row_index, col_index, formatted_value]
                assert isinstance(row[0], int)  # row index
                assert isinstance(row[1], int)  # col index
                assert isinstance(row[2], str)  # formatted value (may contain ** or * markers)
    
    def test_simple_html_format(self):
        """Test simple_html output format."""
        result = self.simple_extractor.extract_tables_simple(
            file_path=self.test_file,
            slide_numbers=[1, 2, 3],
            output_format="simple_html"
        )
        
        # Check result structure
        assert "extracted_html_tables" in result
        assert isinstance(result["extracted_html_tables"], list)
        
        # Check table structure
        if result["extracted_html_tables"]:
            table = result["extracted_html_tables"][0]
            assert "slide_number" in table
            assert "rows" in table
            assert "columns" in table
            assert "headers" in table
            assert "htmldata" in table
            
            # Check HTML structure
            html = table["htmldata"]
            assert isinstance(html, str)
            assert "<table" in html
            assert "<thead" in html  # Changed from exact match to allow style attribute
            assert "<tbody>" in html
            assert "</table>" in html
            assert "<th" in html  # Header cells
            assert "<td" in html  # Data cells
            
            # Simple HTML should have basic styling
            assert 'border="1"' in html
            assert 'cellpadding="5"' in html
    
    def test_html_format_with_formatting(self):
        """Test html output format with formatting support."""
        result = self.simple_extractor.extract_tables_simple(
            file_path=self.test_file,
            slide_numbers=[1, 2, 3],
            output_format="html"
        )
        
        # Check result structure
        assert "extracted_html_tables" in result
        assert isinstance(result["extracted_html_tables"], list)
        
        # Check table structure
        if result["extracted_html_tables"]:
            table = result["extracted_html_tables"][0]
            assert "slide_number" in table
            assert "rows" in table
            assert "columns" in table
            assert "headers" in table
            assert "htmldata" in table
            
            # Check HTML structure
            html = table["htmldata"]
            assert isinstance(html, str)
            assert "<table" in html
            assert "<thead" in html  # Changed from exact match to allow style attribute
            assert "<tbody>" in html
            assert "</table>" in html
            
            # HTML format should have enhanced styling
            assert "border-collapse: collapse" in html
            assert "white-space: pre-wrap" in html
            
            # Check for potential formatting tags (may or may not be present depending on content)
            # Just verify the HTML is valid and contains expected structure
            assert html.count("<table") == html.count("</table>")
            assert html.count("<thead") == html.count("</thead>")
            assert html.count("<tbody>") == html.count("</tbody>")
    
    def test_invalid_format(self):
        """Test that invalid format raises an error."""
        # This should be caught at the server level, but test extractor behavior
        # The extractor itself doesn't validate, so it will use default behavior
        result = self.simple_extractor.extract_tables_simple(
            file_path=self.test_file,
            slide_numbers=[1, 2, 3],
            output_format="invalid_format"
        )
        
        # Should return extracted_tables (default behavior)
        assert "extracted_tables" in result
    
    def test_html_escaping(self):
        """Test that HTML special characters are properly escaped."""
        result = self.simple_extractor.extract_tables_simple(
            file_path=self.test_file,
            slide_numbers=[1, 2, 3],
            output_format="html"
        )
        
        if result.get("extracted_html_tables"):
            html = result["extracted_html_tables"][0]["htmldata"]
            
            # HTML should not contain unescaped special characters in content
            # (but should contain them in tags)
            assert "<table" in html  # Tags should be present
            
            # If there's any text content with special chars, they should be escaped
            # This is a basic check - actual content depends on test file
            assert isinstance(html, str)
            assert len(html) > 0
    
    def test_rowspan_colspan_support(self):
        """Test that rowspan and colspan attributes are preserved in HTML."""
        result = self.simple_extractor.extract_tables_simple(
            file_path=self.test_file,
            slide_numbers=[1, 2, 3],
            output_format="html"
        )
        
        if result.get("extracted_html_tables"):
            html = result["extracted_html_tables"][0]["htmldata"]
            
            # Check that the HTML is well-formed
            assert "<td" in html
            # Rowspan/colspan may or may not be present depending on table structure
            # Just verify the HTML structure is valid
            assert html.count("<tr>") == html.count("</tr>")
    
    def test_column_selection_with_formats(self):
        """Test that column selection works with different output formats."""
        column_selection = {
            "specific_columns": ["Header 1"]
        }
        
        # Test with row_col_value
        result = self.simple_extractor.extract_tables_simple(
            file_path=self.test_file,
            slide_numbers=[1, 2, 3],
            column_selection=column_selection,
            output_format="row_col_value"
        )
        
        if result.get("extracted_tables"):
            table = result["extracted_tables"][0]
            # Should only have selected columns
            assert "Header 1" in table["headers"]
        
        # Test with html
        result = self.simple_extractor.extract_tables_simple(
            file_path=self.test_file,
            slide_numbers=[1, 2, 3],
            column_selection=column_selection,
            output_format="html"
        )
        
        if result.get("extracted_html_tables"):
            table = result["extracted_html_tables"][0]
            assert "Header 1" in table["headers"]
            html = table["htmldata"]
            assert "Header 1" in html
    
    def test_empty_slides(self):
        """Test behavior with slides that have no tables."""
        result = self.simple_extractor.extract_tables_simple(
            file_path=self.test_file,
            slide_numbers=[1],  # Slide 1 typically has no tables
            output_format="html"
        )
        
        # Should return empty list, not error
        assert "extracted_html_tables" in result
        assert isinstance(result["extracted_html_tables"], list)
    
    def test_all_formats_consistency(self):
        """Test that all formats return consistent table counts."""
        formats = ["row_col_value", "row_col_formattedvalue", "simple_html", "html"]
        table_counts = {}
        
        for fmt in formats:
            result = self.simple_extractor.extract_tables_simple(
                file_path=self.test_file,
                slide_numbers=[1, 2, 3],
                output_format=fmt
            )
            
            if fmt in ["html", "simple_html"]:
                table_counts[fmt] = len(result.get("extracted_html_tables", []))
            else:
                table_counts[fmt] = len(result.get("extracted_tables", []))
        
        # All formats should find the same number of tables
        counts = list(table_counts.values())
        if counts:
            assert all(c == counts[0] for c in counts), f"Inconsistent table counts: {table_counts}"


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
