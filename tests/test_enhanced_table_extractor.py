"""
Unit tests for EnhancedTableExtractor.
"""

import pytest
import tempfile
import os
from unittest.mock import Mock, patch, MagicMock

from powerpoint_mcp_server.core.enhanced_table_extractor import (
    EnhancedTableExtractor,
    TableCriteria,
    ColumnSelection,
    FormattingDetection,
    OutputFormat,
    EnhancedTable,
    EnhancedTableCell,
    CellFormatting,
    create_table_criteria_from_dict,
    create_column_selection_from_dict,
    create_formatting_detection_from_dict
)


class TestEnhancedTableExtractor:
    """Test cases for EnhancedTableExtractor."""
    
    @pytest.fixture
    def mock_content_extractor(self):
        """Create a mock content extractor."""
        extractor = Mock()
        
        # Mock XML parser
        xml_parser = Mock()
        xml_parser.parse_xml_string.return_value = Mock()
        xml_parser.find_elements_with_namespace.return_value = []
        xml_parser.find_element_with_namespace.return_value = None
        extractor.xml_parser = xml_parser
        
        # Mock methods
        extractor._extract_graphic_frame_transform.return_value = ((100, 200), (800, 400))
        extractor._extract_cell_text_content.return_value = "Sample text"
        
        return extractor
    
    @pytest.fixture
    def table_extractor(self, mock_content_extractor):
        """Create an EnhancedTableExtractor with mocked dependencies."""
        return EnhancedTableExtractor(mock_content_extractor)
    
    @pytest.fixture
    def sample_enhanced_table(self):
        """Create a sample enhanced table for testing."""
        # Create sample cells with formatting
        cell1 = EnhancedTableCell(
            value="Task Name",
            formatting=CellFormatting(bold=True, font_color="#000000"),
            position=(0, 0)
        )
        cell2 = EnhancedTableCell(
            value="Progress",
            formatting=CellFormatting(bold=True, font_color="#000000"),
            position=(0, 1)
        )
        cell3 = EnhancedTableCell(
            value="Status",
            formatting=CellFormatting(bold=True, font_color="#000000"),
            position=(0, 2)
        )
        
        cell4 = EnhancedTableCell(
            value="System Design",
            formatting=CellFormatting(),
            position=(1, 0)
        )
        cell5 = EnhancedTableCell(
            value="80%",
            formatting=CellFormatting(),
            position=(1, 1)
        )
        cell6 = EnhancedTableCell(
            value="In Progress",
            formatting=CellFormatting(highlight=True, font_color="#FF0000"),
            position=(1, 2)
        )
        
        # Create table data
        data = [
            {"Task Name": cell1, "Progress": cell2, "Status": cell3},
            {"Task Name": cell4, "Progress": cell5, "Status": cell6}
        ]
        
        return EnhancedTable(
            slide_number=1,
            table_index=0,
            rows=2,
            columns=3,
            headers=["Task Name", "Progress", "Status"],
            data=data,
            metadata={"has_formatting": True, "cell_count": 6, "non_empty_cells": 6},
            position=(100, 200),
            size=(800, 400)
        )
    
    def test_table_criteria_creation(self):
        """Test creating table criteria from dictionary."""
        criteria_dict = {
            "min_rows": 2,
            "min_columns": 3,
            "max_rows": 10,
            "header_contains": ["Task", "Progress"],
            "header_patterns": [".*Status.*"]
        }
        
        criteria = create_table_criteria_from_dict(criteria_dict)
        
        assert criteria.min_rows == 2
        assert criteria.min_columns == 3
        assert criteria.max_rows == 10
        assert criteria.header_contains == ["Task", "Progress"]
        assert criteria.header_patterns == [".*Status.*"]
    
    def test_column_selection_creation(self):
        """Test creating column selection from dictionary."""
        selection_dict = {
            "specific_columns": ["Task Name", "Progress"],
            "exclude_columns": ["Notes"],
            "all_columns": False
        }
        
        selection = create_column_selection_from_dict(selection_dict)
        
        assert selection.specific_columns == ["Task Name", "Progress"]
        assert selection.exclude_columns == ["Notes"]
        assert selection.all_columns is False
    
    def test_formatting_detection_creation(self):
        """Test creating formatting detection from dictionary."""
        detection_dict = {
            "detect_bold": True,
            "detect_italic": False,
            "detect_colors": True,
            "preserve_formatting": True
        }
        
        detection = create_formatting_detection_from_dict(detection_dict)
        
        assert detection.detect_bold is True
        assert detection.detect_italic is False
        assert detection.detect_colors is True
        assert detection.preserve_formatting is True
    
    def test_meets_table_criteria_min_rows(self, sample_enhanced_table):
        """Test table criteria checking for minimum rows."""
        extractor = EnhancedTableExtractor()
        
        # Should pass with min_rows = 2
        criteria = TableCriteria(min_rows=2)
        assert extractor._meets_table_criteria(sample_enhanced_table, criteria) is True
        
        # Should fail with min_rows = 5
        criteria = TableCriteria(min_rows=5)
        assert extractor._meets_table_criteria(sample_enhanced_table, criteria) is False
    
    def test_meets_table_criteria_min_columns(self, sample_enhanced_table):
        """Test table criteria checking for minimum columns."""
        extractor = EnhancedTableExtractor()
        
        # Should pass with min_columns = 3
        criteria = TableCriteria(min_columns=3)
        assert extractor._meets_table_criteria(sample_enhanced_table, criteria) is True
        
        # Should fail with min_columns = 5
        criteria = TableCriteria(min_columns=5)
        assert extractor._meets_table_criteria(sample_enhanced_table, criteria) is False
    
    def test_meets_table_criteria_header_contains(self, sample_enhanced_table):
        """Test table criteria checking for header contains."""
        extractor = EnhancedTableExtractor()
        
        # Should pass with existing headers
        criteria = TableCriteria(header_contains=["Task", "Progress"])
        assert extractor._meets_table_criteria(sample_enhanced_table, criteria) is True
        
        # Should fail with non-existing header
        criteria = TableCriteria(header_contains=["NonExistent"])
        assert extractor._meets_table_criteria(sample_enhanced_table, criteria) is False
    
    def test_meets_table_criteria_header_patterns(self, sample_enhanced_table):
        """Test table criteria checking for header patterns."""
        extractor = EnhancedTableExtractor()
        
        # Should pass with regex pattern
        criteria = TableCriteria(header_patterns=[".*Task.*", ".*Status.*"])
        assert extractor._meets_table_criteria(sample_enhanced_table, criteria) is True
        
        # Should fail with non-matching pattern
        criteria = TableCriteria(header_patterns=[".*NonExistent.*"])
        assert extractor._meets_table_criteria(sample_enhanced_table, criteria) is False
    
    def test_apply_column_selection_specific_columns(self, sample_enhanced_table):
        """Test applying column selection with specific columns."""
        extractor = EnhancedTableExtractor()
        
        selection = ColumnSelection(
            specific_columns=["Task Name", "Status"],
            all_columns=False
        )
        
        filtered_table = extractor._apply_column_selection(sample_enhanced_table, selection)
        
        assert len(filtered_table.headers) == 2
        assert "Task Name" in filtered_table.headers
        assert "Status" in filtered_table.headers
        assert "Progress" not in filtered_table.headers
        assert filtered_table.columns == 2
        
        # Check that data is filtered correctly
        for row in filtered_table.data:
            assert "Task Name" in row
            assert "Status" in row
            assert "Progress" not in row
    
    def test_apply_column_selection_exclude_columns(self, sample_enhanced_table):
        """Test applying column selection with excluded columns."""
        extractor = EnhancedTableExtractor()
        
        selection = ColumnSelection(
            exclude_columns=["Progress"],
            all_columns=True
        )
        
        filtered_table = extractor._apply_column_selection(sample_enhanced_table, selection)
        
        assert len(filtered_table.headers) == 2
        assert "Task Name" in filtered_table.headers
        assert "Status" in filtered_table.headers
        assert "Progress" not in filtered_table.headers
    
    def test_has_formatting(self, sample_enhanced_table):
        """Test checking if table has formatting."""
        extractor = EnhancedTableExtractor()
        
        # Sample table has formatting
        assert extractor._has_formatting(sample_enhanced_table.data) is True
        
        # Create table without formatting
        cell_no_format = EnhancedTableCell(value="Plain text", formatting=CellFormatting())
        data_no_format = [{"Column1": cell_no_format}]
        
        assert extractor._has_formatting(data_no_format) is False
    
    def test_count_non_empty_cells(self, sample_enhanced_table):
        """Test counting non-empty cells."""
        extractor = EnhancedTableExtractor()
        
        count = extractor._count_non_empty_cells(sample_enhanced_table.data)
        assert count == 6  # All cells have content
        
        # Create table with empty cells
        empty_cell = EnhancedTableCell(value="", formatting=CellFormatting())
        filled_cell = EnhancedTableCell(value="Content", formatting=CellFormatting())
        data_with_empty = [{"Col1": filled_cell, "Col2": empty_cell}]
        
        count = extractor._count_non_empty_cells(data_with_empty)
        assert count == 1
    
    def test_format_structured_output(self, sample_enhanced_table):
        """Test formatting output in structured format."""
        extractor = EnhancedTableExtractor()
        
        result = extractor._format_structured_output([sample_enhanced_table], include_metadata=True)
        
        assert "extracted_tables" in result
        assert "summary" in result
        assert len(result["extracted_tables"]) == 1
        
        table_dict = result["extracted_tables"][0]
        assert table_dict["slide_number"] == 1
        assert table_dict["table_index"] == 0
        assert table_dict["rows"] == 2
        assert table_dict["columns"] == 3
        assert table_dict["headers"] == ["Task Name", "Progress", "Status"]
        assert "metadata" in table_dict
        assert "position" in table_dict
        assert "size" in table_dict
        
        # Check data structure
        assert len(table_dict["data"]) == 2
        first_row = table_dict["data"][0]
        assert "Task Name" in first_row
        assert "value" in first_row["Task Name"]
        assert "formatting" in first_row["Task Name"]
        
        # Check formatting
        task_cell = first_row["Task Name"]
        assert task_cell["formatting"]["bold"] is True
        assert task_cell["formatting"]["font_color"] == "#000000"
        
        # Check summary
        summary = result["summary"]
        assert summary["total_tables"] == 1
        assert summary["total_rows"] == 2
        assert summary["slides_with_tables"] == 1
        assert "formatting_found" in summary
    
    def test_format_flat_output(self, sample_enhanced_table):
        """Test formatting output in flat format."""
        extractor = EnhancedTableExtractor()
        
        result = extractor._format_flat_output([sample_enhanced_table], include_metadata=True)
        
        assert "data" in result
        assert "summary" in result
        assert len(result["data"]) == 2  # Two rows
        
        first_row = result["data"][0]
        assert first_row["slide_number"] == 1
        assert first_row["table_index"] == 0
        assert first_row["row_index"] == 0
        assert first_row["Task Name"] == "Task Name"
        assert first_row["Progress"] == "Progress"
        assert first_row["Status"] == "Status"
        
        # Check formatting metadata
        assert first_row["Task Name_bold"] is True
        assert first_row["Progress_bold"] is True
        assert first_row["Status_bold"] is True
    
    def test_format_grouped_output(self, sample_enhanced_table):
        """Test formatting output grouped by slide."""
        extractor = EnhancedTableExtractor()
        
        result = extractor._format_grouped_output([sample_enhanced_table], include_metadata=True)
        
        assert "slides" in result
        assert "summary" in result
        assert len(result["slides"]) == 1
        
        slide_data = result["slides"][0]
        assert slide_data["slide_number"] == 1
        assert "tables" in slide_data
        assert len(slide_data["tables"]) == 1
        
        # Check summary
        summary = result["summary"]
        assert summary["total_slides"] == 1
        assert summary["total_tables"] == 1
    
    def test_count_formatted_cells(self, sample_enhanced_table):
        """Test counting formatted cells."""
        extractor = EnhancedTableExtractor()
        
        bold_count = extractor._count_formatted_cells([sample_enhanced_table], "bold")
        assert bold_count == 3  # Header cells are bold
        
        highlight_count = extractor._count_formatted_cells([sample_enhanced_table], "highlight")
        assert highlight_count == 1  # One cell is highlighted
        
        color_count = extractor._count_formatted_cells([sample_enhanced_table], "color")
        assert color_count == 4  # Four cells have colors
    
    def test_extract_color_from_fill(self):
        """Test extracting color from fill element."""
        extractor = EnhancedTableExtractor()
        
        # Mock solid fill with RGB color
        solid_fill = Mock()
        srgb_clr = Mock()
        srgb_clr.get.return_value = "FF0000"
        
        extractor.content_extractor.xml_parser.find_element_with_namespace.side_effect = [
            srgb_clr,  # First call returns RGB color
            None       # Second call returns None for scheme color
        ]
        
        color = extractor._extract_color_from_fill(solid_fill)
        assert color == "#FF0000"
    
    def test_extract_color_from_fill_scheme(self):
        """Test extracting scheme color from fill element."""
        extractor = EnhancedTableExtractor()
        
        # Mock solid fill with scheme color
        solid_fill = Mock()
        scheme_clr = Mock()
        scheme_clr.get.return_value = "accent1"
        
        extractor.content_extractor.xml_parser.find_element_with_namespace.side_effect = [
            None,       # First call returns None for RGB color
            scheme_clr  # Second call returns scheme color
        ]
        
        color = extractor._extract_color_from_fill(solid_fill)
        assert color == "accent1"
    
    def test_cache_operations(self, table_extractor):
        """Test cache operations."""
        # Add something to cache
        table_extractor._table_cache["test_key"] = "test_value"
        assert len(table_extractor._table_cache) == 1
        
        # Clear cache
        table_extractor.clear_cache()
        assert len(table_extractor._table_cache) == 0
    
    @patch('powerpoint_mcp_server.utils.zip_extractor.ZipExtractor')
    def test_extract_tables_integration(self, mock_zip_extractor, table_extractor):
        """Test the main extract_tables method integration."""
        # Mock ZipExtractor
        mock_extractor_instance = Mock()
        mock_zip_extractor.return_value.__enter__.return_value = mock_extractor_instance
        
        mock_extractor_instance.get_slide_xml_files.return_value = ["slide1.xml"]
        mock_extractor_instance.read_xml_content.return_value = "<xml>mock content</xml>"
        
        # Mock the internal extraction method to return a sample table
        sample_table = EnhancedTable(
            slide_number=1,
            table_index=0,
            rows=1,
            columns=2,
            headers=["Col1", "Col2"],
            data=[{"Col1": EnhancedTableCell(value="A"), "Col2": EnhancedTableCell(value="B")}]
        )
        
        table_extractor._extract_tables_from_slide = Mock(return_value=[sample_table])
        
        # Test extraction
        result = table_extractor.extract_tables(
            file_path="test.pptx",
            slide_numbers=[1],
            output_format=OutputFormat.STRUCTURED
        )
        
        assert "extracted_tables" in result
        assert "summary" in result
        assert len(result["extracted_tables"]) == 1
        
        # Verify the mock was called correctly
        table_extractor._extract_tables_from_slide.assert_called_once()


class TestCellFormatting:
    """Test cases for CellFormatting class."""
    
    def test_cell_formatting_defaults(self):
        """Test default values for CellFormatting."""
        formatting = CellFormatting()
        
        assert formatting.bold is False
        assert formatting.italic is False
        assert formatting.underline is False
        assert formatting.highlight is False
        assert formatting.strikethrough is False
        assert formatting.font_color is None
        assert formatting.background_color is None
        assert formatting.font_size is None
        assert formatting.hyperlink is None
    
    def test_cell_formatting_with_values(self):
        """Test CellFormatting with specific values."""
        formatting = CellFormatting(
            bold=True,
            italic=True,
            font_color="#FF0000",
            font_size=12
        )
        
        assert formatting.bold is True
        assert formatting.italic is True
        assert formatting.font_color == "#FF0000"
        assert formatting.font_size == 12


class TestEnhancedTableCell:
    """Test cases for EnhancedTableCell class."""
    
    def test_enhanced_table_cell_defaults(self):
        """Test default values for EnhancedTableCell."""
        cell = EnhancedTableCell(value="Test")
        
        assert cell.value == "Test"
        assert isinstance(cell.formatting, CellFormatting)
        assert cell.row_span == 1
        assert cell.col_span == 1
        assert cell.position == (0, 0)
    
    def test_enhanced_table_cell_with_formatting(self):
        """Test EnhancedTableCell with formatting."""
        formatting = CellFormatting(bold=True, font_color="#FF0000")
        cell = EnhancedTableCell(
            value="Formatted Text",
            formatting=formatting,
            row_span=2,
            col_span=3,
            position=(1, 2)
        )
        
        assert cell.value == "Formatted Text"
        assert cell.formatting.bold is True
        assert cell.formatting.font_color == "#FF0000"
        assert cell.row_span == 2
        assert cell.col_span == 3
        assert cell.position == (1, 2)


class TestEnhancedTable:
    """Test cases for EnhancedTable class."""
    
    def test_enhanced_table_creation(self):
        """Test creating an EnhancedTable."""
        cell = EnhancedTableCell(value="Test")
        data = [{"Column1": cell}]
        
        table = EnhancedTable(
            slide_number=1,
            table_index=0,
            rows=1,
            columns=1,
            headers=["Column1"],
            data=data,
            position=(100, 200),
            size=(800, 400)
        )
        
        assert table.slide_number == 1
        assert table.table_index == 0
        assert table.rows == 1
        assert table.columns == 1
        assert table.headers == ["Column1"]
        assert len(table.data) == 1
        assert table.position == (100, 200)
        assert table.size == (800, 400)
        assert isinstance(table.metadata, dict)


if __name__ == "__main__":
    pytest.main([__file__])