"""
Unit tests for TextFormattingAnalyzer.
"""

import pytest
from unittest.mock import Mock, patch, MagicMock
from collections import defaultdict

from powerpoint_mcp_server.core.text_formatting_analyzer import (
    TextFormattingAnalyzer,
    FormattingFilter,
    FormattedTextElement,
    FormattingAnalysisResult,
    ContentType,
    FormattingType,
    GroupingType,
    create_formatting_filter_from_dict
)


class TestTextFormattingAnalyzer:
    """Test cases for TextFormattingAnalyzer."""
    
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
        extractor._extract_shape_text_content.return_value = "Sample text"
        extractor._extract_shape_transform.return_value = ((100, 200), (300, 400))
        extractor._extract_cell_text_content.return_value = "Cell text"
        
        return extractor
    
    @pytest.fixture
    def formatting_analyzer(self, mock_content_extractor):
        """Create a TextFormattingAnalyzer with mocked dependencies."""
        return TextFormattingAnalyzer(mock_content_extractor)
    
    @pytest.fixture
    def sample_formatted_elements(self):
        """Create sample formatted text elements for testing."""
        elements = [
            FormattedTextElement(
                slide_number=1,
                content_type=ContentType.TITLES,
                element_index=0,
                text_content="Bold Title",
                formatting={
                    'bold_count': 1,
                    'italic_count': 0,
                    'font_colors': ['#000000'],
                    'font_sizes': [24],
                    'has_formatting': True
                },
                position=(100, 200),
                parent_element="titles_0"
            ),
            FormattedTextElement(
                slide_number=1,
                content_type=ContentType.TEXT_BOXES,
                element_index=0,
                text_content="Italic and colored text",
                formatting={
                    'bold_count': 0,
                    'italic_count': 1,
                    'font_colors': ['#FF0000'],
                    'font_sizes': [12],
                    'highlight_count': 1,
                    'has_formatting': True
                },
                position=(200, 300),
                parent_element="text_boxes_0"
            ),
            FormattedTextElement(
                slide_number=2,
                content_type=ContentType.TABLES,
                element_index=0,
                text_content="Table cell with hyperlink",
                formatting={
                    'bold_count': 0,
                    'italic_count': 0,
                    'font_colors': ['#0000FF'],
                    'font_sizes': [10],
                    'hyperlinks': ['link1'],
                    'has_formatting': True
                },
                position=(0, 0),
                parent_element="table_0_row_0"
            ),
            FormattedTextElement(
                slide_number=2,
                content_type=ContentType.BULLETS,
                element_index=0,
                text_content="Plain bullet point",
                formatting={
                    'bold_count': 0,
                    'italic_count': 0,
                    'font_colors': [],
                    'font_sizes': [11],
                    'has_formatting': False
                },
                position=(150, 250),
                parent_element="bullets_0_para_0"
            )
        ]
        return elements
    
    def test_formatting_filter_creation(self):
        """Test creating formatting filter from dictionary."""
        filter_dict = {
            'formatting_types': ['bold', 'italic'],
            'content_types': ['titles', 'text_boxes'],
            'text_contains': 'important',
            'slide_numbers': [1, 2, 3],
            'min_font_size': 10,
            'max_font_size': 20,
            'colors': ['#FF0000', '#00FF00']
        }
        
        formatting_filter = create_formatting_filter_from_dict(filter_dict)
        
        assert len(formatting_filter.formatting_types) == 2
        assert FormattingType.BOLD in formatting_filter.formatting_types
        assert FormattingType.ITALIC in formatting_filter.formatting_types
        assert len(formatting_filter.content_types) == 2
        assert ContentType.TITLES in formatting_filter.content_types
        assert ContentType.TEXT_BOXES in formatting_filter.content_types
        assert formatting_filter.text_contains == 'important'
        assert formatting_filter.slide_numbers == [1, 2, 3]
        assert formatting_filter.min_font_size == 10
        assert formatting_filter.max_font_size == 20
        assert formatting_filter.colors == ['#FF0000', '#00FF00']
    
    def test_has_requested_formatting_bold(self, formatting_analyzer, sample_formatted_elements):
        """Test checking for bold formatting."""
        bold_element = sample_formatted_elements[0]  # Has bold formatting
        plain_element = sample_formatted_elements[3]  # No bold formatting
        
        assert formatting_analyzer._has_requested_formatting(
            bold_element, [FormattingType.BOLD]
        ) is True
        
        assert formatting_analyzer._has_requested_formatting(
            plain_element, [FormattingType.BOLD]
        ) is False
    
    def test_has_requested_formatting_italic(self, formatting_analyzer, sample_formatted_elements):
        """Test checking for italic formatting."""
        italic_element = sample_formatted_elements[1]  # Has italic formatting
        bold_element = sample_formatted_elements[0]   # No italic formatting
        
        assert formatting_analyzer._has_requested_formatting(
            italic_element, [FormattingType.ITALIC]
        ) is True
        
        assert formatting_analyzer._has_requested_formatting(
            bold_element, [FormattingType.ITALIC]
        ) is False
    
    def test_has_requested_formatting_color(self, formatting_analyzer, sample_formatted_elements):
        """Test checking for color formatting."""
        colored_element = sample_formatted_elements[1]  # Has color formatting
        
        assert formatting_analyzer._has_requested_formatting(
            colored_element, [FormattingType.COLOR]
        ) is True
    
    def test_has_requested_formatting_hyperlink(self, formatting_analyzer, sample_formatted_elements):
        """Test checking for hyperlink formatting."""
        hyperlink_element = sample_formatted_elements[2]  # Has hyperlink
        plain_element = sample_formatted_elements[3]      # No hyperlink
        
        assert formatting_analyzer._has_requested_formatting(
            hyperlink_element, [FormattingType.HYPERLINK]
        ) is True
        
        assert formatting_analyzer._has_requested_formatting(
            plain_element, [FormattingType.HYPERLINK]
        ) is False
    
    def test_apply_formatting_filters_slide_numbers(self, formatting_analyzer, sample_formatted_elements):
        """Test filtering by slide numbers."""
        formatting_filter = FormattingFilter(slide_numbers=[1])
        
        filtered = formatting_analyzer._apply_formatting_filters(
            sample_formatted_elements, formatting_filter
        )
        
        assert len(filtered) == 2  # Only elements from slide 1
        assert all(elem.slide_number == 1 for elem in filtered)
    
    def test_apply_formatting_filters_formatting_types(self, formatting_analyzer, sample_formatted_elements):
        """Test filtering by formatting types."""
        formatting_filter = FormattingFilter(
            formatting_types=[FormattingType.BOLD]
        )
        
        filtered = formatting_analyzer._apply_formatting_filters(
            sample_formatted_elements, formatting_filter
        )
        
        assert len(filtered) == 1  # Only the bold element
        assert filtered[0].formatting['bold_count'] > 0
    
    def test_apply_formatting_filters_text_contains(self, formatting_analyzer, sample_formatted_elements):
        """Test filtering by text content."""
        formatting_filter = FormattingFilter(text_contains="Title")
        
        filtered = formatting_analyzer._apply_formatting_filters(
            sample_formatted_elements, formatting_filter
        )
        
        assert len(filtered) == 1  # Only the element with "Title"
        assert "Title" in filtered[0].text_content
    
    def test_apply_formatting_filters_font_size(self, formatting_analyzer, sample_formatted_elements):
        """Test filtering by font size."""
        formatting_filter = FormattingFilter(min_font_size=15)
        
        filtered = formatting_analyzer._apply_formatting_filters(
            sample_formatted_elements, formatting_filter
        )
        
        assert len(filtered) == 1  # Only the element with font size 24
        assert 24 in filtered[0].formatting['font_sizes']
    
    def test_apply_formatting_filters_colors(self, formatting_analyzer, sample_formatted_elements):
        """Test filtering by colors."""
        formatting_filter = FormattingFilter(colors=['#FF0000'])
        
        filtered = formatting_analyzer._apply_formatting_filters(
            sample_formatted_elements, formatting_filter
        )
        
        assert len(filtered) == 1  # Only the element with red color
        assert '#FF0000' in filtered[0].formatting['font_colors']
    
    def test_create_formatting_summary(self, formatting_analyzer, sample_formatted_elements):
        """Test creating formatting summary."""
        summary = formatting_analyzer._create_formatting_summary(sample_formatted_elements)
        
        assert summary['total_elements'] == 4
        assert summary['elements_with_formatting'] == 3  # 3 elements have formatting
        
        # Check formatting counts
        assert summary['formatting_counts']['bold'] == 1
        assert summary['formatting_counts']['italic'] == 1
        assert summary['formatting_counts']['highlight'] == 1
        assert summary['formatting_counts']['colored_text'] == 3
        assert summary['formatting_counts']['hyperlinks'] == 1
        
        # Check font sizes
        assert 24 in summary['font_sizes']['unique_sizes']
        assert 12 in summary['font_sizes']['unique_sizes']
        assert 10 in summary['font_sizes']['unique_sizes']
        assert 11 in summary['font_sizes']['unique_sizes']
        
        # Check colors
        assert '#000000' in summary['colors']['unique_colors']
        assert '#FF0000' in summary['colors']['unique_colors']
        assert '#0000FF' in summary['colors']['unique_colors']
        
        # Check content type distribution
        assert summary['content_type_distribution']['titles'] == 1
        assert summary['content_type_distribution']['text_boxes'] == 1
        assert summary['content_type_distribution']['tables'] == 1
        assert summary['content_type_distribution']['bullets'] == 1
        
        # Check slide distribution
        assert summary['slide_distribution'][1] == 2
        assert summary['slide_distribution'][2] == 2
    
    def test_group_by_slide(self, formatting_analyzer, sample_formatted_elements):
        """Test grouping elements by slide."""
        groups = formatting_analyzer._group_by_slide(sample_formatted_elements)
        
        assert 'slide_1' in groups
        assert 'slide_2' in groups
        assert len(groups['slide_1']) == 2
        assert len(groups['slide_2']) == 2
        
        # Check that elements are properly grouped
        slide_1_elements = groups['slide_1']
        assert any(elem['content_type'] == 'titles' for elem in slide_1_elements)
        assert any(elem['content_type'] == 'text_boxes' for elem in slide_1_elements)
    
    def test_group_by_formatting_type(self, formatting_analyzer, sample_formatted_elements):
        """Test grouping elements by formatting type."""
        groups = formatting_analyzer._group_by_formatting_type(sample_formatted_elements)
        
        assert len(groups['bold']) == 1
        assert len(groups['italic']) == 1
        assert len(groups['colored']) == 3
        assert len(groups['hyperlinks']) == 1
        
        # Check that bold group contains the correct element
        bold_element = groups['bold'][0]
        assert bold_element['slide_number'] == 1
        assert 'Bold Title' in bold_element['text_content']
    
    def test_group_by_content_type(self, formatting_analyzer, sample_formatted_elements):
        """Test grouping elements by content type."""
        groups = formatting_analyzer._group_by_content_type(sample_formatted_elements)
        
        assert 'titles' in groups
        assert 'text_boxes' in groups
        assert 'tables' in groups
        assert 'bullets' in groups
        
        assert len(groups['titles']) == 1
        assert len(groups['text_boxes']) == 1
        assert len(groups['tables']) == 1
        assert len(groups['bullets']) == 1
    
    def test_group_by_color(self, formatting_analyzer, sample_formatted_elements):
        """Test grouping elements by color."""
        groups = formatting_analyzer._group_by_color(sample_formatted_elements)
        
        assert '#000000' in groups
        assert '#FF0000' in groups
        assert '#0000FF' in groups
        assert 'no_color' in groups
        
        assert len(groups['#000000']) == 1
        assert len(groups['#FF0000']) == 1
        assert len(groups['#0000FF']) == 1
        assert len(groups['no_color']) == 1  # The element with no colors
    
    def test_group_by_font_size(self, formatting_analyzer, sample_formatted_elements):
        """Test grouping elements by font size."""
        groups = formatting_analyzer._group_by_font_size(sample_formatted_elements)
        
        assert 'size_24' in groups
        assert 'size_12' in groups
        assert 'size_10' in groups
        assert 'size_11' in groups
        
        # Check that each group has the correct element
        assert len(groups['size_24']) == 1
        assert groups['size_24'][0]['slide_number'] == 1
    
    def test_apply_grouping_by_slide(self, formatting_analyzer, sample_formatted_elements):
        """Test applying slide grouping."""
        groups = formatting_analyzer._apply_grouping(
            sample_formatted_elements, GroupingType.BY_SLIDE
        )
        
        assert 'slide_1' in groups
        assert 'slide_2' in groups
    
    def test_apply_grouping_by_formatting_type(self, formatting_analyzer, sample_formatted_elements):
        """Test applying formatting type grouping."""
        groups = formatting_analyzer._apply_grouping(
            sample_formatted_elements, GroupingType.BY_FORMATTING_TYPE
        )
        
        assert 'bold' in groups
        assert 'italic' in groups
        assert 'colored' in groups
    
    def test_apply_grouping_none(self, formatting_analyzer, sample_formatted_elements):
        """Test applying no grouping."""
        groups = formatting_analyzer._apply_grouping(
            sample_formatted_elements, GroupingType.NONE
        )
        
        assert groups == {}
    
    def test_analyze_text_formatting_in_element(self, formatting_analyzer):
        """Test analyzing text formatting in an element."""
        # Mock element with formatting
        element = Mock()
        
        # Mock runs with formatting
        run1 = Mock()
        run2 = Mock()
        
        # Mock run properties
        r_pr1 = Mock()
        r_pr2 = Mock()
        
        # Mock formatting elements
        bold_elem = Mock()
        bold_elem.get.return_value = '1'
        
        italic_elem = Mock()
        italic_elem.get.return_value = '1'
        
        font_size_elem = Mock()
        font_size_elem.get.return_value = '1200'  # 12pt in hundredths
        
        # Set up the mock chain
        formatting_analyzer.content_extractor.xml_parser.find_elements_with_namespace.return_value = [run1, run2]
        
        def mock_find_element(elem, xpath):
            if elem == run1 and 'rPr' in xpath:
                return r_pr1
            elif elem == run2 and 'rPr' in xpath:
                return r_pr2
            elif elem == r_pr1 and 'b' in xpath:
                return bold_elem
            elif elem == r_pr2 and 'i' in xpath:
                return italic_elem
            elif elem == r_pr1 and 'sz' in xpath:
                return font_size_elem
            return None
        
        formatting_analyzer.content_extractor.xml_parser.find_element_with_namespace.side_effect = mock_find_element
        
        # Test the method
        formatting = formatting_analyzer._analyze_text_formatting_in_element(element)
        
        assert formatting['bold_count'] == 1
        assert formatting['italic_count'] == 1
        assert 12 in formatting['font_sizes']
        assert formatting['has_formatting'] is True
    
    def test_extract_color_from_fill_rgb(self, formatting_analyzer):
        """Test extracting RGB color from fill."""
        solid_fill = Mock()
        srgb_clr = Mock()
        srgb_clr.get.return_value = "FF0000"
        
        formatting_analyzer.content_extractor.xml_parser.find_element_with_namespace.side_effect = [
            srgb_clr,  # First call returns RGB color
            None       # Second call returns None for scheme color
        ]
        
        color = formatting_analyzer._extract_color_from_fill(solid_fill)
        assert color == "#FF0000"
    
    def test_extract_color_from_fill_scheme(self, formatting_analyzer):
        """Test extracting scheme color from fill."""
        solid_fill = Mock()
        scheme_clr = Mock()
        scheme_clr.get.return_value = "accent1"
        
        formatting_analyzer.content_extractor.xml_parser.find_element_with_namespace.side_effect = [
            None,       # First call returns None for RGB color
            scheme_clr  # Second call returns scheme color
        ]
        
        color = formatting_analyzer._extract_color_from_fill(solid_fill)
        assert color == "accent1"
    
    def test_cache_operations(self, formatting_analyzer):
        """Test cache operations."""
        # Add something to cache
        formatting_analyzer._analysis_cache["test_key"] = "test_value"
        assert len(formatting_analyzer._analysis_cache) == 1
        
        # Clear cache
        formatting_analyzer.clear_cache()
        assert len(formatting_analyzer._analysis_cache) == 0
    
    @patch('powerpoint_mcp_server.utils.zip_extractor.ZipExtractor')
    def test_analyze_formatting_integration(self, mock_zip_extractor, formatting_analyzer):
        """Test the main analyze_formatting method integration."""
        # Mock ZipExtractor
        mock_extractor_instance = Mock()
        mock_zip_extractor.return_value.__enter__.return_value = mock_extractor_instance
        
        mock_extractor_instance.get_slide_xml_files.return_value = ["slide1.xml"]
        mock_extractor_instance.read_xml_content.return_value = "<xml>mock content</xml>"
        
        # Mock the internal extraction method
        sample_element = FormattedTextElement(
            slide_number=1,
            content_type=ContentType.TITLES,
            element_index=0,
            text_content="Test Title",
            formatting={'bold_count': 1, 'has_formatting': True}
        )
        
        formatting_analyzer._extract_formatted_elements_from_slide = Mock(
            return_value=[sample_element]
        )
        
        # Test analysis
        result = formatting_analyzer.analyze_formatting(
            file_path="test.pptx",
            slide_numbers=[1],
            grouping=GroupingType.BY_SLIDE
        )
        
        assert isinstance(result, FormattingAnalysisResult)
        assert result.total_elements == 1
        assert len(result.formatted_elements) == 1
        assert 'formatting_summary' in result.__dict__
        assert 'groupings' in result.__dict__
        
        # Verify the mock was called correctly
        formatting_analyzer._extract_formatted_elements_from_slide.assert_called_once()


class TestFormattedTextElement:
    """Test cases for FormattedTextElement class."""
    
    def test_formatted_text_element_creation(self):
        """Test creating a FormattedTextElement."""
        element = FormattedTextElement(
            slide_number=1,
            content_type=ContentType.TITLES,
            element_index=0,
            text_content="Test Title",
            formatting={'bold_count': 1},
            position=(100, 200),
            size=(300, 400),
            parent_element="titles_0"
        )
        
        assert element.slide_number == 1
        assert element.content_type == ContentType.TITLES
        assert element.element_index == 0
        assert element.text_content == "Test Title"
        assert element.formatting == {'bold_count': 1}
        assert element.position == (100, 200)
        assert element.size == (300, 400)
        assert element.parent_element == "titles_0"


class TestFormattingFilter:
    """Test cases for FormattingFilter class."""
    
    def test_formatting_filter_defaults(self):
        """Test default values for FormattingFilter."""
        formatting_filter = FormattingFilter()
        
        assert formatting_filter.formatting_types is None
        assert formatting_filter.content_types is None
        assert formatting_filter.text_contains is None
        assert formatting_filter.text_patterns is None
        assert formatting_filter.slide_numbers is None
        assert formatting_filter.min_font_size is None
        assert formatting_filter.max_font_size is None
        assert formatting_filter.colors is None
    
    def test_formatting_filter_with_values(self):
        """Test FormattingFilter with specific values."""
        formatting_filter = FormattingFilter(
            formatting_types=[FormattingType.BOLD, FormattingType.ITALIC],
            content_types=[ContentType.TITLES],
            text_contains="important",
            slide_numbers=[1, 2, 3],
            min_font_size=10,
            colors=['#FF0000']
        )
        
        assert len(formatting_filter.formatting_types) == 2
        assert FormattingType.BOLD in formatting_filter.formatting_types
        assert FormattingType.ITALIC in formatting_filter.formatting_types
        assert len(formatting_filter.content_types) == 1
        assert ContentType.TITLES in formatting_filter.content_types
        assert formatting_filter.text_contains == "important"
        assert formatting_filter.slide_numbers == [1, 2, 3]
        assert formatting_filter.min_font_size == 10
        assert formatting_filter.colors == ['#FF0000']


if __name__ == "__main__":
    pytest.main([__file__])