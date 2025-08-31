"""
Unit tests for SlideQueryEngine.
"""

import pytest
import tempfile
import os
from unittest.mock import Mock, patch, MagicMock

from powerpoint_mcp_server.core.slide_query_engine import (
    SlideQueryEngine,
    SlideQueryFilters,
    TitleFilter,
    ContentFilter,
    LayoutFilter,
    SlideQueryResult,
    create_filters_from_dict
)


class TestSlideQueryEngine:
    """Test cases for SlideQueryEngine."""
    
    @pytest.fixture
    def mock_content_extractor(self):
        """Create a mock content extractor."""
        extractor = Mock()
        
        # Mock slide info
        mock_slide_info = Mock()
        mock_slide_info.title = "Test Slide Title"
        mock_slide_info.subtitle = "Test Subtitle"
        mock_slide_info.layout_name = "Title and Content"
        mock_slide_info.layout_type = "title_content"
        mock_slide_info.placeholders = []
        mock_slide_info.text_elements = [
            {"content_plain": "Sample text content", "content_formatted": "Sample text content"}
        ]
        mock_slide_info.tables = []
        
        extractor.extract_slide_content.return_value = mock_slide_info
        extractor.extract_presentation_metadata.return_value = {"title": "Test Presentation"}
        extractor._count_slide_objects.return_value = {"shapes": 2, "text_boxes": 1, "images": 0, "tables": 0}
        
        # Mock XML parser
        extractor.xml_parser.parse_xml_string.return_value = Mock()
        
        return extractor
    
    @pytest.fixture
    def query_engine(self, mock_content_extractor):
        """Create a SlideQueryEngine with mocked dependencies."""
        return SlideQueryEngine(mock_content_extractor)
    
    @pytest.fixture
    def sample_slides_data(self):
        """Create sample slide data for testing."""
        return [
            {
                'slide_number': 1,
                'title': 'Introduction',
                'subtitle': 'Welcome to the presentation',
                'layout_name': 'Title Slide',
                'layout_type': 'title',
                'text_elements': [{'content_plain': 'Welcome text'}],
                'tables': [],
                'object_counts': {'shapes': 1, 'text_boxes': 1, 'images': 0, 'tables': 0}
            },
            {
                'slide_number': 2,
                'title': 'Project A Progress',
                'subtitle': 'Current status',
                'layout_name': 'Title and Content',
                'layout_type': 'title_content',
                'text_elements': [{'content_plain': 'Progress details'}],
                'tables': [{'rows': [['Task', 'Status'], ['Task 1', 'Complete']]}],
                'object_counts': {'shapes': 2, 'text_boxes': 1, 'images': 0, 'tables': 1}
            },
            {
                'slide_number': 3,
                'title': 'Project B Progress',
                'subtitle': 'Current status',
                'layout_name': 'Title and Content',
                'layout_type': 'title_content',
                'text_elements': [{'content_plain': 'More progress details'}],
                'tables': [{'rows': [['Task', 'Status'], ['Task 2', 'In Progress']]}],
                'object_counts': {'shapes': 2, 'text_boxes': 1, 'images': 1, 'tables': 1}
            }
        ]
    
    def test_title_filter_contains(self, query_engine, sample_slides_data):
        """Test title filtering with contains condition."""
        query_engine._slide_cache = {"test.pptx:all_slides": sample_slides_data}
        
        filters = SlideQueryFilters(
            title=TitleFilter(contains="Progress")
        )
        
        results = query_engine.query_slides("test.pptx", filters)
        
        assert len(results) == 2
        assert results[0].slide_number == 2
        assert results[1].slide_number == 3
        assert "Progress" in results[0].title
        assert "Progress" in results[1].title
    
    def test_title_filter_starts_with(self, query_engine, sample_slides_data):
        """Test title filtering with starts_with condition."""
        query_engine._slide_cache = {"test.pptx:all_slides": sample_slides_data}
        
        filters = SlideQueryFilters(
            title=TitleFilter(starts_with="Project")
        )
        
        results = query_engine.query_slides("test.pptx", filters)
        
        assert len(results) == 2
        assert all("Project" in result.title for result in results)
    
    def test_title_filter_one_of(self, query_engine, sample_slides_data):
        """Test title filtering with one_of condition."""
        query_engine._slide_cache = {"test.pptx:all_slides": sample_slides_data}
        
        filters = SlideQueryFilters(
            title=TitleFilter(one_of=[".*Project A.*", ".*Introduction.*"])
        )
        
        results = query_engine.query_slides("test.pptx", filters)
        
        assert len(results) == 2
        titles = [result.title for result in results]
        assert "Introduction" in titles
        assert "Project A Progress" in titles
    
    def test_content_filter_has_tables(self, query_engine, sample_slides_data):
        """Test content filtering for slides with tables."""
        query_engine._slide_cache = {"test.pptx:all_slides": sample_slides_data}
        
        filters = SlideQueryFilters(
            content=ContentFilter(has_tables=True)
        )
        
        results = query_engine.query_slides("test.pptx", filters)
        
        assert len(results) == 2
        assert results[0].slide_number == 2
        assert results[1].slide_number == 3
    
    def test_content_filter_has_images(self, query_engine, sample_slides_data):
        """Test content filtering for slides with images."""
        query_engine._slide_cache = {"test.pptx:all_slides": sample_slides_data}
        
        filters = SlideQueryFilters(
            content=ContentFilter(has_images=True)
        )
        
        results = query_engine.query_slides("test.pptx", filters)
        
        assert len(results) == 1
        assert results[0].slide_number == 3
    
    def test_combined_filters(self, query_engine, sample_slides_data):
        """Test combining multiple filters."""
        query_engine._slide_cache = {"test.pptx:all_slides": sample_slides_data}
        
        filters = SlideQueryFilters(
            title=TitleFilter(contains="Project"),
            content=ContentFilter(has_tables=True)
        )
        
        results = query_engine.query_slides("test.pptx", filters)
        
        assert len(results) == 2
        assert all("Project" in result.title for result in results)
    
    def test_slide_numbers_filter(self, query_engine, sample_slides_data):
        """Test filtering by specific slide numbers."""
        query_engine._slide_cache = {"test.pptx:all_slides": sample_slides_data}
        
        filters = SlideQueryFilters(
            slide_numbers=[1, 3]
        )
        
        results = query_engine.query_slides("test.pptx", filters)
        
        assert len(results) == 2
        assert results[0].slide_number == 1
        assert results[1].slide_number == 3
    
    def test_return_fields_selection(self, query_engine, sample_slides_data):
        """Test selecting specific return fields."""
        query_engine._slide_cache = {"test.pptx:all_slides": sample_slides_data}
        
        filters = SlideQueryFilters()
        return_fields = ["slide_number", "title", "object_counts", "table_info"]
        
        results = query_engine.query_slides("test.pptx", filters, return_fields)
        
        assert len(results) == 3
        for result in results:
            assert result.slide_number is not None
            assert result.title is not None
            assert result.object_counts is not None
            # table_info should be present for slides with tables
            if result.slide_number > 1:
                assert result.table_info is not None
    
    def test_limit_results(self, query_engine, sample_slides_data):
        """Test limiting the number of results."""
        query_engine._slide_cache = {"test.pptx:all_slides": sample_slides_data}
        
        filters = SlideQueryFilters()
        
        results = query_engine.query_slides("test.pptx", filters, limit=2)
        
        assert len(results) == 2
    
    def test_regex_title_filter(self, query_engine, sample_slides_data):
        """Test regex pattern matching in title filter."""
        query_engine._slide_cache = {"test.pptx:all_slides": sample_slides_data}
        
        filters = SlideQueryFilters(
            title=TitleFilter(regex=r"Project [AB]")
        )
        
        results = query_engine.query_slides("test.pptx", filters)
        
        assert len(results) == 2
        assert "Project A" in results[0].title or "Project B" in results[0].title
        assert "Project A" in results[1].title or "Project B" in results[1].title
    
    def test_object_count_filter(self, query_engine, sample_slides_data):
        """Test filtering by object count range."""
        query_engine._slide_cache = {"test.pptx:all_slides": sample_slides_data}
        
        filters = SlideQueryFilters(
            content=ContentFilter(object_count_min=3)  # Total objects >= 3
        )
        
        results = query_engine.query_slides("test.pptx", filters)
        
        # Slides 2 and 3 have 4 and 5 total objects respectively
        assert len(results) == 2
        assert results[0].slide_number == 2
        assert results[1].slide_number == 3
    
    def test_preview_text_generation(self, query_engine, sample_slides_data):
        """Test preview text generation."""
        query_engine._slide_cache = {"test.pptx:all_slides": sample_slides_data}
        
        filters = SlideQueryFilters()
        return_fields = ["slide_number", "preview_text"]
        
        results = query_engine.query_slides("test.pptx", filters, return_fields)
        
        assert len(results) == 3
        for result in results:
            assert result.preview_text is not None
            assert "Title:" in result.preview_text
    
    def test_empty_results(self, query_engine, sample_slides_data):
        """Test handling of queries that return no results."""
        query_engine._slide_cache = {"test.pptx:all_slides": sample_slides_data}
        
        filters = SlideQueryFilters(
            title=TitleFilter(contains="NonexistentTitle")
        )
        
        results = query_engine.query_slides("test.pptx", filters)
        
        assert len(results) == 0
    
    def test_cache_clearing(self, query_engine):
        """Test cache clearing functionality."""
        query_engine._slide_cache = {"test_key": "test_value"}
        
        query_engine.clear_cache()
        
        assert len(query_engine._slide_cache) == 0


class TestFilterCreation:
    """Test cases for filter creation utilities."""
    
    def test_create_filters_from_dict_title(self):
        """Test creating title filters from dictionary."""
        filters_dict = {
            "title": {
                "contains": "Progress",
                "starts_with": "Project",
                "one_of": [".*A.*", ".*B.*"]
            }
        }
        
        filters = create_filters_from_dict(filters_dict)
        
        assert filters.title is not None
        assert filters.title.contains == "Progress"
        assert filters.title.starts_with == "Project"
        assert filters.title.one_of == [".*A.*", ".*B.*"]
    
    def test_create_filters_from_dict_content(self):
        """Test creating content filters from dictionary."""
        filters_dict = {
            "content": {
                "has_tables": True,
                "has_images": False,
                "contains_text": "sample text",
                "object_count": {"min": 2, "max": 10}
            }
        }
        
        filters = create_filters_from_dict(filters_dict)
        
        assert filters.content is not None
        assert filters.content.has_tables is True
        assert filters.content.has_images is False
        assert filters.content.contains_text == "sample text"
        assert filters.content.object_count_min == 2
        assert filters.content.object_count_max == 10
    
    def test_create_filters_from_dict_layout(self):
        """Test creating layout filters from dictionary."""
        filters_dict = {
            "layout": {
                "type": "title_content",
                "name": "Title and Content"
            }
        }
        
        filters = create_filters_from_dict(filters_dict)
        
        assert filters.layout is not None
        assert filters.layout.layout_type == "title_content"
        assert filters.layout.layout_name == "Title and Content"
    
    def test_create_filters_from_dict_complete(self):
        """Test creating complete filters from dictionary."""
        filters_dict = {
            "title": {"contains": "Progress"},
            "content": {"has_tables": True},
            "layout": {"type": "title_content"},
            "slide_numbers": [1, 2, 3],
            "section": "Main Section"
        }
        
        filters = create_filters_from_dict(filters_dict)
        
        assert filters.title.contains == "Progress"
        assert filters.content.has_tables is True
        assert filters.layout.layout_type == "title_content"
        assert filters.slide_numbers == [1, 2, 3]
        assert filters.section == "Main Section"


if __name__ == "__main__":
    pytest.main([__file__])