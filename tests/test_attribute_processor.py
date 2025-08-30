"""
Unit tests for AttributeProcessor class.

Tests attribute filtering functionality for PowerPoint content,
including validation and selective data extraction.
"""

import pytest
from powerpoint_mcp_server.core.attribute_processor import AttributeProcessor


class TestAttributeProcessor:
    """Test cases for AttributeProcessor class."""
    
    def setup_method(self):
        """Set up test fixtures."""
        self.processor = AttributeProcessor()
        
        # Sample slide data for testing
        self.sample_slide_data = {
            'slide_number': 1,
            'title': 'Sample Title',
            'subtitle': 'Sample Subtitle',
            'layout_name': 'Title and Content',
            'layout_type': 'titleAndContent',
            'text_elements': [
                {
                    'content_plain': 'Sample text',
                    'content_formatted': '<b>Sample text</b>',
                    'font_sizes': [12],
                    'font_colors': ['#000000']
                }
            ],
            'tables': [
                {
                    'rows': 2,
                    'columns': 2,
                    'cells': [
                        [{'content': 'Cell 1,1'}, {'content': 'Cell 1,2'}],
                        [{'content': 'Cell 2,1'}, {'content': 'Cell 2,2'}]
                    ]
                }
            ],
            'placeholders': [
                {
                    'type': 'title',
                    'position': (100, 200),
                    'size': (800, 150)
                }
            ],
            'notes': 'Speaker notes content',
            'object_counts': {
                'shapes': 3,
                'text_boxes': 2,
                'images': 1,
                'tables': 1,
                'charts': 0,
                'media': 0,
                'connectors': 0,
                'groups': 0
            }
        }
        
        # Sample presentation data
        self.sample_presentation_data = {
            'slide_count': 2,
            'slide_size': {'width': 9144000, 'height': 6858000},
            'sections': [{'name': 'Introduction', 'id': '1'}],
            'metadata': {'author': 'Test Author'},
            'slides': [
                self.sample_slide_data,
                {
                    'slide_number': 2,
                    'title': 'Second Slide',
                    'text_elements': [],
                    'tables': [],
                    'object_counts': {'shapes': 1, 'text_boxes': 1, 'images': 0, 'tables': 0}
                }
            ]
        }
    
    def test_init(self):
        """Test AttributeProcessor initialization."""
        processor = AttributeProcessor()
        assert hasattr(processor, 'VALID_ATTRIBUTES')
        assert 'title' in processor.VALID_ATTRIBUTES
        assert 'subtitle' in processor.VALID_ATTRIBUTES
        assert 'text' in processor.VALID_ATTRIBUTES
    
    def test_get_available_attributes(self):
        """Test getting list of available attributes."""
        attributes = self.processor.get_available_attributes()
        
        assert isinstance(attributes, list)
        assert 'title' in attributes
        assert 'subtitle' in attributes
        assert 'text' in attributes
        assert 'tables' in attributes
        assert 'images' in attributes
        assert 'layout' in attributes
        assert 'size' in attributes
        assert 'sections' in attributes
        assert 'notes' in attributes
        assert 'object_counts' in attributes
        assert attributes == sorted(attributes)  # Should be sorted
    
    def test_validate_attributes_valid(self):
        """Test validation with valid attributes."""
        valid_attrs = ['title', 'subtitle', 'text']
        invalid = self.processor._validate_attributes(valid_attrs)
        
        assert invalid == []
    
    def test_validate_attributes_invalid(self):
        """Test validation with invalid attributes."""
        invalid_attrs = ['title', 'invalid_attr', 'another_invalid']
        invalid = self.processor._validate_attributes(invalid_attrs)
        
        assert 'invalid_attr' in invalid
        assert 'another_invalid' in invalid
        assert 'title' not in invalid
    
    def test_filter_attributes_title_only(self):
        """Test filtering to include only title."""
        result = self.processor.filter_attributes(
            self.sample_slide_data, 
            ['title']
        )
        
        assert 'title' in result
        assert result['title'] == 'Sample Title'
        assert 'subtitle' not in result
        assert 'text_elements' not in result
        assert 'slide_number' in result  # Always included
    
    def test_filter_attributes_multiple(self):
        """Test filtering with multiple attributes."""
        result = self.processor.filter_attributes(
            self.sample_slide_data,
            ['title', 'subtitle', 'notes']
        )
        
        assert 'title' in result
        assert 'subtitle' in result
        assert 'notes' in result
        assert result['title'] == 'Sample Title'
        assert result['subtitle'] == 'Sample Subtitle'
        assert result['notes'] == 'Speaker notes content'
        assert 'text_elements' not in result
        assert 'tables' not in result
    
    def test_filter_attributes_text(self):
        """Test filtering text attributes."""
        result = self.processor.filter_attributes(
            self.sample_slide_data,
            ['text']
        )
        
        assert 'text_elements' in result
        assert len(result['text_elements']) == 1
        assert result['text_elements'][0]['content_plain'] == 'Sample text'
    
    def test_filter_attributes_tables(self):
        """Test filtering table attributes."""
        result = self.processor.filter_attributes(
            self.sample_slide_data,
            ['tables']
        )
        
        assert 'tables' in result
        assert len(result['tables']) == 1
        assert result['tables'][0]['rows'] == 2
        assert result['tables'][0]['columns'] == 2
    
    def test_filter_attributes_layout(self):
        """Test filtering layout attributes."""
        result = self.processor.filter_attributes(
            self.sample_slide_data,
            ['layout']
        )
        
        assert 'layout_name' in result
        assert 'layout_type' in result
        assert 'placeholders' in result
        assert result['layout_name'] == 'Title and Content'
        assert result['layout_type'] == 'titleAndContent'
    
    def test_filter_attributes_object_counts(self):
        """Test filtering object counts."""
        result = self.processor.filter_attributes(
            self.sample_slide_data,
            ['object_counts']
        )
        
        assert 'object_counts' in result
        assert result['object_counts']['shapes'] == 3
        assert result['object_counts']['text_boxes'] == 2
        assert result['object_counts']['images'] == 1
    
    def test_filter_attributes_images(self):
        """Test filtering image attributes."""
        result = self.processor.filter_attributes(
            self.sample_slide_data,
            ['images']
        )
        
        # Should include image count from object_counts
        assert 'object_counts' in result
        assert result['object_counts']['images'] == 1
    
    def test_filter_attributes_presentation_level(self):
        """Test filtering presentation-level attributes."""
        result = self.processor.filter_attributes(
            self.sample_presentation_data,
            ['slide_count', 'slide_size', 'sections']
        )
        
        assert 'slide_count' in result
        assert 'slide_size' in result
        assert 'sections' in result
        assert result['slide_count'] == 2
        assert result['slide_size']['width'] == 9144000
        assert len(result['sections']) == 1
    
    def test_filter_attributes_with_slides(self):
        """Test filtering attributes for presentation with slides."""
        result = self.processor.filter_attributes(
            self.sample_presentation_data,
            ['title', 'object_counts']
        )
        
        assert 'slides' in result
        assert len(result['slides']) == 2
        
        # Check first slide
        slide1 = result['slides'][0]
        assert 'title' in slide1
        assert 'object_counts' in slide1
        assert slide1['title'] == 'Sample Title'
        assert 'subtitle' not in slide1  # Not requested
        
        # Check second slide
        slide2 = result['slides'][1]
        assert 'title' in slide2
        assert slide2['title'] == 'Second Slide'
    
    def test_filter_attributes_empty_list(self):
        """Test filtering with empty attribute list returns all data."""
        result = self.processor.filter_attributes(
            self.sample_slide_data,
            []
        )
        
        # Should return all original data
        assert result == self.sample_slide_data
    
    def test_filter_attributes_invalid_raises_error(self):
        """Test that invalid attributes raise ValueError."""
        with pytest.raises(ValueError) as exc_info:
            self.processor.filter_attributes(
                self.sample_slide_data,
                ['title', 'invalid_attribute']
            )
        
        assert 'Invalid attribute types' in str(exc_info.value)
        assert 'invalid_attribute' in str(exc_info.value)
    
    def test_process_slide_attributes(self):
        """Test processing slide attributes with additional processing."""
        result = self.processor.process_slide_attributes(
            self.sample_slide_data,
            ['title', 'object_counts']
        )
        
        assert 'title' in result
        assert 'object_counts' in result
        assert result['title'] == 'Sample Title'
    
    def test_compute_object_counts(self):
        """Test computing object counts from slide data."""
        slide_data = {
            'text_elements': [{'content': 'text1'}, {'content': 'text2'}],
            'tables': [{'rows': 2}],
            'placeholders': [{'type': 'title'}]
        }
        
        counts = self.processor._compute_object_counts(slide_data)
        
        assert counts['text_boxes'] == 2
        assert counts['tables'] == 1
        assert counts['shapes'] == 1  # From placeholders
    
    def test_create_attribute_summary(self):
        """Test creating attribute summary across slides."""
        result = self.processor.create_attribute_summary(
            self.sample_presentation_data,
            ['object_counts', 'text', 'tables']
        )
        
        assert 'requested_attributes' in result
        assert 'total_slides' in result
        assert 'summary' in result
        
        assert result['total_slides'] == 2
        assert result['requested_attributes'] == ['object_counts', 'text', 'tables']
        
        # Check aggregated counts
        summary = result['summary']
        assert 'total_objects' in summary
        assert summary['total_objects']['shapes'] == 4  # 3 + 1 from both slides
        assert summary['total_objects']['text_boxes'] == 3  # 2 + 1
        assert summary['total_objects']['tables'] == 1  # 1 + 0
        
        assert 'total_text_elements' in summary
        assert summary['total_text_elements'] == 1  # 1 + 0
        
        assert 'total_tables' in summary
        assert summary['total_tables'] == 1  # 1 + 0
    
    def test_filter_slide_attributes_direct(self):
        """Test filtering slide attributes directly."""
        attr_set = {'title', 'notes', 'layout'}
        
        result = self.processor.filter_slide_attributes(
            self.sample_slide_data,
            attr_set
        )
        
        assert 'title' in result
        assert 'notes' in result
        assert 'layout_name' in result
        assert 'layout_type' in result
        assert 'placeholders' in result
        assert 'subtitle' not in result
        assert 'text_elements' not in result
    
    def test_filter_attributes_size(self):
        """Test filtering size-related attributes."""
        slide_data_with_size = {
            **self.sample_slide_data,
            'slide_size': {'width': 1000, 'height': 750},
            'position': (100, 200),
            'size': (800, 600)
        }
        
        result = self.processor.filter_attributes(
            slide_data_with_size,
            ['size']
        )
        
        assert 'slide_size' in result
        assert 'position' in result
        assert 'size' in result
        assert result['slide_size']['width'] == 1000
    
    def test_filter_attributes_placeholders(self):
        """Test filtering placeholder attributes specifically."""
        result = self.processor.filter_attributes(
            self.sample_slide_data,
            ['placeholders']
        )
        
        assert 'placeholders' in result
        assert len(result['placeholders']) == 1
        assert result['placeholders'][0]['type'] == 'title'