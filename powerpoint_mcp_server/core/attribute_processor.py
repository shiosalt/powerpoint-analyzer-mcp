"""
Attribute Processor for PowerPoint content filtering.

This module provides attribute filtering functionality to selectively
extract only requested attributes from PowerPoint content.
"""

from typing import Dict, List, Optional, Any, Set
import logging

logger = logging.getLogger(__name__)


class AttributeProcessor:
    """
    Processor for filtering PowerPoint content by requested attributes.
    
    Supports filtering by title, subtitle, text, tables, images, layout,
    size, sections, notes, and object counts.
    """
    
    # Valid attribute types that can be requested
    VALID_ATTRIBUTES = {
        'title', 'subtitle', 'text', 'tables', 'images', 'layout', 
        'size', 'sections', 'notes', 'object_counts', 'placeholders',
        'text_elements', 'metadata', 'slide_count', 'slide_size'
    }
    
    def __init__(self):
        """Initialize the attribute processor."""
        pass
    
    def filter_attributes(self, data: Dict[str, Any], requested_attributes: List[str]) -> Dict[str, Any]:
        """
        Filter data to include only requested attributes.
        
        Args:
            data: Complete data dictionary
            requested_attributes: List of attribute names to include
            
        Returns:
            Filtered data dictionary containing only requested attributes
            
        Raises:
            ValueError: If invalid attribute types are specified
        """
        try:
            # Validate requested attributes
            invalid_attrs = self._validate_attributes(requested_attributes)
            if invalid_attrs:
                raise ValueError(f"Invalid attribute types: {invalid_attrs}. Valid options: {sorted(self.VALID_ATTRIBUTES)}")
            
            # If no attributes specified, return all data
            if not requested_attributes:
                return data
            
            # Convert to set for faster lookup
            attr_set = set(requested_attributes)
            
            # Filter the data
            filtered_data = {}
            
            # Handle presentation-level data
            if 'slide_count' in attr_set and 'slide_count' in data:
                filtered_data['slide_count'] = data['slide_count']
            
            if 'slide_size' in attr_set and 'slide_size' in data:
                filtered_data['slide_size'] = data['slide_size']
            
            if 'sections' in attr_set and 'sections' in data:
                filtered_data['sections'] = data['sections']
            
            if 'metadata' in attr_set and 'metadata' in data:
                filtered_data['metadata'] = data['metadata']
            
            # Handle slides data
            if 'slides' in data:
                filtered_slides = []
                for slide in data['slides']:
                    filtered_slide = self.filter_slide_attributes(slide, attr_set)
                    filtered_slides.append(filtered_slide)
                filtered_data['slides'] = filtered_slides
            
            # Handle single slide data (when processing individual slides)
            elif any(attr in data for attr in ['title', 'subtitle', 'text', 'tables', 'images', 'layout', 'notes', 'object_counts']):
                filtered_data.update(self.filter_slide_attributes(data, attr_set))
            
            return filtered_data
            
        except Exception as e:
            logger.error(f"Failed to filter attributes: {e}")
            raise
    
    def filter_slide_attributes(self, slide_data: Dict[str, Any], requested_attributes: Set[str]) -> Dict[str, Any]:
        """
        Filter attributes for a single slide.
        
        Args:
            slide_data: Complete slide data dictionary
            requested_attributes: Set of attribute names to include
            
        Returns:
            Filtered slide data dictionary
        """
        try:
            filtered_slide = {}
            
            # Always include slide number if present
            if 'slide_number' in slide_data:
                filtered_slide['slide_number'] = slide_data['slide_number']
            
            # Filter specific attributes
            if 'title' in requested_attributes and 'title' in slide_data:
                filtered_slide['title'] = slide_data['title']
            
            if 'subtitle' in requested_attributes and 'subtitle' in slide_data:
                filtered_slide['subtitle'] = slide_data['subtitle']
            
            if 'text' in requested_attributes:
                # Include text-related data
                if 'text_elements' in slide_data:
                    filtered_slide['text_elements'] = slide_data['text_elements']
                if 'content_plain' in slide_data:
                    filtered_slide['content_plain'] = slide_data['content_plain']
                if 'content_formatted' in slide_data:
                    filtered_slide['content_formatted'] = slide_data['content_formatted']
            
            if 'text_elements' in requested_attributes and 'text_elements' in slide_data:
                filtered_slide['text_elements'] = slide_data['text_elements']
            
            if 'tables' in requested_attributes and 'tables' in slide_data:
                filtered_slide['tables'] = slide_data['tables']
            
            if 'images' in requested_attributes:
                # Include image-related data (always include the key, even if empty)
                filtered_slide['images'] = slide_data.get('images', [])
                # Images are counted in object_counts, so include that info
                if 'object_counts' in slide_data and 'images' in slide_data['object_counts']:
                    if 'object_counts' not in filtered_slide:
                        filtered_slide['object_counts'] = {}
                    filtered_slide['object_counts']['images'] = slide_data['object_counts']['images']
            
            if 'layout' in requested_attributes:
                # Include layout-related data
                if 'layout_name' in slide_data:
                    filtered_slide['layout_name'] = slide_data['layout_name']
                if 'layout_type' in slide_data:
                    filtered_slide['layout_type'] = slide_data['layout_type']
                if 'placeholders' in slide_data:
                    filtered_slide['placeholders'] = slide_data['placeholders']
            
            if 'placeholders' in requested_attributes and 'placeholders' in slide_data:
                filtered_slide['placeholders'] = slide_data['placeholders']
            
            if 'size' in requested_attributes:
                # Include size-related data
                if 'slide_size' in slide_data:
                    filtered_slide['slide_size'] = slide_data['slide_size']
                # Include position and size info from elements
                for key in ['position', 'size']:
                    if key in slide_data:
                        filtered_slide[key] = slide_data[key]
            
            if 'notes' in requested_attributes and 'notes' in slide_data:
                filtered_slide['notes'] = slide_data['notes']
            
            if 'object_counts' in requested_attributes and 'object_counts' in slide_data:
                filtered_slide['object_counts'] = slide_data['object_counts']
            
            return filtered_slide
            
        except Exception as e:
            logger.warning(f"Failed to filter slide attributes: {e}")
            return slide_data
    
    def _validate_attributes(self, requested_attributes: List[str]) -> List[str]:
        """
        Validate that all requested attributes are valid.
        
        Args:
            requested_attributes: List of attribute names to validate
            
        Returns:
            List of invalid attribute names
        """
        invalid_attrs = []
        for attr in requested_attributes:
            if attr not in self.VALID_ATTRIBUTES:
                invalid_attrs.append(attr)
        return invalid_attrs
    
    def get_available_attributes(self) -> List[str]:
        """
        Get list of all available attribute types.
        
        Returns:
            Sorted list of valid attribute names
        """
        return sorted(self.VALID_ATTRIBUTES)
    
    def process_slide_attributes(self, slide_data: Dict[str, Any], attributes: List[str]) -> Dict[str, Any]:
        """
        Process and filter attributes for a single slide with additional processing.
        
        Args:
            slide_data: Complete slide data dictionary
            attributes: List of attribute names to include
            
        Returns:
            Processed and filtered slide data
        """
        try:
            # First filter the attributes
            filtered_data = self.filter_attributes(slide_data, attributes)
            
            # Add computed attributes if requested
            if 'object_counts' in attributes and 'object_counts' not in filtered_data:
                filtered_data['object_counts'] = self._compute_object_counts(slide_data)
            
            return filtered_data
            
        except Exception as e:
            logger.error(f"Failed to process slide attributes: {e}")
            return slide_data
    
    def _compute_object_counts(self, slide_data: Dict[str, Any]) -> Dict[str, int]:
        """
        Compute object counts from slide data if not already present.
        
        Args:
            slide_data: Slide data dictionary
            
        Returns:
            Dictionary with object counts
        """
        try:
            counts = {
                'shapes': 0,
                'text_boxes': 0,
                'images': 0,
                'tables': 0,
                'charts': 0,
                'media': 0,
                'connectors': 0,
                'groups': 0
            }
            
            # Count from available data
            if 'text_elements' in slide_data:
                counts['text_boxes'] = len(slide_data['text_elements'])
            
            if 'tables' in slide_data:
                counts['tables'] = len(slide_data['tables'])
            
            if 'placeholders' in slide_data:
                counts['shapes'] += len(slide_data['placeholders'])
            
            return counts
            
        except Exception as e:
            logger.warning(f"Failed to compute object counts: {e}")
            return {}
    
    def create_attribute_summary(self, data: Dict[str, Any], attributes: List[str]) -> Dict[str, Any]:
        """
        Create a summary of requested attributes across all slides.
        
        Args:
            data: Complete presentation data
            attributes: List of attribute names to summarize
            
        Returns:
            Summary dictionary with aggregated attribute information
        """
        try:
            summary = {
                'requested_attributes': attributes,
                'total_slides': 0,
                'summary': {}
            }
            
            if 'slides' in data:
                summary['total_slides'] = len(data['slides'])
                
                # Initialize counters
                if 'object_counts' in attributes:
                    summary['summary']['total_objects'] = {
                        'shapes': 0,
                        'text_boxes': 0,
                        'images': 0,
                        'tables': 0,
                        'charts': 0,
                        'media': 0,
                        'connectors': 0,
                        'groups': 0
                    }
                
                if 'text' in attributes:
                    summary['summary']['total_text_elements'] = 0
                
                if 'tables' in attributes:
                    summary['summary']['total_tables'] = 0
                
                # Aggregate data from all slides
                for slide in data['slides']:
                    if 'object_counts' in attributes and 'object_counts' in slide:
                        for obj_type, count in slide['object_counts'].items():
                            if obj_type in summary['summary']['total_objects']:
                                summary['summary']['total_objects'][obj_type] += count
                    
                    if 'text' in attributes and 'text_elements' in slide:
                        summary['summary']['total_text_elements'] += len(slide['text_elements'])
                    
                    if 'tables' in attributes and 'tables' in slide:
                        summary['summary']['total_tables'] += len(slide['tables'])
            
            return summary
            
        except Exception as e:
            logger.error(f"Failed to create attribute summary: {e}")
            return {'error': str(e)}