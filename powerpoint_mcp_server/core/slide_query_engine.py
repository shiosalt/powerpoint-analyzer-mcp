"""
Flexible slide query engine for complex filtering and search operations.
"""

import re
import logging
from typing import Dict, List, Any, Optional, Union, Tuple
from dataclasses import dataclass
from enum import Enum

from .content_extractor import ContentExtractor
from ..utils.zip_extractor import ZipExtractor

logger = logging.getLogger(__name__)


class MatchCondition(Enum):
    """Enumeration of available matching conditions."""
    EQUALS = "equals"
    CONTAINS = "contains"
    STARTS_WITH = "starts_with"
    ENDS_WITH = "ends_with"
    REGEX = "regex"
    ONE_OF = "one_of"


@dataclass
class TitleFilter:
    """Filter configuration for slide titles."""
    contains: Optional[str] = None
    starts_with: Optional[str] = None
    ends_with: Optional[str] = None
    regex: Optional[str] = None
    one_of: Optional[List[str]] = None


@dataclass
class ContentFilter:
    """Filter configuration for slide content."""
    contains_text: Optional[str] = None
    has_tables: Optional[bool] = None
    has_charts: Optional[bool] = None
    has_images: Optional[bool] = None
    object_count_min: Optional[int] = None
    object_count_max: Optional[int] = None


@dataclass
class NotesFilter:
    """Filter configuration for speaker notes."""
    contains: Optional[str] = None
    regex: Optional[str] = None
    is_empty: Optional[bool] = None


@dataclass
class LayoutFilter:
    """Filter configuration for slide layout."""
    layout_type: Optional[str] = None
    layout_name: Optional[str] = None


@dataclass
class SlideQueryFilters:
    """Complete filter configuration for slide queries."""
    title: Optional[TitleFilter] = None
    content: Optional[ContentFilter] = None
    layout: Optional[LayoutFilter] = None
    slide_numbers: Optional[List[int]] = None
    section: Optional[str] = None
    notes: Optional[NotesFilter] = None


@dataclass
class SlideQueryResult:
    """Result of a slide query operation."""
    slide_number: int
    title: Optional[str] = None
    subtitle: Optional[str] = None
    layout_name: Optional[str] = None
    layout_type: Optional[str] = None
    object_counts: Optional[Dict[str, int]] = None
    preview_text: Optional[str] = None
    table_info: Optional[List[Dict[str, Any]]] = None
    full_content: Optional[Dict[str, Any]] = None


class SlideQueryEngine:
    """
    Engine for performing flexible slide queries with complex filtering capabilities.
    """
    
    def __init__(self, content_extractor: Optional[ContentExtractor] = None):
        """Initialize the slide query engine."""
        self.content_extractor = content_extractor or ContentExtractor()
        self._slide_cache = {}
        
    def query_slides(
        self,
        file_path: str,
        filters: SlideQueryFilters,
        return_fields: List[str] = None,
        limit: int = 50
    ) -> List[SlideQueryResult]:
        """
        Query slides based on flexible filtering criteria.
        
        Args:
            file_path: Path to the PowerPoint file
            filters: Filter configuration
            return_fields: Fields to include in results
            limit: Maximum number of results to return
            
        Returns:
            List of matching slides with requested fields
        """
        if return_fields is None:
            return_fields = ["slide_number", "title", "object_counts"]
            
        logger.info(f"Querying slides in {file_path} with filters: {filters}")
        
        try:
            # Validate search criteria upfront with comprehensive grammar checking
            validation_result = self._validate_search_criteria(filters, return_fields)
            if not validation_result['is_valid']:
                logger.warning(f"Invalid search criteria: {validation_result['errors']}")
                # Return empty results with error details for invalid criteria
                return []
            
            # Extract all slides if not cached
            cache_key = f"{file_path}:all_slides"
            if cache_key not in self._slide_cache:
                self._slide_cache[cache_key] = self._extract_all_slides(file_path)
            
            all_slides = self._slide_cache[cache_key]
            
            # Apply filters
            filtered_slides = self._apply_filters(all_slides, filters)
            
            # Limit results
            if limit > 0:
                filtered_slides = filtered_slides[:limit]
            
            # Build results with requested fields
            results = []
            for slide_data in filtered_slides:
                result = self._build_slide_result(slide_data, return_fields)
                results.append(result)
            
            logger.info(f"Query returned {len(results)} slides")
            return results
            
        except Exception as e:
            logger.error(f"Error querying slides: {e}")
            raise
    
    def _extract_all_slides(self, file_path: str) -> List[Dict[str, Any]]:
        """Extract basic information from all slides."""
        slides = []
        
        with ZipExtractor(file_path) as extractor:
            # Get presentation metadata including sections
            presentation_xml = extractor.read_xml_content('ppt/presentation.xml')
            presentation_metadata = {}
            if presentation_xml:
                presentation_metadata = self.content_extractor.extract_presentation_metadata(presentation_xml)
                # Extract section information
                sections = self.content_extractor.extract_section_information(presentation_xml)
                presentation_metadata['sections'] = sections
                logger.debug(f"Extracted {len(sections)} sections from presentation")
            
            # Get slide XML files
            slide_files = extractor.get_slide_xml_files()
            
            for i, slide_file in enumerate(slide_files, 1):
                slide_xml = extractor.read_xml_content(slide_file)
                if slide_xml:
                    # Extract slide content
                    slide_info = self.content_extractor.extract_slide_content(slide_xml, i)
                    
                    # Extract notes if available
                    notes_file = f'ppt/notesSlides/notesSlide{i}.xml'
                    notes_content = ""
                    try:
                        notes_xml = extractor.read_xml_content(notes_file)
                        if notes_xml:
                            notes_content = self.content_extractor.extract_slide_notes(notes_xml)
                    except Exception as e:
                        logger.debug(f"No notes found for slide {i}: {e}")
                    
                    # Get object counts
                    object_counts = self.content_extractor._count_slide_objects(
                        self.content_extractor.xml_parser.parse_xml_string(slide_xml)
                    )
                    
                    # Create slide data
                    slide_data = {
                        'slide_number': i,
                        'title': slide_info.title,
                        'subtitle': slide_info.subtitle,
                        'layout_name': slide_info.layout_name,
                        'layout_type': slide_info.layout_type,
                        'placeholders': slide_info.placeholders,
                        'text_elements': slide_info.text_elements,
                        'tables': slide_info.tables,
                        'notes': notes_content,
                        'section_name': slide_info.section_name,
                        'object_counts': object_counts,
                        'slide_xml': slide_xml,  # Keep for detailed analysis
                        'presentation_metadata': presentation_metadata
                    }
                    
                    slides.append(slide_data)
        
        return slides
    
    def _apply_filters(
        self, 
        slides: List[Dict[str, Any]], 
        filters: SlideQueryFilters
    ) -> List[Dict[str, Any]]:
        """Apply all filters to the slide list."""
        filtered_slides = slides.copy()
        
        # Apply slide number filter
        if filters.slide_numbers:
            filtered_slides = [
                slide for slide in filtered_slides 
                if slide['slide_number'] in filters.slide_numbers
            ]
        
        # Apply title filters
        if filters.title:
            filtered_slides = self._apply_title_filters(filtered_slides, filters.title)
        
        # Apply content filters
        if filters.content:
            filtered_slides = self._apply_content_filters(filtered_slides, filters.content)
        
        # Apply layout filters
        if filters.layout:
            filtered_slides = self._apply_layout_filters(filtered_slides, filters.layout)
        
        # Apply section filter
        if filters.section:
            filtered_slides = self._apply_section_filter(filtered_slides, filters.section)
        
        # Apply notes filter
        if filters.notes:
            filtered_slides = self._apply_notes_filter(filtered_slides, filters.notes)
        
        return filtered_slides
    
    def _apply_title_filters(
        self, 
        slides: List[Dict[str, Any]], 
        title_filter: TitleFilter
    ) -> List[Dict[str, Any]]:
        """Apply title-based filters."""
        filtered_slides = []
        
        for slide in slides:
            title = slide.get('title', '') or ''
            
            # Check each title condition
            if self._matches_title_condition(title, title_filter):
                filtered_slides.append(slide)
        
        return filtered_slides
    
    def _matches_title_condition(self, title: str, title_filter: TitleFilter) -> bool:
        """Check if title matches any of the specified conditions."""
        if not title:
            title = ''
        
        # Contains condition
        if title_filter.contains:
            if title_filter.contains.lower() not in title.lower():
                return False
        
        # Starts with condition
        if title_filter.starts_with:
            if not title.lower().startswith(title_filter.starts_with.lower()):
                return False
        
        # Ends with condition
        if title_filter.ends_with:
            if not title.lower().endswith(title_filter.ends_with.lower()):
                return False
        
        # Regex condition
        if title_filter.regex:
            try:
                if not re.search(title_filter.regex, title, re.IGNORECASE):
                    return False
            except re.error as e:
                logger.warning(f"Invalid regex pattern '{title_filter.regex}': {e}")
                return False
        
        # One of condition (OR logic)
        if title_filter.one_of:
            match_found = False
            for pattern in title_filter.one_of:
                try:
                    if re.search(pattern, title, re.IGNORECASE):
                        match_found = True
                        break
                except re.error:
                    # Fallback to simple string matching
                    if pattern.lower() in title.lower():
                        match_found = True
                        break
            
            if not match_found:
                return False
        
        return True
    
    def _apply_content_filters(
        self, 
        slides: List[Dict[str, Any]], 
        content_filter: ContentFilter
    ) -> List[Dict[str, Any]]:
        """Apply content-based filters."""
        filtered_slides = []
        
        for slide in slides:
            if self._matches_content_condition(slide, content_filter):
                filtered_slides.append(slide)
        
        return filtered_slides
    
    def _matches_content_condition(self, slide: Dict[str, Any], content_filter: ContentFilter) -> bool:
        """Check if slide content matches the specified conditions."""
        
        # Has tables condition
        if content_filter.has_tables is not None:
            has_tables = len(slide.get('tables', [])) > 0
            if content_filter.has_tables != has_tables:
                return False
        
        # Has charts condition
        if content_filter.has_charts is not None:
            object_counts = slide.get('object_counts', {})
            has_charts = object_counts.get('charts', 0) > 0
            if content_filter.has_charts != has_charts:
                return False
        
        # Has images condition
        if content_filter.has_images is not None:
            object_counts = slide.get('object_counts', {})
            has_images = object_counts.get('images', 0) > 0
            if content_filter.has_images != has_images:
                return False
        
        # Object count conditions
        if content_filter.object_count_min is not None or content_filter.object_count_max is not None:
            object_counts = slide.get('object_counts', {})
            total_objects = sum(object_counts.values())
            
            if content_filter.object_count_min is not None:
                if total_objects < content_filter.object_count_min:
                    return False
            
            if content_filter.object_count_max is not None:
                if total_objects > content_filter.object_count_max:
                    return False
        
        # Contains text condition
        if content_filter.contains_text:
            text_found = False
            
            # Check title
            title = slide.get('title', '') or ''
            if content_filter.contains_text.lower() in title.lower():
                text_found = True
            
            # Check text elements
            if not text_found:
                text_elements = slide.get('text_elements', [])
                for text_elem in text_elements:
                    if isinstance(text_elem, dict):
                        content = text_elem.get('content_plain', '') or ''
                        if content_filter.contains_text.lower() in content.lower():
                            text_found = True
                            break
            
            if not text_found:
                return False
        
        return True
    
    def _apply_layout_filters(
        self, 
        slides: List[Dict[str, Any]], 
        layout_filter: LayoutFilter
    ) -> List[Dict[str, Any]]:
        """Apply layout-based filters."""
        filtered_slides = []
        
        for slide in slides:
            if self._matches_layout_condition(slide, layout_filter):
                filtered_slides.append(slide)
        
        return filtered_slides
    
    def _matches_layout_condition(self, slide: Dict[str, Any], layout_filter: LayoutFilter) -> bool:
        """Check if slide layout matches the specified conditions."""
        
        # Layout type condition
        if layout_filter.layout_type:
            layout_type = slide.get('layout_type', '') or ''
            if layout_filter.layout_type.lower() not in layout_type.lower():
                return False
        
        # Layout name condition
        if layout_filter.layout_name:
            layout_name = slide.get('layout_name', '') or ''
            if layout_filter.layout_name.lower() not in layout_name.lower():
                return False
        
        return True
    
    def _apply_section_filter(
        self, 
        slides: List[Dict[str, Any]], 
        section_name: str
    ) -> List[Dict[str, Any]]:
        """Apply section-based filtering."""
        try:
            filtered_slides = []
            
            # Extract section information from presentation metadata
            if not slides:
                return filtered_slides
            
            # Get presentation metadata from the first slide
            presentation_metadata = slides[0].get('presentation_metadata', {})
            sections = presentation_metadata.get('sections', [])
            
            logger.debug(f"Found {len(sections)} sections in presentation")
            
            # If no sections are defined, return empty results
            if not sections:
                logger.info(f"No sections found in presentation, section filter '{section_name}' returns no results")
                return []
            
            # Find the requested section
            target_section = None
            for section in sections:
                if isinstance(section, dict):
                    if section.get('name', '').lower() == section_name.lower():
                        target_section = section
                        break
            
            if target_section is None:
                logger.info(f"Section '{section_name}' not found in presentation")
                return []
            
            # Get slide range for the section
            slide_range = target_section.get('slide_range', [])
            if len(slide_range) >= 2:
                start_slide, end_slide = slide_range[0], slide_range[1]
                
                # Filter slides within the section range
                for slide in slides:
                    slide_number = slide.get('slide_number', 0)
                    if start_slide <= slide_number <= end_slide:
                        filtered_slides.append(slide)
                
                logger.info(f"Section '{section_name}' contains slides {start_slide}-{end_slide}, found {len(filtered_slides)} matching slides")
            else:
                logger.warning(f"Invalid slide range for section '{section_name}': {slide_range}")
            
            return filtered_slides
            
        except Exception as e:
            logger.error(f"Error applying section filter: {e}")
            return []  # Return empty results on error
    
    def _apply_notes_filter(
        self, 
        slides: List[Dict[str, Any]], 
        notes_filter: NotesFilter
    ) -> List[Dict[str, Any]]:
        """Apply notes-based filtering."""
        try:
            filtered_slides = []
            
            for slide in slides:
                notes = slide.get('notes', '') or ''
                
                if self._matches_notes_condition(notes, notes_filter):
                    filtered_slides.append(slide)
            
            logger.info(f"Notes filter matched {len(filtered_slides)} slides")
            return filtered_slides
            
        except Exception as e:
            logger.error(f"Error applying notes filter: {e}")
            return []  # Return empty results on error
    
    def _matches_notes_condition(self, notes: str, notes_filter: NotesFilter) -> bool:
        """Check if notes match the specified conditions."""
        try:
            # Is empty condition
            if notes_filter.is_empty is not None:
                is_empty = not notes.strip()
                if notes_filter.is_empty != is_empty:
                    return False
            
            # Contains condition
            if notes_filter.contains:
                if notes_filter.contains.lower() not in notes.lower():
                    return False
            
            # Regex condition
            if notes_filter.regex:
                try:
                    if not re.search(notes_filter.regex, notes, re.IGNORECASE):
                        return False
                except re.error as e:
                    logger.warning(f"Invalid regex pattern '{notes_filter.regex}': {e}")
                    return False
            
            return True
            
        except Exception as e:
            logger.warning(f"Error matching notes condition: {e}")
            return False
    
    def _build_slide_result(
        self, 
        slide_data: Dict[str, Any], 
        return_fields: List[str]
    ) -> SlideQueryResult:
        """Build a slide result with only the requested fields."""
        result_data = {}
        
        # Always include slide number
        result_data['slide_number'] = slide_data['slide_number']
        
        # Add requested fields
        for field in return_fields:
            if field == 'slide_number':
                continue  # Already added
            elif field == 'title':
                result_data['title'] = slide_data.get('title')
            elif field == 'subtitle':
                result_data['subtitle'] = slide_data.get('subtitle')
            elif field == 'layout':
                result_data['layout_name'] = slide_data.get('layout_name')
                result_data['layout_type'] = slide_data.get('layout_type')
            elif field == 'object_counts':
                result_data['object_counts'] = slide_data.get('object_counts')
            elif field == 'preview_text':
                result_data['preview_text'] = self._generate_preview_text(slide_data)
            elif field == 'table_info':
                result_data['table_info'] = self._generate_table_info(slide_data)
            elif field == 'full_content':
                result_data['full_content'] = self._generate_full_content(slide_data)
        
        return SlideQueryResult(**result_data)
    
    def _generate_preview_text(self, slide_data: Dict[str, Any]) -> str:
        """Generate preview text from slide content."""
        preview_parts = []
        
        # Add title
        if slide_data.get('title'):
            preview_parts.append(f"Title: {slide_data['title']}")
        
        # Add first few text elements
        text_elements = slide_data.get('text_elements', [])
        for i, text_elem in enumerate(text_elements[:3]):  # First 3 text elements
            if isinstance(text_elem, dict):
                content = text_elem.get('content_plain', '') or ''
                if content:
                    # Truncate long text
                    if len(content) > 100:
                        content = content[:97] + "..."
                    preview_parts.append(f"Text {i+1}: {content}")
        
        return " | ".join(preview_parts)
    
    def _generate_table_info(self, slide_data: Dict[str, Any]) -> List[Dict[str, Any]]:
        """Generate table information summary."""
        tables = slide_data.get('tables', [])
        table_info = []
        
        for i, table in enumerate(tables):
            if isinstance(table, dict):
                # Handle different table data structures
                rows_data = table.get('rows', [])
                if isinstance(rows_data, list) and rows_data:
                    row_count = len(rows_data)
                    # Get column count from first row
                    first_row = rows_data[0] if rows_data else []
                    if isinstance(first_row, list):
                        col_count = len(first_row)
                        headers = first_row
                    else:
                        col_count = 0
                        headers = []
                elif isinstance(rows_data, int):
                    # Handle case where 'rows' is just a count
                    row_count = rows_data
                    col_count = table.get('columns', 0)
                    headers = []
                else:
                    row_count = 0
                    col_count = 0
                    headers = []
                
                info = {
                    'table_index': i,
                    'rows': row_count,
                    'columns': col_count,
                    'headers': headers
                }
                table_info.append(info)
        
        return table_info
    
    def _generate_full_content(self, slide_data: Dict[str, Any]) -> Dict[str, Any]:
        """Generate full content representation."""
        return {
            'title': slide_data.get('title'),
            'subtitle': slide_data.get('subtitle'),
            'layout_name': slide_data.get('layout_name'),
            'layout_type': slide_data.get('layout_type'),
            'text_elements': slide_data.get('text_elements', []),
            'tables': slide_data.get('tables', []),
            'object_counts': slide_data.get('object_counts', {}),
            'placeholders': slide_data.get('placeholders', [])
        }
    
    def _validate_search_criteria(self, filters: SlideQueryFilters, return_fields: List[str]) -> Dict[str, Any]:
        """Validate search criteria before processing with comprehensive grammar checking."""
        try:
            validation_result = {
                'is_valid': True,
                'errors': [],
                'warnings': []
            }
            
            # Validate that filters is a proper SlideQueryFilters object
            if not isinstance(filters, SlideQueryFilters):
                validation_result['errors'].append("Invalid search criteria structure. Expected SlideQueryFilters object.")
                validation_result['is_valid'] = False
                return validation_result
            
            # Check for None filters (this can happen with malformed input)
            if filters is None:
                validation_result['errors'].append("Search criteria cannot be None")
                validation_result['is_valid'] = False
                return validation_result
            
            # Validate return fields
            valid_return_fields = [
                'slide_number', 'title', 'subtitle', 'layout', 'object_counts',
                'preview_text', 'table_info', 'full_content', 'notes', 'section'
            ]
            
            for field in return_fields:
                if field not in valid_return_fields:
                    validation_result['errors'].append(f"Invalid return field: '{field}'. Valid fields: {valid_return_fields}")
                    validation_result['is_valid'] = False
            
            # Validate title filters
            if filters.title:
                if filters.title.regex:
                    try:
                        re.compile(filters.title.regex)
                    except re.error as e:
                        validation_result['errors'].append(f"Invalid regex pattern in title filter: {e}")
                        validation_result['is_valid'] = False
                
                if filters.title.one_of:
                    for pattern in filters.title.one_of:
                        try:
                            re.compile(pattern)
                        except re.error:
                            # This is okay, we'll fall back to string matching
                            pass
            
            # Validate content filters
            if filters.content:
                if filters.content.object_count_min is not None and filters.content.object_count_min < 0:
                    validation_result['errors'].append("object_count_min must be non-negative")
                    validation_result['is_valid'] = False
                
                if filters.content.object_count_max is not None and filters.content.object_count_max < 0:
                    validation_result['errors'].append("object_count_max must be non-negative")
                    validation_result['is_valid'] = False
                
                if (filters.content.object_count_min is not None and 
                    filters.content.object_count_max is not None and
                    filters.content.object_count_min > filters.content.object_count_max):
                    validation_result['errors'].append("object_count_min cannot be greater than object_count_max")
                    validation_result['is_valid'] = False
            
            # Validate slide numbers
            if filters.slide_numbers:
                for slide_num in filters.slide_numbers:
                    if not isinstance(slide_num, int) or slide_num < 1:
                        validation_result['errors'].append(f"Invalid slide number: {slide_num}. Slide numbers must be positive integers")
                        validation_result['is_valid'] = False
            
            # Validate layout filters
            if filters.layout:
                # Layout validation is less strict since layout names can vary
                pass
            
            # Validate section filter
            if filters.section:
                if not isinstance(filters.section, str) or not filters.section.strip():
                    validation_result['errors'].append("Section name must be a non-empty string")
                    validation_result['is_valid'] = False
            
            # Validate notes filter
            if filters.notes:
                if filters.notes.regex:
                    try:
                        re.compile(filters.notes.regex)
                    except re.error as e:
                        validation_result['errors'].append(f"Invalid regex pattern in notes filter: {e}")
                        validation_result['is_valid'] = False
                
                if filters.notes.contains and not isinstance(filters.notes.contains, str):
                    validation_result['errors'].append("Notes contains filter must be a string")
                    validation_result['is_valid'] = False
                
                if filters.notes.is_empty is not None and not isinstance(filters.notes.is_empty, bool):
                    validation_result['errors'].append("Notes is_empty filter must be a boolean")
                    validation_result['is_valid'] = False
            
            return validation_result
            
        except Exception as e:
            logger.error(f"Error validating search criteria: {e}")
            return {
                'is_valid': False,
                'errors': [f"Validation error: {str(e)}"],
                'warnings': []
            }
    
    def validate_search_criteria_dict(self, criteria_dict: Dict[str, Any]) -> Dict[str, Any]:
        """Validate search criteria dictionary for grammar and syntax errors."""
        try:
            validation_result = {
                'is_valid': True,
                'errors': [],
                'warnings': []
            }
            
            # Check if criteria_dict is a valid dictionary
            if not isinstance(criteria_dict, dict):
                validation_result['errors'].append("Search criteria must be a dictionary/object")
                validation_result['is_valid'] = False
                return validation_result
            
            # Validate known field names
            valid_top_level_fields = [
                'title', 'content', 'layout', 'slide_numbers', 'section', 'notes'
            ]
            
            for field_name in criteria_dict.keys():
                if field_name not in valid_top_level_fields:
                    validation_result['errors'].append(f"Unknown field '{field_name}'. Valid fields: {valid_top_level_fields}")
                    validation_result['is_valid'] = False
            
            # Validate title field structure
            if 'title' in criteria_dict:
                title_validation = self._validate_title_field(criteria_dict['title'])
                validation_result['errors'].extend(title_validation['errors'])
                if not title_validation['is_valid']:
                    validation_result['is_valid'] = False
            
            # Validate content field structure
            if 'content' in criteria_dict:
                content_validation = self._validate_content_field(criteria_dict['content'])
                validation_result['errors'].extend(content_validation['errors'])
                if not content_validation['is_valid']:
                    validation_result['is_valid'] = False
            
            # Validate layout field structure
            if 'layout' in criteria_dict:
                layout_validation = self._validate_layout_field(criteria_dict['layout'])
                validation_result['errors'].extend(layout_validation['errors'])
                if not layout_validation['is_valid']:
                    validation_result['is_valid'] = False
            
            # Validate notes field structure
            if 'notes' in criteria_dict:
                notes_validation = self._validate_notes_field(criteria_dict['notes'])
                validation_result['errors'].extend(notes_validation['errors'])
                if not notes_validation['is_valid']:
                    validation_result['is_valid'] = False
            
            # Validate slide_numbers field
            if 'slide_numbers' in criteria_dict:
                slide_numbers = criteria_dict['slide_numbers']
                if not isinstance(slide_numbers, list):
                    validation_result['errors'].append("slide_numbers must be a list")
                    validation_result['is_valid'] = False
                else:
                    for i, slide_num in enumerate(slide_numbers):
                        if not isinstance(slide_num, int) or slide_num < 1:
                            validation_result['errors'].append(f"slide_numbers[{i}] must be a positive integer, got: {slide_num}")
                            validation_result['is_valid'] = False
            
            # Validate section field
            if 'section' in criteria_dict:
                section = criteria_dict['section']
                if not isinstance(section, str) or not section.strip():
                    validation_result['errors'].append("section must be a non-empty string")
                    validation_result['is_valid'] = False
            
            return validation_result
            
        except Exception as e:
            logger.error(f"Error validating search criteria dictionary: {e}")
            return {
                'is_valid': False,
                'errors': [f"Grammar validation error: {str(e)}"],
                'warnings': []
            }
    
    def _validate_title_field(self, title_field: Any) -> Dict[str, Any]:
        """Validate title field structure."""
        result = {'is_valid': True, 'errors': []}
        
        if not isinstance(title_field, dict):
            result['errors'].append("title field must be an object/dictionary")
            result['is_valid'] = False
            return result
        
        valid_title_fields = ['contains', 'starts_with', 'ends_with', 'regex', 'one_of']
        for field_name in title_field.keys():
            if field_name not in valid_title_fields:
                result['errors'].append(f"Unknown title field '{field_name}'. Valid fields: {valid_title_fields}")
                result['is_valid'] = False
        
        # Validate regex if present
        if 'regex' in title_field:
            try:
                re.compile(title_field['regex'])
            except re.error as e:
                result['errors'].append(f"Invalid regex in title.regex: {e}")
                result['is_valid'] = False
        
        # Validate one_of if present
        if 'one_of' in title_field:
            if not isinstance(title_field['one_of'], list):
                result['errors'].append("title.one_of must be a list")
                result['is_valid'] = False
        
        return result
    
    def _validate_content_field(self, content_field: Any) -> Dict[str, Any]:
        """Validate content field structure."""
        result = {'is_valid': True, 'errors': []}
        
        if not isinstance(content_field, dict):
            result['errors'].append("content field must be an object/dictionary")
            result['is_valid'] = False
            return result
        
        valid_content_fields = [
            'contains_text', 'has_tables', 'has_charts', 'has_images', 
            'object_count_min', 'object_count_max', 'object_count'
        ]
        for field_name in content_field.keys():
            if field_name not in valid_content_fields:
                result['errors'].append(f"Unknown content field '{field_name}'. Valid fields: {valid_content_fields}")
                result['is_valid'] = False
        
        # Validate boolean fields
        for bool_field in ['has_tables', 'has_charts', 'has_images']:
            if bool_field in content_field:
                if not isinstance(content_field[bool_field], bool):
                    result['errors'].append(f"content.{bool_field} must be a boolean")
                    result['is_valid'] = False
        
        # Validate numeric fields
        for num_field in ['object_count_min', 'object_count_max']:
            if num_field in content_field:
                if not isinstance(content_field[num_field], int) or content_field[num_field] < 0:
                    result['errors'].append(f"content.{num_field} must be a non-negative integer")
                    result['is_valid'] = False
        
        return result
    
    def _validate_layout_field(self, layout_field: Any) -> Dict[str, Any]:
        """Validate layout field structure."""
        result = {'is_valid': True, 'errors': []}
        
        if not isinstance(layout_field, dict):
            result['errors'].append("layout field must be an object/dictionary")
            result['is_valid'] = False
            return result
        
        valid_layout_fields = ['type', 'name']
        for field_name in layout_field.keys():
            if field_name not in valid_layout_fields:
                result['errors'].append(f"Unknown layout field '{field_name}'. Valid fields: {valid_layout_fields}")
                result['is_valid'] = False
        
        return result
    
    def _validate_notes_field(self, notes_field: Any) -> Dict[str, Any]:
        """Validate notes field structure."""
        result = {'is_valid': True, 'errors': []}
        
        if not isinstance(notes_field, dict):
            result['errors'].append("notes field must be an object/dictionary")
            result['is_valid'] = False
            return result
        
        valid_notes_fields = ['contains', 'regex', 'is_empty']
        for field_name in notes_field.keys():
            if field_name not in valid_notes_fields:
                result['errors'].append(f"Unknown notes field '{field_name}'. Valid fields: {valid_notes_fields}")
                result['is_valid'] = False
        
        # Validate regex if present
        if 'regex' in notes_field:
            try:
                re.compile(notes_field['regex'])
            except re.error as e:
                result['errors'].append(f"Invalid regex in notes.regex: {e}")
                result['is_valid'] = False
        
        # Validate is_empty if present
        if 'is_empty' in notes_field:
            if not isinstance(notes_field['is_empty'], bool):
                result['errors'].append("notes.is_empty must be a boolean")
                result['is_valid'] = False
        
        return result
    
    def clear_cache(self):
        """Clear the internal slide cache."""
        self._slide_cache.clear()
        logger.debug("Slide query cache cleared")


def create_filters_from_dict(filters_dict: Dict[str, Any]) -> SlideQueryFilters:
    """Create SlideQueryFilters from a dictionary representation."""
    filters = SlideQueryFilters()
    
    # Parse title filters
    if 'title' in filters_dict:
        title_dict = filters_dict['title']
        filters.title = TitleFilter(
            contains=title_dict.get('contains'),
            starts_with=title_dict.get('starts_with'),
            ends_with=title_dict.get('ends_with'),
            regex=title_dict.get('regex'),
            one_of=title_dict.get('one_of')
        )
    
    # Parse content filters
    if 'content' in filters_dict:
        content_dict = filters_dict['content']
        filters.content = ContentFilter(
            contains_text=content_dict.get('contains_text'),
            has_tables=content_dict.get('has_tables'),
            has_charts=content_dict.get('has_charts'),
            has_images=content_dict.get('has_images'),
            object_count_min=content_dict.get('object_count', {}).get('min') if content_dict.get('object_count') else None,
            object_count_max=content_dict.get('object_count', {}).get('max') if content_dict.get('object_count') else None
        )
    
    # Parse layout filters
    if 'layout' in filters_dict:
        layout_dict = filters_dict['layout']
        filters.layout = LayoutFilter(
            layout_type=layout_dict.get('type'),
            layout_name=layout_dict.get('name')
        )
    
    # Parse notes filters
    if 'notes' in filters_dict:
        notes_dict = filters_dict['notes']
        filters.notes = NotesFilter(
            contains=notes_dict.get('contains'),
            regex=notes_dict.get('regex'),
            is_empty=notes_dict.get('is_empty')
        )
    
    # Parse other filters
    filters.slide_numbers = filters_dict.get('slide_numbers')
    filters.section = filters_dict.get('section')
    
    return filters