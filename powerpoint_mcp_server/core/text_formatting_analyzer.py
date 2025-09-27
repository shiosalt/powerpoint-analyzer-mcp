"""
Text formatting analysis system for PowerPoint slides.
"""

import re
import logging
from typing import Dict, List, Any, Optional, Union, Tuple
from dataclasses import dataclass, field
from enum import Enum
from collections import defaultdict

from .content_extractor import ContentExtractor
from ..utils.zip_extractor import ZipExtractor

logger = logging.getLogger(__name__)


class ContentType(Enum):
    """Enumeration of content types for analysis."""
    TABLES = "tables"
    TEXT_BOXES = "text_boxes"
    TITLES = "titles"
    BULLETS = "bullets"
    ALL = "all"


class FormattingType(Enum):
    """Enumeration of formatting types."""
    BOLD = "bold"
    ITALIC = "italic"
    UNDERLINE = "underline"
    HIGHLIGHT = "highlight"
    STRIKETHROUGH = "strikethrough"
    COLOR = "color"
    FONT_SIZE = "font_size"
    HYPERLINK = "hyperlink"
    ALL = "all"


class GroupingType(Enum):
    """Enumeration of grouping types."""
    BY_SLIDE = "by_slide"
    BY_FORMATTING_TYPE = "by_formatting_type"
    BY_CONTENT_TYPE = "by_content_type"
    BY_COLOR = "by_color"
    BY_FONT_SIZE = "by_font_size"
    NONE = "none"


@dataclass
class FormattingFilter:
    """Filter configuration for formatting analysis."""
    formatting_types: Optional[List[FormattingType]] = None
    content_types: Optional[List[ContentType]] = None
    text_contains: Optional[str] = None
    text_patterns: Optional[List[str]] = None
    slide_numbers: Optional[List[int]] = None
    min_font_size: Optional[int] = None
    max_font_size: Optional[int] = None
    colors: Optional[List[str]] = None


@dataclass
class FormattedTextElement:
    """A text element with formatting information."""
    slide_number: int
    content_type: ContentType
    element_index: int
    text_content: str
    formatting: Dict[str, Any]
    position: Tuple[int, int] = (0, 0)
    size: Tuple[int, int] = (0, 0)
    parent_element: Optional[str] = None  # e.g., "table_0", "text_box_1"


@dataclass
class FormattingAnalysisResult:
    """Result of formatting analysis."""
    total_elements: int
    formatted_elements: List[FormattedTextElement]
    formatting_summary: Dict[str, Any]
    groupings: Dict[str, Any] = field(default_factory=dict)


class TextFormattingAnalyzer:
    """
    Analyzer for text formatting in PowerPoint slides.
    """
    
    def __init__(self, content_extractor: Optional[ContentExtractor] = None, server=None):
        """Initialize the text formatting analyzer."""
        self.content_extractor = content_extractor or ContentExtractor()
        self.server = server  # Reference to PowerPointMCPServer for accessing _extract_powerpoint_content
        self._analysis_cache = {}
    
    def analyze_formatting(
        self,
        file_path: str,
        slide_numbers: Optional[List[int]] = None,
        formatting_filter: Optional[FormattingFilter] = None,
        grouping: GroupingType = GroupingType.NONE
    ) -> FormattingAnalysisResult:
        """
        Analyze text formatting across specified slides.
        
        Args:
            file_path: Path to the PowerPoint file
            slide_numbers: List of slide numbers to analyze (None for all)
            formatting_filter: Filter configuration
            grouping: How to group the results
            
        Returns:
            FormattingAnalysisResult with analysis data
        """
        if formatting_filter is None:
            formatting_filter = FormattingFilter()
        
        logger.info(f"Analyzing text formatting in {file_path}")
        
        try:
            # Extract formatted text elements
            formatted_elements = self._extract_formatted_elements(
                file_path, slide_numbers, formatting_filter
            )
            
            # Apply filters
            filtered_elements = self._apply_formatting_filters(
                formatted_elements, formatting_filter
            )
            
            # Create formatting summary
            formatting_summary = self._create_formatting_summary(filtered_elements)
            
            # Add sections and notes information to summary
            if hasattr(self, '_sections_info'):
                formatting_summary['sections'] = self._sections_info
            if hasattr(self, '_notes_info'):
                formatting_summary['notes'] = self._notes_info
            
            # Apply grouping
            groupings = {}
            if grouping != GroupingType.NONE:
                groupings = self._apply_grouping(filtered_elements, grouping)
            
            result = FormattingAnalysisResult(
                total_elements=len(filtered_elements),
                formatted_elements=filtered_elements,
                formatting_summary=formatting_summary,
                groupings=groupings
            )
            
            logger.info(f"Analyzed {len(filtered_elements)} formatted text elements")
            return result
            
        except Exception as e:
            logger.error(f"Error analyzing text formatting: {e}")
            raise
    
    def _extract_formatted_elements(
        self,
        file_path: str,
        slide_numbers: Optional[List[int]],
        formatting_filter: FormattingFilter
    ) -> List[FormattedTextElement]:
        """Extract formatted text elements from slides using PowerPointMCPServer."""
        try:
            formatted_elements = []
            
            # Use ContentExtractor directly with ZipExtractor
            from ..utils.zip_extractor import ZipExtractor
            
            with ZipExtractor(file_path) as extractor:
                # Get presentation XML for metadata and sections
                presentation_xml = extractor.read_xml_content('ppt/presentation.xml')
                sections = []
                slide_to_section = {}
                
                if presentation_xml:
                    sections = self.content_extractor.extract_presentation_sections(presentation_xml)
                    # Create slide ID mapping for sections
                    slide_id_mapping = {}
                    slide_files_dict_for_sections = extractor.get_slide_xml_files()
                    
                    # Type check to prevent the error
                    if isinstance(slide_files_dict_for_sections, dict):
                        for i, slide_file in enumerate(sorted(slide_files_dict_for_sections.keys()), 1):
                            # Extract slide ID from presentation.xml if needed
                            slide_id_mapping[f"slide{i}"] = i
                        slide_to_section = self.content_extractor.map_slides_to_sections(sections, slide_id_mapping)
                    else:
                        logger.warning(f"Expected dict from get_slide_xml_files() for sections, got {type(slide_files_dict_for_sections)}")
                        slide_to_section = {}
                
                # Get slide XML files
                slide_files_dict = extractor.get_slide_xml_files()
                
                # Type check to prevent the error
                if not isinstance(slide_files_dict, dict):
                    logger.error(f"Expected dict from get_slide_xml_files(), got {type(slide_files_dict)}: {slide_files_dict}")
                    return []
                
                slide_files = sorted(slide_files_dict.keys())
                
                slides = []
                for i, slide_file in enumerate(slide_files, 1):
                    slide_xml = extractor.read_xml_content(slide_file)
                    if slide_xml:
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
                        
                        # Resolve hyperlinks for this slide
                        self.content_extractor._resolve_hyperlink_relationships(
                            extractor, i, slide_info.text_elements
                        )
                        
                        # Convert SlideInfo to dict format
                        slide_data = {
                            'slide_number': i,
                            'title': slide_info.title,
                            'subtitle': slide_info.subtitle,
                            'text_elements': slide_info.text_elements,
                            'tables': slide_info.tables,
                            'notes': notes_content,
                            'section_name': slide_to_section.get(i, None)
                        }
                        slides.append(slide_data)
                
                content_result = {'slides': slides}
                
                # Store sections and notes information for summary
                self._sections_info = {
                    'total_sections': len(sections),
                    'sections': [{'name': s['name'], 'slide_count': s['slide_count']} for s in sections]
                }
                
                notes_stats = {
                    'slides_with_notes': sum(1 for slide in slides if slide.get('notes', '').strip()),
                    'total_notes_length': sum(len(slide.get('notes', '')) for slide in slides),
                    'average_notes_length': 0
                }
                if notes_stats['slides_with_notes'] > 0:
                    notes_stats['average_notes_length'] = notes_stats['total_notes_length'] / notes_stats['slides_with_notes']
                
                self._notes_info = notes_stats
            
            if not content_result or 'slides' not in content_result:
                logger.warning("No slide content found")
                return []
            
            slides = content_result['slides']
            logger.debug(f"Found {len(slides)} slides to analyze")
            
            # Determine which slides to analyze
            if slide_numbers is None or len(slide_numbers) == 0:
                slides_to_analyze = list(range(1, len(slides) + 1))
            else:
                slides_to_analyze = [s for s in slide_numbers if s <= len(slides)]
            
            logger.debug(f"Analyzing slides: {slides_to_analyze}")
            
            for slide_num in slides_to_analyze:
                slide_data = slides[slide_num - 1]  # Convert to 0-based index
                logger.debug(f"Processing slide {slide_num}: {len(slide_data.get('text_elements', []))} text elements")
                
                elements = self._extract_formatted_elements_from_slide_data(
                    slide_data, slide_num, formatting_filter
                )
                logger.debug(f"Extracted {len(elements)} formatted elements from slide {slide_num}")
                formatted_elements.extend(elements)
            
            return formatted_elements
            
        except Exception as e:
            logger.warning(f"Failed to extract formatted elements: {e}")
            import traceback
            logger.warning(f"Traceback: {traceback.format_exc()}")
            return []
    
    def _extract_formatted_elements_from_slide_data(
        self,
        slide_data: Dict[str, Any],
        slide_number: int,
        formatting_filter: FormattingFilter
    ) -> List[FormattedTextElement]:
        """Extract formatted text elements from slide data provided by ContentExtractor."""
        try:
            logger.debug(f"_extract_formatted_elements_from_slide_data called for slide {slide_number}")
            elements = []
            
            # Extract from different content types based on filter
            content_types = formatting_filter.content_types or [ContentType.ALL]
            logger.debug(f"Content types to analyze: {content_types}")
            
            # Process text elements from ContentExtractor
            text_elements = slide_data.get('text_elements', [])
            logger.debug(f"Found {len(text_elements)} text elements in slide {slide_number}")
            
            for element_index, text_element in enumerate(text_elements):
                logger.debug(f"Processing text element {element_index} from slide {slide_number}: {text_element.get('content_plain', '')[:50]}...")
                # Create FormattedTextElement from ContentExtractor data
                formatted_element = self._create_formatted_element_from_text_element(
                    text_element, slide_number, element_index, content_types
                )
                if formatted_element:
                    elements.append(formatted_element)
                    logger.debug(f"Added formatted element from slide {slide_number}, element {element_index}")
                else:
                    logger.debug(f"No formatted element created for slide {slide_number}, element {element_index}")
            
            # Also process title and subtitle if they have formatting
            if ContentType.ALL in content_types or ContentType.TITLES in content_types:
                title = slide_data.get('title')
                if title:
                    title_element = self._create_formatted_element_from_title(
                        title, slide_number, 'title'
                    )
                    if title_element:
                        elements.append(title_element)
                
                subtitle = slide_data.get('subtitle')
                if subtitle:
                    subtitle_element = self._create_formatted_element_from_title(
                        subtitle, slide_number, 'subtitle'
                    )
                    if subtitle_element:
                        elements.append(subtitle_element)
            
            return elements
            
        except Exception as e:
            logger.warning(f"Failed to extract formatted elements from slide {slide_number}: {e}")
            return []
    
    def _create_formatted_element_from_text_element(
        self,
        text_element: Dict[str, Any],
        slide_number: int,
        element_index: int,
        content_types: List[ContentType]
    ) -> Optional[FormattedTextElement]:
        """Create FormattedTextElement from ContentExtractor text element data."""
        try:
            logger.debug(f"Creating formatted element from text element: {text_element}")
            content_plain = text_element.get('content_plain', '')
            if not content_plain.strip():
                logger.debug(f"No content_plain found or empty: '{content_plain}'")
                return None
            
            # Create formatting counts from ContentExtractor data
            formatting_counts = {
                'bold': text_element.get('bolded', 0),
                'italic': text_element.get('italicized', 0),
                'underline': text_element.get('underlined', 0),
                'strikethrough': text_element.get('strikethrough', 0),
                'highlight': 0,  # ContentExtractor doesn't track this separately yet
                'colored_text': len(text_element.get('font_colors', [])),
                'hyperlinks': len(text_element.get('hyperlinks', []))
            }
            
            # Get position and size
            position = text_element.get('position', [0, 0])
            size = text_element.get('size', [0, 0])
            
            # Create formatting dictionary
            formatting = {
                'counts': formatting_counts,
                'font_sizes': text_element.get('font_sizes', []),
                'font_colors': text_element.get('font_colors', []),
                'hyperlinks': self._extract_hyperlink_urls(text_element.get('hyperlinks', []))
            }
            
            # Create FormattedTextElement
            formatted_element = FormattedTextElement(
                slide_number=slide_number,
                element_index=element_index,
                content_type=ContentType.TEXT_BOXES,
                text_content=content_plain,
                formatting=formatting,
                position=tuple(position),
                size=tuple(size)
            )
            
            logger.debug(f"Successfully created formatted element for slide {slide_number}, element {element_index}")
            return formatted_element
            
        except Exception as e:
            logger.warning(f"Failed to create formatted element from text element: {e}")
            import traceback
            logger.warning(f"Traceback: {traceback.format_exc()}")
            return None
    
    def _create_formatted_element_from_title(
        self,
        title_text: str,
        slide_number: int,
        title_type: str
    ) -> Optional[FormattedTextElement]:
        """Create FormattedTextElement from title/subtitle text."""
        try:
            if not title_text.strip():
                return None
            
            # For titles, we assume no special formatting for now
            # This could be enhanced to analyze title formatting if needed
            formatting_counts = {
                'bold': 0,
                'italic': 0,
                'underline': 0,
                'strikethrough': 0,
                'highlight': 0,
                'colored_text': 0,
                'hyperlinks': 0
            }
            
            content_type = ContentType.TITLES if title_type == 'title' else ContentType.TEXT_BOXES
            
            # Create formatting dictionary
            formatting = {
                'counts': formatting_counts,
                'font_sizes': [],
                'font_colors': [],
                'hyperlinks': []
            }
            
            formatted_element = FormattedTextElement(
                slide_number=slide_number,
                element_index=0,
                content_type=content_type,
                text_content=title_text,
                formatting=formatting,
                position=(0, 0),
                size=(0, 0)
            )
            
            return formatted_element
            
        except Exception as e:
            logger.warning(f"Failed to create formatted element from title: {e}")
            return None
    
    def _extract_hyperlink_urls(self, hyperlinks) -> List[str]:
        """Extract URLs from hyperlinks data."""
        try:
            urls = []
            if isinstance(hyperlinks, list):
                for hl in hyperlinks:
                    if isinstance(hl, dict):
                        # Dictionary format: {'url': 'http://...', 'display_text': '...'}
                        urls.append(hl.get('url', ''))
                    elif isinstance(hl, str):
                        # String format: direct URL
                        urls.append(hl)
            return urls
        except Exception as e:
            logger.warning(f"Failed to extract hyperlink URLs: {e}")
            return []
    
    def _extract_title_formatting(self, root, slide_number: int) -> List[FormattedTextElement]:
        """Extract formatting from slide titles."""
        try:
            elements = []
            
            # Find title placeholders
            shapes = self.content_extractor.xml_parser.find_elements_with_namespace(
                root, './/p:sp'
            )
            
            for shape_index, shape in enumerate(shapes):
                # Check if this is a title placeholder
                nv_sp_pr = self.content_extractor.xml_parser.find_element_with_namespace(
                    shape, './/p:nvSpPr'
                )
                if nv_sp_pr is not None:
                    ph = self.content_extractor.xml_parser.find_element_with_namespace(
                        nv_sp_pr, './/p:ph'
                    )
                    if ph is not None and ph.get('type') == 'title':
                        element = self._analyze_shape_text_formatting(
                            shape, slide_number, ContentType.TITLES, shape_index
                        )
                        if element:
                            elements.append(element)
            
            return elements
            
        except Exception as e:
            logger.warning(f"Failed to extract title formatting: {e}")
            return []
    
    def _extract_text_box_formatting(self, root, slide_number: int) -> List[FormattedTextElement]:
        """Extract formatting from text boxes."""
        try:
            elements = []
            
            # Find text box shapes (shapes with text body but not title/subtitle)
            shapes = self.content_extractor.xml_parser.find_elements_with_namespace(
                root, './/p:sp'
            )
            
            for shape_index, shape in enumerate(shapes):
                # Check if this shape has text but is not a title/subtitle
                tx_body = self.content_extractor.xml_parser.find_element_with_namespace(
                    shape, './/p:txBody'
                )
                
                if tx_body is not None:
                    # Check if it's not a title/subtitle placeholder
                    nv_sp_pr = self.content_extractor.xml_parser.find_element_with_namespace(
                        shape, './/p:nvSpPr'
                    )
                    is_title_placeholder = False
                    
                    if nv_sp_pr is not None:
                        ph = self.content_extractor.xml_parser.find_element_with_namespace(
                            nv_sp_pr, './/p:ph'
                        )
                        if ph is not None:
                            ph_type = ph.get('type', '')
                            if ph_type in ['title', 'subTitle', 'subtitle']:
                                is_title_placeholder = True
                    
                    if not is_title_placeholder:
                        element = self._analyze_shape_text_formatting(
                            shape, slide_number, ContentType.TEXT_BOXES, shape_index
                        )
                        if element:
                            elements.append(element)
            
            return elements
            
        except Exception as e:
            logger.warning(f"Failed to extract text box formatting: {e}")
            return []
    
    def _extract_table_text_formatting(self, root, slide_number: int) -> List[FormattedTextElement]:
        """Extract formatting from table text."""
        try:
            elements = []
            
            # Find table elements
            graphic_frames = self.content_extractor.xml_parser.find_elements_with_namespace(
                root, './/p:graphicFrame'
            )
            
            for frame_index, frame in enumerate(graphic_frames):
                table_elem = self.content_extractor.xml_parser.find_element_with_namespace(
                    frame, './/a:tbl'
                )
                
                if table_elem is not None:
                    # Analyze text formatting in table cells
                    rows = self.content_extractor.xml_parser.find_elements_with_namespace(
                        table_elem, './/a:tr'
                    )
                    
                    for row_index, row in enumerate(rows):
                        cells = self.content_extractor.xml_parser.find_elements_with_namespace(
                            row, './/a:tc'
                        )
                        
                        for cell_index, cell in enumerate(cells):
                            element = self._analyze_table_cell_formatting(
                                cell, slide_number, frame_index, row_index, cell_index
                            )
                            if element:
                                elements.append(element)
            
            return elements
            
        except Exception as e:
            logger.warning(f"Failed to extract table text formatting: {e}")
            return []
    
    def _extract_bullet_formatting(self, root, slide_number: int) -> List[FormattedTextElement]:
        """Extract formatting from bullet points."""
        try:
            elements = []
            
            # Find shapes with bullet lists
            shapes = self.content_extractor.xml_parser.find_elements_with_namespace(
                root, './/p:sp'
            )
            
            for shape_index, shape in enumerate(shapes):
                tx_body = self.content_extractor.xml_parser.find_element_with_namespace(
                    shape, './/p:txBody'
                )
                
                if tx_body is not None:
                    # Look for paragraphs with bullet formatting
                    paragraphs = self.content_extractor.xml_parser.find_elements_with_namespace(
                        tx_body, './/a:p'
                    )
                    
                    for para_index, paragraph in enumerate(paragraphs):
                        # Check if paragraph has bullet properties
                        p_pr = self.content_extractor.xml_parser.find_element_with_namespace(
                            paragraph, './/a:pPr'
                        )
                        
                        if p_pr is not None:
                            # Look for bullet properties
                            bullet_props = self.content_extractor.xml_parser.find_elements_with_namespace(
                                p_pr, './/a:buFont | .//a:buChar | .//a:buAutoNum'
                            )
                            
                            if bullet_props:
                                element = self._analyze_paragraph_formatting(
                                    paragraph, slide_number, ContentType.BULLETS, 
                                    shape_index, para_index
                                )
                                if element:
                                    elements.append(element)
            
            return elements
            
        except Exception as e:
            logger.warning(f"Failed to extract bullet formatting: {e}")
            return []
    
    def _analyze_shape_text_formatting(
        self,
        shape,
        slide_number: int,
        content_type: ContentType,
        element_index: int
    ) -> Optional[FormattedTextElement]:
        """Analyze text formatting in a shape."""
        try:
            # Extract text content
            text_content = self.content_extractor._extract_shape_text_content(shape)
            if not text_content or not text_content.strip():
                return None
            
            # Extract position and size
            position, size = self.content_extractor._extract_shape_transform(shape)
            
            # Analyze formatting
            formatting = self._analyze_text_formatting_in_element(shape)
            
            return FormattedTextElement(
                slide_number=slide_number,
                content_type=content_type,
                element_index=element_index,
                text_content=text_content,
                formatting=formatting,
                position=position,
                size=size,
                parent_element=f"{content_type.value}_{element_index}"
            )
            
        except Exception as e:
            logger.warning(f"Failed to analyze shape text formatting: {e}")
            return None
    
    def _analyze_table_cell_formatting(
        self,
        cell,
        slide_number: int,
        table_index: int,
        row_index: int,
        cell_index: int
    ) -> Optional[FormattedTextElement]:
        """Analyze text formatting in a table cell."""
        try:
            # Extract cell text content
            text_content = self.content_extractor._extract_cell_text_content(cell)
            if not text_content or not text_content.strip():
                return None
            
            # Analyze formatting
            formatting = self._analyze_text_formatting_in_element(cell)
            
            return FormattedTextElement(
                slide_number=slide_number,
                content_type=ContentType.TABLES,
                element_index=cell_index,
                text_content=text_content,
                formatting=formatting,
                position=(row_index, cell_index),
                size=(0, 0),
                parent_element=f"table_{table_index}_row_{row_index}"
            )
            
        except Exception as e:
            logger.warning(f"Failed to analyze table cell formatting: {e}")
            return None
    
    def _analyze_paragraph_formatting(
        self,
        paragraph,
        slide_number: int,
        content_type: ContentType,
        shape_index: int,
        para_index: int
    ) -> Optional[FormattedTextElement]:
        """Analyze text formatting in a paragraph."""
        try:
            # Extract paragraph text
            runs = self.content_extractor.xml_parser.find_elements_with_namespace(
                paragraph, './/a:r'
            )
            
            text_parts = []
            for run in runs:
                text_elem = self.content_extractor.xml_parser.find_element_with_namespace(
                    run, './/a:t'
                )
                if text_elem is not None and text_elem.text:
                    text_parts.append(text_elem.text)
            
            text_content = ''.join(text_parts)
            if not text_content.strip():
                return None
            
            # Analyze formatting
            formatting = self._analyze_text_formatting_in_element(paragraph)
            
            return FormattedTextElement(
                slide_number=slide_number,
                content_type=content_type,
                element_index=para_index,
                text_content=text_content,
                formatting=formatting,
                position=(0, 0),
                size=(0, 0),
                parent_element=f"{content_type.value}_{shape_index}_para_{para_index}"
            )
            
        except Exception as e:
            logger.warning(f"Failed to analyze paragraph formatting: {e}")
            return None
    
    def _analyze_text_formatting_in_element(self, element) -> Dict[str, Any]:
        """Analyze text formatting within an element."""
        try:
            formatting = {
                'bold_count': 0,
                'italic_count': 0,
                'underline_count': 0,
                'highlight_count': 0,
                'strikethrough_count': 0,
                'font_sizes': [],
                'font_colors': [],
                'hyperlinks': [],
                'has_formatting': False
            }
            
            # Use the XML parser's namespace-aware methods for better reliability
            # Find all text runs for run-level formatting
            runs = self.content_extractor.xml_parser.find_elements_with_namespace(
                element, './/a:r'
            )
            
            logger.debug(f"Found {len(runs)} text runs in element")
            
            for run in runs:
                r_pr = self.content_extractor.xml_parser.find_element_with_namespace(
                    run, './/a:rPr'
                )
                
                if r_pr is not None:
                    # Check for bold formatting
                    bold_elem = self.content_extractor.xml_parser.find_element_with_namespace(
                        r_pr, './/a:b'
                    )
                    if bold_elem is not None:
                        bold_val = bold_elem.get('val', '1')
                        if bold_val != '0':
                            formatting['bold_count'] += 1
                            formatting['has_formatting'] = True
                            logger.debug(f"Found bold formatting in run")
                    
                    # Check for italic formatting
                    italic_elem = self.content_extractor.xml_parser.find_element_with_namespace(
                        r_pr, './/a:i'
                    )
                    if italic_elem is not None:
                        italic_val = italic_elem.get('val', '1')
                        if italic_val != '0':
                            formatting['italic_count'] += 1
                            formatting['has_formatting'] = True
                            logger.debug(f"Found italic formatting in run")
                    
                    # Check for underline formatting
                    underline_elem = self.content_extractor.xml_parser.find_element_with_namespace(
                        r_pr, './/a:u'
                    )
                    if underline_elem is not None:
                        underline_val = underline_elem.get('val', 'sng')
                        if underline_val != 'none':
                            formatting['underline_count'] += 1
                            formatting['has_formatting'] = True
                            logger.debug(f"Found underline formatting in run")
                    
                    # Check for strikethrough formatting
                    strike_elem = self.content_extractor.xml_parser.find_element_with_namespace(
                        r_pr, './/a:strike'
                    )
                    if strike_elem is not None:
                        strike_val = strike_elem.get('val', 'sngStrike')
                        if strike_val != 'noStrike':
                            formatting['strikethrough_count'] += 1
                            formatting['has_formatting'] = True
                            logger.debug(f"Found strikethrough formatting in run")
                    
                    # Check for highlight formatting
                    highlight_elem = self.content_extractor.xml_parser.find_element_with_namespace(
                        r_pr, './/a:highlight'
                    )
                    if highlight_elem is not None:
                        formatting['highlight_count'] += 1
                        formatting['has_formatting'] = True
                        logger.debug(f"Found highlight formatting in run")
                    
                    # Extract font size
                    font_size_elem = self.content_extractor.xml_parser.find_element_with_namespace(
                        r_pr, './/a:sz'
                    )
                    if font_size_elem is not None:
                        sz = font_size_elem.get('val')
                        if sz:
                            try:
                                font_size = float(sz) / 100.0
                                formatting['font_sizes'].append(font_size)
                                logger.debug(f"Extracted font size: {font_size} from sz value: {sz}")
                            except (ValueError, TypeError) as e:
                                logger.warning(f"Failed to parse font size '{sz}': {e}")
                    
                    # Extract font color from solidFill
                    solid_fill = self.content_extractor.xml_parser.find_element_with_namespace(
                        r_pr, './/a:solidFill'
                    )
                    if solid_fill is not None:
                        color = self._extract_color_from_fill(solid_fill)
                        if color:
                            formatting['font_colors'].append(color)
                            formatting['has_formatting'] = True
                            logger.debug(f"Found font color: {color}")
            
            # Check for paragraph-level default formatting
            paragraphs = self.content_extractor.xml_parser.find_elements_with_namespace(
                element, './/a:p'
            )
            
            for paragraph in paragraphs:
                p_pr = self.content_extractor.xml_parser.find_element_with_namespace(
                    paragraph, './/a:pPr'
                )
                if p_pr is not None:
                    # Check default run properties in paragraph
                    def_r_pr = self.content_extractor.xml_parser.find_element_with_namespace(
                        p_pr, './/a:defRPr'
                    )
                    if def_r_pr is not None:
                        # Check for bold in default run properties
                        bold_elem = self.content_extractor.xml_parser.find_element_with_namespace(
                            def_r_pr, './/a:b'
                        )
                        if bold_elem is not None:
                            bold_val = bold_elem.get('val', '1')
                            if bold_val != '0':
                                formatting['bold_count'] += 1
                                formatting['has_formatting'] = True
                                logger.debug(f"Found bold in paragraph default properties")
                        
                        # Check for italic in default run properties
                        italic_elem = self.content_extractor.xml_parser.find_element_with_namespace(
                            def_r_pr, './/a:i'
                        )
                        if italic_elem is not None:
                            italic_val = italic_elem.get('val', '1')
                            if italic_val != '0':
                                formatting['italic_count'] += 1
                                formatting['has_formatting'] = True
                                logger.debug(f"Found italic in paragraph default properties")
            
            # Check for hyperlinks
            hyperlinks = self.content_extractor.xml_parser.find_elements_with_namespace(
                element, './/a:hlinkClick'
            )
            if hyperlinks:
                # Extract relationship IDs from hyperlinks
                hyperlink_ids = []
                for hl in hyperlinks:
                    # Try different attribute names for the relationship ID
                    r_id = hl.get('id') or hl.get('r:id') or hl.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                    if r_id:
                        hyperlink_ids.append(r_id)
                        logger.debug(f"Found hyperlink with relationship ID: {r_id}")
                    else:
                        hyperlink_ids.append('unknown')
                        logger.debug("Found hyperlink but could not extract relationship ID")
                
                formatting['hyperlinks'] = hyperlink_ids
                formatting['has_formatting'] = True
            
            # Remove duplicates
            formatting['font_sizes'] = list(set(formatting['font_sizes']))
            formatting['font_colors'] = list(set(formatting['font_colors']))
            
            logger.debug(f"Formatting analysis complete: {formatting}")
            return formatting
            
        except Exception as e:
            logger.warning(f"Failed to analyze text formatting in element: {e}")
            import traceback
            logger.warning(f"Traceback: {traceback.format_exc()}")
            return {
                'bold_count': 0,
                'italic_count': 0,
                'underline_count': 0,
                'highlight_count': 0,
                'strikethrough_count': 0,
                'font_sizes': [],
                'font_colors': [],
                'hyperlinks': [],
                'has_formatting': False
            }
    
    def _extract_color_from_fill(self, solid_fill) -> Optional[str]:
        """Extract color value from a solid fill element."""
        try:
            # Look for RGB color
            srgb_clr = self.content_extractor.xml_parser.find_element_with_namespace(
                solid_fill, './/a:srgbClr'
            )
            if srgb_clr is not None:
                color_val = srgb_clr.get('val')
                if color_val:
                    return f"#{color_val}"
            
            # Look for scheme color
            scheme_clr = self.content_extractor.xml_parser.find_element_with_namespace(
                solid_fill, './/a:schemeClr'
            )
            if scheme_clr is not None:
                color_val = scheme_clr.get('val')
                if color_val:
                    return color_val
            
            return None
            
        except Exception as e:
            logger.warning(f"Failed to extract color from fill: {e}")
            return None
    
    def _apply_formatting_filters(
        self,
        elements: List[FormattedTextElement],
        formatting_filter: FormattingFilter
    ) -> List[FormattedTextElement]:
        """Apply filters to formatted text elements."""
        try:
            filtered_elements = elements.copy()
            
            # Filter by slide numbers
            if formatting_filter.slide_numbers:
                filtered_elements = [
                    elem for elem in filtered_elements
                    if elem.slide_number in formatting_filter.slide_numbers
                ]
            
            # Filter by formatting types
            if formatting_filter.formatting_types and FormattingType.ALL not in formatting_filter.formatting_types:
                filtered_elements = [
                    elem for elem in filtered_elements
                    if self._has_requested_formatting(elem, formatting_filter.formatting_types)
                ]
            
            # Filter by text content
            if formatting_filter.text_contains:
                filtered_elements = [
                    elem for elem in filtered_elements
                    if formatting_filter.text_contains.lower() in elem.text_content.lower()
                ]
            
            # Filter by text patterns
            if formatting_filter.text_patterns:
                pattern_filtered = []
                for elem in filtered_elements:
                    for pattern in formatting_filter.text_patterns:
                        try:
                            if re.search(pattern, elem.text_content, re.IGNORECASE):
                                pattern_filtered.append(elem)
                                break
                        except re.error:
                            # Fallback to simple string matching
                            if pattern.lower() in elem.text_content.lower():
                                pattern_filtered.append(elem)
                                break
                filtered_elements = pattern_filtered
            
            # Filter by font size
            if formatting_filter.min_font_size is not None or formatting_filter.max_font_size is not None:
                size_filtered = []
                for elem in filtered_elements:
                    font_sizes = elem.formatting.get('font_sizes', [])
                    if font_sizes:
                        avg_size = sum(font_sizes) / len(font_sizes)
                        if formatting_filter.min_font_size is not None and avg_size < formatting_filter.min_font_size:
                            continue
                        if formatting_filter.max_font_size is not None and avg_size > formatting_filter.max_font_size:
                            continue
                    size_filtered.append(elem)
                filtered_elements = size_filtered
            
            # Filter by colors
            if formatting_filter.colors:
                color_filtered = []
                for elem in filtered_elements:
                    font_colors = elem.formatting.get('font_colors', [])
                    if any(color in font_colors for color in formatting_filter.colors):
                        color_filtered.append(elem)
                filtered_elements = color_filtered
            
            return filtered_elements
            
        except Exception as e:
            logger.warning(f"Failed to apply formatting filters: {e}")
            return elements
    
    def _has_requested_formatting(
        self,
        element: FormattedTextElement,
        formatting_types: List[FormattingType]
    ) -> bool:
        """Check if element has any of the requested formatting types."""
        try:
            formatting = element.formatting
            
            for fmt_type in formatting_types:
                if fmt_type == FormattingType.BOLD and formatting.get('bold_count', 0) > 0:
                    return True
                elif fmt_type == FormattingType.ITALIC and formatting.get('italic_count', 0) > 0:
                    return True
                elif fmt_type == FormattingType.UNDERLINE and formatting.get('underline_count', 0) > 0:
                    return True
                elif fmt_type == FormattingType.HIGHLIGHT and formatting.get('highlight_count', 0) > 0:
                    return True
                elif fmt_type == FormattingType.STRIKETHROUGH and formatting.get('strikethrough_count', 0) > 0:
                    return True
                elif fmt_type == FormattingType.COLOR and formatting.get('font_colors'):
                    return True
                elif fmt_type == FormattingType.FONT_SIZE and formatting.get('font_sizes'):
                    return True
                elif fmt_type == FormattingType.HYPERLINK and formatting.get('hyperlinks'):
                    return True
            
            return False
            
        except Exception as e:
            logger.warning(f"Failed to check requested formatting: {e}")
            return True  # Default to including the element
    
    def _create_formatting_summary(
        self,
        elements: List[FormattedTextElement]
    ) -> Dict[str, Any]:
        """Create a summary of formatting statistics."""
        try:
            summary = {
                'total_elements': len(elements),
                'elements_with_formatting': 0,
                'formatting_counts': {
                    'bold': 0,
                    'italic': 0,
                    'underline': 0,
                    'highlight': 0,
                    'strikethrough': 0,
                    'colored_text': 0,
                    'hyperlinks': 0
                },
                'font_sizes': {
                    'unique_sizes': set(),
                    'size_distribution': defaultdict(int)
                },
                'colors': {
                    'unique_colors': set(),
                    'color_distribution': defaultdict(int)
                },
                'content_type_distribution': defaultdict(int),
                'slide_distribution': defaultdict(int)
            }
            
            for element in elements:
                formatting = element.formatting
                
                # Count elements with formatting
                if formatting.get('has_formatting', False):
                    summary['elements_with_formatting'] += 1
                
                # Count specific formatting types
                summary['formatting_counts']['bold'] += formatting.get('bold_count', 0)
                summary['formatting_counts']['italic'] += formatting.get('italic_count', 0)
                summary['formatting_counts']['underline'] += formatting.get('underline_count', 0)
                summary['formatting_counts']['highlight'] += formatting.get('highlight_count', 0)
                summary['formatting_counts']['strikethrough'] += formatting.get('strikethrough_count', 0)
                
                # Count colored text
                if formatting.get('font_colors'):
                    summary['formatting_counts']['colored_text'] += 1
                
                # Count hyperlinks
                if formatting.get('hyperlinks'):
                    summary['formatting_counts']['hyperlinks'] += len(formatting['hyperlinks'])
                
                # Collect font sizes
                font_sizes = formatting.get('font_sizes', [])
                for size in font_sizes:
                    summary['font_sizes']['unique_sizes'].add(size)
                    summary['font_sizes']['size_distribution'][size] += 1
                
                # Collect colors
                colors = formatting.get('font_colors', [])
                for color in colors:
                    summary['colors']['unique_colors'].add(color)
                    summary['colors']['color_distribution'][color] += 1
                
                # Count by content type
                summary['content_type_distribution'][element.content_type.value] += 1
                
                # Count by slide
                summary['slide_distribution'][element.slide_number] += 1
            
            # Convert sets to lists for JSON serialization
            summary['font_sizes']['unique_sizes'] = list(summary['font_sizes']['unique_sizes'])
            summary['colors']['unique_colors'] = list(summary['colors']['unique_colors'])
            
            # Convert defaultdicts to regular dicts
            summary['font_sizes']['size_distribution'] = dict(summary['font_sizes']['size_distribution'])
            summary['colors']['color_distribution'] = dict(summary['colors']['color_distribution'])
            summary['content_type_distribution'] = dict(summary['content_type_distribution'])
            summary['slide_distribution'] = dict(summary['slide_distribution'])
            
            return summary
            
        except Exception as e:
            logger.warning(f"Failed to create formatting summary: {e}")
            return {'total_elements': len(elements)}
    
    def _apply_grouping(
        self,
        elements: List[FormattedTextElement],
        grouping: GroupingType
    ) -> Dict[str, Any]:
        """Apply grouping to formatted text elements."""
        try:
            if grouping == GroupingType.BY_SLIDE:
                return self._group_by_slide(elements)
            elif grouping == GroupingType.BY_FORMATTING_TYPE:
                return self._group_by_formatting_type(elements)
            elif grouping == GroupingType.BY_CONTENT_TYPE:
                return self._group_by_content_type(elements)
            elif grouping == GroupingType.BY_COLOR:
                return self._group_by_color(elements)
            elif grouping == GroupingType.BY_FONT_SIZE:
                return self._group_by_font_size(elements)
            else:
                return {}
                
        except Exception as e:
            logger.warning(f"Failed to apply grouping: {e}")
            return {}
    
    def _group_by_slide(self, elements: List[FormattedTextElement]) -> Dict[str, Any]:
        """Group elements by slide number."""
        groups = defaultdict(list)
        
        for element in elements:
            groups[f"slide_{element.slide_number}"].append({
                'content_type': element.content_type.value,
                'element_index': element.element_index,
                'text_content': element.text_content[:100] + "..." if len(element.text_content) > 100 else element.text_content,
                'formatting': element.formatting,
                'parent_element': element.parent_element
            })
        
        return dict(groups)
    
    def _group_by_formatting_type(self, elements: List[FormattedTextElement]) -> Dict[str, Any]:
        """Group elements by formatting type."""
        groups = {
            'bold': [],
            'italic': [],
            'underline': [],
            'highlight': [],
            'strikethrough': [],
            'colored': [],
            'hyperlinks': []
        }
        
        for element in elements:
            formatting = element.formatting
            element_info = {
                'slide_number': element.slide_number,
                'content_type': element.content_type.value,
                'text_content': element.text_content[:100] + "..." if len(element.text_content) > 100 else element.text_content,
                'parent_element': element.parent_element
            }
            
            if formatting.get('bold_count', 0) > 0:
                groups['bold'].append(element_info)
            if formatting.get('italic_count', 0) > 0:
                groups['italic'].append(element_info)
            if formatting.get('underline_count', 0) > 0:
                groups['underline'].append(element_info)
            if formatting.get('highlight_count', 0) > 0:
                groups['highlight'].append(element_info)
            if formatting.get('strikethrough_count', 0) > 0:
                groups['strikethrough'].append(element_info)
            if formatting.get('font_colors'):
                groups['colored'].append({**element_info, 'colors': formatting['font_colors']})
            if formatting.get('hyperlinks'):
                groups['hyperlinks'].append({**element_info, 'hyperlinks': formatting['hyperlinks']})
        
        return groups
    
    def _group_by_content_type(self, elements: List[FormattedTextElement]) -> Dict[str, Any]:
        """Group elements by content type."""
        groups = defaultdict(list)
        
        for element in elements:
            groups[element.content_type.value].append({
                'slide_number': element.slide_number,
                'element_index': element.element_index,
                'text_content': element.text_content[:100] + "..." if len(element.text_content) > 100 else element.text_content,
                'formatting': element.formatting,
                'parent_element': element.parent_element
            })
        
        return dict(groups)
    
    def _group_by_color(self, elements: List[FormattedTextElement]) -> Dict[str, Any]:
        """Group elements by font color."""
        groups = defaultdict(list)
        
        for element in elements:
            colors = element.formatting.get('font_colors', [])
            if colors:
                for color in colors:
                    groups[color].append({
                        'slide_number': element.slide_number,
                        'content_type': element.content_type.value,
                        'text_content': element.text_content[:100] + "..." if len(element.text_content) > 100 else element.text_content,
                        'parent_element': element.parent_element
                    })
            else:
                groups['no_color'].append({
                    'slide_number': element.slide_number,
                    'content_type': element.content_type.value,
                    'text_content': element.text_content[:100] + "..." if len(element.text_content) > 100 else element.text_content,
                    'parent_element': element.parent_element
                })
        
        return dict(groups)
    
    def _group_by_font_size(self, elements: List[FormattedTextElement]) -> Dict[str, Any]:
        """Group elements by font size."""
        groups = defaultdict(list)
        
        for element in elements:
            font_sizes = element.formatting.get('font_sizes', [])
            if font_sizes:
                # Use average font size for grouping
                avg_size = sum(font_sizes) / len(font_sizes)
                size_key = f"size_{int(avg_size)}"
                groups[size_key].append({
                    'slide_number': element.slide_number,
                    'content_type': element.content_type.value,
                    'text_content': element.text_content[:100] + "..." if len(element.text_content) > 100 else element.text_content,
                    'font_sizes': font_sizes,
                    'parent_element': element.parent_element
                })
            else:
                groups['no_size_info'].append({
                    'slide_number': element.slide_number,
                    'content_type': element.content_type.value,
                    'text_content': element.text_content[:100] + "..." if len(element.text_content) > 100 else element.text_content,
                    'parent_element': element.parent_element
                })
        
        return dict(groups)
    
    def clear_cache(self):
        """Clear the analysis cache."""
        self._analysis_cache.clear()
        logger.debug("Text formatting analysis cache cleared")


def create_formatting_filter_from_dict(filter_dict: Dict[str, Any]) -> FormattingFilter:
    """Create FormattingFilter from a dictionary representation."""
    # Convert string enums to enum objects
    formatting_types = None
    if 'formatting_types' in filter_dict:
        formatting_types = [FormattingType(ft) for ft in filter_dict['formatting_types']]
    
    content_types = None
    if 'content_types' in filter_dict:
        content_types = [ContentType(ct) for ct in filter_dict['content_types']]
    
    return FormattingFilter(
        formatting_types=formatting_types,
        content_types=content_types,
        text_contains=filter_dict.get('text_contains'),
        text_patterns=filter_dict.get('text_patterns'),
        slide_numbers=filter_dict.get('slide_numbers'),
        min_font_size=filter_dict.get('min_font_size'),
        max_font_size=filter_dict.get('max_font_size'),
        colors=filter_dict.get('colors')
    )

def create_formatting_filter_from_dict(filter_dict: Dict[str, Any]) -> FormattingFilter:
    """Create FormattingFilter from a dictionary representation."""
    try:
        formatting_filter = FormattingFilter()
        
        # Parse formatting types
        if 'formatting_types' in filter_dict:
            formatting_types = []
            for fmt_type in filter_dict['formatting_types']:
                if isinstance(fmt_type, str):
                    try:
                        formatting_types.append(FormattingType(fmt_type))
                    except ValueError:
                        pass  # Skip invalid formatting types
            formatting_filter.formatting_types = formatting_types
        
        # Parse content types
        if 'content_types' in filter_dict:
            content_types = []
            for content_type in filter_dict['content_types']:
                if isinstance(content_type, str):
                    try:
                        content_types.append(ContentType(content_type))
                    except ValueError:
                        pass  # Skip invalid content types
            formatting_filter.content_types = content_types
        
        # Parse other filter options
        formatting_filter.text_contains = filter_dict.get('text_contains')
        formatting_filter.text_patterns = filter_dict.get('text_patterns')
        formatting_filter.slide_numbers = filter_dict.get('slide_numbers')
        formatting_filter.min_font_size = filter_dict.get('min_font_size')
        formatting_filter.max_font_size = filter_dict.get('max_font_size')
        formatting_filter.colors = filter_dict.get('colors')
        
        return formatting_filter
        
    except Exception as e:
        logger.warning(f"Failed to create formatting filter from dict: {e}")
        return FormattingFilter()