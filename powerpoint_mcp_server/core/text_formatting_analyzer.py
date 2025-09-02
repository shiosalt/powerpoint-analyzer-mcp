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
    
    def __init__(self, content_extractor: Optional[ContentExtractor] = None):
        """Initialize the text formatting analyzer."""
        self.content_extractor = content_extractor or ContentExtractor()
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
        """Extract formatted text elements from slides."""
        try:
            formatted_elements = []
            
            with ZipExtractor(file_path) as extractor:
                # Get slide XML files
                slide_files_dict = extractor.get_slide_xml_files()
                slide_files = sorted(slide_files_dict.keys())  # Convert to sorted list
                
                # Determine which slides to analyze
                if slide_numbers is None:
                    slides_to_analyze = list(range(1, len(slide_files) + 1))
                else:
                    slides_to_analyze = [s for s in slide_numbers if s <= len(slide_files)]
                
                for slide_num in slides_to_analyze:
                    slide_file = slide_files[slide_num - 1]
                    slide_xml = extractor.read_xml_content(slide_file)
                    
                    if slide_xml:
                        elements = self._extract_formatted_elements_from_slide(
                            slide_xml, slide_num, formatting_filter
                        )
                        formatted_elements.extend(elements)
            
            return formatted_elements
            
        except Exception as e:
            logger.warning(f"Failed to extract formatted elements: {e}")
            return []
    
    def _extract_formatted_elements_from_slide(
        self,
        slide_xml: str,
        slide_number: int,
        formatting_filter: FormattingFilter
    ) -> List[FormattedTextElement]:
        """Extract formatted text elements from a single slide."""
        try:
            root = self.content_extractor.xml_parser.parse_xml_string(slide_xml)
            if root is None:
                return []
            
            elements = []
            
            # Extract from different content types based on filter
            content_types = formatting_filter.content_types or [ContentType.ALL]
            
            if ContentType.ALL in content_types or ContentType.TITLES in content_types:
                elements.extend(self._extract_title_formatting(root, slide_number))
            
            if ContentType.ALL in content_types or ContentType.TEXT_BOXES in content_types:
                elements.extend(self._extract_text_box_formatting(root, slide_number))
            
            if ContentType.ALL in content_types or ContentType.TABLES in content_types:
                elements.extend(self._extract_table_text_formatting(root, slide_number))
            
            if ContentType.ALL in content_types or ContentType.BULLETS in content_types:
                elements.extend(self._extract_bullet_formatting(root, slide_number))
            
            return elements
            
        except Exception as e:
            logger.warning(f"Failed to extract formatted elements from slide {slide_number}: {e}")
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
            
            # First check for paragraph-level default formatting using manual navigation
            all_elements = list(element.iter())
            paragraphs = [elem for elem in all_elements if elem.tag.endswith('}p')]
            
            logger.debug(f"Found {len(paragraphs)} paragraphs in element")
            
            for paragraph in paragraphs:
                # Find paragraph properties manually
                para_props = [child for child in paragraph if child.tag.endswith('}pPr')]
                
                for p_pr in para_props:
                    # Find default run properties manually
                    def_r_prs = [child for child in p_pr if child.tag.endswith('}defRPr')]
                    
                    for def_r_pr in def_r_prs:
                        # Check for bold in default run properties
                        for child in def_r_pr:
                            local_name = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                            if local_name == 'b':
                                bold_val = child.get('val', '1')
                                if bold_val != '0':
                                    logger.debug(f"Found explicit bold in paragraph defRPr")
                                    formatting['bold_count'] += 1
                                    formatting['has_formatting'] = True
                        
                        # Check Panose numbers for bold indication
                        latin_elems = [child for child in def_r_pr if child.tag.endswith('}latin')]
                        for latin_elem in latin_elems:
                            panose = latin_elem.get('panose', '')
                            logger.debug(f"Found panose in paragraph defRPr: {panose}")
                            if len(panose) >= 4:
                                # Panose weight is the 3rd and 4th characters (2nd byte)
                                weight_hex = panose[2:4]
                                try:
                                    weight = int(weight_hex, 16)
                                    logger.debug(f"Panose weight: {weight_hex} = {weight}")
                                    # Values >= 7 typically indicate bold (0x07-0x0F)
                                    if weight >= 7:
                                        logger.debug(f"Detected bold from panose weight: {weight}")
                                        formatting['bold_count'] += 1
                                        formatting['has_formatting'] = True
                                except ValueError:
                                    logger.debug(f"Invalid panose weight hex: {weight_hex}")
                                    pass
            
            # Check list style formatting using manual navigation
            all_elements = list(element.iter())
            lst_styles = [elem for elem in all_elements if elem.tag.endswith('}lstStyle')]
            
            logger.debug(f"List style found: {len(lst_styles) > 0}")
            
            for lst_style in lst_styles:
                # Find level properties manually
                level_props = [child for child in lst_style if 'lvl' in child.tag and child.tag.endswith('pPr')]
                logger.debug(f"Found {len(level_props)} level properties")
                
                for level_prop in level_props:
                    # Find default run properties in level
                    def_r_prs = [child for child in level_prop if child.tag.endswith('}defRPr')]
                    
                    for def_r_pr in def_r_prs:
                        # Check for bold in level default run properties
                        for child in def_r_pr:
                            local_name = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                            if local_name == 'b':
                                bold_val = child.get('val', '1')
                                if bold_val != '0':
                                    logger.debug(f"Found explicit bold in level defRPr")
                                    formatting['bold_count'] += 1
                                    formatting['has_formatting'] = True
                        
                        # Check Panose numbers in level properties
                        latin_elems = [child for child in def_r_pr if child.tag.endswith('}latin')]
                        for latin_elem in latin_elems:
                            panose = latin_elem.get('panose', '')
                            logger.debug(f"Found panose in level defRPr: {panose}")
                            if len(panose) >= 4:
                                weight_hex = panose[2:4]
                                try:
                                    weight = int(weight_hex, 16)
                                    logger.debug(f"Level panose weight: {weight_hex} = {weight}")
                                    if weight >= 7:
                                        logger.debug(f"Detected bold from level panose weight: {weight}")
                                        formatting['bold_count'] += 1
                                        formatting['has_formatting'] = True
                                except ValueError:
                                    logger.debug(f"Invalid level panose weight hex: {weight_hex}")
                                    pass
            
            # Find all text runs for run-level formatting
            runs = self.content_extractor.xml_parser.find_elements_with_namespace(
                element, './/a:r'
            )
            
            for run in runs:
                r_pr = self.content_extractor.xml_parser.find_element_with_namespace(
                    run, './/a:rPr'
                )
                
                if r_pr is not None:
                    # Check for bold - need to be more specific about the element name
                    # Look for direct child elements with local name 'b'
                    bold_elem = None
                    for child in r_pr:
                        local_name = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                        if local_name == 'b':
                            bold_elem = child
                            break
                    
                    if bold_elem is not None:
                        bold_val = bold_elem.get('val', '1')
                        if bold_val != '0':
                            formatting['bold_count'] += 1
                            formatting['has_formatting'] = True
                    
                    # Check for italic - look for direct child with local name 'i'
                    italic_elem = None
                    for child in r_pr:
                        local_name = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                        if local_name == 'i':
                            italic_elem = child
                            break
                    
                    if italic_elem is not None:
                        italic_val = italic_elem.get('val', '1')
                        if italic_val != '0':
                            formatting['italic_count'] += 1
                            formatting['has_formatting'] = True
                    
                    # Check for underline - look for direct child with local name 'u'
                    underline_elem = None
                    for child in r_pr:
                        local_name = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                        if local_name == 'u':
                            underline_elem = child
                            break
                    
                    if underline_elem is not None:
                        underline_val = underline_elem.get('val', 'sng')
                        if underline_val != 'none':
                            formatting['underline_count'] += 1
                            formatting['has_formatting'] = True
                    
                    # Check for strikethrough - look for direct child with local name 'strike'
                    strike_elem = None
                    for child in r_pr:
                        local_name = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                        if local_name == 'strike':
                            strike_elem = child
                            break
                    
                    if strike_elem is not None:
                        strike_val = strike_elem.get('val', 'sngStrike')
                        if strike_val != 'noStrike':
                            formatting['strikethrough_count'] += 1
                            formatting['has_formatting'] = True
                    
                    # Check for highlight - look for direct child with local name 'highlight'
                    highlight_elem = None
                    for child in r_pr:
                        local_name = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                        if local_name == 'highlight':
                            highlight_elem = child
                            break
                    
                    if highlight_elem is not None:
                        formatting['highlight_count'] += 1
                        formatting['has_formatting'] = True
                    
                    # Extract font size - look for direct child with local name 'sz'
                    font_size_elem = None
                    for child in r_pr:
                        local_name = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                        if local_name == 'sz':
                            font_size_elem = child
                            break
                    
                    # Extract font size - check both attribute and child element
                    sz = r_pr.get('sz')  # Check as attribute first
                    if not sz and font_size_elem is not None:
                        sz = font_size_elem.get('val')  # Check as child element
                    
                    if sz:
                        try:
                            font_size = float(sz) / 100.0
                            formatting['font_sizes'].append(font_size)
                            logger.debug(f"Extracted font size: {font_size} from sz value: {sz}")
                        except (ValueError, TypeError) as e:
                            logger.warning(f"Failed to parse font size '{sz}': {e}")
                    # Don't add default font size here - let the calling code handle defaults
                    
                    # Extract font color - look for solidFill child
                    solid_fill = None
                    for child in r_pr:
                        local_name = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                        if local_name == 'solidFill':
                            solid_fill = child
                            break
                    
                    if solid_fill is not None:
                        color = self._extract_color_from_fill(solid_fill)
                        if color:
                            formatting['font_colors'].append(color)
                            formatting['has_formatting'] = True
            
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
            
            return formatting
            
        except Exception as e:
            logger.warning(f"Failed to analyze text formatting in element: {e}")
            return {'has_formatting': False}
    
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