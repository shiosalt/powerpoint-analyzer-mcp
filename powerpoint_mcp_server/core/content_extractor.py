"""
Content Extractor for PowerPoint files.

This module provides content extraction functionality for PowerPoint slides,
including slide layout, placeholder information, and basic slide content.
"""

import xml.etree.ElementTree as ET
from typing import Dict, List, Optional, Any, Tuple
from dataclasses import dataclass, field
import logging
import re

from .xml_parser import XMLParser
from ..utils.cache_manager import get_global_cache

logger = logging.getLogger(__name__)


@dataclass
class SlideInfo:
    """Information about a single slide."""
    slide_number: int
    layout_name: Optional[str] = None
    layout_type: Optional[str] = None
    title: Optional[str] = None
    subtitle: Optional[str] = None
    placeholders: List[Dict[str, Any]] = None
    text_elements: List[Dict[str, Any]] = None
    tables: List[Dict[str, Any]] = None
    
    def __post_init__(self):
        if self.placeholders is None:
            self.placeholders = []
        if self.text_elements is None:
            self.text_elements = []
        if self.tables is None:
            self.tables = []


@dataclass
class PlaceholderInfo:
    """Information about a slide placeholder."""
    placeholder_type: str
    position: Tuple[int, int]
    size: Tuple[int, int]
    content: Optional[str] = None


@dataclass
class TextElement:
    """Information about a text element with formatting."""
    content_plain: str
    content_formatted: str
    font_sizes: List[float] = field(default_factory=list)
    font_colors: List[str] = field(default_factory=list)
    hyperlinks: List[str] = field(default_factory=list)
    bolded: int = 0
    italic: int = 0
    underlined: int = 0
    highlighted: int = 0
    strikethrough: int = 0
    position: Tuple[int, int] = (0, 0)
    size: Tuple[int, int] = (0, 0)


@dataclass
class TableCell:
    """Information about a table cell."""
    content: str
    row_span: int = 1
    col_span: int = 1
    formatting: Optional[Dict[str, Any]] = None


@dataclass
class Table:
    """Information about a table."""
    rows: int
    columns: int
    cells: List[List[TableCell]]
    position: Tuple[int, int] = (0, 0)
    size: Tuple[int, int] = (0, 0)


class ContentExtractor:
    """
    Content extractor for PowerPoint slides.
    
    Extracts slide content including layout information, placeholders,
    and basic slide structure from PowerPoint XML data.
    """
    
    def __init__(self, enable_caching: bool = True):
        """
        Initialize the content extractor with XML parser and caching.
        
        Args:
            enable_caching: Enable caching of extraction results
        """
        self.xml_parser = XMLParser(enable_performance_mode=True)
        self.enable_caching = enable_caching
        self.cache_manager = get_global_cache() if enable_caching else None
    
    def extract_slide_content(self, slide_xml_content: str, slide_number: int) -> SlideInfo:
        """
        Extract content from a single slide XML.
        
        Args:
            slide_xml_content: XML content of the slide
            slide_number: Slide number (1-based)
            
        Returns:
            SlideInfo object containing extracted slide information
            
        Raises:
            ET.ParseError: If the XML is malformed
        """
        # Check cache first if caching is enabled
        if self.enable_caching and self.cache_manager:
            import hashlib
            cache_key = f"slide_content_{slide_number}_{hashlib.md5(slide_xml_content.encode()).hexdigest()}"
            cached_result = self.cache_manager.get(cache_key)
            if cached_result is not None:
                logger.debug(f"Retrieved slide {slide_number} content from cache")
                return cached_result
        
        try:
            root = self.xml_parser.parse_xml_string(slide_xml_content)
            if root is None:
                logger.warning(f"Failed to parse slide {slide_number} XML")
                return SlideInfo(slide_number=slide_number)
            
            slide_info = SlideInfo(slide_number=slide_number)
            
            # Extract slide layout information
            self._extract_layout_info(root, slide_info)
            
            # Extract placeholder information
            self._extract_placeholder_info(root, slide_info)
            
            # Extract title and subtitle
            self._extract_title_subtitle(root, slide_info)
            
            # Extract text elements with formatting
            self._extract_text_elements(root, slide_info)
            
            # Extract table data
            self._extract_tables(root, slide_info)
            
            logger.debug(f"Successfully extracted content for slide {slide_number}")
            
            # Cache the result if caching is enabled
            if self.enable_caching and self.cache_manager:
                import hashlib
                cache_key = f"slide_content_{slide_number}_{hashlib.md5(slide_xml_content.encode()).hexdigest()}"
                self.cache_manager.put(cache_key, slide_info, ttl=3600)  # Cache for 1 hour
                logger.debug(f"Cached slide {slide_number} content")
            
            return slide_info
            
        except Exception as e:
            logger.error(f"Failed to extract slide {slide_number} content: {e}")
            return SlideInfo(slide_number=slide_number)
    
    def _extract_layout_info(self, root: ET.Element, slide_info: SlideInfo) -> None:
        """
        Extract layout information from slide XML.
        
        Args:
            root: Root element of slide XML
            slide_info: SlideInfo object to populate
        """
        try:
            # Look for slide layout reference
            slide_layout = self.xml_parser.find_element_with_namespace(
                root, './/p:cSld', 
            )
            
            if slide_layout is not None:
                # Extract layout name from name attribute if available
                name_attr = slide_layout.get('name')
                if name_attr:
                    slide_info.layout_name = name_attr
                
                # Try to determine layout type based on structure
                slide_info.layout_type = self._determine_layout_type(root)
            
        except Exception as e:
            logger.warning(f"Failed to extract layout info for slide {slide_info.slide_number}: {e}")
    
    def _extract_placeholder_info(self, root: ET.Element, slide_info: SlideInfo) -> None:
        """
        Extract placeholder information from slide XML.
        
        Args:
            root: Root element of slide XML
            slide_info: SlideInfo object to populate
        """
        try:
            # Find all shape elements that might be placeholders
            shapes = self.xml_parser.find_elements_with_namespace(root, './/p:sp')
            
            for shape in shapes:
                placeholder_info = self._extract_single_placeholder(shape)
                if placeholder_info:
                    slide_info.placeholders.append({
                        'type': placeholder_info.placeholder_type,
                        'position': placeholder_info.position,
                        'size': placeholder_info.size,
                        'content': placeholder_info.content
                    })
            
        except Exception as e:
            logger.warning(f"Failed to extract placeholder info for slide {slide_info.slide_number}: {e}")
    
    def _extract_single_placeholder(self, shape: ET.Element) -> Optional[PlaceholderInfo]:
        """
        Extract information from a single placeholder shape.
        
        Args:
            shape: Shape element that might be a placeholder
            
        Returns:
            PlaceholderInfo object if the shape is a placeholder, None otherwise
        """
        try:
            # Check if this shape has placeholder properties
            nv_sp_pr = self.xml_parser.find_element_with_namespace(shape, './/p:nvSpPr')
            if nv_sp_pr is None:
                return None
            
            # Look for placeholder type
            ph = self.xml_parser.find_element_with_namespace(nv_sp_pr, './/p:ph')
            if ph is None:
                return None
            
            # Extract placeholder type
            placeholder_type = ph.get('type', 'content')
            
            # Extract position and size from transform
            position, size = self._extract_shape_transform(shape)
            
            # Extract content if available
            content = self._extract_shape_text_content(shape)
            
            return PlaceholderInfo(
                placeholder_type=placeholder_type,
                position=position,
                size=size,
                content=content
            )
            
        except Exception as e:
            logger.warning(f"Failed to extract placeholder info from shape: {e}")
            return None
    
    def _extract_shape_transform(self, shape: ET.Element) -> Tuple[Tuple[int, int], Tuple[int, int]]:
        """
        Extract position and size from shape transform.
        
        Args:
            shape: Shape element
            
        Returns:
            Tuple of (position, size) where each is (x, y) or (width, height)
        """
        try:
            # Find transform element
            xfrm = self.xml_parser.find_element_with_namespace(shape, './/a:xfrm')
            if xfrm is None:
                return (0, 0), (0, 0)
            
            # Extract offset (position)
            off = self.xml_parser.find_element_with_namespace(xfrm, './/a:off')
            position = (0, 0)
            if off is not None:
                x = int(off.get('x', '0'))
                y = int(off.get('y', '0'))
                position = (x, y)
            
            # Extract extent (size)
            ext = self.xml_parser.find_element_with_namespace(xfrm, './/a:ext')
            size = (0, 0)
            if ext is not None:
                cx = int(ext.get('cx', '0'))
                cy = int(ext.get('cy', '0'))
                size = (cx, cy)
            
            return position, size
            
        except Exception as e:
            logger.warning(f"Failed to extract shape transform: {e}")
            return (0, 0), (0, 0)
    
    def _extract_shape_text_content(self, shape: ET.Element) -> Optional[str]:
        """
        Extract text content from a shape.
        
        Args:
            shape: Shape element
            
        Returns:
            Text content if available, None otherwise
        """
        try:
            # Find text body
            tx_body = self.xml_parser.find_element_with_namespace(shape, './/p:txBody')
            if tx_body is None:
                return None
            
            # Extract all text from paragraphs
            text_parts = []
            paragraphs = self.xml_parser.find_elements_with_namespace(tx_body, './/a:p')
            
            for paragraph in paragraphs:
                # Get text from all runs in the paragraph
                runs = self.xml_parser.find_elements_with_namespace(paragraph, './/a:r')
                paragraph_text = []
                
                for run in runs:
                    text_elem = self.xml_parser.find_element_with_namespace(run, './/a:t')
                    if text_elem is not None and text_elem.text:
                        paragraph_text.append(text_elem.text)
                
                if paragraph_text:
                    text_parts.append(''.join(paragraph_text))
            
            return '\n'.join(text_parts) if text_parts else None
            
        except Exception as e:
            logger.warning(f"Failed to extract shape text content: {e}")
            return None
    
    def _extract_title_subtitle(self, root: ET.Element, slide_info: SlideInfo) -> None:
        """
        Extract title and subtitle from slide placeholders.
        
        Args:
            root: Root element of slide XML
            slide_info: SlideInfo object to populate
        """
        try:
            # Look for title and subtitle in placeholders
            for placeholder in slide_info.placeholders:
                if placeholder['type'] in ['title', 'ctrTitle'] and placeholder['content']:
                    slide_info.title = placeholder['content']
                elif placeholder['type'] in ['subTitle', 'subtitle'] and placeholder['content']:
                    slide_info.subtitle = placeholder['content']
            
        except Exception as e:
            logger.warning(f"Failed to extract title/subtitle for slide {slide_info.slide_number}: {e}")
    
    def _determine_layout_type(self, root: ET.Element) -> str:
        """
        Determine the layout type based on slide structure.
        
        Args:
            root: Root element of slide XML
            
        Returns:
            Layout type string
        """
        try:
            # Count different types of elements to guess layout
            shapes = self.xml_parser.find_elements_with_namespace(root, './/p:sp')
            
            has_title = False
            has_content = False
            has_two_content = False
            content_count = 0
            
            for shape in shapes:
                nv_sp_pr = self.xml_parser.find_element_with_namespace(shape, './/p:nvSpPr')
                if nv_sp_pr is not None:
                    ph = self.xml_parser.find_element_with_namespace(nv_sp_pr, './/p:ph')
                    if ph is not None:
                        ph_type = ph.get('type', 'content')
                        if ph_type == 'title':
                            has_title = True
                        elif ph_type in ['body', 'obj', 'content']:
                            content_count += 1
                            has_content = True
                            if content_count >= 2:
                                has_two_content = True
            
            # Determine layout based on placeholder types
            if has_title and has_two_content:
                return 'twoContent'
            elif has_title and has_content:
                return 'titleAndContent'
            elif has_title:
                return 'titleOnly'
            elif has_content:
                return 'contentOnly'
            else:
                return 'blank'
                
        except Exception as e:
            logger.warning(f"Failed to determine layout type: {e}")
            return 'unknown'
    
    def extract_slide_layout_info(self, layout_xml_content: str) -> Dict[str, Any]:
        """
        Extract layout information from slide layout XML.
        
        Args:
            layout_xml_content: XML content of the slide layout
            
        Returns:
            Dictionary containing layout information
        """
        try:
            root = self.xml_parser.parse_xml_string(layout_xml_content)
            if root is None:
                return {}
            
            layout_info = {
                'name': None,
                'type': None,
                'placeholders': []
            }
            
            # Extract layout name
            cSld = self.xml_parser.find_element_with_namespace(root, './/p:cSld')
            if cSld is not None:
                layout_info['name'] = cSld.get('name', 'Unknown Layout')
            
            # Extract placeholder definitions
            shapes = self.xml_parser.find_elements_with_namespace(root, './/p:sp')
            for shape in shapes:
                placeholder_info = self._extract_single_placeholder(shape)
                if placeholder_info:
                    layout_info['placeholders'].append({
                        'type': placeholder_info.placeholder_type,
                        'position': placeholder_info.position,
                        'size': placeholder_info.size
                    })
            
            # Determine layout type
            layout_info['type'] = self._determine_layout_type(root)
            
            return layout_info
            
        except Exception as e:
            logger.error(f"Failed to extract layout info: {e}")
            return {}
    
    def extract_basic_slide_info(self, slide_xml_content: str, slide_number: int) -> Dict[str, Any]:
        """
        Extract basic slide information for quick access.
        
        Args:
            slide_xml_content: XML content of the slide
            slide_number: Slide number (1-based)
            
        Returns:
            Dictionary containing basic slide information
        """
        try:
            slide_info = self.extract_slide_content(slide_xml_content, slide_number)
            
            return {
                'slide_number': slide_info.slide_number,
                'layout_name': slide_info.layout_name,
                'layout_type': slide_info.layout_type,
                'title': slide_info.title,
                'subtitle': slide_info.subtitle,
                'placeholder_count': len(slide_info.placeholders),
                'placeholder_types': [p['type'] for p in slide_info.placeholders]
            }
            
        except Exception as e:
            logger.error(f"Failed to extract basic slide info for slide {slide_number}: {e}")
            return {
                'slide_number': slide_number,
                'error': str(e)
            }
    
    def _extract_text_elements(self, root: ET.Element, slide_info: SlideInfo) -> None:
        """
        Extract text elements with formatting information from slide XML.
        
        Args:
            root: Root element of slide XML
            slide_info: SlideInfo object to populate
        """
        try:
            # Find all shapes that contain text
            shapes = self.xml_parser.find_elements_with_namespace(root, './/p:sp')
            
            for shape in shapes:
                text_element = self._extract_text_element_from_shape(shape)
                if text_element and (text_element.content_plain.strip() or text_element.hyperlinks):
                    slide_info.text_elements.append({
                        'content_plain': text_element.content_plain,
                        'content_formatted': text_element.content_formatted,
                        'font_sizes': text_element.font_sizes,
                        'font_colors': text_element.font_colors,
                        'hyperlinks': text_element.hyperlinks,
                        'bolded': text_element.bolded,
                        'italic': text_element.italic,
                        'underlined': text_element.underlined,
                        'highlighted': text_element.highlighted,
                        'strikethrough': text_element.strikethrough,
                        'position': text_element.position,
                        'size': text_element.size
                    })
            
        except Exception as e:
            logger.warning(f"Failed to extract text elements for slide {slide_info.slide_number}: {e}")
    
    def _extract_text_element_from_shape(self, shape: ET.Element) -> Optional[TextElement]:
        """
        Extract text element with formatting from a single shape.
        
        Args:
            shape: Shape element that might contain text
            
        Returns:
            TextElement object if the shape contains text, None otherwise
        """
        try:
            # Find text body
            tx_body = self.xml_parser.find_element_with_namespace(shape, './/p:txBody')
            if tx_body is None:
                return None
            
            # Extract position and size
            position, size = self._extract_shape_transform(shape)
            
            # Initialize text element
            text_element = TextElement(
                content_plain="",
                content_formatted="",
                position=position,
                size=size
            )
            
            # Extract text with formatting from all paragraphs
            paragraphs = self.xml_parser.find_elements_with_namespace(tx_body, './/a:p')
            paragraph_texts_plain = []
            paragraph_texts_formatted = []
            
            for paragraph in paragraphs:
                para_plain, para_formatted = self._extract_paragraph_text(paragraph, text_element)
                if para_plain or para_formatted:
                    paragraph_texts_plain.append(para_plain)
                    paragraph_texts_formatted.append(para_formatted)
            
            # Combine all paragraphs
            text_element.content_plain = '\n'.join(paragraph_texts_plain)
            text_element.content_formatted = '\n'.join(paragraph_texts_formatted)
            
            # Debug: Log font sizes before deduplication
            logger.debug(f"Font sizes before deduplication: {text_element.font_sizes}")
            
            # Add context-aware default font size if none found
            if not text_element.font_sizes:
                # Determine default based on context (title vs content)
                default_size = self._get_default_font_size(shape)
                text_element.font_sizes.append(default_size)
                logger.debug(f"Added context-aware default font size: {default_size}pt")
            
            # Remove duplicates from lists
            text_element.font_sizes = list(set(text_element.font_sizes))
            text_element.font_colors = list(set(text_element.font_colors))
            text_element.hyperlinks = list(set(text_element.hyperlinks))
            
            # Debug: Log font sizes after deduplication
            logger.debug(f"Font sizes after deduplication: {text_element.font_sizes}")
            
            return text_element if text_element.content_plain.strip() or text_element.hyperlinks else None
            
        except Exception as e:
            logger.warning(f"Failed to extract text element from shape: {e}")
            return None
    
    def _extract_paragraph_text(self, paragraph: ET.Element, text_element: TextElement) -> Tuple[str, str]:
        """
        Extract text from a paragraph with formatting information.
        
        Args:
            paragraph: Paragraph element
            text_element: TextElement to accumulate formatting info
            
        Returns:
            Tuple of (plain_text, formatted_text)
        """
        try:
            plain_parts = []
            formatted_parts = []
            
            # Extract text from all runs in the paragraph
            runs = self.xml_parser.find_elements_with_namespace(paragraph, './/a:r')
            
            for run in runs:
                run_plain, run_formatted = self._extract_run_text(run, text_element)
                if run_plain:
                    plain_parts.append(run_plain)
                    formatted_parts.append(run_formatted)
            
            # Check for hyperlinks in the paragraph
            hyperlinks = self.xml_parser.find_elements_with_namespace(paragraph, './/a:hlinkClick')
            for hyperlink in hyperlinks:
                r_id = self.xml_parser.get_attribute_with_namespace(
                    hyperlink, 'id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
                )
                if r_id:
                    # Store the relationship ID for now - we'll resolve it later if needed
                    text_element.hyperlinks.append(r_id)
                    logger.debug(f"Found hyperlink with relationship ID: {r_id}")
            
            return ''.join(plain_parts), ''.join(formatted_parts)
            
        except Exception as e:
            logger.warning(f"Failed to extract paragraph text: {e}")
            return "", ""
    
    def _extract_run_text(self, run: ET.Element, text_element: TextElement) -> Tuple[str, str]:
        """
        Extract text from a run with formatting information.
        
        Args:
            run: Run element
            text_element: TextElement to accumulate formatting info
            
        Returns:
            Tuple of (plain_text, formatted_text)
        """
        try:
            # Extract text content
            text_elem = self.xml_parser.find_element_with_namespace(run, './/a:t')
            if text_elem is None or not text_elem.text:
                return "", ""
            
            text_content = text_elem.text
            formatted_text = text_content
            
            # Extract run properties
            r_pr = self.xml_parser.find_element_with_namespace(run, './/a:rPr')
            if r_pr is not None:
                formatted_text = self._apply_text_formatting(text_content, r_pr, text_element)
            # Note: If no run properties found, we don't add a default font size here
            # The default will be added at the text element level if no font sizes are found
            
            return text_content, formatted_text
            
        except Exception as e:
            logger.warning(f"Failed to extract run text: {e}")
            return "", ""
    
    def _apply_text_formatting(self, text: str, r_pr: ET.Element, text_element: TextElement) -> str:
        """
        Apply formatting to text and accumulate formatting statistics.
        
        Args:
            text: Text content
            r_pr: Run properties element
            text_element: TextElement to accumulate formatting info
            
        Returns:
            Formatted text with HTML-like tags
        """
        try:
            formatted_text = text
            formatting_tags = []
            
            # Debug: Log the run properties XML for troubleshooting
            if logger.isEnabledFor(logging.DEBUG):
                import xml.etree.ElementTree as ET
                logger.debug(f"Processing run properties for text: '{text[:50]}...'")
                logger.debug(f"Run properties XML: {ET.tostring(r_pr, encoding='unicode')}")
            
            # Extract font size - check both attribute and child element
            sz = r_pr.get('sz')  # Check as attribute first
            if not sz:
                # Check as child element
                font_size_elem = self.xml_parser.find_element_with_namespace(r_pr, './/a:sz')
                if font_size_elem is not None:
                    sz = font_size_elem.get('val')
            
            if sz:
                try:
                    # Font size in PowerPoint is in hundredths of a point
                    font_size = float(sz) / 100.0
                    text_element.font_sizes.append(font_size)
                    logger.debug(f"Extracted font size: {font_size} from sz value: {sz}")
                except (ValueError, TypeError) as e:
                    logger.warning(f"Failed to parse font size '{sz}': {e}")
            # Note: If no explicit font size found, we don't add a default here
            # The default will be added at the text element level if no font sizes are found
            
            # Extract font color
            solid_fill = self.xml_parser.find_element_with_namespace(r_pr, './/a:solidFill')
            if solid_fill is not None:
                # Look for RGB color
                srgb_clr = self.xml_parser.find_element_with_namespace(solid_fill, './/a:srgbClr')
                if srgb_clr is not None:
                    color_val = srgb_clr.get('val')
                    if color_val:
                        text_element.font_colors.append(f"#{color_val}")
                
                # Look for scheme color
                scheme_clr = self.xml_parser.find_element_with_namespace(solid_fill, './/a:schemeClr')
                if scheme_clr is not None:
                    color_val = scheme_clr.get('val')
                    if color_val:
                        text_element.font_colors.append(color_val)
            
            # Check for bold - can be either attribute or child element
            bold_attr = r_pr.get('b')
            if bold_attr is not None and bold_attr != '0':
                text_element.bolded += 1
                formatting_tags.append('b')
                logger.debug(f"Applied bold formatting (attribute) to text: '{text[:30]}...'")
            else:
                # Also check for bold as child element
                bold_elem = self.xml_parser.find_element_with_namespace(r_pr, './/a:b')
                if bold_elem is not None:
                    bold_val = bold_elem.get('val', '1')
                    logger.debug(f"Found bold element with val='{bold_val}' for text: '{text[:30]}...'")
                    if bold_val != '0':
                        text_element.bolded += 1
                        formatting_tags.append('b')
                        logger.debug(f"Applied bold formatting (element) to text: '{text[:30]}...'")
                else:
                    logger.debug(f"No bold formatting found for text: '{text[:30]}...'")
            
            # Check for italic - can be either attribute or child element
            italic_attr = r_pr.get('i')
            if italic_attr is not None and italic_attr != '0':
                text_element.italic += 1
                formatting_tags.append('i')
                logger.debug(f"Applied italic formatting (attribute) to text: '{text[:30]}...'")
            else:
                # Also check for italic as child element
                italic_elem = self.xml_parser.find_element_with_namespace(r_pr, './/a:i')
                if italic_elem is not None:
                    italic_val = italic_elem.get('val', '1')
                    logger.debug(f"Found italic element with val='{italic_val}' for text: '{text[:30]}...'")
                    if italic_val != '0':
                        text_element.italic += 1
                        formatting_tags.append('i')
                        logger.debug(f"Applied italic formatting (element) to text: '{text[:30]}...'")
                else:
                    logger.debug(f"No italic formatting found for text: '{text[:30]}...'")
            
            # Check for underline - can be either attribute or child element
            underline_attr = r_pr.get('u')
            if underline_attr is not None and underline_attr != 'none':
                text_element.underlined += 1
                formatting_tags.append('u')
                logger.debug(f"Applied underline formatting (attribute) to text: '{text[:30]}...'")
            else:
                # Also check for underline as child element
                underline_elem = self.xml_parser.find_element_with_namespace(r_pr, './/a:u')
                if underline_elem is not None:
                    underline_val = underline_elem.get('val', 'sng')
                    logger.debug(f"Found underline element with val='{underline_val}' for text: '{text[:30]}...'")
                    if underline_val != 'none':
                        text_element.underlined += 1
                        formatting_tags.append('u')
                        logger.debug(f"Applied underline formatting (element) to text: '{text[:30]}...'")
                else:
                    logger.debug(f"No underline formatting found for text: '{text[:30]}...'")
            
            # Check for strikethrough - can be either attribute or child element
            strike_attr = r_pr.get('strike')
            if strike_attr is not None and strike_attr != 'noStrike':
                text_element.strikethrough += 1
                formatting_tags.append('s')
                logger.debug(f"Applied strikethrough formatting (attribute) to text: '{text[:30]}...'")
            else:
                # Also check for strikethrough as child element
                strike_elem = self.xml_parser.find_element_with_namespace(r_pr, './/a:strike')
                if strike_elem is not None:
                    strike_val = strike_elem.get('val', 'sngStrike')
                    logger.debug(f"Found strikethrough element with val='{strike_val}' for text: '{text[:30]}...'")
                    if strike_val != 'noStrike':
                        text_element.strikethrough += 1
                        formatting_tags.append('s')
                        logger.debug(f"Applied strikethrough formatting (element) to text: '{text[:30]}...'")
                else:
                    logger.debug(f"No strikethrough formatting found for text: '{text[:30]}...'")
            
            # Check for highlight (background fill)
            highlight_elem = self.xml_parser.find_element_with_namespace(r_pr, './/a:highlight')
            if highlight_elem is not None:
                text_element.highlighted += 1
                formatting_tags.append('mark')
            
            # Apply formatting tags
            if formatting_tags:
                for tag in formatting_tags:
                    formatted_text = f"<{tag}>{formatted_text}</{tag}>"
            
            return formatted_text
            
        except Exception as e:
            logger.warning(f"Failed to apply text formatting: {e}")
            return text
    
    def extract_text_elements(self, slide_xml_content: str, slide_number: int) -> List[Dict[str, Any]]:
        """
        Extract all text elements with formatting from a slide.
        
        Args:
            slide_xml_content: XML content of the slide
            slide_number: Slide number (1-based)
            
        Returns:
            List of text element dictionaries
        """
        try:
            slide_info = self.extract_slide_content(slide_xml_content, slide_number)
            return slide_info.text_elements
            
        except Exception as e:
            logger.error(f"Failed to extract text elements for slide {slide_number}: {e}")
            return []
    
    def _resolve_hyperlink_relationships(self, extractor, slide_number: int, text_elements: List[Dict[str, Any]]) -> None:
        """
        Resolve hyperlink relationship IDs to actual URLs.
        
        Args:
            extractor: ZipExtractor instance
            slide_number: Slide number (1-based)
            text_elements: List of text elements to update
        """
        try:
            # Read the slide relationships file
            rels_file = f'ppt/slides/_rels/slide{slide_number}.xml.rels'
            rels_content = extractor.read_xml_content(rels_file)
            
            if not rels_content:
                logger.debug(f"No relationships file found for slide {slide_number}")
                return
            
            # Parse relationships
            rels_root = self.xml_parser.parse_xml_string(rels_content)
            relationships = {}
            
            # Extract all relationships
            rel_elements = self.xml_parser.find_elements_with_namespace(
                rels_root, './/r:Relationship',
                {'r': 'http://schemas.openxmlformats.org/package/2006/relationships'}
            )
            
            for rel in rel_elements:
                rel_id = rel.get('Id')
                target = rel.get('Target')
                rel_type = rel.get('Type')
                
                if rel_id and target:
                    relationships[rel_id] = {
                        'target': target,
                        'type': rel_type
                    }
                    logger.debug(f"Found relationship {rel_id} -> {target}")
            
            # Resolve hyperlinks in text elements
            for text_elem in text_elements:
                if 'hyperlinks' in text_elem and text_elem['hyperlinks']:
                    resolved_links = []
                    for link in text_elem['hyperlinks']:
                        if link in relationships:
                            target = relationships[link]['target']
                            resolved_links.append(target)
                            logger.debug(f"Resolved hyperlink {link} to {target}")
                        else:
                            # Keep original if not found in relationships
                            resolved_links.append(link)
                            logger.debug(f"Could not resolve hyperlink {link}")
                    text_elem['hyperlinks'] = resolved_links
                    
        except Exception as e:
            logger.warning(f"Failed to resolve hyperlink relationships for slide {slide_number}: {e}")

    def _extract_cell_text_content(self, cell) -> Optional[str]:
        """
        Extract text content from a table cell.
        
        Args:
            cell: Table cell element
            
        Returns:
            Text content if available, None otherwise
        """
        try:
            # Find text body in cell
            tx_body = self.xml_parser.find_element_with_namespace(cell, './/a:txBody')
            if tx_body is None:
                return None
            
            # Extract all text from paragraphs
            text_parts = []
            paragraphs = self.xml_parser.find_elements_with_namespace(tx_body, './/a:p')
            
            for paragraph in paragraphs:
                # Get text from all runs in the paragraph
                runs = self.xml_parser.find_elements_with_namespace(paragraph, './/a:r')
                paragraph_text = []
                
                for run in runs:
                    text_elem = self.xml_parser.find_element_with_namespace(run, './/a:t')
                    if text_elem is not None and text_elem.text:
                        paragraph_text.append(text_elem.text)
                
                if paragraph_text:
                    text_parts.append(''.join(paragraph_text))
            
            return '\n'.join(text_parts) if text_parts else None
            
        except Exception as e:
            logger.warning(f"Failed to extract cell text content: {e}")
            return None

    def extract_formatted_text(self, slide_xml_content: str) -> Dict[str, Any]:
        """
        Extract formatted and plain text content from a slide.
        
        Args:
            slide_xml_content: XML content of the slide
            
        Returns:
            Dictionary containing formatted and plain text content
        """
        try:
            root = self.xml_parser.parse_xml_string(slide_xml_content)
            if root is None:
                return {'plain_text': '', 'formatted_text': '', 'text_elements': []}
            
            # Create temporary slide info to collect text elements
            temp_slide_info = SlideInfo(slide_number=0)
            self._extract_text_elements(root, temp_slide_info)
            
            # Combine all text elements
            all_plain_text = []
            all_formatted_text = []
            
            for text_elem in temp_slide_info.text_elements:
                if text_elem['content_plain'].strip():
                    all_plain_text.append(text_elem['content_plain'])
                    all_formatted_text.append(text_elem['content_formatted'])
            
            return {
                'plain_text': '\n\n'.join(all_plain_text),
                'formatted_text': '\n\n'.join(all_formatted_text),
                'text_elements': temp_slide_info.text_elements
            }
            
        except Exception as e:
            logger.error(f"Failed to extract formatted text: {e}")
            return {'plain_text': '', 'formatted_text': '', 'text_elements': []}
    
    def _extract_tables(self, root: ET.Element, slide_info: SlideInfo) -> None:
        """
        Extract table data from slide XML.
        
        Args:
            root: Root element of slide XML
            slide_info: SlideInfo object to populate
        """
        try:
            # Find all graphic frames that might contain tables
            graphic_frames = self.xml_parser.find_elements_with_namespace(root, './/p:graphicFrame')
            
            for graphic_frame in graphic_frames:
                table_data = self._extract_table_from_graphic_frame(graphic_frame)
                if table_data:
                    slide_info.tables.append(table_data)
            
        except Exception as e:
            logger.warning(f"Failed to extract tables for slide {slide_info.slide_number}: {e}")
    
    def _extract_table_from_graphic_frame(self, graphic_frame: ET.Element) -> Optional[Dict[str, Any]]:
        """
        Extract table data from a graphic frame element.
        
        Args:
            graphic_frame: Graphic frame element that might contain a table
            
        Returns:
            Dictionary containing table data if found, None otherwise
        """
        try:
            # Check if this graphic frame contains a table
            table_elem = self.xml_parser.find_element_with_namespace(
                graphic_frame, './/a:tbl'
            )
            
            if table_elem is None:
                return None
            
            # Extract position and size from transform
            position, size = self._extract_graphic_frame_transform(graphic_frame)
            
            # Extract table structure
            table = self._parse_table_structure(table_elem)
            if table is None:
                return None
            
            return {
                'rows': table.rows,
                'columns': table.columns,
                'cells': [[{
                    'content': cell.content,
                    'row_span': cell.row_span,
                    'col_span': cell.col_span,
                    'formatting': cell.formatting
                } for cell in row] for row in table.cells],
                'position': position,
                'size': size
            }
            
        except Exception as e:
            logger.warning(f"Failed to extract table from graphic frame: {e}")
            return None
    
    def _extract_graphic_frame_transform(self, graphic_frame: ET.Element) -> Tuple[Tuple[int, int], Tuple[int, int]]:
        """
        Extract position and size from graphic frame transform.
        
        Args:
            graphic_frame: Graphic frame element
            
        Returns:
            Tuple of (position, size) where each is (x, y) or (width, height)
        """
        try:
            # Find transform element - it might be directly under graphicFrame
            xfrm = self.xml_parser.find_element_with_namespace(graphic_frame, './/p:xfrm')
            if xfrm is None:
                # Try alternative path
                xfrm = self.xml_parser.find_element_with_namespace(graphic_frame, './/a:xfrm')
            
            if xfrm is None:
                return (0, 0), (0, 0)
            
            # Extract offset (position)
            off = self.xml_parser.find_element_with_namespace(xfrm, './/a:off')
            position = (0, 0)
            if off is not None:
                x = int(off.get('x', '0'))
                y = int(off.get('y', '0'))
                position = (x, y)
            
            # Extract extent (size)
            ext = self.xml_parser.find_element_with_namespace(xfrm, './/a:ext')
            size = (0, 0)
            if ext is not None:
                cx = int(ext.get('cx', '0'))
                cy = int(ext.get('cy', '0'))
                size = (cx, cy)
            
            return position, size
            
        except Exception as e:
            logger.warning(f"Failed to extract graphic frame transform: {e}")
            return (0, 0), (0, 0)
    
    def _parse_table_structure(self, table_elem: ET.Element) -> Optional[Table]:
        """
        Parse table structure from table element.
        
        Args:
            table_elem: Table element
            
        Returns:
            Table object with parsed structure
        """
        try:
            # Find all table rows
            rows = self.xml_parser.find_elements_with_namespace(table_elem, './/a:tr')
            if not rows:
                return None
            
            table_rows = []
            max_columns = 0
            
            for row_elem in rows:
                # Find all cells in this row
                cells = self.xml_parser.find_elements_with_namespace(row_elem, './/a:tc')
                row_cells = []
                
                for cell_elem in cells:
                    cell = self._parse_table_cell(cell_elem)
                    row_cells.append(cell)
                
                table_rows.append(row_cells)
                max_columns = max(max_columns, len(row_cells))
            
            # Pad rows to have consistent column count
            for row in table_rows:
                while len(row) < max_columns:
                    row.append(TableCell(content="", row_span=1, col_span=1))
            
            return Table(
                rows=len(table_rows),
                columns=max_columns,
                cells=table_rows
            )
            
        except Exception as e:
            logger.warning(f"Failed to parse table structure: {e}")
            return None
    
    def _parse_table_cell(self, cell_elem: ET.Element) -> TableCell:
        """
        Parse a single table cell.
        
        Args:
            cell_elem: Table cell element
            
        Returns:
            TableCell object with parsed content and formatting
        """
        try:
            # Extract cell content
            content = self._extract_cell_text_content(cell_elem)
            
            # Extract row span and column span
            row_span = int(cell_elem.get('rowSpan', '1'))
            col_span = int(cell_elem.get('gridSpan', '1'))
            
            # Extract cell formatting
            formatting = self._extract_cell_formatting(cell_elem)
            
            return TableCell(
                content=content,
                row_span=row_span,
                col_span=col_span,
                formatting=formatting
            )
            
        except Exception as e:
            logger.warning(f"Failed to parse table cell: {e}")
            return TableCell(content="", row_span=1, col_span=1)
    
    def _extract_cell_text_content(self, cell_elem: ET.Element) -> str:
        """
        Extract text content from a table cell.
        
        Args:
            cell_elem: Table cell element
            
        Returns:
            Text content of the cell
        """
        try:
            # Find text body in the cell
            tx_body = self.xml_parser.find_element_with_namespace(cell_elem, './/a:txBody')
            if tx_body is None:
                return ""
            
            # Extract all text from paragraphs
            text_parts = []
            paragraphs = self.xml_parser.find_elements_with_namespace(tx_body, './/a:p')
            
            for paragraph in paragraphs:
                # Get text from all runs in the paragraph
                runs = self.xml_parser.find_elements_with_namespace(paragraph, './/a:r')
                paragraph_text = []
                
                for run in runs:
                    text_elem = self.xml_parser.find_element_with_namespace(run, './/a:t')
                    if text_elem is not None and text_elem.text:
                        paragraph_text.append(text_elem.text)
                
                if paragraph_text:
                    text_parts.append(''.join(paragraph_text))
            
            return '\n'.join(text_parts) if text_parts else ""
            
        except Exception as e:
            logger.warning(f"Failed to extract cell text content: {e}")
            return ""
    
    def _extract_cell_formatting(self, cell_elem: ET.Element) -> Dict[str, Any]:
        """
        Extract formatting information from a table cell.
        
        Args:
            cell_elem: Table cell element
            
        Returns:
            Dictionary containing formatting information
        """
        try:
            formatting = {}
            
            # Extract cell properties
            tc_pr = self.xml_parser.find_element_with_namespace(cell_elem, './/a:tcPr')
            if tc_pr is not None:
                # Extract fill color
                solid_fill = self.xml_parser.find_element_with_namespace(tc_pr, './/a:solidFill')
                if solid_fill is not None:
                    # Look for RGB color
                    srgb_clr = self.xml_parser.find_element_with_namespace(solid_fill, './/a:srgbClr')
                    if srgb_clr is not None:
                        color_val = srgb_clr.get('val')
                        if color_val:
                            formatting['fill_color'] = f"#{color_val}"
                    
                    # Look for scheme color
                    scheme_clr = self.xml_parser.find_element_with_namespace(solid_fill, './/a:schemeClr')
                    if scheme_clr is not None:
                        color_val = scheme_clr.get('val')
                        if color_val:
                            formatting['fill_color'] = color_val
                
                # Extract border information
                borders = ['lnL', 'lnR', 'lnT', 'lnB']  # left, right, top, bottom
                border_info = {}
                
                for border in borders:
                    border_elem = self.xml_parser.find_element_with_namespace(tc_pr, f'.//a:{border}')
                    if border_elem is not None:
                        width = border_elem.get('w', '0')
                        border_info[border] = {'width': int(width)}
                
                if border_info:
                    formatting['borders'] = border_info
            
            return formatting
            
        except Exception as e:
            logger.warning(f"Failed to extract cell formatting: {e}")
            return {}
    
    def extract_table_data(self, slide_xml_content: str, slide_number: int) -> List[Dict[str, Any]]:
        """
        Extract all table data from a slide.
        
        Args:
            slide_xml_content: XML content of the slide
            slide_number: Slide number (1-based)
            
        Returns:
            List of table dictionaries
        """
        try:
            slide_info = self.extract_slide_content(slide_xml_content, slide_number)
            return slide_info.tables
            
        except Exception as e:
            logger.error(f"Failed to extract table data for slide {slide_number}: {e}")
            return []
    
    def extract_tables_with_structure(self, slide_xml_content: str) -> Dict[str, Any]:
        """
        Extract table data with detailed structure information.
        
        Args:
            slide_xml_content: XML content of the slide
            
        Returns:
            Dictionary containing table structure and content
        """
        try:
            root = self.xml_parser.parse_xml_string(slide_xml_content)
            if root is None:
                return {'tables': [], 'table_count': 0}
            
            # Create temporary slide info to collect tables
            temp_slide_info = SlideInfo(slide_number=0)
            self._extract_tables(root, temp_slide_info)
            
            return {
                'tables': temp_slide_info.tables,
                'table_count': len(temp_slide_info.tables)
            }
            
        except Exception as e:
            logger.error(f"Failed to extract table structure: {e}")
            return {'tables': [], 'table_count': 0}   
 
    def extract_presentation_metadata(self, presentation_xml_content: str) -> Dict[str, Any]:
        """
        Extract presentation-level metadata from presentation.xml.
        
        Args:
            presentation_xml_content: XML content of presentation.xml
            
        Returns:
            Dictionary containing presentation metadata
        """
        try:
            presentation_data = self.xml_parser.parse_presentation_xml(presentation_xml_content)
            
            metadata = {
                'slide_count': len(presentation_data.get('slide_ids', [])),
                'slide_size': presentation_data.get('slide_size'),
                'slide_master_count': len(presentation_data.get('slide_master_ids', [])),
                'has_notes_master': presentation_data.get('notes_master_id') is not None,
                'has_handout_master': presentation_data.get('handout_master_id') is not None,
                'slide_ids': presentation_data.get('slide_ids', []),
                'slide_master_ids': presentation_data.get('slide_master_ids', [])
            }
            
            return metadata
            
        except Exception as e:
            logger.error(f"Failed to extract presentation metadata: {e}")
            return {}
    
    def extract_slide_metadata(self, slide_xml_content: str, slide_number: int, notes_xml_content: str = None) -> Dict[str, Any]:
        """
        Extract metadata for a single slide including notes and object counts.
        
        Args:
            slide_xml_content: XML content of the slide
            slide_number: Slide number (1-based)
            notes_xml_content: Optional XML content of slide notes
            
        Returns:
            Dictionary containing slide metadata
        """
        try:
            root = self.xml_parser.parse_xml_string(slide_xml_content)
            if root is None:
                return {'slide_number': slide_number, 'error': 'Failed to parse slide XML'}
            
            # Extract basic slide info
            slide_info = self.extract_slide_content(slide_xml_content, slide_number)
            
            # Count objects on the slide
            object_counts = self._count_slide_objects(root)
            
            # Extract notes content if provided
            notes_content = ""
            if notes_xml_content:
                notes_content = self._extract_notes_content(notes_xml_content)
            
            metadata = {
                'slide_number': slide_number,
                'layout_name': slide_info.layout_name,
                'layout_type': slide_info.layout_type,
                'title': slide_info.title,
                'subtitle': slide_info.subtitle,
                'notes': notes_content,
                'object_counts': object_counts,
                'placeholder_count': len(slide_info.placeholders),
                'text_element_count': len(slide_info.text_elements),
                'table_count': len(slide_info.tables)
            }
            
            return metadata
            
        except Exception as e:
            logger.error(f"Failed to extract slide metadata for slide {slide_number}: {e}")
            return {'slide_number': slide_number, 'error': str(e)}
    
    def _count_slide_objects(self, root: ET.Element) -> Dict[str, int]:
        """
        Count different types of objects on a slide.
        
        Args:
            root: Root element of slide XML
            
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
            
            # Count shapes (text boxes, basic shapes) - exclude shapes in groups
            shapes = self.xml_parser.find_elements_with_namespace(root, './/p:spTree/p:sp')
            for shape in shapes:
                counts['shapes'] += 1
                
                # Check if it's a text box (has text body)
                tx_body = self.xml_parser.find_element_with_namespace(shape, './/p:txBody')
                if tx_body is not None:
                    counts['text_boxes'] += 1
            
            # Count images
            pics = self.xml_parser.find_elements_with_namespace(root, './/p:pic')
            counts['images'] = len(pics)
            
            # Count tables
            tables = self.xml_parser.find_elements_with_namespace(root, './/a:tbl')
            counts['tables'] = len(tables)
            
            # Count charts (look for chart elements in graphic data)
            graphic_frames = self.xml_parser.find_elements_with_namespace(root, './/p:graphicFrame')
            for frame in graphic_frames:
                # Check if this frame contains a chart
                graphic_data = self.xml_parser.find_element_with_namespace(frame, './/a:graphicData')
                if graphic_data is not None:
                    # Look for chart elements (they might have different namespaces)
                    chart_elems = graphic_data.findall('.//*')
                    for elem in chart_elems:
                        if 'chart' in elem.tag.lower():
                            counts['charts'] += 1
                            break
            
            # Count media objects (audio, video)
            media = self.xml_parser.find_elements_with_namespace(root, './/p:media')
            counts['media'] = len(media)
            
            # Count connectors
            connectors = self.xml_parser.find_elements_with_namespace(root, './/p:cxnSp')
            counts['connectors'] = len(connectors)
            
            # Count groups
            groups = self.xml_parser.find_elements_with_namespace(root, './/p:grpSp')
            counts['groups'] = len(groups)
            
            return counts
            
        except Exception as e:
            logger.warning(f"Failed to count slide objects: {e}")
            return {}
    
    def _extract_notes_content(self, notes_xml_content: str) -> str:
        """
        Extract speaker notes content from notes XML.
        
        Args:
            notes_xml_content: XML content of slide notes
            
        Returns:
            Notes content as string
        """
        try:
            root = self.xml_parser.parse_xml_string(notes_xml_content)
            if root is None:
                return ""
            
            # Find all text in the notes
            text_parts = []
            
            # Look for text in shapes
            shapes = self.xml_parser.find_elements_with_namespace(root, './/p:sp')
            for shape in shapes:
                # Skip the slide thumbnail shape
                nv_sp_pr = self.xml_parser.find_element_with_namespace(shape, './/p:nvSpPr')
                if nv_sp_pr is not None:
                    ph = self.xml_parser.find_element_with_namespace(nv_sp_pr, './/p:ph')
                    if ph is not None:
                        ph_type = ph.get('type')
                        # Skip slide image placeholder
                        if ph_type == 'sldImg':
                            continue
                
                # Extract text content
                content = self._extract_shape_text_content(shape)
                if content and content.strip():
                    text_parts.append(content.strip())
            
            return '\n\n'.join(text_parts)
            
        except Exception as e:
            logger.warning(f"Failed to extract notes content: {e}")
            return ""
    
    def extract_section_information(self, presentation_xml_content: str) -> List[Dict[str, Any]]:
        """
        Extract section information from presentation XML.
        
        Args:
            presentation_xml_content: XML content of presentation.xml
            
        Returns:
            List of section dictionaries
        """
        try:
            # Parse XML directly with ElementTree
            root = ET.fromstring(presentation_xml_content)
            
            sections = []
            
            # Define namespaces for section detection
            namespaces = {
                'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
                'p14': 'http://schemas.microsoft.com/office/powerpoint/2010/main',
                'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
                'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
            }
            
            # Look for section list in both standard and PowerPoint 2010+ namespaces
            section_list = None
            
            # Try standard namespace first
            section_list = root.find('.//p:sectionLst', namespaces)
            
            # If not found, try PowerPoint 2010+ namespace
            if section_list is None:
                section_list = root.find('.//p14:sectionLst', namespaces)
            
            # Also try searching without namespace prefix (in case of namespace issues)
            if section_list is None:
                for elem in root.iter():
                    if elem.tag.endswith('}sectionLst') or elem.tag == 'sectionLst':
                        section_list = elem
                        break
            
            if section_list is not None:
                logger.debug(f"Found section list: {section_list.tag}")
                
                # Try both namespaces for section elements
                section_elements = section_list.findall('.//p:section', namespaces)
                if not section_elements:
                    section_elements = section_list.findall('.//p14:section', namespaces)
                
                # Also try searching without namespace prefix
                if not section_elements:
                    for elem in section_list.iter():
                        if elem.tag.endswith('}section') or elem.tag == 'section':
                            section_elements.append(elem)
                
                logger.debug(f"Found {len(section_elements)} section elements")
                
                for section_elem in section_elements:
                    section_name = section_elem.get('name', 'Unnamed Section')
                    section_id = section_elem.get('id', '')
                    
                    logger.debug(f"Processing section: name='{section_name}', id='{section_id}'")
                    
                    # Look for slide references in this section
                    slide_refs = section_elem.findall('.//p:sldId', namespaces)
                    if not slide_refs:
                        slide_refs = section_elem.findall('.//p14:sldId', namespaces)
                    
                    # Also try searching without namespace prefix
                    if not slide_refs:
                        for elem in section_elem.iter():
                            if elem.tag.endswith('}sldId') or elem.tag == 'sldId':
                                slide_refs.append(elem)
                    
                    slide_count = len(slide_refs)
                    slide_ids = []
                    for slide_ref in slide_refs:
                        slide_id = slide_ref.get('id', '')
                        r_id = slide_ref.get('r:id', '')
                        if slide_id or r_id:
                            slide_ids.append({
                                'id': slide_id,
                                'r_id': r_id
                            })
                    
                    sections.append({
                        'name': section_name,
                        'id': section_id,
                        'slide_count': slide_count,
                        'slide_ids': slide_ids
                    })
            else:
                logger.debug("No section list found in presentation XML")
            
            return sections
            
        except Exception as e:
            logger.warning(f"Failed to extract section information: {e}")
            import traceback
            logger.debug(f"Section extraction traceback: {traceback.format_exc()}")
            return []
    
    def get_slide_size_info(self, presentation_xml_content: str) -> Dict[str, Any]:
        """
        Extract slide size information from presentation XML.
        
        Args:
            presentation_xml_content: XML content of presentation.xml
            
        Returns:
            Dictionary containing slide size information
        """
        try:
            presentation_data = self.xml_parser.parse_presentation_xml(presentation_xml_content)
            slide_size = presentation_data.get('slide_size')
            
            if slide_size:
                # Convert from EMUs (English Metric Units) to more readable formats
                width_emu = slide_size['width']
                height_emu = slide_size['height']
                
                # Convert to inches (1 inch = 914400 EMUs)
                width_inches = width_emu / 914400
                height_inches = height_emu / 914400
                
                # Convert to centimeters (1 inch = 2.54 cm)
                width_cm = width_inches * 2.54
                height_cm = height_inches * 2.54
                
                # Convert to points (1 inch = 72 points)
                width_points = width_inches * 72
                height_points = height_inches * 72
                
                return {
                    'width_emu': width_emu,
                    'height_emu': height_emu,
                    'width_inches': round(width_inches, 2),
                    'height_inches': round(height_inches, 2),
                    'width_cm': round(width_cm, 2),
                    'height_cm': round(height_cm, 2),
                    'width_points': round(width_points, 1),
                    'height_points': round(height_points, 1),
                    'aspect_ratio': round(width_inches / height_inches, 2) if height_inches > 0 else 0
                }
            
            return {}
            
        except Exception as e:
            logger.warning(f"Failed to extract slide size info: {e}")
            return {}
    
    def clear_cache(self) -> None:
        """Clear the content extraction cache."""
        if self.enable_caching and self.cache_manager:
            self.cache_manager.clear()
            logger.debug("Cleared content extraction cache")
    
    def get_cache_stats(self) -> Dict[str, Any]:
        """
        Get cache statistics for content extraction.
        
        Returns:
            Dictionary with cache statistics
        """
        if self.enable_caching and self.cache_manager:
            stats = self.cache_manager.get_cache_stats()
            xml_stats = self.xml_parser.get_cache_stats()
            return {
                'content_cache': stats,
                'xml_parser_cache': xml_stats,
                'caching_enabled': True
            }
        else:
            return {
                'content_cache': {'total_entries': 0},
                'xml_parser_cache': {'cached_elements': 0},
                'caching_enabled': False
            }
    
    def cleanup_expired_cache(self) -> int:
        """
        Clean up expired cache entries.
        
        Returns:
            Number of entries removed
        """
        if self.enable_caching and self.cache_manager:
            removed = self.cache_manager.cleanup_expired()
            self.xml_parser.clear_element_cache()
            logger.debug(f"Cleaned up {removed} expired cache entries")
            return removed
        return 0
    
    def resolve_hyperlinks(self, slide_info: SlideInfo, slide_rels_content: Optional[str]) -> None:
        """
        Resolve hyperlink relationship IDs to actual URLs.
        
        Args:
            slide_info: SlideInfo object to update with resolved hyperlinks
            slide_rels_content: XML content of slide relationships file
        """
        if not slide_rels_content:
            return
            
        try:
            # Parse the relationships XML
            rels_root = ET.fromstring(slide_rels_content)
            
            # Build a mapping of relationship IDs to targets
            rel_map = {}
            for rel in rels_root.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                rel_id = rel.get('Id')
                target = rel.get('Target')
                rel_type = rel.get('Type')
                
                if rel_id and target and 'hyperlink' in rel_type.lower():
                    rel_map[rel_id] = target
            
            # Resolve hyperlinks in text elements
            for text_element in slide_info.text_elements:
                if text_element.get('hyperlinks'):
                    resolved_links = []
                    for link_id in text_element['hyperlinks']:
                        if link_id in rel_map:
                            resolved_links.append(rel_map[link_id])
                        else:
                            resolved_links.append(link_id)  # Keep original if not found
                    text_element['hyperlinks'] = resolved_links
                    
        except Exception as e:
            logger.warning(f"Failed to resolve hyperlinks: {e}")
    
    def _resolve_hyperlink_relationships(self, extractor, slide_number: int, text_elements: List[Dict[str, Any]]) -> None:
        """
        Resolve hyperlink relationship IDs to actual URLs using slide relationships.
        
        Args:
            extractor: ZipExtractor instance
            slide_number: Slide number (1-based)
            text_elements: List of text elements to update with resolved hyperlinks
        """
        try:
            # Get slide relationships file
            rels_file = f'ppt/slides/_rels/slide{slide_number}.xml.rels'
            rels_content = extractor.read_xml_content(rels_file)
            
            if not rels_content:
                logger.debug(f"No relationships file found for slide {slide_number}")
                return
            
            # Parse the relationships XML
            rels_root = ET.fromstring(rels_content)
            
            # Build a mapping of relationship IDs to targets
            rel_map = {}
            for rel in rels_root.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                rel_id = rel.get('Id')
                target = rel.get('Target')
                rel_type = rel.get('Type')
                
                if rel_id and target and 'hyperlink' in rel_type.lower():
                    rel_map[rel_id] = target
                    logger.debug(f"Found hyperlink relationship: {rel_id} -> {target}")
            
            # Resolve hyperlinks in text elements
            for text_element in text_elements:
                # Handle both dictionary and object formats
                hyperlinks = None
                if isinstance(text_element, dict):
                    hyperlinks = text_element.get('hyperlinks')
                else:
                    # Assume it's a TextElement object
                    hyperlinks = getattr(text_element, 'hyperlinks', None)
                
                if hyperlinks:
                    resolved_links = []
                    for link_id in hyperlinks:
                        if link_id in rel_map:
                            resolved_links.append(rel_map[link_id])
                            logger.debug(f"Resolved hyperlink {link_id} to {rel_map[link_id]}")
                        else:
                            resolved_links.append(link_id)  # Keep original if not found
                            logger.debug(f"Could not resolve hyperlink {link_id}")
                    
                    # Update the hyperlinks
                    if isinstance(text_element, dict):
                        text_element['hyperlinks'] = resolved_links
                    else:
                        text_element.hyperlinks = resolved_links
                    
        except Exception as e:
            logger.warning(f"Failed to resolve hyperlink relationships for slide {slide_number}: {e}")
    
    def _get_default_font_size(self, shape: ET.Element) -> float:
        """
        Get context-aware default font size based on shape type.
        
        Args:
            shape: Shape element
            
        Returns:
            Default font size in points
        """
        try:
            # Check if this is a title placeholder
            nv_sp_pr = self.xml_parser.find_element_with_namespace(shape, './/p:nvSpPr')
            if nv_sp_pr is not None:
                ph = self.xml_parser.find_element_with_namespace(nv_sp_pr, './/p:ph')
                if ph is not None:
                    placeholder_type = ph.get('type', 'content')
                    if placeholder_type in ['title', 'ctrTitle']:
                        return 44.0  # Default title font size
                    elif placeholder_type in ['subTitle']:
                        return 24.0  # Default subtitle font size
            
            # Default content font size
            return 18.0
            
        except Exception as e:
            logger.debug(f"Failed to determine context for default font size: {e}")
            return 18.0
    
    def extract_notes(self, extractor) -> List[Dict[str, Any]]:
        """
        Extract notes from the PowerPoint file.
        
        Args:
            extractor: ZipExtractor instance
            
        Returns:
            List of note dictionaries
        """
        notes = []
        
        try:
            logger.info("Starting notes extraction process")
            # Build a mapping of notes files to slide numbers using relationship files
            notes_to_slide_map = self._build_notes_slide_mapping(extractor)
            
            # Look for notes files
            notes_files = []
            for filename in extractor.list_archive_contents():
                if filename.startswith('ppt/notesSlides/notesSlide') and filename.endswith('.xml'):
                    notes_files.append(filename)
                    logger.info(f"Found notes file: {filename}")
            
            for notes_file in notes_files:
                notes_content = extractor.read_xml_content(notes_file)
                if notes_content:
                    # Get the correct slide number for this notes file using relationship mapping
                    slide_number = notes_to_slide_map.get(notes_file)
                    if slide_number is None:
                        # Fallback: extract slide number from filename if mapping fails
                        notes_number_match = re.search(r'notesSlide(\d+)\.xml', notes_file)
                        if notes_number_match:
                            slide_number = int(notes_number_match.group(1))
                        else:
                            slide_number = 1
                    
                    parsed_notes = self._parse_notes_content(notes_content, slide_number)
                    if parsed_notes:
                        notes.append({
                            'slide_number': slide_number,
                            'content': parsed_notes
                        })
            
            # Also check for comments embedded in slide files
            slide_files = []
            for filename in extractor.list_archive_contents():
                if filename.startswith('ppt/slides/slide') and filename.endswith('.xml'):
                    slide_files.append(filename)
            
            for slide_file in slide_files:
                slide_content = extractor.read_xml_content(slide_file)
                if slide_content:
                    # Extract slide number from filename
                    slide_number_match = re.search(r'slide(\d+)\.xml$', slide_file)
                    if slide_number_match:
                        slide_number = int(slide_number_match.group(1))

            
        except Exception as e:
            logger.warning(f"Failed to extract notes: {e}")
        
        logger.info(f"Notes extraction completed. Found {len(notes)} notes")
        return notes
    
    def _build_notes_slide_mapping(self, extractor) -> Dict[str, int]:
        """
        Build a mapping of comment files to slide numbers by examining slide relationships.
        
        Args:
            extractor: ZipExtractor instance
            
        Returns:
            Dictionary mapping comment file paths to slide numbers
        """
        comment_to_slide_map = {}
        
        try:
            # Examine each slide's relationship file
            for filename in extractor.list_archive_contents():
                if filename.startswith('ppt/slides/_rels/slide') and filename.endswith('.xml.rels'):
                    # Extract slide number from filename (e.g., 'ppt/slides/_rels/slide3.xml.rels' -> 3)
                    slide_filename = filename
                    slide_number_match = re.search(r'slide(\d+)\.xml\.rels$', slide_filename)
                    if not slide_number_match:
                        continue
                    
                    slide_number = int(slide_number_match.group(1))
                    
                    # Read the relationships file
                    rels_content = extractor.read_xml_content(slide_filename)
                    if not rels_content:
                        continue
                    
                    # Parse relationships and look for comment references
                    try:
                        rels_root = ET.fromstring(rels_content)
                        for rel in rels_root.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                            rel_type = rel.get('Type', '')
                            target = rel.get('Target', '')
                            
                            # Check if this is a comment relationship
                            if 'comments' in rel_type.lower() and target:
                                # Convert relative path to absolute path
                                # Target is like '../comments/comment1.xml'
                                if target.startswith('../'):
                                    comment_file_path = 'ppt/' + target[3:]  # Remove '../' and add 'ppt/'
                                else:
                                    comment_file_path = target
                                
                                comment_to_slide_map[comment_file_path] = slide_number
                                logger.debug(f"Found comment relationship: {comment_file_path} -> slide {slide_number}")
                    
                    except Exception as e:
                        logger.warning(f"Failed to parse relationships file {slide_filename}: {e}")
        
        except Exception as e:
            logger.warning(f"Failed to build comment-slide mapping: {e}")
        
        return comment_to_slide_map
    
    def _parse_comment_file(self, comment_content: str, slide_number: int = 1) -> List[Dict[str, Any]]:
        """
        Parse a comment XML file and extract comment information.
        
        Args:
            comment_content: XML content of the comment file
            slide_number: The slide number this comment belongs to
            
        Returns:
            List of comment dictionaries
        """
        comments = []
        
        try:
            root = ET.fromstring(comment_content)
            
            # Find all comment elements (support multiple formats)
            comment_patterns = [
                './/{http://schemas.openxmlformats.org/presentationml/2006/main}cm',
                './/{http://schemas.microsoft.com/office/powerpoint/2018/main}threadedComment',
                './/{http://schemas.openxmlformats.org/presentationml/2006/main}comment'
            ]
            
            for pattern in comment_patterns:
                for cm in root.findall(pattern):
                    comment_data = {
                        'slide_number': slide_number,  # Use the correct slide number
                        'author_id': cm.get('authorId', cm.get('author', '')),
                        'datetime': cm.get('dt', cm.get('created', '')),
                        'index': cm.get('idx', cm.get('id', '')),
                        'position': {'x': 0, 'y': 0},
                        'text': ''
                    }
                    
                    # Extract position
                    pos = cm.find('.//{http://schemas.openxmlformats.org/presentationml/2006/main}pos')
                    if pos is not None:
                        comment_data['position'] = {
                            'x': int(pos.get('x', 0)),
                            'y': int(pos.get('y', 0))
                        }
                    
                    # Extract text from various possible locations
                    text_patterns = [
                        './/{http://schemas.openxmlformats.org/presentationml/2006/main}text',
                        './/{http://schemas.microsoft.com/office/powerpoint/2018/main}text',
                        './/text'
                    ]
                    
                    for text_pattern in text_patterns:
                        text_elem = cm.find(text_pattern)
                        if text_elem is not None and text_elem.text:
                            comment_data['text'] = text_elem.text
                            break
                    
                    if comment_data['text']:  # Only add if we found text
                        comments.append(comment_data)
        
        except Exception as e:
            logger.warning(f"Failed to parse comment file: {e}")
        
        return comments
    
    def _parse_embedded_comments(self, slide_content: str, slide_number: int) -> List[Dict[str, Any]]:
        """
        Parse comments that are embedded directly in slide XML.
        
        Args:
            slide_content: XML content of the slide
            slide_number: The slide number
            
        Returns:
            List of comment dictionaries
        """
        comments = []
        
        try:
            root = ET.fromstring(slide_content)
            
            # Look for various comment element patterns
            comment_patterns = [
                './/{http://schemas.openxmlformats.org/presentationml/2006/main}cm',
                './/{http://schemas.openxmlformats.org/presentationml/2006/main}comment',
                './/{http://schemas.microsoft.com/office/powerpoint/2018/main}threadedComment'
            ]
            
            for pattern in comment_patterns:
                comment_elements = root.findall(pattern)
                for cm in comment_elements:
                    comment_data = {
                        'slide_number': slide_number,
                        'author_id': cm.get('authorId', cm.get('author', '')),
                        'datetime': cm.get('dt', cm.get('created', '')),
                        'index': cm.get('idx', cm.get('id', '')),
                        'position': {'x': 0, 'y': 0},
                        'text': ''
                    }
                    
                    # Extract position
                    pos = cm.find('.//{http://schemas.openxmlformats.org/presentationml/2006/main}pos')
                    if pos is not None:
                        comment_data['position'] = {
                            'x': int(pos.get('x', 0)),
                            'y': int(pos.get('y', 0))
                        }
                    
                    # Extract text from various possible locations
                    text_patterns = [
                        './/{http://schemas.openxmlformats.org/presentationml/2006/main}text',
                        './/{http://schemas.microsoft.com/office/powerpoint/2018/main}text',
                        './/text'
                    ]
                    
                    for text_pattern in text_patterns:
                        text_elem = cm.find(text_pattern)
                        if text_elem is not None and text_elem.text:
                            comment_data['text'] = text_elem.text
                            break
                    
                    if comment_data['text']:  # Only add if we found text
                        comments.append(comment_data)
                        logger.debug(f"Found embedded comment on slide {slide_number}: {comment_data['text']}")
        
        except Exception as e:
            logger.warning(f"Failed to parse embedded comments for slide {slide_number}: {e}")
        
        return comments    
    def _build_notes_slide_mapping(self, extractor) -> Dict[str, int]:
        """
        Build a mapping of notes slide files to slide numbers by examining notes relationships.
        
        Args:
            extractor: ZipExtractor instance
            
        Returns:
            Dictionary mapping notes slide file paths to slide numbers
        """
        notes_to_slide_map = {}
        
        try:
            # Examine each notes slide's relationship file
            for filename in extractor.list_archive_contents():
                if filename.startswith('ppt/notesSlides/_rels/notesSlide') and filename.endswith('.xml.rels'):
                    # Extract notes slide number from filename
                    notes_filename = filename
                    notes_number_match = re.search(r'notesSlide(\d+)\.xml\.rels$', notes_filename)
                    if not notes_number_match:
                        continue
                    
                    notes_number = int(notes_number_match.group(1))
                    
                    # Read the relationships file
                    rels_content = extractor.read_xml_content(notes_filename)
                    if not rels_content:
                        continue
                    
                    # Parse relationships and look for slide references
                    try:
                        rels_root = ET.fromstring(rels_content)
                        for rel in rels_root.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                            rel_type = rel.get('Type', '')
                            target = rel.get('Target', '')
                            
                            # Check if this is a slide relationship
                            if 'slide' in rel_type.lower() and 'slide' in target and target.endswith('.xml'):
                                # Extract slide number from target (e.g., "../slides/slide3.xml" -> 3)
                                slide_match = re.search(r'slide(\d+)\.xml$', target)
                                if slide_match:
                                    slide_number = int(slide_match.group(1))
                                    notes_file_path = f'ppt/notesSlides/notesSlide{notes_number}.xml'
                                    notes_to_slide_map[notes_file_path] = slide_number
                                    logger.debug(f"Found notes-slide relationship: {notes_file_path} -> slide {slide_number}")
                    
                    except Exception as e:
                        logger.warning(f"Failed to parse notes relationships file {notes_filename}: {e}")
        
        except Exception as e:
            logger.warning(f"Failed to build notes-slide mapping: {e}")
        
        return notes_to_slide_map
    
    def _parse_notes_content(self, notes_content: str, slide_number: int) -> str:
        """
        Parse text content from notes slide XML.
        
        Args:
            notes_content: XML content of the notes slide
            slide_number: The slide number this notes belongs to
            
        Returns:
            Notes text content as string
        """
        try:
            root = ET.fromstring(notes_content)
            
            # Look for text content in notes
            # Notes slides contain text in paragraph elements
            text_elements = root.findall('.//{http://schemas.openxmlformats.org/drawingml/2006/main}t')
            
            text_parts = []
            for text_elem in text_elements:
                if text_elem.text:
                    text_parts.append(text_elem.text)
            
            if text_parts:
                # Combine all text parts to form the notes content
                full_text = ''.join(text_parts)
                logger.debug(f"Found notes content for slide {slide_number}: {full_text[:50]}...")
                return full_text
            
            return ""
        
        except Exception as e:
            logger.warning(f"Failed to parse notes content for slide {slide_number}: {e}")
            return ""