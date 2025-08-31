"""
XML Parser for PowerPoint files.

This module provides XML parsing functionality for Office Open XML format
used in .pptx files. It handles namespace management and provides methods
to parse various XML components of PowerPoint presentations.
"""

import xml.etree.ElementTree as ET
from typing import Dict, Optional, List, Any, Iterator
from pathlib import Path
import logging
import io
from contextlib import contextmanager

logger = logging.getLogger(__name__)


class XMLParser:
    """
    XML parser for PowerPoint Office Open XML files.
    
    Handles parsing of presentation.xml, slide XML files, and other
    Office Open XML components with proper namespace management.
    """
    
    # Office Open XML namespaces commonly used in PowerPoint files
    NAMESPACES = {
        'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
        'dc': 'http://purl.org/dc/elements/1.1/',
        'dcterms': 'http://purl.org/dc/terms/',
        'dcmitype': 'http://purl.org/dc/dcmitype/',
        'xsi': 'http://www.w3.org/2001/XMLSchema-instance',
        'rel': 'http://schemas.openxmlformats.org/package/2006/relationships'
    }
    
    def __init__(self, enable_performance_mode: bool = True):
        """
        Initialize the XML parser with namespace registration.
        
        Args:
            enable_performance_mode: Enable optimizations for large files
        """
        self._register_namespaces()
        self.enable_performance_mode = enable_performance_mode
        self._element_cache = {}  # Cache for frequently accessed elements
        self.namespaces = self.NAMESPACES  # Make namespaces accessible
    
    def _register_namespaces(self) -> None:
        """Register XML namespaces for ElementTree parsing."""
        for prefix, uri in self.NAMESPACES.items():
            ET.register_namespace(prefix, uri)
    
    def parse_xml_file(self, file_path: str) -> Optional[ET.Element]:
        """
        Parse an XML file and return the root element.
        
        Args:
            file_path: Path to the XML file to parse
            
        Returns:
            Root element of the parsed XML, or None if parsing fails
            
        Raises:
            FileNotFoundError: If the XML file doesn't exist
            ET.ParseError: If the XML is malformed
        """
        try:
            path = Path(file_path)
            if not path.exists():
                raise FileNotFoundError(f"XML file not found: {file_path}")
            
            # Performance optimization: use iterparse for large files
            if self.enable_performance_mode and path.stat().st_size > 1024 * 1024:  # 1MB threshold
                return self._parse_large_xml_file(file_path)
            else:
                tree = ET.parse(file_path)
                root = tree.getroot()
                logger.debug(f"Successfully parsed XML file: {file_path}")
                return root
            
        except ET.ParseError as e:
            logger.error(f"Failed to parse XML file {file_path}: {e}")
            raise
        except Exception as e:
            logger.error(f"Unexpected error parsing XML file {file_path}: {e}")
            raise
    
    def parse_xml_string(self, xml_content: str) -> Optional[ET.Element]:
        """
        Parse XML content from a string and return the root element.
        
        Args:
            xml_content: XML content as string
            
        Returns:
            Root element of the parsed XML, or None if parsing fails
            
        Raises:
            ET.ParseError: If the XML is malformed
        """
        try:
            # Performance optimization: use iterparse for large XML strings
            if self.enable_performance_mode and len(xml_content) > 1024 * 1024:  # 1MB threshold
                return self._parse_large_xml_string(xml_content)
            else:
                root = ET.fromstring(xml_content)
                logger.debug("Successfully parsed XML from string")
                return root
            
        except ET.ParseError as e:
            logger.error(f"Failed to parse XML string: {e}")
            raise
        except Exception as e:
            logger.error(f"Unexpected error parsing XML string: {e}")
            raise
    
    def parse_presentation_xml(self, xml_content: str) -> Dict[str, Any]:
        """
        Parse presentation.xml content and extract presentation structure.
        
        Args:
            xml_content: Content of presentation.xml file
            
        Returns:
            Dictionary containing presentation structure information
            
        Raises:
            ET.ParseError: If the XML is malformed
        """
        try:
            root = self.parse_xml_string(xml_content)
            if root is None:
                return {}
            
            presentation_data = {
                'slide_master_ids': [],
                'slide_ids': [],
                'slide_size': None,
                'notes_master_id': None,
                'handout_master_id': None
            }
            
            # Extract slide master IDs
            slide_master_id_list = root.find('.//p:sldMasterIdLst', self.NAMESPACES)
            if slide_master_id_list is not None:
                for slide_master_id in slide_master_id_list.findall('p:sldMasterId', self.NAMESPACES):
                    r_id = slide_master_id.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                    if r_id:
                        presentation_data['slide_master_ids'].append(r_id)
            
            # Extract slide IDs
            slide_id_list = root.find('.//p:sldIdLst', self.NAMESPACES)
            if slide_id_list is not None:
                for slide_id in slide_id_list.findall('p:sldId', self.NAMESPACES):
                    r_id = slide_id.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                    slide_num_id = slide_id.get('id')
                    if r_id:
                        presentation_data['slide_ids'].append({
                            'r_id': r_id,
                            'id': slide_num_id
                        })
            
            # Extract slide size
            slide_size = root.find('.//p:sldSz', self.NAMESPACES)
            if slide_size is not None:
                cx = slide_size.get('cx')
                cy = slide_size.get('cy')
                if cx and cy:
                    presentation_data['slide_size'] = {
                        'width': int(cx),
                        'height': int(cy)
                    }
            
            # Extract notes master ID
            notes_master_id_list = root.find('.//p:notesMasterIdLst', self.NAMESPACES)
            if notes_master_id_list is not None:
                notes_master_id = notes_master_id_list.find('p:notesMasterId', self.NAMESPACES)
                if notes_master_id is not None:
                    r_id = notes_master_id.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                    if r_id:
                        presentation_data['notes_master_id'] = r_id
            
            # Extract handout master ID
            handout_master_id_list = root.find('.//p:handoutMasterIdLst', self.NAMESPACES)
            if handout_master_id_list is not None:
                handout_master_id = handout_master_id_list.find('p:handoutMasterId', self.NAMESPACES)
                if handout_master_id is not None:
                    r_id = handout_master_id.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                    if r_id:
                        presentation_data['handout_master_id'] = r_id
            
            logger.debug("Successfully parsed presentation.xml structure")
            return presentation_data
            
        except Exception as e:
            logger.error(f"Failed to parse presentation.xml: {e}")
            raise
    
    def find_elements_with_namespace(self, root: ET.Element, xpath: str) -> List[ET.Element]:
        """
        Find elements using XPath with namespace support.
        
        Args:
            root: Root element to search from
            xpath: XPath expression with namespace prefixes
            
        Returns:
            List of matching elements
        """
        try:
            return root.findall(xpath, self.NAMESPACES)
        except Exception as e:
            logger.error(f"Failed to find elements with XPath {xpath}: {e}")
            return []
    
    def find_element_with_namespace(self, root: ET.Element, xpath: str) -> Optional[ET.Element]:
        """
        Find a single element using XPath with namespace support.
        
        Args:
            root: Root element to search from
            xpath: XPath expression with namespace prefixes
            
        Returns:
            First matching element, or None if not found
        """
        try:
            return root.find(xpath, self.NAMESPACES)
        except Exception as e:
            logger.error(f"Failed to find element with XPath {xpath}: {e}")
            return None
    
    def get_element_text(self, element: ET.Element) -> str:
        """
        Get text content from an element, handling None cases.
        
        Args:
            element: Element to extract text from
            
        Returns:
            Text content of the element, or empty string if None
        """
        if element is not None and element.text is not None:
            return element.text.strip()
        return ""
    
    def get_attribute_with_namespace(self, element: ET.Element, attr_name: str, namespace: str = None) -> Optional[str]:
        """
        Get attribute value with namespace support.
        
        Args:
            element: Element to get attribute from
            attr_name: Attribute name
            namespace: Namespace URI (optional)
            
        Returns:
            Attribute value, or None if not found
        """
        try:
            if namespace:
                full_attr_name = f"{{{namespace}}}{attr_name}"
                return element.get(full_attr_name)
            else:
                return element.get(attr_name)
        except Exception as e:
            logger.error(f"Failed to get attribute {attr_name}: {e}")
            return None
    
    def _parse_large_xml_file(self, file_path: str) -> Optional[ET.Element]:
        """
        Parse large XML files using iterparse for memory efficiency.
        
        Args:
            file_path: Path to the XML file to parse
            
        Returns:
            Root element of the parsed XML
        """
        try:
            logger.debug(f"Using performance mode for large XML file: {file_path}")
            
            # Use iterparse to build the tree incrementally
            with open(file_path, 'rb') as file:
                events = ET.iterparse(file, events=('start', 'end'))
                root = None
                
                for event, elem in events:
                    if event == 'start' and root is None:
                        root = elem
                    elif event == 'end':
                        # Clear processed elements to save memory
                        elem.clear()
                
                return root
                
        except Exception as e:
            logger.error(f"Failed to parse large XML file {file_path}: {e}")
            raise
    
    def _parse_large_xml_string(self, xml_content: str) -> Optional[ET.Element]:
        """
        Parse large XML strings using iterparse for memory efficiency.
        
        Args:
            xml_content: XML content as string
            
        Returns:
            Root element of the parsed XML
        """
        try:
            logger.debug("Using performance mode for large XML string")
            
            # Use StringIO to create a file-like object from string
            xml_file = io.StringIO(xml_content)
            events = ET.iterparse(xml_file, events=('start', 'end'))
            root = None
            
            for event, elem in events:
                if event == 'start' and root is None:
                    root = elem
                elif event == 'end':
                    # Clear processed elements to save memory
                    elem.clear()
            
            return root
            
        except Exception as e:
            logger.error(f"Failed to parse large XML string: {e}")
            raise
    
    @contextmanager
    def cached_element_lookup(self, cache_key: str):
        """
        Context manager for caching frequently accessed elements.
        
        Args:
            cache_key: Key to use for caching
        """
        if cache_key in self._element_cache:
            yield self._element_cache[cache_key]
        else:
            element = yield None
            if element is not None:
                self._element_cache[cache_key] = element
    
    def clear_element_cache(self) -> None:
        """Clear the element cache to free memory."""
        self._element_cache.clear()
        logger.debug("Cleared XML element cache")
    
    def get_cache_stats(self) -> Dict[str, int]:
        """
        Get statistics about the element cache.
        
        Returns:
            Dictionary with cache statistics
        """
        return {
            'cached_elements': len(self._element_cache),
            'performance_mode': self.enable_performance_mode
        }
    
    def parse_xml_iteratively(self, file_path: str, target_elements: List[str]) -> Iterator[ET.Element]:
        """
        Parse XML file iteratively, yielding only target elements.
        Useful for processing large files without loading everything into memory.
        
        Args:
            file_path: Path to the XML file
            target_elements: List of element tag names to yield
            
        Yields:
            Elements matching the target element names
        """
        try:
            logger.debug(f"Parsing XML iteratively for elements: {target_elements}")
            
            with open(file_path, 'rb') as file:
                events = ET.iterparse(file, events=('start', 'end'))
                
                for event, elem in events:
                    if event == 'end' and elem.tag in target_elements:
                        # Create a copy of the element before yielding to preserve data
                        elem_copy = ET.Element(elem.tag, elem.attrib)
                        elem_copy.text = elem.text
                        elem_copy.tail = elem.tail
                        # Copy children
                        for child in elem:
                            elem_copy.append(child)
                        
                        yield elem_copy
                        
                        # Clear the original element to save memory
                        elem.clear()
                        
        except Exception as e:
            logger.error(f"Failed to parse XML iteratively from {file_path}: {e}")
            raise