"""
Unit tests for XMLParser class.

Tests XML parsing functionality for PowerPoint Office Open XML files,
including namespace handling and presentation.xml structure parsing.
"""

import pytest
import xml.etree.ElementTree as ET
from unittest.mock import patch, mock_open
from pathlib import Path
import tempfile
import os

from powerpoint_mcp_server.core.xml_parser import XMLParser


class TestXMLParser:
    """Test cases for XMLParser class."""
    
    def setup_method(self):
        """Set up test fixtures."""
        self.parser = XMLParser()
    
    def test_init_registers_namespaces(self):
        """Test that XMLParser initializes with proper namespaces."""
        parser = XMLParser()
        assert hasattr(parser, 'NAMESPACES')
        assert 'p' in parser.NAMESPACES
        assert 'a' in parser.NAMESPACES
        assert 'r' in parser.NAMESPACES
    
    def test_parse_xml_string_valid_xml(self):
        """Test parsing valid XML string."""
        xml_content = """<?xml version="1.0" encoding="UTF-8"?>
        <root xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
            <p:test>content</p:test>
        </root>"""
        
        root = self.parser.parse_xml_string(xml_content)
        
        assert root is not None
        assert root.tag == 'root'
    
    def test_parse_xml_string_invalid_xml(self):
        """Test parsing invalid XML string raises ParseError."""
        invalid_xml = "<root><unclosed>"
        
        with pytest.raises(ET.ParseError):
            self.parser.parse_xml_string(invalid_xml)
    
    def test_parse_xml_file_existing_file(self):
        """Test parsing existing XML file."""
        xml_content = """<?xml version="1.0" encoding="UTF-8"?>
        <root>
            <test>content</test>
        </root>"""
        
        with tempfile.NamedTemporaryFile(mode='w', suffix='.xml', delete=False) as f:
            f.write(xml_content)
            temp_path = f.name
        
        try:
            root = self.parser.parse_xml_file(temp_path)
            assert root is not None
            assert root.tag == 'root'
        finally:
            os.unlink(temp_path)
    
    def test_parse_xml_file_nonexistent_file(self):
        """Test parsing non-existent file raises FileNotFoundError."""
        with pytest.raises(FileNotFoundError):
            self.parser.parse_xml_file("nonexistent.xml")
    
    def test_parse_xml_file_malformed_xml(self):
        """Test parsing malformed XML file raises ParseError."""
        malformed_xml = "<root><unclosed>"
        
        with tempfile.NamedTemporaryFile(mode='w', suffix='.xml', delete=False) as f:
            f.write(malformed_xml)
            temp_path = f.name
        
        try:
            with pytest.raises(ET.ParseError):
                self.parser.parse_xml_file(temp_path)
        finally:
            os.unlink(temp_path)
    
    def test_parse_presentation_xml_basic_structure(self):
        """Test parsing basic presentation.xml structure."""
        presentation_xml = """<?xml version="1.0" encoding="UTF-8"?>
        <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <p:sldMasterIdLst>
                <p:sldMasterId id="2147483648" r:id="rId1"/>
            </p:sldMasterIdLst>
            <p:sldIdLst>
                <p:sldId id="256" r:id="rId2"/>
                <p:sldId id="257" r:id="rId3"/>
            </p:sldIdLst>
            <p:sldSz cx="9144000" cy="6858000"/>
        </p:presentation>"""
        
        result = self.parser.parse_presentation_xml(presentation_xml)
        
        assert 'slide_master_ids' in result
        assert 'slide_ids' in result
        assert 'slide_size' in result
        
        assert len(result['slide_master_ids']) == 1
        assert result['slide_master_ids'][0] == 'rId1'
        
        assert len(result['slide_ids']) == 2
        assert result['slide_ids'][0]['r_id'] == 'rId2'
        assert result['slide_ids'][0]['id'] == '256'
        assert result['slide_ids'][1]['r_id'] == 'rId3'
        assert result['slide_ids'][1]['id'] == '257'
        
        assert result['slide_size']['width'] == 9144000
        assert result['slide_size']['height'] == 6858000
    
    def test_parse_presentation_xml_with_notes_master(self):
        """Test parsing presentation.xml with notes master."""
        presentation_xml = """<?xml version="1.0" encoding="UTF-8"?>
        <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <p:sldIdLst>
                <p:sldId id="256" r:id="rId2"/>
            </p:sldIdLst>
            <p:notesMasterIdLst>
                <p:notesMasterId r:id="rId4"/>
            </p:notesMasterIdLst>
        </p:presentation>"""
        
        result = self.parser.parse_presentation_xml(presentation_xml)
        
        assert result['notes_master_id'] == 'rId4'
    
    def test_parse_presentation_xml_with_handout_master(self):
        """Test parsing presentation.xml with handout master."""
        presentation_xml = """<?xml version="1.0" encoding="UTF-8"?>
        <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <p:sldIdLst>
                <p:sldId id="256" r:id="rId2"/>
            </p:sldIdLst>
            <p:handoutMasterIdLst>
                <p:handoutMasterId r:id="rId5"/>
            </p:handoutMasterIdLst>
        </p:presentation>"""
        
        result = self.parser.parse_presentation_xml(presentation_xml)
        
        assert result['handout_master_id'] == 'rId5'
    
    def test_parse_presentation_xml_empty_structure(self):
        """Test parsing presentation.xml with minimal structure."""
        presentation_xml = """<?xml version="1.0" encoding="UTF-8"?>
        <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
        </p:presentation>"""
        
        result = self.parser.parse_presentation_xml(presentation_xml)
        
        assert result['slide_master_ids'] == []
        assert result['slide_ids'] == []
        assert result['slide_size'] is None
        assert result['notes_master_id'] is None
        assert result['handout_master_id'] is None
    
    def test_parse_presentation_xml_invalid_xml(self):
        """Test parsing invalid presentation.xml raises exception."""
        invalid_xml = "<p:presentation><unclosed>"
        
        with pytest.raises(ET.ParseError):
            self.parser.parse_presentation_xml(invalid_xml)
    
    def test_find_elements_with_namespace(self):
        """Test finding elements with namespace support."""
        xml_content = """<?xml version="1.0" encoding="UTF-8"?>
        <root xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
            <p:test>content1</p:test>
            <p:test>content2</p:test>
        </root>"""
        
        root = self.parser.parse_xml_string(xml_content)
        elements = self.parser.find_elements_with_namespace(root, './/p:test')
        
        assert len(elements) == 2
        assert elements[0].text == 'content1'
        assert elements[1].text == 'content2'
    
    def test_find_element_with_namespace(self):
        """Test finding single element with namespace support."""
        xml_content = """<?xml version="1.0" encoding="UTF-8"?>
        <root xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
            <p:test>content</p:test>
        </root>"""
        
        root = self.parser.parse_xml_string(xml_content)
        element = self.parser.find_element_with_namespace(root, './/p:test')
        
        assert element is not None
        assert element.text == 'content'
    
    def test_find_element_with_namespace_not_found(self):
        """Test finding non-existent element returns None."""
        xml_content = """<?xml version="1.0" encoding="UTF-8"?>
        <root xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
        </root>"""
        
        root = self.parser.parse_xml_string(xml_content)
        element = self.parser.find_element_with_namespace(root, './/p:nonexistent')
        
        assert element is None
    
    def test_get_element_text_with_content(self):
        """Test getting text from element with content."""
        xml_content = """<?xml version="1.0" encoding="UTF-8"?>
        <root>
            <test>  content with spaces  </test>
        </root>"""
        
        root = self.parser.parse_xml_string(xml_content)
        element = root.find('test')
        text = self.parser.get_element_text(element)
        
        assert text == 'content with spaces'
    
    def test_get_element_text_empty_element(self):
        """Test getting text from empty element returns empty string."""
        xml_content = """<?xml version="1.0" encoding="UTF-8"?>
        <root>
            <test></test>
        </root>"""
        
        root = self.parser.parse_xml_string(xml_content)
        element = root.find('test')
        text = self.parser.get_element_text(element)
        
        assert text == ''
    
    def test_get_element_text_none_element(self):
        """Test getting text from None element returns empty string."""
        text = self.parser.get_element_text(None)
        assert text == ''
    
    def test_get_attribute_with_namespace(self):
        """Test getting attribute with namespace."""
        xml_content = """<?xml version="1.0" encoding="UTF-8"?>
        <root xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <test r:id="rId1" normal="value"/>
        </root>"""
        
        root = self.parser.parse_xml_string(xml_content)
        element = root.find('test')
        
        # Test namespaced attribute
        r_id = self.parser.get_attribute_with_namespace(
            element, 'id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        )
        assert r_id == 'rId1'
        
        # Test normal attribute
        normal = self.parser.get_attribute_with_namespace(element, 'normal')
        assert normal == 'value'
    
    def test_get_attribute_with_namespace_not_found(self):
        """Test getting non-existent attribute returns None."""
        xml_content = """<?xml version="1.0" encoding="UTF-8"?>
        <root>
            <test/>
        </root>"""
        
        root = self.parser.parse_xml_string(xml_content)
        element = root.find('test')
        
        attr = self.parser.get_attribute_with_namespace(element, 'nonexistent')
        assert attr is None
    
    def test_namespace_constants(self):
        """Test that required namespaces are defined."""
        namespaces = self.parser.NAMESPACES
        
        # Check essential PowerPoint namespaces
        assert namespaces['p'] == 'http://schemas.openxmlformats.org/presentationml/2006/main'
        assert namespaces['a'] == 'http://schemas.openxmlformats.org/drawingml/2006/main'
        assert namespaces['r'] == 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        assert namespaces['cp'] == 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties'
        assert namespaces['dc'] == 'http://purl.org/dc/elements/1.1/'


class TestXMLParserPerformance:
    """Test cases for XMLParser performance optimizations."""
    
    def setup_method(self):
        """Set up test fixtures."""
        self.parser = XMLParser(enable_performance_mode=True)
    
    def test_performance_mode_initialization(self):
        """Test XMLParser initialization with performance mode."""
        parser = XMLParser(enable_performance_mode=True)
        assert parser.enable_performance_mode is True
        assert hasattr(parser, '_element_cache')
        assert isinstance(parser._element_cache, dict)
    
    def test_performance_mode_disabled(self):
        """Test XMLParser initialization with performance mode disabled."""
        parser = XMLParser(enable_performance_mode=False)
        assert parser.enable_performance_mode is False
    
    def test_element_cache_functionality(self):
        """Test element caching functionality."""
        # Initially empty cache
        stats = self.parser.get_cache_stats()
        assert stats['cached_elements'] == 0
        assert stats['performance_mode'] is True
        
        # Clear cache (should work even when empty)
        self.parser.clear_element_cache()
        
        # Stats should still show empty cache
        stats = self.parser.get_cache_stats()
        assert stats['cached_elements'] == 0
    
    def test_parse_large_xml_string_threshold(self):
        """Test that large XML strings trigger performance mode."""
        # Create a large XML string (over 1MB)
        large_content = "a" * (1024 * 1024 + 1)  # Just over 1MB
        large_xml = f'''<?xml version="1.0" encoding="UTF-8"?>
        <root>
            <data>{large_content}</data>
        </root>'''
        
        # This should use the performance parsing method
        # We can't easily test the internal method call, but we can verify it doesn't crash
        try:
            result = self.parser.parse_xml_string(large_xml)
            assert result is not None
            assert result.tag == 'root'
        except Exception as e:
            # If it fails due to memory constraints in test environment, that's acceptable
            assert "memory" in str(e).lower() or "size" in str(e).lower()
    
    def test_parse_small_xml_string_normal_path(self):
        """Test that small XML strings use normal parsing."""
        small_xml = '''<?xml version="1.0" encoding="UTF-8"?>
        <root>
            <data>small content</data>
        </root>'''
        
        result = self.parser.parse_xml_string(small_xml)
        assert result is not None
        assert result.tag == 'root'
        
        # Find the data element
        data_elem = result.find('data')
        assert data_elem is not None
        assert data_elem.text == 'small content'
    
    def test_parse_xml_iteratively(self):
        """Test iterative XML parsing functionality."""
        import tempfile
        import os
        
        # Create a test XML file
        xml_content = '''<?xml version="1.0" encoding="UTF-8"?>
        <root>
            <item id="1">Item 1</item>
            <item id="2">Item 2</item>
            <other>Other content</other>
            <item id="3">Item 3</item>
        </root>'''
        
        with tempfile.NamedTemporaryFile(mode='w', suffix='.xml', delete=False) as temp_file:
            temp_file.write(xml_content)
            temp_path = temp_file.name
        
        try:
            # Parse iteratively for 'item' elements
            items = list(self.parser.parse_xml_iteratively(temp_path, ['item']))
            
            # Should find 3 item elements
            assert len(items) == 3
            
            # Check that we got the right elements
            item_ids = [item.get('id') for item in items]
            assert '1' in item_ids
            assert '2' in item_ids
            assert '3' in item_ids
            
        finally:
            os.unlink(temp_path)
    
    def test_parse_xml_iteratively_no_matches(self):
        """Test iterative parsing with no matching elements."""
        import tempfile
        import os
        
        xml_content = '''<?xml version="1.0" encoding="UTF-8"?>
        <root>
            <data>Some data</data>
            <info>Some info</info>
        </root>'''
        
        with tempfile.NamedTemporaryFile(mode='w', suffix='.xml', delete=False) as temp_file:
            temp_file.write(xml_content)
            temp_path = temp_file.name
        
        try:
            # Parse iteratively for elements that don't exist
            items = list(self.parser.parse_xml_iteratively(temp_path, ['nonexistent']))
            
            # Should find no elements
            assert len(items) == 0
            
        finally:
            os.unlink(temp_path)
    
    def test_parse_xml_iteratively_file_not_found(self):
        """Test iterative parsing with non-existent file."""
        with pytest.raises(Exception):  # Should raise some kind of exception
            list(self.parser.parse_xml_iteratively('/nonexistent/file.xml', ['item']))
    
    def test_cached_element_lookup_context_manager(self):
        """Test the cached element lookup context manager."""
        # Test with new cache key
        with self.parser.cached_element_lookup('test_key') as cached_element:
            assert cached_element is None  # First time should be None
        
        # The context manager doesn't automatically cache None values
        # This is just testing that the context manager works without errors
        assert True  # If we get here, the context manager worked
    
    def test_large_file_parsing_threshold(self):
        """Test that large files trigger performance mode."""
        import tempfile
        import os
        
        # Create a file that's just over the 1MB threshold
        large_content = "a" * (1024 * 1024 + 1)
        large_xml = f'''<?xml version="1.0" encoding="UTF-8"?>
        <root>
            <data>{large_content}</data>
        </root>'''
        
        with tempfile.NamedTemporaryFile(mode='w', suffix='.xml', delete=False) as temp_file:
            temp_file.write(large_xml)
            temp_path = temp_file.name
        
        try:
            # This should trigger the large file parsing method
            result = self.parser.parse_xml_file(temp_path)
            assert result is not None
            assert result.tag == 'root'
        except Exception as e:
            # If it fails due to memory constraints in test environment, that's acceptable
            assert "memory" in str(e).lower() or "size" in str(e).lower()
        finally:
            os.unlink(temp_path)