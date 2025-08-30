"""
Unit tests for ContentExtractor class.

Tests slide content extraction functionality including layout information,
placeholder extraction, and basic slide structure parsing.
"""

import pytest
import xml.etree.ElementTree as ET
from unittest.mock import Mock, patch

from powerpoint_mcp_server.core.content_extractor import ContentExtractor, SlideInfo, PlaceholderInfo, TextElement


class TestContentExtractor:
    """Test cases for ContentExtractor class."""
    
    def setup_method(self):
        """Set up test fixtures."""
        self.extractor = ContentExtractor()
    
    def test_init_creates_xml_parser(self):
        """Test that ContentExtractor initializes with XMLParser."""
        extractor = ContentExtractor()
        assert hasattr(extractor, 'xml_parser')
        assert extractor.xml_parser is not None
    
    def test_extract_slide_content_basic_slide(self):
        """Test extracting content from basic slide XML."""
        slide_xml = """<?xml version="1.0" encoding="UTF-8"?>
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <p:cSld name="Title Slide">
                <p:spTree>
                    <p:sp>
                        <p:nvSpPr>
                            <p:cNvPr id="1" name="Title 1"/>
                            <p:cNvSpPr>
                                <p:spLocks noGrp="1"/>
                            </p:cNvSpPr>
                            <p:nvPr>
                                <p:ph type="title"/>
                            </p:nvPr>
                        </p:nvSpPr>
                        <p:spPr>
                            <a:xfrm>
                                <a:off x="1000" y="2000"/>
                                <a:ext cx="8000" cy="1500"/>
                            </a:xfrm>
                        </p:spPr>
                        <p:txBody>
                            <a:p>
                                <a:r>
                                    <a:t>Sample Title</a:t>
                                </a:r>
                            </a:p>
                        </p:txBody>
                    </p:sp>
                </p:spTree>
            </p:cSld>
        </p:sld>"""
        
        result = self.extractor.extract_slide_content(slide_xml, 1)
        
        assert isinstance(result, SlideInfo)
        assert result.slide_number == 1
        assert result.layout_name == "Title Slide"
        assert result.title == "Sample Title"
        assert len(result.placeholders) == 1
        assert result.placeholders[0]['type'] == 'title'
        assert result.placeholders[0]['content'] == 'Sample Title'
        assert result.placeholders[0]['position'] == (1000, 2000)
        assert result.placeholders[0]['size'] == (8000, 1500)
    
    def test_extract_slide_content_with_subtitle(self):
        """Test extracting slide with title and subtitle."""
        slide_xml = """<?xml version="1.0" encoding="UTF-8"?>
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <p:cSld>
                <p:spTree>
                    <p:sp>
                        <p:nvSpPr>
                            <p:nvPr>
                                <p:ph type="title"/>
                            </p:nvPr>
                        </p:nvSpPr>
                        <p:spPr>
                            <a:xfrm>
                                <a:off x="0" y="0"/>
                                <a:ext cx="100" cy="50"/>
                            </a:xfrm>
                        </p:spPr>
                        <p:txBody>
                            <a:p>
                                <a:r>
                                    <a:t>Main Title</a:t>
                                </a:r>
                            </a:p>
                        </p:txBody>
                    </p:sp>
                    <p:sp>
                        <p:nvSpPr>
                            <p:nvPr>
                                <p:ph type="subTitle"/>
                            </p:nvPr>
                        </p:nvSpPr>
                        <p:spPr>
                            <a:xfrm>
                                <a:off x="0" y="100"/>
                                <a:ext cx="100" cy="30"/>
                            </a:xfrm>
                        </p:spPr>
                        <p:txBody>
                            <a:p>
                                <a:r>
                                    <a:t>Subtitle Text</a:t>
                                </a:r>
                            </a:p>
                        </p:txBody>
                    </p:sp>
                </p:spTree>
            </p:cSld>
        </p:sld>"""
        
        result = self.extractor.extract_slide_content(slide_xml, 2)
        
        assert result.slide_number == 2
        assert result.title == "Main Title"
        assert result.subtitle == "Subtitle Text"
        assert len(result.placeholders) == 2
    
    def test_extract_slide_content_multiple_paragraphs(self):
        """Test extracting text with multiple paragraphs."""
        slide_xml = """<?xml version="1.0" encoding="UTF-8"?>
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <p:cSld>
                <p:spTree>
                    <p:sp>
                        <p:nvSpPr>
                            <p:nvPr>
                                <p:ph type="body"/>
                            </p:nvPr>
                        </p:nvSpPr>
                        <p:spPr>
                            <a:xfrm>
                                <a:off x="0" y="0"/>
                                <a:ext cx="100" cy="100"/>
                            </a:xfrm>
                        </p:spPr>
                        <p:txBody>
                            <a:p>
                                <a:r>
                                    <a:t>First paragraph</a:t>
                                </a:r>
                            </a:p>
                            <a:p>
                                <a:r>
                                    <a:t>Second paragraph</a:t>
                                </a:r>
                            </a:p>
                        </p:txBody>
                    </p:sp>
                </p:spTree>
            </p:cSld>
        </p:sld>"""
        
        result = self.extractor.extract_slide_content(slide_xml, 3)
        
        assert len(result.placeholders) == 1
        assert result.placeholders[0]['content'] == "First paragraph\nSecond paragraph"
    
    def test_extract_slide_content_multiple_runs(self):
        """Test extracting text with multiple runs in same paragraph."""
        slide_xml = """<?xml version="1.0" encoding="UTF-8"?>
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <p:cSld>
                <p:spTree>
                    <p:sp>
                        <p:nvSpPr>
                            <p:nvPr>
                                <p:ph type="body"/>
                            </p:nvPr>
                        </p:nvSpPr>
                        <p:spPr>
                            <a:xfrm>
                                <a:off x="0" y="0"/>
                                <a:ext cx="100" cy="100"/>
                            </a:xfrm>
                        </p:spPr>
                        <p:txBody>
                            <a:p>
                                <a:r>
                                    <a:t>Bold </a:t>
                                </a:r>
                                <a:r>
                                    <a:t>and normal text</a:t>
                                </a:r>
                            </a:p>
                        </p:txBody>
                    </p:sp>
                </p:spTree>
            </p:cSld>
        </p:sld>"""
        
        result = self.extractor.extract_slide_content(slide_xml, 4)
        
        assert len(result.placeholders) == 1
        assert result.placeholders[0]['content'] == "Bold and normal text"
    
    def test_extract_slide_content_no_placeholders(self):
        """Test extracting slide with no placeholders."""
        slide_xml = """<?xml version="1.0" encoding="UTF-8"?>
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
            <p:cSld>
                <p:spTree>
                </p:spTree>
            </p:cSld>
        </p:sld>"""
        
        result = self.extractor.extract_slide_content(slide_xml, 5)
        
        assert result.slide_number == 5
        assert len(result.placeholders) == 0
        assert result.title is None
        assert result.subtitle is None
    
    def test_extract_slide_content_invalid_xml(self):
        """Test extracting content from invalid XML."""
        invalid_xml = "<p:sld><unclosed>"
        
        result = self.extractor.extract_slide_content(invalid_xml, 6)
        
        assert result.slide_number == 6
        assert len(result.placeholders) == 0
    
    def test_determine_layout_type_title_and_content(self):
        """Test layout type determination for title and content layout."""
        slide_xml = """<?xml version="1.0" encoding="UTF-8"?>
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
            <p:cSld>
                <p:spTree>
                    <p:sp>
                        <p:nvSpPr>
                            <p:nvPr>
                                <p:ph type="title"/>
                            </p:nvPr>
                        </p:nvSpPr>
                    </p:sp>
                    <p:sp>
                        <p:nvSpPr>
                            <p:nvPr>
                                <p:ph type="body"/>
                            </p:nvPr>
                        </p:nvSpPr>
                    </p:sp>
                </p:spTree>
            </p:cSld>
        </p:sld>"""
        
        result = self.extractor.extract_slide_content(slide_xml, 1)
        assert result.layout_type == "titleAndContent"
    
    def test_determine_layout_type_two_content(self):
        """Test layout type determination for two content layout."""
        slide_xml = """<?xml version="1.0" encoding="UTF-8"?>
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
            <p:cSld>
                <p:spTree>
                    <p:sp>
                        <p:nvSpPr>
                            <p:nvPr>
                                <p:ph type="title"/>
                            </p:nvPr>
                        </p:nvSpPr>
                    </p:sp>
                    <p:sp>
                        <p:nvSpPr>
                            <p:nvPr>
                                <p:ph type="body"/>
                            </p:nvPr>
                        </p:nvSpPr>
                    </p:sp>
                    <p:sp>
                        <p:nvSpPr>
                            <p:nvPr>
                                <p:ph type="obj"/>
                            </p:nvPr>
                        </p:nvSpPr>
                    </p:sp>
                </p:spTree>
            </p:cSld>
        </p:sld>"""
        
        result = self.extractor.extract_slide_content(slide_xml, 1)
        assert result.layout_type == "twoContent"
    
    def test_determine_layout_type_title_only(self):
        """Test layout type determination for title only layout."""
        slide_xml = """<?xml version="1.0" encoding="UTF-8"?>
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
            <p:cSld>
                <p:spTree>
                    <p:sp>
                        <p:nvSpPr>
                            <p:nvPr>
                                <p:ph type="title"/>
                            </p:nvPr>
                        </p:nvSpPr>
                    </p:sp>
                </p:spTree>
            </p:cSld>
        </p:sld>"""
        
        result = self.extractor.extract_slide_content(slide_xml, 1)
        assert result.layout_type == "titleOnly"
    
    def test_determine_layout_type_blank(self):
        """Test layout type determination for blank layout."""
        slide_xml = """<?xml version="1.0" encoding="UTF-8"?>
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
            <p:cSld>
                <p:spTree>
                </p:spTree>
            </p:cSld>
        </p:sld>"""
        
        result = self.extractor.extract_slide_content(slide_xml, 1)
        assert result.layout_type == "blank"
    
    def test_extract_slide_layout_info(self):
        """Test extracting slide layout information."""
        layout_xml = """<?xml version="1.0" encoding="UTF-8"?>
        <p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                     xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <p:cSld name="Title and Content">
                <p:spTree>
                    <p:sp>
                        <p:nvSpPr>
                            <p:nvPr>
                                <p:ph type="title"/>
                            </p:nvPr>
                        </p:nvSpPr>
                        <p:spPr>
                            <a:xfrm>
                                <a:off x="100" y="200"/>
                                <a:ext cx="800" cy="150"/>
                            </a:xfrm>
                        </p:spPr>
                    </p:sp>
                    <p:sp>
                        <p:nvSpPr>
                            <p:nvPr>
                                <p:ph type="body" idx="1"/>
                            </p:nvPr>
                        </p:nvSpPr>
                        <p:spPr>
                            <a:xfrm>
                                <a:off x="100" y="400"/>
                                <a:ext cx="800" cy="600"/>
                            </a:xfrm>
                        </p:spPr>
                    </p:sp>
                </p:spTree>
            </p:cSld>
        </p:sldLayout>"""
        
        result = self.extractor.extract_slide_layout_info(layout_xml)
        
        assert result['name'] == "Title and Content"
        assert result['type'] == "titleAndContent"
        assert len(result['placeholders']) == 2
        assert result['placeholders'][0]['type'] == 'title'
        assert result['placeholders'][1]['type'] == 'body'
    
    def test_extract_basic_slide_info(self):
        """Test extracting basic slide information."""
        slide_xml = """<?xml version="1.0" encoding="UTF-8"?>
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <p:cSld name="Sample Layout">
                <p:spTree>
                    <p:sp>
                        <p:nvSpPr>
                            <p:nvPr>
                                <p:ph type="title"/>
                            </p:nvPr>
                        </p:nvSpPr>
                        <p:txBody>
                            <a:p>
                                <a:r>
                                    <a:t>Test Title</a:t>
                                </a:r>
                            </a:p>
                        </p:txBody>
                    </p:sp>
                </p:spTree>
            </p:cSld>
        </p:sld>"""
        
        result = self.extractor.extract_basic_slide_info(slide_xml, 7)
        
        assert result['slide_number'] == 7
        assert result['layout_name'] == "Sample Layout"
        assert result['layout_type'] == "titleOnly"
        assert result['title'] == "Test Title"
        assert result['placeholder_count'] == 1
        assert result['placeholder_types'] == ['title']
    
    def test_extract_shape_transform_missing_transform(self):
        """Test extracting transform from shape without transform element."""
        shape_xml = """<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
        </p:sp>"""
        
        root = ET.fromstring(shape_xml)
        position, size = self.extractor._extract_shape_transform(root)
        
        assert position == (0, 0)
        assert size == (0, 0)
    
    def test_extract_shape_text_content_no_text(self):
        """Test extracting text from shape without text body."""
        shape_xml = """<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
        </p:sp>"""
        
        root = ET.fromstring(shape_xml)
        content = self.extractor._extract_shape_text_content(root)
        
        assert content is None
    
    def test_extract_single_placeholder_not_placeholder(self):
        """Test extracting placeholder info from non-placeholder shape."""
        shape_xml = """<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
            <p:nvSpPr>
                <p:cNvPr id="1" name="Shape 1"/>
            </p:nvSpPr>
        </p:sp>"""
        
        root = ET.fromstring(shape_xml)
        result = self.extractor._extract_single_placeholder(root)
        
        assert result is None
    
    def test_slide_info_dataclass(self):
        """Test SlideInfo dataclass initialization."""
        slide_info = SlideInfo(slide_number=1)
        
        assert slide_info.slide_number == 1
        assert slide_info.layout_name is None
        assert slide_info.layout_type is None
        assert slide_info.title is None
        assert slide_info.subtitle is None
        assert slide_info.placeholders == []
    
    def test_placeholder_info_dataclass(self):
        """Test PlaceholderInfo dataclass initialization."""
        placeholder_info = PlaceholderInfo(
            placeholder_type="title",
            position=(100, 200),
            size=(800, 150),
            content="Test Content"
        )
        
        assert placeholder_info.placeholder_type == "title"
        assert placeholder_info.position == (100, 200)
        assert placeholder_info.size == (800, 150)
        assert placeholder_info.content == "Test Content"
    
    def test_extract_text_elements_basic(self):
        """Test extracting basic text elements."""
        slide_xml = """<?xml version="1.0" encoding="UTF-8"?>
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <p:cSld>
                <p:spTree>
                    <p:sp>
                        <p:spPr>
                            <a:xfrm>
                                <a:off x="100" y="200"/>
                                <a:ext cx="800" cy="150"/>
                            </a:xfrm>
                        </p:spPr>
                        <p:txBody>
                            <a:p>
                                <a:r>
                                    <a:t>Simple text</a:t>
                                </a:r>
                            </a:p>
                        </p:txBody>
                    </p:sp>
                </p:spTree>
            </p:cSld>
        </p:sld>"""
        
        result = self.extractor.extract_slide_content(slide_xml, 1)
        
        assert len(result.text_elements) == 1
        text_elem = result.text_elements[0]
        assert text_elem['content_plain'] == "Simple text"
        assert text_elem['content_formatted'] == "Simple text"
        assert text_elem['position'] == (100, 200)
        assert text_elem['size'] == (800, 150)
    
    def test_extract_text_elements_with_formatting(self):
        """Test extracting text elements with formatting."""
        slide_xml = """<?xml version="1.0" encoding="UTF-8"?>
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <p:cSld>
                <p:spTree>
                    <p:sp>
                        <p:spPr>
                            <a:xfrm>
                                <a:off x="0" y="0"/>
                                <a:ext cx="100" cy="50"/>
                            </a:xfrm>
                        </p:spPr>
                        <p:txBody>
                            <a:p>
                                <a:r>
                                    <a:rPr>
                                        <a:b val="1"/>
                                        <a:i val="1"/>
                                        <a:u val="sng"/>
                                        <a:sz val="2400"/>
                                        <a:solidFill>
                                            <a:srgbClr val="FF0000"/>
                                        </a:solidFill>
                                    </a:rPr>
                                    <a:t>Formatted text</a:t>
                                </a:r>
                            </a:p>
                        </p:txBody>
                    </p:sp>
                </p:spTree>
            </p:cSld>
        </p:sld>"""
        
        result = self.extractor.extract_slide_content(slide_xml, 1)
        
        assert len(result.text_elements) == 1
        text_elem = result.text_elements[0]
        assert text_elem['content_plain'] == "Formatted text"
        assert text_elem['content_formatted'] == "<u><i><b>Formatted text</b></i></u>"
        assert text_elem['bolded'] == 1
        assert text_elem['italic'] == 1
        assert text_elem['underlined'] == 1
        assert 24 in text_elem['font_sizes']  # 2400/100 = 24
        assert "#FF0000" in text_elem['font_colors']
    
    def test_extract_text_elements_multiple_runs(self):
        """Test extracting text with multiple runs and different formatting."""
        slide_xml = """<?xml version="1.0" encoding="UTF-8"?>
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <p:cSld>
                <p:spTree>
                    <p:sp>
                        <p:txBody>
                            <a:p>
                                <a:r>
                                    <a:rPr>
                                        <a:b val="1"/>
                                    </a:rPr>
                                    <a:t>Bold </a:t>
                                </a:r>
                                <a:r>
                                    <a:rPr>
                                        <a:i val="1"/>
                                    </a:rPr>
                                    <a:t>italic </a:t>
                                </a:r>
                                <a:r>
                                    <a:t>normal</a:t>
                                </a:r>
                            </a:p>
                        </p:txBody>
                    </p:sp>
                </p:spTree>
            </p:cSld>
        </p:sld>"""
        
        result = self.extractor.extract_slide_content(slide_xml, 1)
        
        assert len(result.text_elements) == 1
        text_elem = result.text_elements[0]
        assert text_elem['content_plain'] == "Bold italic normal"
        assert text_elem['content_formatted'] == "<b>Bold </b><i>italic </i>normal"
        assert text_elem['bolded'] == 1
        assert text_elem['italic'] == 1
    
    def test_extract_text_elements_with_strikethrough_highlight(self):
        """Test extracting text with strikethrough and highlight formatting."""
        slide_xml = """<?xml version="1.0" encoding="UTF-8"?>
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <p:cSld>
                <p:spTree>
                    <p:sp>
                        <p:txBody>
                            <a:p>
                                <a:r>
                                    <a:rPr>
                                        <a:strike val="sngStrike"/>
                                        <a:highlight>
                                            <a:srgbClr val="FFFF00"/>
                                        </a:highlight>
                                    </a:rPr>
                                    <a:t>Strike and highlight</a:t>
                                </a:r>
                            </a:p>
                        </p:txBody>
                    </p:sp>
                </p:spTree>
            </p:cSld>
        </p:sld>"""
        
        result = self.extractor.extract_slide_content(slide_xml, 1)
        
        assert len(result.text_elements) == 1
        text_elem = result.text_elements[0]
        assert text_elem['content_plain'] == "Strike and highlight"
        assert text_elem['content_formatted'] == "<mark><s>Strike and highlight</s></mark>"
        assert text_elem['strikethrough'] == 1
        assert text_elem['highlighted'] == 1
    
    def test_extract_text_elements_scheme_color(self):
        """Test extracting text with scheme color."""
        slide_xml = """<?xml version="1.0" encoding="UTF-8"?>
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <p:cSld>
                <p:spTree>
                    <p:sp>
                        <p:txBody>
                            <a:p>
                                <a:r>
                                    <a:rPr>
                                        <a:solidFill>
                                            <a:schemeClr val="accent1"/>
                                        </a:solidFill>
                                    </a:rPr>
                                    <a:t>Scheme color text</a:t>
                                </a:r>
                            </a:p>
                        </p:txBody>
                    </p:sp>
                </p:spTree>
            </p:cSld>
        </p:sld>"""
        
        result = self.extractor.extract_slide_content(slide_xml, 1)
        
        assert len(result.text_elements) == 1
        text_elem = result.text_elements[0]
        assert "accent1" in text_elem['font_colors']
    
    def test_extract_text_elements_multiple_paragraphs(self):
        """Test extracting text with multiple paragraphs."""
        slide_xml = """<?xml version="1.0" encoding="UTF-8"?>
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <p:cSld>
                <p:spTree>
                    <p:sp>
                        <p:txBody>
                            <a:p>
                                <a:r>
                                    <a:t>First paragraph</a:t>
                                </a:r>
                            </a:p>
                            <a:p>
                                <a:r>
                                    <a:t>Second paragraph</a:t>
                                </a:r>
                            </a:p>
                        </p:txBody>
                    </p:sp>
                </p:spTree>
            </p:cSld>
        </p:sld>"""
        
        result = self.extractor.extract_slide_content(slide_xml, 1)
        
        assert len(result.text_elements) == 1
        text_elem = result.text_elements[0]
        assert text_elem['content_plain'] == "First paragraph\nSecond paragraph"
    
    def test_extract_text_elements_with_hyperlink(self):
        """Test extracting text with hyperlinks."""
        slide_xml = """<?xml version="1.0" encoding="UTF-8"?>
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <p:cSld>
                <p:spTree>
                    <p:sp>
                        <p:txBody>
                            <a:p>
                                <a:r>
                                    <a:t>Click </a:t>
                                </a:r>
                                <a:r>
                                    <a:t>here</a:t>
                                </a:r>
                                <a:hlinkClick r:id="rId1"/>
                            </a:p>
                        </p:txBody>
                    </p:sp>
                </p:spTree>
            </p:cSld>
        </p:sld>"""
        
        result = self.extractor.extract_slide_content(slide_xml, 1)
        
        assert len(result.text_elements) == 1
        text_elem = result.text_elements[0]
        assert text_elem['content_plain'] == "Click here"
        assert "rId1" in text_elem['hyperlinks']
    
    def test_extract_text_elements_empty_shape(self):
        """Test extracting from shape with no text."""
        slide_xml = """<?xml version="1.0" encoding="UTF-8"?>
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
            <p:cSld>
                <p:spTree>
                    <p:sp>
                        <p:spPr/>
                    </p:sp>
                </p:spTree>
            </p:cSld>
        </p:sld>"""
        
        result = self.extractor.extract_slide_content(slide_xml, 1)
        
        assert len(result.text_elements) == 0
    
    def test_extract_formatted_text(self):
        """Test extracting formatted text summary."""
        slide_xml = """<?xml version="1.0" encoding="UTF-8"?>
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <p:cSld>
                <p:spTree>
                    <p:sp>
                        <p:txBody>
                            <a:p>
                                <a:r>
                                    <a:rPr>
                                        <a:b val="1"/>
                                    </a:rPr>
                                    <a:t>Bold text</a:t>
                                </a:r>
                            </a:p>
                        </p:txBody>
                    </p:sp>
                    <p:sp>
                        <p:txBody>
                            <a:p>
                                <a:r>
                                    <a:t>Normal text</a:t>
                                </a:r>
                            </a:p>
                        </p:txBody>
                    </p:sp>
                </p:spTree>
            </p:cSld>
        </p:sld>"""
        
        result = self.extractor.extract_formatted_text(slide_xml)
        
        assert result['plain_text'] == "Bold text\n\nNormal text"
        assert result['formatted_text'] == "<b>Bold text</b>\n\nNormal text"
        assert len(result['text_elements']) == 2
    
    def test_extract_text_elements_method(self):
        """Test the extract_text_elements method."""
        slide_xml = """<?xml version="1.0" encoding="UTF-8"?>
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <p:cSld>
                <p:spTree>
                    <p:sp>
                        <p:txBody>
                            <a:p>
                                <a:r>
                                    <a:t>Test text</a:t>
                                </a:r>
                            </a:p>
                        </p:txBody>
                    </p:sp>
                </p:spTree>
            </p:cSld>
        </p:sld>"""
        
        result = self.extractor.extract_text_elements(slide_xml, 1)
        
        assert len(result) == 1
        assert result[0]['content_plain'] == "Test text"    

    def test_extract_table_basic(self):
        """Test extracting basic table structure."""
        slide_xml = """<?xml version="1.0" encoding="UTF-8"?>
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <p:cSld>
                <p:spTree>
                    <p:graphicFrame>
                        <p:xfrm>
                            <a:off x="100" y="200"/>
                            <a:ext cx="800" cy="400"/>
                        </p:xfrm>
                        <a:graphic>
                            <a:graphicData>
                                <a:tbl>
                                    <a:tr>
                                        <a:tc>
                                            <a:txBody>
                                                <a:p>
                                                    <a:r>
                                                        <a:t>Cell 1,1</a:t>
                                                    </a:r>
                                                </a:p>
                                            </a:txBody>
                                        </a:tc>
                                        <a:tc>
                                            <a:txBody>
                                                <a:p>
                                                    <a:r>
                                                        <a:t>Cell 1,2</a:t>
                                                    </a:r>
                                                </a:p>
                                            </a:txBody>
                                        </a:tc>
                                    </a:tr>
                                    <a:tr>
                                        <a:tc>
                                            <a:txBody>
                                                <a:p>
                                                    <a:r>
                                                        <a:t>Cell 2,1</a:t>
                                                    </a:r>
                                                </a:p>
                                            </a:txBody>
                                        </a:tc>
                                        <a:tc>
                                            <a:txBody>
                                                <a:p>
                                                    <a:r>
                                                        <a:t>Cell 2,2</a:t>
                                                    </a:r>
                                                </a:p>
                                            </a:txBody>
                                        </a:tc>
                                    </a:tr>
                                </a:tbl>
                            </a:graphicData>
                        </a:graphic>
                    </p:graphicFrame>
                </p:spTree>
            </p:cSld>
        </p:sld>"""
        
        result = self.extractor.extract_slide_content(slide_xml, 1)
        
        assert len(result.tables) == 1
        table = result.tables[0]
        assert table['rows'] == 2
        assert table['columns'] == 2
        assert table['position'] == (100, 200)
        assert table['size'] == (800, 400)
        assert table['cells'][0][0]['content'] == "Cell 1,1"
        assert table['cells'][0][1]['content'] == "Cell 1,2"
        assert table['cells'][1][0]['content'] == "Cell 2,1"
        assert table['cells'][1][1]['content'] == "Cell 2,2"
    
    def test_extract_table_with_merged_cells(self):
        """Test extracting table with merged cells."""
        slide_xml = """<?xml version="1.0" encoding="UTF-8"?>
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <p:cSld>
                <p:spTree>
                    <p:graphicFrame>
                        <a:graphic>
                            <a:graphicData>
                                <a:tbl>
                                    <a:tr>
                                        <a:tc gridSpan="2" rowSpan="1">
                                            <a:txBody>
                                                <a:p>
                                                    <a:r>
                                                        <a:t>Merged Cell</a:t>
                                                    </a:r>
                                                </a:p>
                                            </a:txBody>
                                        </a:tc>
                                    </a:tr>
                                    <a:tr>
                                        <a:tc>
                                            <a:txBody>
                                                <a:p>
                                                    <a:r>
                                                        <a:t>Cell 2,1</a:t>
                                                    </a:r>
                                                </a:p>
                                            </a:txBody>
                                        </a:tc>
                                        <a:tc>
                                            <a:txBody>
                                                <a:p>
                                                    <a:r>
                                                        <a:t>Cell 2,2</a:t>
                                                    </a:r>
                                                </a:p>
                                            </a:txBody>
                                        </a:tc>
                                    </a:tr>
                                </a:tbl>
                            </a:graphicData>
                        </a:graphic>
                    </p:graphicFrame>
                </p:spTree>
            </p:cSld>
        </p:sld>"""
        
        result = self.extractor.extract_slide_content(slide_xml, 1)
        
        assert len(result.tables) == 1
        table = result.tables[0]
        assert table['rows'] == 2
        assert table['columns'] == 2
        assert table['cells'][0][0]['content'] == "Merged Cell"
        assert table['cells'][0][0]['col_span'] == 2
        assert table['cells'][0][0]['row_span'] == 1
    
    def test_extract_table_with_formatting(self):
        """Test extracting table with cell formatting."""
        slide_xml = """<?xml version="1.0" encoding="UTF-8"?>
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <p:cSld>
                <p:spTree>
                    <p:graphicFrame>
                        <a:graphic>
                            <a:graphicData>
                                <a:tbl>
                                    <a:tr>
                                        <a:tc>
                                            <a:tcPr>
                                                <a:solidFill>
                                                    <a:srgbClr val="FF0000"/>
                                                </a:solidFill>
                                                <a:lnL w="12700"/>
                                                <a:lnR w="12700"/>
                                                <a:lnT w="12700"/>
                                                <a:lnB w="12700"/>
                                            </a:tcPr>
                                            <a:txBody>
                                                <a:p>
                                                    <a:r>
                                                        <a:t>Formatted Cell</a:t>
                                                    </a:r>
                                                </a:p>
                                            </a:txBody>
                                        </a:tc>
                                    </a:tr>
                                </a:tbl>
                            </a:graphicData>
                        </a:graphic>
                    </p:graphicFrame>
                </p:spTree>
            </p:cSld>
        </p:sld>"""
        
        result = self.extractor.extract_slide_content(slide_xml, 1)
        
        assert len(result.tables) == 1
        table = result.tables[0]
        cell = table['cells'][0][0]
        assert cell['content'] == "Formatted Cell"
        assert cell['formatting']['fill_color'] == "#FF0000"
        assert 'borders' in cell['formatting']
        assert cell['formatting']['borders']['lnL']['width'] == 12700
    
    def test_extract_table_multiple_paragraphs(self):
        """Test extracting table cell with multiple paragraphs."""
        slide_xml = """<?xml version="1.0" encoding="UTF-8"?>
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <p:cSld>
                <p:spTree>
                    <p:graphicFrame>
                        <a:graphic>
                            <a:graphicData>
                                <a:tbl>
                                    <a:tr>
                                        <a:tc>
                                            <a:txBody>
                                                <a:p>
                                                    <a:r>
                                                        <a:t>First paragraph</a:t>
                                                    </a:r>
                                                </a:p>
                                                <a:p>
                                                    <a:r>
                                                        <a:t>Second paragraph</a:t>
                                                    </a:r>
                                                </a:p>
                                            </a:txBody>
                                        </a:tc>
                                    </a:tr>
                                </a:tbl>
                            </a:graphicData>
                        </a:graphic>
                    </p:graphicFrame>
                </p:spTree>
            </p:cSld>
        </p:sld>"""
        
        result = self.extractor.extract_slide_content(slide_xml, 1)
        
        assert len(result.tables) == 1
        table = result.tables[0]
        assert table['cells'][0][0]['content'] == "First paragraph\nSecond paragraph"
    
    def test_extract_table_empty_cells(self):
        """Test extracting table with empty cells."""
        slide_xml = """<?xml version="1.0" encoding="UTF-8"?>
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <p:cSld>
                <p:spTree>
                    <p:graphicFrame>
                        <a:graphic>
                            <a:graphicData>
                                <a:tbl>
                                    <a:tr>
                                        <a:tc>
                                            <a:txBody>
                                                <a:p/>
                                            </a:txBody>
                                        </a:tc>
                                        <a:tc>
                                            <a:txBody>
                                                <a:p>
                                                    <a:r>
                                                        <a:t>Not empty</a:t>
                                                    </a:r>
                                                </a:p>
                                            </a:txBody>
                                        </a:tc>
                                    </a:tr>
                                </a:tbl>
                            </a:graphicData>
                        </a:graphic>
                    </p:graphicFrame>
                </p:spTree>
            </p:cSld>
        </p:sld>"""
        
        result = self.extractor.extract_slide_content(slide_xml, 1)
        
        assert len(result.tables) == 1
        table = result.tables[0]
        assert table['cells'][0][0]['content'] == ""
        assert table['cells'][0][1]['content'] == "Not empty"
    
    def test_extract_table_no_tables(self):
        """Test extracting from slide with no tables."""
        slide_xml = """<?xml version="1.0" encoding="UTF-8"?>
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
            <p:cSld>
                <p:spTree>
                    <p:sp>
                        <p:txBody>
                            <a:p>
                                <a:r>
                                    <a:t>Just text, no tables</a:t>
                                </a:r>
                            </a:p>
                        </p:txBody>
                    </p:sp>
                </p:spTree>
            </p:cSld>
        </p:sld>"""
        
        result = self.extractor.extract_slide_content(slide_xml, 1)
        
        assert len(result.tables) == 0
    
    def test_extract_table_data_method(self):
        """Test the extract_table_data method."""
        slide_xml = """<?xml version="1.0" encoding="UTF-8"?>
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <p:cSld>
                <p:spTree>
                    <p:graphicFrame>
                        <a:graphic>
                            <a:graphicData>
                                <a:tbl>
                                    <a:tr>
                                        <a:tc>
                                            <a:txBody>
                                                <a:p>
                                                    <a:r>
                                                        <a:t>Test</a:t>
                                                    </a:r>
                                                </a:p>
                                            </a:txBody>
                                        </a:tc>
                                    </a:tr>
                                </a:tbl>
                            </a:graphicData>
                        </a:graphic>
                    </p:graphicFrame>
                </p:spTree>
            </p:cSld>
        </p:sld>"""
        
        result = self.extractor.extract_table_data(slide_xml, 1)
        
        assert len(result) == 1
        assert result[0]['cells'][0][0]['content'] == "Test"
    
    def test_extract_tables_with_structure(self):
        """Test the extract_tables_with_structure method."""
        slide_xml = """<?xml version="1.0" encoding="UTF-8"?>
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <p:cSld>
                <p:spTree>
                    <p:graphicFrame>
                        <a:graphic>
                            <a:graphicData>
                                <a:tbl>
                                    <a:tr>
                                        <a:tc>
                                            <a:txBody>
                                                <a:p>
                                                    <a:r>
                                                        <a:t>Table 1</a:t>
                                                    </a:r>
                                                </a:p>
                                            </a:txBody>
                                        </a:tc>
                                    </a:tr>
                                </a:tbl>
                            </a:graphicData>
                        </a:graphic>
                    </p:graphicFrame>
                    <p:graphicFrame>
                        <a:graphic>
                            <a:graphicData>
                                <a:tbl>
                                    <a:tr>
                                        <a:tc>
                                            <a:txBody>
                                                <a:p>
                                                    <a:r>
                                                        <a:t>Table 2</a:t>
                                                    </a:r>
                                                </a:p>
                                            </a:txBody>
                                        </a:tc>
                                    </a:tr>
                                </a:tbl>
                            </a:graphicData>
                        </a:graphic>
                    </p:graphicFrame>
                </p:spTree>
            </p:cSld>
        </p:sld>"""
        
        result = self.extractor.extract_tables_with_structure(slide_xml)
        
        assert result['table_count'] == 2
        assert len(result['tables']) == 2
        assert result['tables'][0]['cells'][0][0]['content'] == "Table 1"
        assert result['tables'][1]['cells'][0][0]['content'] == "Table 2"
    
    def test_table_cell_dataclass(self):
        """Test TableCell dataclass initialization."""
        from powerpoint_mcp_server.core.content_extractor import TableCell
        
        cell = TableCell(
            content="Test content",
            row_span=2,
            col_span=3,
            formatting={'fill_color': '#FF0000'}
        )
        
        assert cell.content == "Test content"
        assert cell.row_span == 2
        assert cell.col_span == 3
        assert cell.formatting['fill_color'] == '#FF0000'
    
    def test_table_dataclass(self):
        """Test Table dataclass initialization."""
        from powerpoint_mcp_server.core.content_extractor import Table, TableCell
        
        cells = [[TableCell(content="Cell 1"), TableCell(content="Cell 2")]]
        table = Table(
            rows=1,
            columns=2,
            cells=cells,
            position=(100, 200),
            size=(800, 400)
        )
        
        assert table.rows == 1
        assert table.columns == 2
        assert len(table.cells) == 1
        assert len(table.cells[0]) == 2
        assert table.position == (100, 200)
        assert table.size == (800, 400)    

    def test_extract_presentation_metadata(self):
        """Test extracting presentation-level metadata."""
        presentation_xml = """<?xml version="1.0" encoding="UTF-8"?>
        <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <p:sldMasterIdLst>
                <p:sldMasterId id="2147483648" r:id="rId1"/>
            </p:sldMasterIdLst>
            <p:sldIdLst>
                <p:sldId id="256" r:id="rId2"/>
                <p:sldId id="257" r:id="rId3"/>
                <p:sldId id="258" r:id="rId4"/>
            </p:sldIdLst>
            <p:sldSz cx="9144000" cy="6858000"/>
            <p:notesMasterIdLst>
                <p:notesMasterId r:id="rId5"/>
            </p:notesMasterIdLst>
        </p:presentation>"""
        
        result = self.extractor.extract_presentation_metadata(presentation_xml)
        
        assert result['slide_count'] == 3
        assert result['slide_size']['width'] == 9144000
        assert result['slide_size']['height'] == 6858000
        assert result['slide_master_count'] == 1
        assert result['has_notes_master'] == True
        assert result['has_handout_master'] == False
        assert len(result['slide_ids']) == 3
        assert len(result['slide_master_ids']) == 1
    
    def test_extract_slide_metadata(self):
        """Test extracting slide metadata with object counts."""
        slide_xml = """<?xml version="1.0" encoding="UTF-8"?>
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <p:cSld name="Title Slide">
                <p:spTree>
                    <p:sp>
                        <p:nvSpPr>
                            <p:nvPr>
                                <p:ph type="title"/>
                            </p:nvPr>
                        </p:nvSpPr>
                        <p:txBody>
                            <a:p>
                                <a:r>
                                    <a:t>Sample Title</a:t>
                                </a:r>
                            </a:p>
                        </p:txBody>
                    </p:sp>
                    <p:sp>
                        <p:txBody>
                            <a:p>
                                <a:r>
                                    <a:t>Text box</a:t>
                                </a:r>
                            </a:p>
                        </p:txBody>
                    </p:sp>
                    <p:pic>
                        <p:nvPicPr>
                            <p:cNvPr id="3" name="Image"/>
                        </p:nvPicPr>
                    </p:pic>
                </p:spTree>
            </p:cSld>
        </p:sld>"""
        
        notes_xml = """<?xml version="1.0" encoding="UTF-8"?>
        <p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <p:cSld>
                <p:spTree>
                    <p:sp>
                        <p:nvSpPr>
                            <p:nvPr>
                                <p:ph type="body"/>
                            </p:nvPr>
                        </p:nvSpPr>
                        <p:txBody>
                            <a:p>
                                <a:r>
                                    <a:t>These are speaker notes</a:t>
                                </a:r>
                            </a:p>
                        </p:txBody>
                    </p:sp>
                </p:spTree>
            </p:cSld>
        </p:notes>"""
        
        result = self.extractor.extract_slide_metadata(slide_xml, 1, notes_xml)
        
        assert result['slide_number'] == 1
        assert result['layout_name'] == "Title Slide"
        assert result['title'] == "Sample Title"
        assert result['notes'] == "These are speaker notes"
        assert result['object_counts']['shapes'] == 2
        assert result['object_counts']['text_boxes'] == 2
        assert result['object_counts']['images'] == 1
        assert result['placeholder_count'] == 1
        assert result['text_element_count'] == 2
    
    def test_count_slide_objects(self):
        """Test counting different types of objects on a slide."""
        slide_xml = """<?xml version="1.0" encoding="UTF-8"?>
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
               xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <p:cSld>
                <p:spTree>
                    <p:sp>
                        <p:txBody>
                            <a:p>
                                <a:r>
                                    <a:t>Text shape</a:t>
                                </a:r>
                            </a:p>
                        </p:txBody>
                    </p:sp>
                    <p:pic/>
                    <p:pic/>
                    <p:graphicFrame>
                        <a:graphic>
                            <a:graphicData>
                                <a:tbl>
                                    <a:tr>
                                        <a:tc/>
                                    </a:tr>
                                </a:tbl>
                            </a:graphicData>
                        </a:graphic>
                    </p:graphicFrame>
                    <p:graphicFrame>
                        <a:graphic>
                            <a:graphicData>
                                <c:chart/>
                            </a:graphicData>
                        </a:graphic>
                    </p:graphicFrame>
                    <p:cxnSp/>
                    <p:grpSp>
                        <p:sp/>
                        <p:sp/>
                    </p:grpSp>
                </p:spTree>
            </p:cSld>
        </p:sld>"""
        
        root = self.extractor.xml_parser.parse_xml_string(slide_xml)
        counts = self.extractor._count_slide_objects(root)
        
        assert counts['shapes'] == 1
        assert counts['text_boxes'] == 1
        assert counts['images'] == 2
        assert counts['tables'] == 1
        assert counts['charts'] == 1
        assert counts['connectors'] == 1
        assert counts['groups'] == 1
    
    def test_extract_notes_content(self):
        """Test extracting speaker notes content."""
        notes_xml = """<?xml version="1.0" encoding="UTF-8"?>
        <p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <p:cSld>
                <p:spTree>
                    <p:sp>
                        <p:nvSpPr>
                            <p:nvPr>
                                <p:ph type="sldImg"/>
                            </p:nvPr>
                        </p:nvSpPr>
                    </p:sp>
                    <p:sp>
                        <p:nvSpPr>
                            <p:nvPr>
                                <p:ph type="body"/>
                            </p:nvPr>
                        </p:nvSpPr>
                        <p:txBody>
                            <a:p>
                                <a:r>
                                    <a:t>First note paragraph</a:t>
                                </a:r>
                            </a:p>
                        </p:txBody>
                    </p:sp>
                    <p:sp>
                        <p:txBody>
                            <a:p>
                                <a:r>
                                    <a:t>Second note paragraph</a:t>
                                </a:r>
                            </a:p>
                        </p:txBody>
                    </p:sp>
                </p:spTree>
            </p:cSld>
        </p:notes>"""
        
        result = self.extractor._extract_notes_content(notes_xml)
        
        assert "First note paragraph" in result
        assert "Second note paragraph" in result
        # Should skip the slide image placeholder
        assert len(result.split('\n\n')) == 2
    
    def test_extract_section_information(self):
        """Test extracting section information from presentation."""
        presentation_xml = """<?xml version="1.0" encoding="UTF-8"?>
        <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
            <p:sectionLst>
                <p:section name="Introduction" id="1">
                </p:section>
                <p:section name="Main Content" id="2">
                </p:section>
                <p:section name="Conclusion" id="3">
                </p:section>
            </p:sectionLst>
        </p:presentation>"""
        
        result = self.extractor.extract_section_information(presentation_xml)
        
        assert len(result) == 3
        assert result[0]['name'] == "Introduction"
        assert result[0]['id'] == "1"
        assert result[1]['name'] == "Main Content"
        assert result[2]['name'] == "Conclusion"
    
    def test_get_slide_size_info(self):
        """Test extracting slide size information with unit conversions."""
        presentation_xml = """<?xml version="1.0" encoding="UTF-8"?>
        <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
            <p:sldSz cx="9144000" cy="6858000"/>
        </p:presentation>"""
        
        result = self.extractor.get_slide_size_info(presentation_xml)
        
        assert result['width_emu'] == 9144000
        assert result['height_emu'] == 6858000
        assert result['width_inches'] == 10.0  # 9144000 / 914400
        assert result['height_inches'] == 7.5   # 6858000 / 914400
        assert result['width_cm'] == 25.4      # 10.0 * 2.54
        assert result['height_cm'] == 19.05    # 7.5 * 2.54
        assert result['width_points'] == 720.0 # 10.0 * 72
        assert result['height_points'] == 540.0 # 7.5 * 72
        assert result['aspect_ratio'] == 1.33  # 10.0 / 7.5
    
    def test_extract_slide_metadata_no_notes(self):
        """Test extracting slide metadata without notes."""
        slide_xml = """<?xml version="1.0" encoding="UTF-8"?>
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <p:cSld>
                <p:spTree>
                    <p:sp>
                        <p:txBody>
                            <a:p>
                                <a:r>
                                    <a:t>Simple slide</a:t>
                                </a:r>
                            </a:p>
                        </p:txBody>
                    </p:sp>
                </p:spTree>
            </p:cSld>
        </p:sld>"""
        
        result = self.extractor.extract_slide_metadata(slide_xml, 5)
        
        assert result['slide_number'] == 5
        assert result['notes'] == ""
        assert 'object_counts' in result
    
    def test_extract_presentation_metadata_minimal(self):
        """Test extracting metadata from minimal presentation XML."""
        presentation_xml = """<?xml version="1.0" encoding="UTF-8"?>
        <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
        </p:presentation>"""
        
        result = self.extractor.extract_presentation_metadata(presentation_xml)
        
        assert result['slide_count'] == 0
        assert result['slide_master_count'] == 0
        assert result['has_notes_master'] == False
        assert result['has_handout_master'] == False
    
    def test_extract_notes_content_empty(self):
        """Test extracting notes from empty notes XML."""
        notes_xml = """<?xml version="1.0" encoding="UTF-8"?>
        <p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
            <p:cSld>
                <p:spTree>
                </p:spTree>
            </p:cSld>
        </p:notes>"""
        
        result = self.extractor._extract_notes_content(notes_xml)
        
        assert result == ""


class TestContentExtractorCaching:
    """Test cases for ContentExtractor caching functionality."""
    
    def setup_method(self):
        """Set up test fixtures."""
        from powerpoint_mcp_server.utils.cache_manager import reset_global_cache
        reset_global_cache()  # Reset cache before each test
        self.extractor = ContentExtractor(enable_caching=True)
    
    def teardown_method(self):
        """Clean up after tests."""
        from powerpoint_mcp_server.utils.cache_manager import reset_global_cache
        reset_global_cache()
    
    def test_caching_enabled_initialization(self):
        """Test ContentExtractor initialization with caching enabled."""
        extractor = ContentExtractor(enable_caching=True)
        assert extractor.enable_caching is True
        assert extractor.cache_manager is not None
    
    def test_caching_disabled_initialization(self):
        """Test ContentExtractor initialization with caching disabled."""
        extractor = ContentExtractor(enable_caching=False)
        assert extractor.enable_caching is False
        assert extractor.cache_manager is None
    
    def test_slide_content_caching(self):
        """Test that slide content extraction results are cached."""
        slide_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <p:cSld>
                <p:spTree>
                    <p:sp>
                        <p:nvSpPr>
                            <p:cNvPr id="1" name="Title 1"/>
                            <p:cNvSpPr>
                                <a:spLocks noGrp="1"/>
                            </p:cNvSpPr>
                            <p:nvPr>
                                <p:ph type="title"/>
                            </p:nvPr>
                        </p:nvSpPr>
                        <p:txBody>
                            <a:p>
                                <a:r>
                                    <a:t>Test Title</a:t>
                                </a:r>
                            </a:p>
                        </p:txBody>
                    </p:sp>
                </p:spTree>
            </p:cSld>
        </p:sld>'''
        
        # First extraction should cache the result
        result1 = self.extractor.extract_slide_content(slide_xml, 1)
        assert result1.slide_number == 1
        assert result1.title == "Test Title"
        
        # Second extraction should use cached result
        result2 = self.extractor.extract_slide_content(slide_xml, 1)
        assert result2.slide_number == 1
        assert result2.title == "Test Title"
        
        # Results should be identical (from cache)
        assert result1.title == result2.title
        assert result1.slide_number == result2.slide_number
    
    def test_different_slides_different_cache_keys(self):
        """Test that different slides get different cache keys."""
        slide_xml1 = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <p:cSld>
                <p:spTree>
                    <p:sp>
                        <p:nvSpPr>
                            <p:cNvPr id="1" name="Title 1"/>
                            <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
                            <p:nvPr><p:ph type="title"/></p:nvPr>
                        </p:nvSpPr>
                        <p:txBody>
                            <a:p><a:r><a:t>Title 1</a:t></a:r></a:p>
                        </p:txBody>
                    </p:sp>
                </p:spTree>
            </p:cSld>
        </p:sld>'''
        
        slide_xml2 = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <p:cSld>
                <p:spTree>
                    <p:sp>
                        <p:nvSpPr>
                            <p:cNvPr id="1" name="Title 1"/>
                            <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
                            <p:nvPr><p:ph type="title"/></p:nvPr>
                        </p:nvSpPr>
                        <p:txBody>
                            <a:p><a:r><a:t>Title 2</a:t></a:r></a:p>
                        </p:txBody>
                    </p:sp>
                </p:spTree>
            </p:cSld>
        </p:sld>'''
        
        # Extract different slides
        result1 = self.extractor.extract_slide_content(slide_xml1, 1)
        result2 = self.extractor.extract_slide_content(slide_xml2, 2)
        
        # Should have different titles
        assert result1.title == "Title 1"
        assert result2.title == "Title 2"
        
        # Cache should contain both entries
        stats = self.extractor.get_cache_stats()
        assert stats['caching_enabled'] is True
        assert stats['content_cache']['total_entries'] >= 2
    
    def test_cache_stats(self):
        """Test cache statistics functionality."""
        # Initially empty cache
        stats = self.extractor.get_cache_stats()
        assert stats['caching_enabled'] is True
        assert stats['content_cache']['total_entries'] == 0
        
        # Add some cached content
        slide_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
            <p:cSld><p:spTree></p:spTree></p:cSld>
        </p:sld>'''
        
        self.extractor.extract_slide_content(slide_xml, 1)
        
        # Check stats after caching
        stats = self.extractor.get_cache_stats()
        assert stats['content_cache']['total_entries'] >= 1
    
    def test_clear_cache(self):
        """Test cache clearing functionality."""
        slide_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
            <p:cSld><p:spTree></p:spTree></p:cSld>
        </p:sld>'''
        
        # Add content to cache
        self.extractor.extract_slide_content(slide_xml, 1)
        
        # Verify cache has content
        stats = self.extractor.get_cache_stats()
        assert stats['content_cache']['total_entries'] >= 1
        
        # Clear cache
        self.extractor.clear_cache()
        
        # Verify cache is empty
        stats = self.extractor.get_cache_stats()
        assert stats['content_cache']['total_entries'] == 0
    
    def test_cleanup_expired_cache(self):
        """Test cleanup of expired cache entries."""
        # This test verifies the method exists and can be called
        removed_count = self.extractor.cleanup_expired_cache()
        assert isinstance(removed_count, int)
        assert removed_count >= 0
    
    def test_caching_disabled_stats(self):
        """Test cache statistics when caching is disabled."""
        extractor = ContentExtractor(enable_caching=False)
        stats = extractor.get_cache_stats()
        
        assert stats['caching_enabled'] is False
        assert stats['content_cache']['total_entries'] == 0
        assert stats['xml_parser_cache']['cached_elements'] == 0