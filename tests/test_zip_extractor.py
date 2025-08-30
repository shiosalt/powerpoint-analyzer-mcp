"""Unit tests for ZipExtractor class."""

import os
import tempfile
import zipfile
from pathlib import Path
import pytest

from powerpoint_mcp_server.utils.zip_extractor import ZipExtractor, ZipExtractionError
from powerpoint_mcp_server.utils.file_validator import FileValidationError


class TestZipExtractor:
    """Test cases for ZipExtractor class."""
    
    def create_test_pptx(self, include_slides=True, include_layouts=True, include_notes=True):
        """Helper method to create a test .pptx file."""
        tmp = tempfile.NamedTemporaryFile(suffix='.pptx', delete=False)
        tmp_path = tmp.name
        tmp.close()
        
        with zipfile.ZipFile(tmp_path, 'w') as zf:
            # Required files
            zf.writestr('[Content_Types].xml', '<?xml version="1.0"?><Types/>')
            zf.writestr('_rels/.rels', '<?xml version="1.0"?><Relationships/>')
            zf.writestr('ppt/presentation.xml', '<?xml version="1.0"?><p:presentation/>')
            
            # Optional files based on parameters
            if include_slides:
                zf.writestr('ppt/slides/slide1.xml', '<?xml version="1.0"?><p:sld/>')
                zf.writestr('ppt/slides/slide2.xml', '<?xml version="1.0"?><p:sld/>')
                zf.writestr('ppt/slides/_rels/slide1.xml.rels', '<?xml version="1.0"?><Relationships/>')
            
            if include_layouts:
                zf.writestr('ppt/slideLayouts/slideLayout1.xml', '<?xml version="1.0"?><p:sldLayout/>')
                zf.writestr('ppt/slideLayouts/slideLayout2.xml', '<?xml version="1.0"?><p:sldLayout/>')
            
            if include_notes:
                zf.writestr('ppt/notesSlides/notesSlide1.xml', '<?xml version="1.0"?><p:notes/>')
            
            # Add some media files
            zf.writestr('ppt/media/image1.png', b'fake image data')
        
        return tmp_path
    
    def test_init_with_valid_file(self):
        """Test initialization with valid .pptx file."""
        tmp_path = self.create_test_pptx()
        try:
            extractor = ZipExtractor(tmp_path)
            assert extractor.file_path == tmp_path
            assert extractor.temp_dir is None
            assert len(extractor._extracted_files) == 0
        finally:
            os.unlink(tmp_path)
    
    def test_init_with_invalid_file(self):
        """Test initialization with invalid file raises exception."""
        with pytest.raises(FileValidationError):
            ZipExtractor("nonexistent.pptx")
    
    def test_extract_archive_context_manager(self):
        """Test extract_archive context manager."""
        tmp_path = self.create_test_pptx()
        try:
            extractor = ZipExtractor(tmp_path)
            
            with extractor.extract_archive() as extracted:
                assert extracted.temp_dir is not None
                assert len(extracted._extracted_files) > 0
                assert os.path.exists(extracted.temp_dir)
            
            # After context manager, temp files should be cleaned up
            assert extractor.temp_dir is None or not os.path.exists(extractor.temp_dir)
        finally:
            os.unlink(tmp_path)
    
    def test_with_statement_support(self):
        """Test 'with' statement support."""
        tmp_path = self.create_test_pptx()
        try:
            with ZipExtractor(tmp_path) as extractor:
                assert extractor.temp_dir is not None
                assert len(extractor._extracted_files) > 0
            
            # After with statement, temp files should be cleaned up
            assert extractor.temp_dir is None or not os.path.exists(extractor.temp_dir)
        finally:
            os.unlink(tmp_path)
    
    def test_get_xml_files(self):
        """Test getting XML files from extracted archive."""
        tmp_path = self.create_test_pptx()
        try:
            with ZipExtractor(tmp_path) as extractor:
                xml_files = extractor.get_xml_files()
                
                assert len(xml_files) > 0
                assert any('presentation.xml' in path for path in xml_files.keys())
                assert any('[Content_Types].xml' in path for path in xml_files.keys())
                assert any('.rels' in path for path in xml_files.keys())
                
                # Check that all returned paths exist
                for extracted_path in xml_files.values():
                    assert os.path.exists(extracted_path)
        finally:
            os.unlink(tmp_path)
    
    def test_get_xml_files_without_extraction(self):
        """Test that get_xml_files raises error when archive not extracted."""
        tmp_path = self.create_test_pptx()
        try:
            extractor = ZipExtractor(tmp_path)
            with pytest.raises(ZipExtractionError):
                extractor.get_xml_files()
        finally:
            os.unlink(tmp_path)
    
    def test_get_specific_xml(self):
        """Test getting specific XML file."""
        tmp_path = self.create_test_pptx()
        try:
            with ZipExtractor(tmp_path) as extractor:
                # Test existing file
                presentation_path = extractor.get_specific_xml('ppt/presentation.xml')
                assert presentation_path is not None
                assert os.path.exists(presentation_path)
                
                # Test non-existing file
                nonexistent_path = extractor.get_specific_xml('nonexistent.xml')
                assert nonexistent_path is None
        finally:
            os.unlink(tmp_path)
    
    def test_read_xml_content(self):
        """Test reading XML content."""
        tmp_path = self.create_test_pptx()
        try:
            with ZipExtractor(tmp_path) as extractor:
                content = extractor.read_xml_content('ppt/presentation.xml')
                assert content is not None
                assert '<?xml version="1.0"?>' in content
                assert 'p:presentation' in content
        finally:
            os.unlink(tmp_path)
    
    def test_read_xml_content_nonexistent(self):
        """Test reading non-existent XML file raises error."""
        tmp_path = self.create_test_pptx()
        try:
            with ZipExtractor(tmp_path) as extractor:
                with pytest.raises(ZipExtractionError):
                    extractor.read_xml_content('nonexistent.xml')
        finally:
            os.unlink(tmp_path)
    
    def test_list_archive_contents(self):
        """Test listing archive contents."""
        tmp_path = self.create_test_pptx()
        try:
            with ZipExtractor(tmp_path) as extractor:
                contents = extractor.list_archive_contents()
                
                assert len(contents) > 0
                assert '[Content_Types].xml' in contents
                assert 'ppt/presentation.xml' in contents
                assert '_rels/.rels' in contents
        finally:
            os.unlink(tmp_path)
    
    def test_get_slide_xml_files(self):
        """Test getting slide XML files."""
        tmp_path = self.create_test_pptx()
        try:
            with ZipExtractor(tmp_path) as extractor:
                slide_files = extractor.get_slide_xml_files()
                
                assert len(slide_files) == 2  # We created 2 slides
                assert 'ppt/slides/slide1.xml' in slide_files
                assert 'ppt/slides/slide2.xml' in slide_files
                
                # Check that paths exist
                for path in slide_files.values():
                    assert os.path.exists(path)
        finally:
            os.unlink(tmp_path)
    
    def test_get_slide_layout_xml_files(self):
        """Test getting slide layout XML files."""
        tmp_path = self.create_test_pptx()
        try:
            with ZipExtractor(tmp_path) as extractor:
                layout_files = extractor.get_slide_layout_xml_files()
                
                assert len(layout_files) == 2  # We created 2 layouts
                assert 'ppt/slideLayouts/slideLayout1.xml' in layout_files
                assert 'ppt/slideLayouts/slideLayout2.xml' in layout_files
        finally:
            os.unlink(tmp_path)
    
    def test_get_notes_xml_files(self):
        """Test getting notes XML files."""
        tmp_path = self.create_test_pptx()
        try:
            with ZipExtractor(tmp_path) as extractor:
                notes_files = extractor.get_notes_xml_files()
                
                assert len(notes_files) == 1  # We created 1 notes slide
                assert 'ppt/notesSlides/notesSlide1.xml' in notes_files
        finally:
            os.unlink(tmp_path)
    
    def test_get_slide_xml_files_empty(self):
        """Test getting slide XML files when none exist."""
        tmp_path = self.create_test_pptx(include_slides=False)
        try:
            with ZipExtractor(tmp_path) as extractor:
                slide_files = extractor.get_slide_xml_files()
                assert len(slide_files) == 0
        finally:
            os.unlink(tmp_path)
    
    def test_cleanup_temp_files(self):
        """Test cleanup of temporary files."""
        tmp_path = self.create_test_pptx()
        try:
            extractor = ZipExtractor(tmp_path)
            extractor._extract_to_temp()
            
            temp_dir = extractor.temp_dir
            assert temp_dir is not None
            assert os.path.exists(temp_dir)
            
            extractor.cleanup_temp_files()
            
            assert extractor.temp_dir is None
            assert not os.path.exists(temp_dir)
            assert len(extractor._extracted_files) == 0
        finally:
            os.unlink(tmp_path)
    
    def test_get_archive_info(self):
        """Test getting archive information without extraction."""
        tmp_path = self.create_test_pptx()
        try:
            info = ZipExtractor.get_archive_info(tmp_path)
            
            assert 'total_files' in info
            assert 'xml_files_count' in info
            assert 'slide_count' in info
            assert 'layout_count' in info
            assert 'notes_count' in info
            assert 'media_count' in info
            
            assert info['slide_count'] == 2
            assert info['layout_count'] == 2
            assert info['notes_count'] == 1
            assert info['media_count'] == 1
            assert info['has_presentation_xml']
            assert info['has_content_types']
            
            assert len(info['slide_files']) == 2
            assert len(info['layout_files']) == 2
            assert len(info['notes_files']) == 1
        finally:
            os.unlink(tmp_path)
    
    def test_get_archive_info_invalid_file(self):
        """Test getting archive info for invalid file."""
        info = ZipExtractor.get_archive_info("nonexistent.pptx")
        assert 'error' in info
    
    def test_extraction_error_handling(self):
        """Test error handling during extraction."""
        # Create invalid ZIP file
        with tempfile.NamedTemporaryFile(suffix='.pptx', delete=False) as tmp:
            tmp.write(b"not a zip file")
            tmp_path = tmp.name
        
        try:
            # This should pass validation but fail during extraction
            # We need to bypass validation for this test
            extractor = ZipExtractor.__new__(ZipExtractor)
            extractor.file_path = tmp_path
            extractor.temp_dir = None
            extractor._extracted_files = {}
            
            with pytest.raises(ZipExtractionError):
                extractor._extract_to_temp()
        finally:
            os.unlink(tmp_path)