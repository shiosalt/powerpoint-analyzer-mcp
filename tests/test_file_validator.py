"""Unit tests for FileValidator class."""

import os
import tempfile
import zipfile
from pathlib import Path
import pytest

from powerpoint_mcp_server.utils.file_validator import FileValidator, FileValidationError


class TestFileValidator:
    """Test cases for FileValidator class."""
    
    def test_validate_nonexistent_file(self):
        """Test validation of non-existent file."""
        is_valid, error = FileValidator.validate_file("nonexistent.pptx")
        assert not is_valid
        assert "does not exist" in error
    
    def test_validate_file_strict_raises_exception(self):
        """Test that validate_file_strict raises exception for invalid file."""
        with pytest.raises(FileValidationError):
            FileValidator.validate_file_strict("nonexistent.pptx")
    
    def test_validate_wrong_extension(self):
        """Test validation of file with wrong extension."""
        with tempfile.NamedTemporaryFile(suffix='.txt', delete=False) as tmp:
            tmp.write(b"test content")
            tmp_path = tmp.name
        
        try:
            is_valid, error = FileValidator.validate_file(tmp_path)
            assert not is_valid
            assert "Unsupported file format" in error
        finally:
            os.unlink(tmp_path)
    
    def test_validate_empty_file(self):
        """Test validation of empty file."""
        with tempfile.NamedTemporaryFile(suffix='.pptx', delete=False) as tmp:
            tmp_path = tmp.name
        
        try:
            is_valid, error = FileValidator.validate_file(tmp_path)
            assert not is_valid
            assert "empty" in error
        finally:
            os.unlink(tmp_path)
    
    def test_validate_oversized_file(self):
        """Test validation of oversized file."""
        with tempfile.NamedTemporaryFile(suffix='.pptx', delete=False) as tmp:
            # Create a file larger than the limit
            large_content = b'x' * (FileValidator.MAX_FILE_SIZE + 1)
            tmp.write(large_content)
            tmp_path = tmp.name
        
        try:
            is_valid, error = FileValidator.validate_file(tmp_path)
            assert not is_valid
            assert "too large" in error
        finally:
            os.unlink(tmp_path)
    
    def test_validate_invalid_zip_file(self):
        """Test validation of invalid ZIP file."""
        with tempfile.NamedTemporaryFile(suffix='.pptx', delete=False) as tmp:
            tmp.write(b"not a zip file")
            tmp_path = tmp.name
        
        try:
            is_valid, error = FileValidator.validate_file(tmp_path)
            assert not is_valid
            assert "not a valid ZIP archive" in error
        finally:
            os.unlink(tmp_path)
    
    def test_validate_zip_missing_required_files(self):
        """Test validation of ZIP file missing required .pptx files."""
        with tempfile.NamedTemporaryFile(suffix='.pptx', delete=False) as tmp:
            tmp_path = tmp.name
        
        try:
            # Create a valid ZIP but without required .pptx structure
            with zipfile.ZipFile(tmp_path, 'w') as zf:
                zf.writestr('some_file.txt', 'content')
            
            is_valid, error = FileValidator.validate_file(tmp_path)
            assert not is_valid
            assert "missing required files" in error
        finally:
            os.unlink(tmp_path)
    
    def test_validate_valid_pptx_structure(self):
        """Test validation of valid .pptx file structure."""
        with tempfile.NamedTemporaryFile(suffix='.pptx', delete=False) as tmp:
            tmp_path = tmp.name
        
        try:
            # Create a ZIP with minimal required .pptx structure
            with zipfile.ZipFile(tmp_path, 'w') as zf:
                zf.writestr('[Content_Types].xml', '<?xml version="1.0"?><Types/>')
                zf.writestr('_rels/.rels', '<?xml version="1.0"?><Relationships/>')
                zf.writestr('ppt/presentation.xml', '<?xml version="1.0"?><p:presentation/>')
            
            is_valid, error = FileValidator.validate_file(tmp_path)
            assert is_valid
            assert error is None
        finally:
            os.unlink(tmp_path)
    
    def test_check_file_exists(self):
        """Test file existence check."""
        with tempfile.NamedTemporaryFile() as tmp:
            assert FileValidator._check_file_exists(tmp.name)
        
        assert not FileValidator._check_file_exists("nonexistent.pptx")
    
    def test_check_file_extension(self):
        """Test file extension check."""
        assert FileValidator._check_file_extension("test.pptx")
        assert FileValidator._check_file_extension("test.PPTX")  # Case insensitive
        assert not FileValidator._check_file_extension("test.ppt")
        assert not FileValidator._check_file_extension("test.txt")
        assert not FileValidator._check_file_extension("test")
    
    def test_check_file_size(self):
        """Test file size check."""
        with tempfile.NamedTemporaryFile() as tmp:
            tmp.write(b"test content")
            tmp.flush()
            
            is_valid, error = FileValidator._check_file_size(tmp.name)
            assert is_valid
            assert error is None
    
    def test_get_file_info_existing_file(self):
        """Test getting file info for existing file."""
        with tempfile.NamedTemporaryFile(suffix='.pptx') as tmp:
            tmp.write(b"test content")
            tmp.flush()
            
            info = FileValidator.get_file_info(tmp.name)
            assert info['exists']
            assert info['is_file']
            assert info['extension'] == '.pptx'
            assert info['size_bytes'] > 0
            assert 'path' in info
            assert 'name' in info
    
    def test_get_file_info_nonexistent_file(self):
        """Test getting file info for non-existent file."""
        info = FileValidator.get_file_info("nonexistent.pptx")
        assert not info['exists']
        assert 'error' in info
    
    def test_supported_extensions(self):
        """Test that only .pptx is supported."""
        assert '.pptx' in FileValidator.SUPPORTED_EXTENSIONS
        assert '.ppt' not in FileValidator.SUPPORTED_EXTENSIONS
        assert '.txt' not in FileValidator.SUPPORTED_EXTENSIONS
    
    def test_required_pptx_files(self):
        """Test that required .pptx files are defined."""
        required = FileValidator.REQUIRED_PPTX_FILES
        assert '[Content_Types].xml' in required
        assert '_rels/.rels' in required
        assert 'ppt/presentation.xml' in required
    
    def test_max_file_size_constant(self):
        """Test that max file size is reasonable."""
        assert FileValidator.MAX_FILE_SIZE == 100 * 1024 * 1024  # 100MB