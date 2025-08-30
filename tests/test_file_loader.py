"""Tests for file loading functionality."""

import pytest
import tempfile
import os
from pathlib import Path
from powerpoint_mcp_server.core.file_loader import FileLoader


class TestFileLoader:
    """Test cases for FileLoader."""
    
    def setup_method(self):
        """Set up test fixtures."""
        self.file_loader = FileLoader()
    
    def test_validate_pptx_format_valid(self):
        """Test validation of valid .pptx file extension."""
        assert self.file_loader.validate_pptx_format("test.pptx") is True
        assert self.file_loader.validate_pptx_format("presentation.PPTX") is True
    
    def test_validate_pptx_format_invalid(self):
        """Test validation of invalid file extensions."""
        assert self.file_loader.validate_pptx_format("test.ppt") is False
        assert self.file_loader.validate_pptx_format("test.txt") is False
        assert self.file_loader.validate_pptx_format("test.docx") is False
    
    def test_validate_file_not_found(self):
        """Test validation of non-existent file."""
        with pytest.raises(FileNotFoundError):
            self.file_loader.validate_file("nonexistent.pptx")
    
    def test_validate_file_wrong_format(self):
        """Test validation of file with wrong format."""
        with tempfile.NamedTemporaryFile(suffix=".txt", delete=False) as tmp:
            tmp.write(b"test content")
            tmp_path = tmp.name
        
        try:
            with pytest.raises(ValueError, match="Unsupported file format"):
                self.file_loader.validate_file(tmp_path)
        finally:
            os.unlink(tmp_path)
    
    def test_get_file_metadata(self):
        """Test file metadata extraction."""
        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
            tmp.write(b"test content")
            tmp_path = tmp.name
        
        try:
            metadata = self.file_loader.get_file_metadata(tmp_path)
            assert "name" in metadata
            assert "size" in metadata
            assert "modified" in metadata
            assert "extension" in metadata
            assert metadata["extension"] == ".pptx"
            assert metadata["size"] > 0
        finally:
            os.unlink(tmp_path)