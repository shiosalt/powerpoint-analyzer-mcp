"""ZIP extraction utilities for PowerPoint files."""

import os
import tempfile
import zipfile
from pathlib import Path
from typing import Dict, List, Optional, Set
from contextlib import contextmanager

from .file_validator import FileValidator, FileValidationError


class ZipExtractionError(Exception):
    """Custom exception for ZIP extraction errors."""
    pass


class ZipExtractor:
    """Extracts and manages .pptx file contents as ZIP archives."""
    
    def __init__(self, file_path: str):
        """
        Initialize ZipExtractor with a .pptx file.
        
        Args:
            file_path: Path to the .pptx file
            
        Raises:
            FileValidationError: If file is invalid
        """
        self.file_path = file_path
        self.temp_dir: Optional[str] = None
        self._extracted_files: Dict[str, str] = {}
        
        # Validate file before proceeding
        FileValidator.validate_file_strict(file_path)
    
    @contextmanager
    def extract_archive(self):
        """
        Context manager for extracting .pptx archive.
        Automatically cleans up temporary files when done.
        
        Yields:
            ZipExtractor instance with extracted files
        """
        try:
            self._extract_to_temp()
            yield self
        finally:
            self.cleanup_temp_files()
    
    def _extract_to_temp(self) -> None:
        """Extract .pptx archive to temporary directory."""
        try:
            # Create temporary directory
            self.temp_dir = tempfile.mkdtemp(prefix='pptx_extract_')
            
            # Extract ZIP archive
            with zipfile.ZipFile(self.file_path, 'r') as zip_file:
                zip_file.extractall(self.temp_dir)
                
                # Build mapping of extracted files
                for file_info in zip_file.filelist:
                    if not file_info.is_dir():
                        extracted_path = os.path.join(self.temp_dir, file_info.filename)
                        self._extracted_files[file_info.filename] = extracted_path
                        
        except zipfile.BadZipFile as e:
            raise ZipExtractionError(f"Invalid ZIP file: {str(e)}")
        except Exception as e:
            self.cleanup_temp_files()
            raise ZipExtractionError(f"Failed to extract archive: {str(e)}")
    
    def get_xml_files(self) -> Dict[str, str]:
        """
        Get mapping of XML files in the extracted archive.
        
        Returns:
            Dictionary mapping XML file paths to their extracted locations
        """
        if not self._extracted_files:
            raise ZipExtractionError("Archive not extracted. Use extract_archive() context manager.")
        
        xml_files = {}
        for archive_path, extracted_path in self._extracted_files.items():
            if archive_path.endswith('.xml') or archive_path.endswith('.rels'):
                xml_files[archive_path] = extracted_path
        
        return xml_files
    
    def get_specific_xml(self, xml_path: str) -> Optional[str]:
        """
        Get the extracted path for a specific XML file.
        
        Args:
            xml_path: Path within the archive (e.g., 'ppt/presentation.xml')
            
        Returns:
            Extracted file path or None if not found
        """
        if not self._extracted_files:
            raise ZipExtractionError("Archive not extracted. Use extract_archive() context manager.")
        
        return self._extracted_files.get(xml_path)
    
    def read_xml_content(self, xml_path: str) -> str:
        """
        Read content of a specific XML file.
        
        Args:
            xml_path: Path within the archive (e.g., 'ppt/presentation.xml')
            
        Returns:
            XML content as string
            
        Raises:
            ZipExtractionError: If file not found or cannot be read
        """
        extracted_path = self.get_specific_xml(xml_path)
        if not extracted_path:
            raise ZipExtractionError(f"XML file not found in archive: {xml_path}")
        
        try:
            with open(extracted_path, 'r', encoding='utf-8') as f:
                return f.read()
        except Exception as e:
            raise ZipExtractionError(f"Failed to read XML file {xml_path}: {str(e)}")
    
    def list_archive_contents(self) -> List[str]:
        """
        List all files in the extracted archive.
        
        Returns:
            List of file paths within the archive
        """
        if not self._extracted_files:
            raise ZipExtractionError("Archive not extracted. Use extract_archive() context manager.")
        
        return list(self._extracted_files.keys())
    
    def get_slide_xml_files(self) -> Dict[str, str]:
        """
        Get mapping of slide XML files.
        
        Returns:
            Dictionary mapping slide XML paths to extracted locations
        """
        xml_files = self.get_xml_files()
        slide_files = {}
        
        for archive_path, extracted_path in xml_files.items():
            if archive_path.startswith('ppt/slides/slide') and archive_path.endswith('.xml'):
                slide_files[archive_path] = extracted_path
        
        return slide_files
    
    def get_slide_layout_xml_files(self) -> Dict[str, str]:
        """
        Get mapping of slide layout XML files.
        
        Returns:
            Dictionary mapping layout XML paths to extracted locations
        """
        xml_files = self.get_xml_files()
        layout_files = {}
        
        for archive_path, extracted_path in xml_files.items():
            if archive_path.startswith('ppt/slideLayouts/slideLayout') and archive_path.endswith('.xml'):
                layout_files[archive_path] = extracted_path
        
        return layout_files
    
    def get_notes_xml_files(self) -> Dict[str, str]:
        """
        Get mapping of notes XML files.
        
        Returns:
            Dictionary mapping notes XML paths to extracted locations
        """
        xml_files = self.get_xml_files()
        notes_files = {}
        
        for archive_path, extracted_path in xml_files.items():
            if archive_path.startswith('ppt/notesSlides/notesSlide') and archive_path.endswith('.xml'):
                notes_files[archive_path] = extracted_path
        
        return notes_files
    
    def cleanup_temp_files(self) -> None:
        """Clean up temporary extracted files."""
        if self.temp_dir and os.path.exists(self.temp_dir):
            try:
                import shutil
                shutil.rmtree(self.temp_dir)
                self.temp_dir = None
                self._extracted_files.clear()
            except Exception as e:
                # Log error but don't raise - cleanup is best effort
                import logging
                logging.warning(f"Failed to cleanup temporary files: {str(e)}")
    
    def __enter__(self):
        """Support for 'with' statement."""
        self._extract_to_temp()
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """Support for 'with' statement."""
        self.cleanup_temp_files()
    
    @staticmethod
    def get_archive_info(file_path: str) -> Dict:
        """
        Get information about a .pptx archive without extracting it.
        
        Args:
            file_path: Path to the .pptx file
            
        Returns:
            Dictionary with archive information
        """
        try:
            FileValidator.validate_file_strict(file_path)
            
            with zipfile.ZipFile(file_path, 'r') as zip_file:
                files = zip_file.namelist()
                
                # Count different file types
                xml_files = [f for f in files if f.endswith('.xml') or f.endswith('.rels')]
                slide_files = [f for f in files if f.startswith('ppt/slides/slide') and f.endswith('.xml')]
                layout_files = [f for f in files if f.startswith('ppt/slideLayouts/') and f.endswith('.xml')]
                notes_files = [f for f in files if f.startswith('ppt/notesSlides/') and f.endswith('.xml')]
                media_files = [f for f in files if f.startswith('ppt/media/')]
                
                return {
                    'total_files': len(files),
                    'xml_files_count': len(xml_files),
                    'slide_count': len(slide_files),
                    'layout_count': len(layout_files),
                    'notes_count': len(notes_files),
                    'media_count': len(media_files),
                    'slide_files': slide_files,
                    'layout_files': layout_files,
                    'notes_files': notes_files,
                    'has_presentation_xml': 'ppt/presentation.xml' in files,
                    'has_content_types': '[Content_Types].xml' in files
                }
                
        except Exception as e:
            return {'error': str(e)}