"""File validation utilities for PowerPoint files."""

import os
import zipfile
from pathlib import Path
from typing import Optional, Tuple


class FileValidationError(Exception):
    """Custom exception for file validation errors."""
    pass


class FileValidator:
    """Validates PowerPoint files for processing."""
    
    # Maximum file size in bytes (100MB)
    MAX_FILE_SIZE = 100 * 1024 * 1024
    
    # Supported file extensions
    SUPPORTED_EXTENSIONS = {'.pptx'}
    
    # Required files in a valid .pptx archive
    REQUIRED_PPTX_FILES = {
        '[Content_Types].xml',
        '_rels/.rels',
        'ppt/presentation.xml'
    }
    
    @classmethod
    def validate_file(cls, file_path: str) -> Tuple[bool, Optional[str]]:
        """
        Validate a PowerPoint file for processing.
        
        Args:
            file_path: Path to the file to validate
            
        Returns:
            Tuple of (is_valid, error_message)
        """
        try:
            # Check file existence
            if not cls._check_file_exists(file_path):
                return False, f"File does not exist: {file_path}"
            
            # Check file extension
            if not cls._check_file_extension(file_path):
                return False, f"Unsupported file format. Only .pptx files are supported."
            
            # Check file size
            size_valid, size_error = cls._check_file_size(file_path)
            if not size_valid:
                return False, size_error
            
            # Check if file is a valid .pptx format
            format_valid, format_error = cls._check_pptx_format(file_path)
            if not format_valid:
                return False, format_error
                
            return True, None
            
        except Exception as e:
            return False, f"Validation error: {str(e)}"
    
    @classmethod
    def validate_file_strict(cls, file_path: str) -> None:
        """
        Validate a PowerPoint file and raise exception if invalid.
        
        Args:
            file_path: Path to the file to validate
            
        Raises:
            FileValidationError: If file is invalid
        """
        is_valid, error_message = cls.validate_file(file_path)
        if not is_valid:
            raise FileValidationError(error_message)
    
    @classmethod
    def _check_file_exists(cls, file_path: str) -> bool:
        """Check if file exists and is accessible."""
        try:
            path = Path(file_path)
            return path.exists() and path.is_file()
        except (OSError, ValueError):
            return False
    
    @classmethod
    def _check_file_extension(cls, file_path: str) -> bool:
        """Check if file has a supported extension."""
        try:
            path = Path(file_path)
            return path.suffix.lower() in cls.SUPPORTED_EXTENSIONS
        except (OSError, ValueError):
            return False
    
    @classmethod
    def _check_file_size(cls, file_path: str) -> Tuple[bool, Optional[str]]:
        """Check if file size is within acceptable limits."""
        try:
            file_size = os.path.getsize(file_path)
            if file_size > cls.MAX_FILE_SIZE:
                size_mb = file_size / (1024 * 1024)
                max_mb = cls.MAX_FILE_SIZE / (1024 * 1024)
                return False, f"File too large: {size_mb:.1f}MB (max: {max_mb}MB)"
            
            if file_size == 0:
                return False, "File is empty"
                
            return True, None
            
        except (OSError, ValueError) as e:
            return False, f"Cannot determine file size: {str(e)}"
    
    @classmethod
    def _check_pptx_format(cls, file_path: str) -> Tuple[bool, Optional[str]]:
        """Check if file is a valid .pptx format (ZIP archive with required structure)."""
        try:
            with zipfile.ZipFile(file_path, 'r') as zip_file:
                # Check if it's a valid ZIP file
                if zip_file.testzip() is not None:
                    return False, "File appears to be corrupted (ZIP test failed)"
                
                # Get list of files in the archive
                archive_files = set(zip_file.namelist())
                
                # Check for required .pptx files
                missing_files = cls.REQUIRED_PPTX_FILES - archive_files
                if missing_files:
                    return False, f"Invalid .pptx format: missing required files {missing_files}"
                
                return True, None
                
        except zipfile.BadZipFile:
            return False, "File is not a valid ZIP archive"
        except Exception as e:
            return False, f"Error validating .pptx format: {str(e)}"
    
    @classmethod
    def get_file_info(cls, file_path: str) -> dict:
        """
        Get basic information about a file.
        
        Args:
            file_path: Path to the file
            
        Returns:
            Dictionary with file information
        """
        try:
            path = Path(file_path)
            stat = path.stat()
            
            return {
                'path': str(path.absolute()),
                'name': path.name,
                'size_bytes': stat.st_size,
                'size_mb': round(stat.st_size / (1024 * 1024), 2),
                'extension': path.suffix.lower(),
                'exists': True,
                'is_file': path.is_file(),
                'modified_time': stat.st_mtime
            }
        except Exception as e:
            return {
                'path': file_path,
                'error': str(e),
                'exists': False
            }