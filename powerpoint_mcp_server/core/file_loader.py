"""File loading and validation functionality."""

import os
from pathlib import Path
from typing import Optional, Dict, Any


class FileLoader:
    """Handles PowerPoint file loading and basic validation."""
    
    def __init__(self):
        """Initialize the file loader."""
        pass
    
    def load_presentation(self, file_path: str) -> Dict[str, Any]:
        """Load a PowerPoint presentation file.
        
        Args:
            file_path: Path to the PowerPoint file
            
        Returns:
            Dictionary containing file metadata and status
            
        Raises:
            FileNotFoundError: If file doesn't exist
            ValueError: If file format is not supported
        """
        # TODO: Implement in later tasks
        self.validate_file(file_path)
        return {
            "file_path": file_path,
            "status": "loaded",
            "metadata": self.get_file_metadata(file_path)
        }
    
    def validate_file(self, file_path: str) -> bool:
        """Validate PowerPoint file existence and format.
        
        Args:
            file_path: Path to the file to validate
            
        Returns:
            True if file is valid
            
        Raises:
            FileNotFoundError: If file doesn't exist
            ValueError: If file format is not supported
        """
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"File not found: {file_path}")
        
        if not self.validate_pptx_format(file_path):
            raise ValueError(f"Unsupported file format. Only .pptx files are supported: {file_path}")
        
        return True
    
    def get_file_metadata(self, file_path: str) -> Dict[str, Any]:
        """Get basic file metadata.
        
        Args:
            file_path: Path to the file
            
        Returns:
            Dictionary containing file metadata
        """
        path = Path(file_path)
        stat = path.stat()
        
        return {
            "name": path.name,
            "size": stat.st_size,
            "modified": stat.st_mtime,
            "extension": path.suffix.lower()
        }
    
    def validate_pptx_format(self, file_path: str) -> bool:
        """Validate that file is a .pptx format.
        
        Args:
            file_path: Path to the file
            
        Returns:
            True if file has .pptx extension
        """
        return Path(file_path).suffix.lower() == '.pptx'