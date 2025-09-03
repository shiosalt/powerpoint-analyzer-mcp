#!/usr/bin/env python3
"""Main entry point for the PowerPoint Analyzer MCP using FastMCP 2.0."""

import asyncio
import json
import logging
import sys
import os
from typing import Any, Dict, List, Optional
from pathlib import Path
from contextlib import asynccontextmanager

# Use the correct FastMCP 2.0 import
from fastmcp import FastMCP
from powerpoint_mcp_server.server import PowerPointMCPServer
from powerpoint_mcp_server.config import get_config, get_config_manager

# Configure logging
config = get_config()

# Create log file handler
log_file = "powerpoint_mcp_server.log"
file_handler = logging.FileHandler(log_file, mode='w', encoding='utf-8')
file_handler.setLevel(logging.DEBUG)

# Create console handler
console_handler = logging.StreamHandler(sys.stderr)
console_handler.setLevel(getattr(logging, config.log_level.upper(), logging.INFO))

# Create formatter
formatter = logging.Formatter(config.log_format)
file_handler.setFormatter(formatter)
console_handler.setFormatter(formatter)

# Configure root logger
logging.basicConfig(
    level=logging.DEBUG,
    handlers=[file_handler, console_handler]
)

logger = logging.getLogger(__name__)

# Initialize global PowerPoint server instance
powerpoint_server: Optional[PowerPointMCPServer] = None

@asynccontextmanager
async def lifespan(app):
    """Lifespan context manager for the FastMCP server."""
    global powerpoint_server
    
    # Startup
    logger.info("Initializing PowerPoint Analyzer MCP...")
    powerpoint_server = PowerPointMCPServer()
    logger.info("PowerPoint Analyzer MCP initialized successfully")
    
    yield
    
    # Shutdown
    logger.info("Shutting down PowerPoint Analyzer MCP...")
    powerpoint_server = None

def get_powerpoint_server() -> PowerPointMCPServer:
    """Get the PowerPoint server instance."""
    global powerpoint_server
    if powerpoint_server is None:
        raise RuntimeError("PowerPoint server not initialized")
    return powerpoint_server

# Create FastMCP instance with lifespan
mcp = FastMCP(config.server_name, lifespan=lifespan)

@mcp.tool
async def extract_powerpoint_content(file_path: str) -> str:
    """Extract complete structured content from a PowerPoint file.
    
    This is the primary content extraction tool that analyzes a PowerPoint presentation
    and returns comprehensive information about all slides, including text content,
    formatting, layout information, tables, metadata, and structural elements.
    
    Args:
        file_path: Path to the PowerPoint file (.pptx). Must be a valid PowerPoint file.
                  Supports both relative and absolute paths.
                  Example: "presentation.pptx" or "/path/to/slides.pptx"
        
    Returns:
        JSON string containing the complete structured content with the following format:
        {
            "file_path": str,                    # Path to the analyzed file
            "slides": [                          # Array of slide objects
                {
                    "slide_number": int,         # Slide number (1-based)
                    "title": str,                # Slide title (null if no title)
                    "subtitle": str,             # Slide subtitle (null if no subtitle)
                    "layout_name": str,          # Name of the slide layout
                    "layout_type": str,          # Type of layout (e.g., "Title Slide", "Content")
                    "placeholders": [            # Array of placeholder objects
                        {
                            "type": str,         # Placeholder type (e.g., "title", "content")
                            "text": str,         # Text content in placeholder
                            "position": [int, int], # [x, y] coordinates
                            "size": [int, int]   # [width, height] dimensions
                        }
                    ],
                    "text_elements": [           # Array of text element objects
                        {
                            "content_plain": str,    # Plain text content
                            "content_formatted": str, # Formatted text content
                            "bolded": int,           # Number of bold text runs
                            "italicized": int,       # Number of italic text runs
                            "underlined": int,       # Number of underlined text runs
                            "font_sizes": [int],     # Array of font sizes used
                            "font_colors": [str],    # Array of font colors (hex format)
                            "hyperlinks": [          # Array of hyperlink objects
                                {
                                    "url": str,      # Hyperlink URL
                                    "display_text": str # Display text for the link
                                }
                            ],
                            "position": [int, int],  # [x, y] coordinates
                            "size": [int, int]       # [width, height] dimensions
                        }
                    ],
                    "tables": [                  # Array of table objects
                        {
                            "rows": int,         # Number of rows
                            "columns": int,      # Number of columns
                            "cells": [           # Array of cell data
                                {
                                    "row": int,  # Row index (0-based)
                                    "col": int,  # Column index (0-based)
                                    "text": str  # Cell text content
                                }
                            ]
                        }
                    ],
                    "notes": str,                # Speaker notes content
                    "object_counts": {           # Count of different object types
                        "text_boxes": int,       # Number of text boxes
                        "tables": int,           # Number of tables
                        "images": int,           # Number of images
                        "shapes": int            # Number of shapes
                    }
                }
            ],
            "metadata": {                        # Presentation metadata
                "slide_count": int,              # Total number of slides
                "title": str,                    # Presentation title
                "author": str,                   # Presentation author
                "created_date": str,             # Creation date
                "modified_date": str             # Last modification date
            },
            "slide_size": {                      # Slide dimensions
                "width_emu": int,                # Width in EMUs (English Metric Units)
                "height_emu": int,               # Height in EMUs
                "width_inches": float,           # Width in inches
                "height_inches": float           # Height in inches
            },
            "sections": [                        # Presentation sections (if any)
                {
                    "name": str,                 # Section name
                    "slide_range": [int, int]    # [start_slide, end_slide]
                }
            ],
            "notes": [                           # Speaker notes for all slides
                {
                    "slide_number": int,         # Slide number
                    "notes_content": str         # Notes text content
                }
            ]
        }
        
        If an error occurs, returns:
        {
            "error": str                         # Error message describing what went wrong
        }
    
    Example Usage:
        extract_powerpoint_content("quarterly_report.pptx")
        # Returns complete analysis of all slides, text, formatting, and metadata
    
    Note: This tool provides the most comprehensive analysis available and serves as the
    foundation for other specialized extraction tools. For focused analysis of specific
    attributes, consider using get_powerpoint_attributes() instead.
    """
    logger.info(f"extract_powerpoint_content called with file_path: {file_path}")
    
    try:
        server = get_powerpoint_server()
        arguments = {"file_path": file_path}
        logger.debug(f"Calling server._extract_powerpoint_content with arguments: {arguments}")
        
        # Call the async method directly
        result = await server._extract_powerpoint_content(arguments)
        
        # Extract text content from CallToolResult
        content_text = ""
        if result.content:
            for content_item in result.content:
                if hasattr(content_item, 'text'):
                    content_text += content_item.text
        
        logger.info(f"extract_powerpoint_content completed successfully, content length: {len(content_text)}")
        return content_text
        
    except Exception as e:
        logger.error(f"Error in extract_powerpoint_content: {e}")
        import traceback
        logger.error(f"Traceback: {traceback.format_exc()}")
        return f"Error: {str(e)}"

@mcp.tool
async def get_powerpoint_attributes(file_path: str, attributes: List[str]) -> str:
    """Get specific attributes from PowerPoint slides with selective extraction.
    
    This tool provides efficient extraction of only the requested attributes from a PowerPoint
    presentation, reducing processing time and response size when you don't need complete
    content analysis. It's ideal for focused analysis or when working with large presentations.
    
    Args:
        file_path: Path to the PowerPoint file (.pptx). Must be a valid PowerPoint file.
                  Example: "presentation.pptx" or "/path/to/slides.pptx"
        
        attributes: List of specific attributes to extract. Valid attribute names are:
            - "title": Slide titles only
            - "subtitle": Slide subtitles only  
            - "text": All text content (same as text_elements)
            - "text_elements": Detailed text elements with formatting information
            - "tables": Table structure and content
            - "images": Image information and metadata
            - "layout": Layout names and types
            - "size": Slide dimensions and size information
            - "sections": Presentation section information
            - "notes": Speaker notes content
            - "object_counts": Count of different object types per slide
            - "placeholders": Placeholder information and content
        
    Returns:
        JSON string containing only the requested attributes with the following structure:
        {
            "file_path": str,                    # Path to the analyzed file
            "slides": [                          # Array of slide objects with only requested attributes
                {
                    "slide_number": int,         # Always included for reference
                    # Only requested attributes will be present:
                    "title": str,                # If "title" was requested
                    "subtitle": str,             # If "subtitle" was requested
                    "text_elements": [...],      # If "text" or "text_elements" was requested
                    "tables": [...],             # If "tables" was requested
                    "object_counts": {...},      # If "object_counts" was requested
                    "layout_name": str,          # If "layout" was requested
                    "layout_type": str,          # If "layout" was requested
                    "placeholders": [...],       # If "placeholders" was requested
                    "notes": str                 # If "notes" was requested
                }
            ],
            # These may be included based on requested attributes:
            "slide_size": {...},                 # If "size" was requested
            "sections": [...],                   # If "sections" was requested
            "metadata": {...}                    # Basic metadata always included
        }
        
        If an error occurs, returns:
        {
            "error": str                         # Error message describing what went wrong
        }
    
    Example Usage:
        get_powerpoint_attributes("slides.pptx", ["title", "subtitle"])
        # Returns only slide titles and subtitles
        
        get_powerpoint_attributes("slides.pptx", ["text_elements", "object_counts"])
        # Returns text content and object counts for analysis
        
        get_powerpoint_attributes("slides.pptx", ["tables", "notes"])
        # Returns table data and speaker notes only
    
    Performance Note: This tool is more efficient than extract_powerpoint_content() when you
    only need specific attributes, especially for large presentations. It processes only the
    requested content types and returns smaller response payloads.
    """
    logger.info(f"get_powerpoint_attributes called with file_path: {file_path}, attributes: {attributes}")
    
    try:
        server = get_powerpoint_server()
        arguments = {"file_path": file_path, "attributes": attributes}
        logger.debug(f"Calling server._get_powerpoint_attributes with arguments: {arguments}")
        
        # Call the async method directly
        result = await server._get_powerpoint_attributes(arguments)
        
        # Extract text content from CallToolResult
        content_text = ""
        if result.content:
            for content_item in result.content:
                if hasattr(content_item, 'text'):
                    content_text += content_item.text
        
        logger.info(f"get_powerpoint_attributes completed successfully, content length: {len(content_text)}")
        return content_text
        
    except Exception as e:
        logger.error(f"Error in get_powerpoint_attributes: {e}")
        import traceback
        logger.error(f"Traceback: {traceback.format_exc()}")
        return f"Error: {str(e)}"

@mcp.tool
async def get_slide_info(file_path: str, slide_number: int) -> str:
    """Get information for a specific slide.
    
    Args:
        file_path: Path to the PowerPoint file (.pptx)
        slide_number: Slide number (1-based)
        
    Returns:
        JSON string containing information about the specified slide
    """
    logger.info(f"get_slide_info called with file_path: {file_path}, slide_number: {slide_number}")
    
    try:
        server = get_powerpoint_server()
        arguments = {"file_path": file_path, "slide_number": slide_number}
        
        # Call the async method directly
        result = await server._get_slide_info(arguments)
        
        # Extract text content from CallToolResult
        content_text = ""
        if result.content:
            for content_item in result.content:
                if hasattr(content_item, 'text'):
                    content_text += content_item.text
        
        return content_text
        
    except Exception as e:
        logger.error(f"Error in get_slide_info: {e}")
        return f"Error: {str(e)}"

@mcp.tool
async def query_slides(file_path: str, search_criteria: Dict[str, Any], return_fields: Optional[List[str]] = None, limit: int = 50) -> str:
    """Query slides with flexible filtering criteria.
    
    Args:
        file_path: Path to the PowerPoint file (.pptx)
        search_criteria: Search criteria for filtering slides
        return_fields: Fields to include in results (optional)
        limit: Maximum number of results (default: 50)
        
    Returns:
        JSON string containing the filtered slide results
    """
    logger.info(f"query_slides called with file_path: {file_path}, search_criteria: {search_criteria}")
    
    try:
        server = get_powerpoint_server()
        arguments = {
            "file_path": file_path,
            "search_criteria": search_criteria,
            "return_fields": return_fields or ["slide_number", "title", "object_counts"],
            "limit": limit
        }
        
        # Call the async method directly
        result = await server._query_slides(arguments)
        
        # Extract text content from CallToolResult
        content_text = ""
        if result.content:
            for content_item in result.content:
                if hasattr(content_item, 'text'):
                    content_text += content_item.text
        
        return content_text
        
    except Exception as e:
        logger.error(f"Error in query_slides: {e}")
        import traceback
        logger.error(f"Traceback: {traceback.format_exc()}")
        return json.dumps({
            "error": str(e),
            "error_type": "query_slides_error",
            "file_path": file_path,
            "search_criteria": search_criteria
        }, indent=2)

@mcp.tool
async def extract_table_data(file_path: str, slide_numbers: Optional[List[int]] = None, table_criteria: Optional[Dict[str, Any]] = None, 
                      column_selection: Optional[Dict[str, Any]] = None, formatting_detection: Optional[Dict[str, Any]] = None,
                      output_format: str = "structured", include_metadata: bool = True) -> str:
    """Extract table data with flexible selection and formatting detection.
    
    Args:
        file_path: Path to the PowerPoint file (.pptx)
        slide_numbers: Slide numbers to extract tables from (1-based indexing). 
                       If not provided or None, analyzes all slides in the presentation.
                       Example: [1, 3, 5] to analyze only slides 1, 3, and 5
        
        table_criteria: Criteria for selecting tables (optional). Dictionary with keys:
            - min_rows: int - Minimum number of rows required
            - max_rows: int - Maximum number of rows allowed
            - min_columns: int - Minimum number of columns required
            - max_columns: int - Maximum number of columns allowed
            - header_contains: List[str] - Headers must contain these strings
            - header_patterns: List[str] - Headers must match these regex patterns
            Example: {"min_rows": 2, "header_contains": ["Name", "Date"]}
        
        column_selection: Configuration for column selection (optional). Dictionary with keys:
            - specific_columns: List[str] - Include only these specific column names
            - column_patterns: List[str] - Include columns matching these regex patterns
            - exclude_columns: List[str] - Exclude these column names
            - all_columns: bool - Include all columns (default: True)
            Example: {"specific_columns": ["Name", "Age"], "exclude_columns": ["ID"]}
        
        formatting_detection: Configuration for formatting detection (optional). Dictionary with keys:
            - detect_bold: bool - Detect bold text formatting (default: True)
            - detect_italic: bool - Detect italic text formatting (default: True)
            - detect_underline: bool - Detect underlined text formatting (default: True)
            - detect_highlight: bool - Detect highlighted text formatting (default: True)
            - detect_colors: bool - Detect font and background colors (default: True)
            - detect_hyperlinks: bool - Detect hyperlinks in cells (default: True)
            - preserve_formatting: bool - Preserve formatting in output (default: False)
            Example: {"detect_bold": True, "detect_colors": False}
        
        output_format: Output format for extracted data. Valid values:
            - "structured": Hierarchical structure with metadata (default)
            - "flat": Flattened array of all table data
            - "grouped_by_slide": Tables grouped by slide number
        
        include_metadata: Whether to include table metadata (position, size, formatting stats)
        
    Returns:
        JSON string containing the extracted table data with structure:
        {
            "summary": {
                "total_tables_found": int,
                "slides_processed": int,
                "slides_with_tables": int
            },
            "extracted_tables": [
                {
                    "slide_number": int,
                    "table_index": int,
                    "rows": int,
                    "columns": int,
                    "headers": List[str],
                    "data": [
                        {
                            "column_name": {
                                "value": str,
                                "formatting": {
                                    "bold": bool,
                                    "italic": bool,
                                    "underline": bool,
                                    "highlight": bool,
                                    "font_color": str,
                                    "background_color": str,
                                    "hyperlink": str
                                }
                            }
                        }
                    ],
                    "metadata": {...}  // if include_metadata is True
                }
            ]
        }
    
    Example Usage:
        # Basic table extraction from all slides
        extract_table_data("presentation.pptx")
        
        # Extract tables from specific slides
        extract_table_data("presentation.pptx", slide_numbers=[1, 2])
        
        # Extract tables with specific criteria
        extract_table_data("presentation.pptx", slide_numbers=[1, 2], 
                          table_criteria={"min_rows": 3, "header_contains": ["Name"]})
        
        # Extract specific columns with formatting from all slides
        extract_table_data("presentation.pptx", 
                          column_selection={"specific_columns": ["Name", "Age"]},
                          formatting_detection={"detect_bold": True, "detect_colors": True})
    """
    logger.info(f"extract_table_data called with file_path: {file_path}, slide_numbers: {slide_numbers}")
    
    try:
        server = get_powerpoint_server()
        arguments = {
            "file_path": file_path,
            "slide_numbers": slide_numbers,
            "table_criteria": table_criteria,
            "column_selection": column_selection,
            "formatting_detection": formatting_detection,
            "output_format": output_format,
            "include_metadata": include_metadata
        }
        
        # Call the async method directly
        result = await server._extract_table_data(arguments)
        
        # Extract text content from CallToolResult
        content_text = ""
        if result.content:
            for content_item in result.content:
                if hasattr(content_item, 'text'):
                    content_text += content_item.text
        
        return content_text
        
    except Exception as e:
        logger.error(f"Error in extract_table_data: {e}")
        import traceback
        logger.error(f"Traceback: {traceback.format_exc()}")
        return json.dumps({
            "error": str(e),
            "error_type": "extract_table_data_error",
            "file_path": file_path,
            "slide_numbers": slide_numbers
        }, indent=2)

@mcp.tool
async def clear_cache(file_path: Optional[str] = None) -> str:
    """Clear the analysis cache for a specific file or all files.
    
    Args:
        file_path: Path to the PowerPoint file to clear from cache (optional, clears all if not specified)
        
    Returns:
        JSON string containing the cache clear result
    """
    logger.info(f"clear_cache called with file_path: {file_path}")
    
    try:
        server = get_powerpoint_server()
        
        # Clear presentation analyzer cache
        if hasattr(server, 'presentation_analyzer') and hasattr(server.presentation_analyzer, 'clear_cache'):
            server.presentation_analyzer.clear_cache()
        
        # Clear content extractor cache if it exists
        if hasattr(server, 'content_extractor') and hasattr(server.content_extractor, '_cache'):
            if file_path:
                # Clear specific file from cache
                if file_path in server.content_extractor._cache:
                    del server.content_extractor._cache[file_path]
                    result = {"status": "success", "message": f"Cache cleared for file: {file_path}"}
                else:
                    result = {"status": "info", "message": f"File not found in cache: {file_path}"}
            else:
                # Clear all cache
                server.content_extractor._cache.clear()
                result = {"status": "success", "message": "All cache cleared"}
        else:
            result = {"status": "info", "message": "No cache to clear"}
        
        return json.dumps(result, indent=2)
        
    except Exception as e:
        logger.error(f"Error in clear_cache: {e}")
        return f"Error: {str(e)}"

@mcp.tool
async def reload_file_content(file_path: str, clear_cache: bool = True) -> str:
    """Reload file content by clearing cache and re-extracting.
    
    Args:
        file_path: Path to the PowerPoint file (.pptx)
        clear_cache: Whether to clear cache before reloading (default: True)
        
    Returns:
        JSON string containing the reloaded file content
    """
    logger.info(f"reload_file_content called with file_path: {file_path}, clear_cache: {clear_cache}")
    
    try:
        server = get_powerpoint_server()
        
        # Clear cache if requested
        if clear_cache:
            # Clear presentation analyzer cache
            if hasattr(server, 'presentation_analyzer') and hasattr(server.presentation_analyzer, 'clear_cache'):
                server.presentation_analyzer.clear_cache()
            
            # Clear content extractor cache if it exists
            if hasattr(server, 'content_extractor') and hasattr(server.content_extractor, '_cache'):
                if file_path in server.content_extractor._cache:
                    del server.content_extractor._cache[file_path]
        
        # Re-extract content
        arguments = {"file_path": file_path}
        result = await server._extract_powerpoint_content(arguments)
        
        # Extract text content from CallToolResult
        content_text = ""
        if result.content:
            for content_item in result.content:
                if hasattr(content_item, 'text'):
                    content_text += content_item.text
        
        logger.info(f"reload_file_content completed successfully, content length: {len(content_text)}")
        return content_text
        
    except Exception as e:
        logger.error(f"Error in reload_file_content: {e}")
        return f"Error: {str(e)}"

@mcp.tool
async def get_presentation_overview(file_path: str, analysis_depth: str = "basic", include_sample_content: bool = True) -> str:
    """Get comprehensive presentation overview and analysis.
    
    Args:
        file_path: Path to the PowerPoint file (.pptx)
        analysis_depth: Analysis depth (basic, detailed, comprehensive)
        include_sample_content: Whether to include sample content
        
    Returns:
        JSON string containing the presentation overview
    """
    logger.info(f"get_presentation_overview called with file_path: {file_path}, analysis_depth: {analysis_depth}")
    
    try:
        server = get_powerpoint_server()
        arguments = {
            "file_path": file_path,
            "analysis_depth": analysis_depth,
            "include_sample_content": include_sample_content
        }
        
        # Call the async method directly
        result = await server._get_presentation_overview(arguments)
        
        # Extract text content from CallToolResult
        content_text = ""
        if result.content:
            for content_item in result.content:
                if hasattr(content_item, 'text'):
                    content_text += content_item.text
        
        return content_text
        
    except Exception as e:
        logger.error(f"Error in get_presentation_overview: {e}")
        return f"Error: {str(e)}"

@mcp.tool
async def analyze_text_formatting(file_path: str, slide_numbers: Optional[List[int]] = None, 
                                 include_bold_analysis: bool = True, include_formatting_details: bool = True) -> str:
    """Analyze text formatting patterns across slides with detailed bold text detection.
    
    Args:
        file_path: Path to the PowerPoint file (.pptx)
        slide_numbers: Optional list of specific slide numbers to analyze (1-based indexing).
                      If not provided or None, analyzes all slides in the presentation.
                      Example: [1, 3, 5] to analyze only slides 1, 3, and 5
        include_bold_analysis: Whether to include detailed bold text analysis
        include_formatting_details: Whether to include detailed formatting information
        
    Returns:
        JSON string containing detailed text formatting analysis
    """
    logger.info(f"analyze_text_formatting called with file_path: {file_path}, slide_numbers: {slide_numbers}")
    
    try:
        server = get_powerpoint_server()
        arguments = {
            "file_path": file_path,
            "slide_numbers": slide_numbers or [],
            "formatting_filter": {
                "include_bold": include_bold_analysis,
                "include_formatting_details": include_formatting_details
            },
            "grouping": "by_formatting_type"
        }
        
        # Call the async method directly
        result = await server._analyze_text_formatting(arguments)
        
        # Extract text content from CallToolResult
        content_text = ""
        if result.content:
            for content_item in result.content:
                if hasattr(content_item, 'text'):
                    content_text += content_item.text
        
        return content_text
        
    except Exception as e:
        logger.error(f"Error in analyze_text_formatting: {e}")
        import traceback
        logger.error(f"Traceback: {traceback.format_exc()}")
        return json.dumps({
            "error": str(e),
            "error_type": "analyze_text_formatting_error",
            "file_path": file_path
        }, indent=2)

@mcp.tool
async def extract_text_formatting(file_path: str, formatting_type: str, slide_numbers: Optional[List[int]] = None) -> str:
    """Extract text with specific formatting attributes from PowerPoint slides.
    
    This tool provides a generalized interface for extracting various types of text formatting
    from PowerPoint presentations. It analyzes slides and returns both complete text content
    and specific formatted segments with position information.
    
    Args:
        file_path: Path to the PowerPoint file (.pptx). Must be a valid PowerPoint file.
                  Example: "presentation.pptx" or "/path/to/slides.pptx"
        
        formatting_type: Type of formatting to extract. Valid values are:
            - "bold": Extract bold text segments and their positions
            - "italic": Extract italic text segments and their positions
            - "underlined": Extract underlined text segments and their positions
            - "highlighted": Extract highlighted text segments and their positions
            - "strikethrough": Extract strikethrough text segments and their positions
            - "hyperlinks": Extract hyperlink text, URLs, and link types (external/internal/email)
            - "font_sizes": Extract text segments with their font size information
            - "font_colors": Extract text segments with their color information (hex format)
        
        slide_numbers: Optional list of specific slide numbers to analyze (1-based indexing).
                      If not provided, analyzes all slides in the presentation.
                      Example: [1, 3, 5] to analyze only slides 1, 3, and 5
        
    Returns:
        JSON string with the following structure:
        {
            "file_path": str,                    # Path to the analyzed file
            "formatting_type": str,              # Type of formatting that was extracted
            "summary": {
                "total_slides_analyzed": int,    # Number of slides that were analyzed
                "slides_with_formatting": int,   # Number of slides containing the requested formatting
                "total_formatted_segments": int # Total number of formatted text segments found
            },
            "results_by_slide": [
                {
                    "slide_number": int,         # Slide number (1-based)
                    "title": str,                # Slide title (empty string if no title)
                    "complete_text": str,        # Complete text content from all text elements
                    "format": str,               # Formatting type (same as input parameter)
                    "formatted_segments": [      # Array of formatted text segments
                        {
                            "text": str,         # The formatted text content
                            "start_position": int # Character position where formatted text starts
                        }
                    ]
                }
            ]
        }
        
        If an error occurs, returns:
        {
            "error": str                         # Error message describing what went wrong
        }
    
    Example Usage:
        extract_text_formatting("slides.pptx", "bold") 
        # Returns all bold text from all slides
        
        extract_text_formatting("slides.pptx", "hyperlinks", [1, 2])
        # Returns hyperlinks from slides 1 and 2 only
    """
    logger.info(f"extract_text_formatting called with file_path: {file_path}, formatting_type: {formatting_type}, slide_numbers: {slide_numbers}")
    
    try:
        server = get_powerpoint_server()
        arguments = {
            "file_path": file_path,
            "formatting_type": formatting_type,
            "slide_numbers": slide_numbers
        }
        
        # Call the server method directly
        result = await server._extract_text_formatting(arguments)
        
        # Extract text content from CallToolResult
        content_text = ""
        if result.content:
            for content_item in result.content:
                if hasattr(content_item, 'text'):
                    content_text += content_item.text
        
        return content_text
        
    except Exception as e:
        logger.error(f"Error in extract_text_formatting: {e}")
        import traceback
        logger.error(f"Traceback: {traceback.format_exc()}")
        return json.dumps({
            "error": str(e),
            "error_type": "extract_text_formatting_error",
            "file_path": file_path,
            "formatting_type": formatting_type
        }, indent=2)

@mcp.tool
async def extract_bold_text(file_path: str, slide_numbers: Optional[List[int]] = None) -> str:
    """Extract all bold text from slides with enhanced location and formatting information.
    
    This tool provides backward-compatible bold text extraction with enhanced response format
    that includes both legacy fields and new position-aware segments. It analyzes PowerPoint
    slides to identify bold text formatting and returns detailed information about bold content.
    
    Args:
        file_path: Path to the PowerPoint file (.pptx). Must be a valid PowerPoint file.
                  Example: "presentation.pptx" or "/path/to/slides.pptx"
        
        slide_numbers: Optional list of specific slide numbers to analyze (1-based indexing).
                      If not provided, analyzes all slides in the presentation.
                      Example: [1, 3, 5] to analyze only slides 1, 3, and 5
        
    Returns:
        JSON string with the following structure:
        {
            "file_path": str,                    # Path to the analyzed file
            "bold_text_summary": {
                "total_slides_analyzed": int,    # Number of slides that were analyzed
                "slides_with_bold_text": int,    # Number of slides containing bold text
                "total_bold_elements": int       # Total number of bold text elements found
            },
            "bold_text_by_slide": [
                {
                    "slide_number": int,         # Slide number (1-based)
                    "title": str,                # Slide title (empty string if no title)
                    "complete_text": str,        # NEW: Complete text content from all text elements
                    "bold_elements": [           # LEGACY: Original format for backward compatibility
                        {
                            "content": str,      # Full text content of element containing bold text
                            "bold_count": int,   # Number of bold text runs in this element
                            "font_sizes": list,  # Font sizes found in this element
                            "font_colors": list, # Font colors found in this element
                            "position": list,    # Element position [x, y] coordinates
                            "size": list,        # Element size [width, height]
                            "content_type": str, # Type of content ("text_elements")
                            "element_index": int # Index of this element within the slide
                        }
                    ],
                    "bold_segments": [           # NEW: Enhanced format with position information
                        {
                            "text": str,         # The bold text content
                            "start_position": int, # Character position where bold text starts
                            "end_position": int  # Character position where bold text ends
                        }
                    ],
                    "bold_count": int            # Total count of bold elements on this slide
                }
            ]
        }
        
        If an error occurs, returns:
        {
            "error": str                         # Error message describing what went wrong
        }
    
    Example Usage:
        extract_bold_text("slides.pptx")
        # Returns all bold text from all slides with both legacy and enhanced formats
        
        extract_bold_text("slides.pptx", [1, 2])
        # Returns bold text from slides 1 and 2 only
    
    Note: This tool maintains backward compatibility by including both the original 'bold_elements'
    format and the new 'bold_segments' format with position information. New applications should
    use the 'bold_segments' array for more precise text positioning.
    """
    logger.info(f"extract_bold_text called with file_path: {file_path}, slide_numbers: {slide_numbers}")
    
    try:
        server = get_powerpoint_server()
        
        # First get the full content
        arguments = {"file_path": file_path}
        result = await server._extract_powerpoint_content(arguments)
        
        # Extract text content from CallToolResult
        content_text = ""
        if result.content:
            for content_item in result.content:
                if hasattr(content_item, 'text'):
                    content_text += content_item.text
        
        # Parse the JSON content to extract bold text information
        content_data = json.loads(content_text)
        
        # Import the new formatting extractor
        from powerpoint_mcp_server.core.formatting_extractor import FormattingExtractor
        formatting_extractor = FormattingExtractor(server.content_extractor)
        
        bold_text_analysis = {
            "file_path": file_path,
            "bold_text_summary": {
                "total_slides_analyzed": 0,
                "slides_with_bold_text": 0,
                "total_bold_elements": 0
            },
            "bold_text_by_slide": []
        }
        
        for slide_data in content_data.get('slides', []):
            slide_num = slide_data.get('slide_number', 0)
            
            if slide_numbers and slide_num not in slide_numbers:
                continue
                
            bold_text_analysis["bold_text_summary"]["total_slides_analyzed"] += 1
            
            # Build complete text from all text elements
            text_elements = slide_data.get('text_elements', [])
            complete_text = ' '.join([elem.get('content_plain', '') for elem in text_elements])
            
            # Extract bold segments using the new extractor
            bold_segments = formatting_extractor.extract_formatting_segments(
                text_elements, "bold", slide_num
            )
            
            slide_bold_info = {
                "slide_number": slide_num,
                "title": slide_data.get('title', ''),
                "complete_text": complete_text,  # New field: complete text content
                "bold_elements": [],  # Legacy format for backward compatibility
                "bold_segments": [],  # New field: array of bold text segments
                "bold_count": 0
            }
            
            # Legacy format: Analyze text elements for bold formatting
            for idx, text_elem in enumerate(slide_data.get('text_elements', [])):
                if text_elem.get('bolded', 0) > 0:
                    bold_element = {
                        "content": text_elem.get('content_plain', ''),
                        "bold_count": text_elem.get('bolded', 0),
                        "font_sizes": text_elem.get('font_sizes', []),
                        "font_colors": text_elem.get('font_colors', []),
                        "position": text_elem.get('position', []),
                        "size": text_elem.get('size', []),
                        "content_type": "text_elements",
                        "element_index": idx
                    }
                    slide_bold_info["bold_elements"].append(bold_element)
                    slide_bold_info["bold_count"] += text_elem.get('bolded', 0)
            
            # New format: Add bold segments with position information
            for segment in bold_segments:
                slide_bold_info["bold_segments"].append({
                    "text": segment.text,
                    "start_position": segment.start_position,
                    "end_position": segment.end_position
                })
            
            if slide_bold_info["bold_count"] > 0 or bold_segments:
                bold_text_analysis["bold_text_summary"]["slides_with_bold_text"] += 1
                bold_text_analysis["bold_text_summary"]["total_bold_elements"] += max(slide_bold_info["bold_count"], len(bold_segments))
            
            # Include all slides if no filter, or only slides with bold text if filtering
            if not slide_numbers or slide_bold_info["bold_elements"] or bold_segments:
                bold_text_analysis["bold_text_by_slide"].append(slide_bold_info)
        
        return json.dumps(bold_text_analysis, indent=2, ensure_ascii=False)
        
    except Exception as e:
        logger.error(f"Error in extract_bold_text: {e}")
        return f"Error: {str(e)}"

@mcp.tool
async def tool_help(tool_name: str) -> str:
    """Get detailed help and documentation for MCP tools.
    
    Args:
        tool_name: Name of the tool to get help for
        
    Returns:
        Formatted help text with detailed documentation including:
        - Tool description and purpose
        - Parameter specifications with types and requirements  
        - Detailed schema for complex parameters
        - Usage examples with real scenarios
        - Important notes and best practices
    """
    logger.info(f"tool_help called with tool_name: {tool_name}")
    
    try:
        from powerpoint_mcp_server.tools.tool_help import get_tool_help
        
        # Get help text for the specified tool
        help_text = get_tool_help(tool_name)
        
        if not help_text or "No help available" in help_text:
            return f"No help available for tool: {tool_name}"
        
        logger.info(f"tool_help completed successfully for tool: {tool_name}")
        return help_text
        
    except Exception as e:
        logger.error(f"Error in tool_help: {e}")
        return f"Error getting help for tool '{tool_name}': {str(e)}"

def main():
    """Main entry point for the FastMCP PowerPoint server."""
    logger.info(f"Starting PowerPoint Analyzer MCP using FastMCP 2.0: {config.server_name} v{config.server_version}")
    logger.info(f"Log file: {log_file}")
    
    # Enable debug logging for FastMCP
    fastmcp_logger = logging.getLogger('fastmcp')
    fastmcp_logger.setLevel(logging.DEBUG)
    
    logger.info("FastMCP 2.0 server configured with tools")
    
    try:
        # Run the FastMCP server
        logger.info("Starting FastMCP 2.0 server...")
        mcp.run()
    except Exception as e:
        logger.error(f"Server error: {e}")
        import traceback
        logger.error(f"Traceback: {traceback.format_exc()}")
        raise

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        logger.info("Server stopped by user")
        sys.exit(0)
    except Exception as e:
        logger.error(f"Fatal error: {e}")
        sys.exit(1)