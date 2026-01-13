#!/usr/bin/env python3
"""Main entry point for the PowerPoint Analyzer MCP using FastMCP 2.0."""

import asyncio
import json
import logging
import sys
import os
from typing import Any, Dict, List, Optional, Union, Annotated
from pathlib import Path
from contextlib import asynccontextmanager

# Import FastMCP
from fastmcp import FastMCP
from powerpoint_mcp_server.server import PowerPointMCPServer
from powerpoint_mcp_server.config import get_config, get_config_manager

# Configure logging
config = get_config()

# Create log file handler
log_file = "powerpoint_mcp_server.log"
file_handler = logging.FileHandler(log_file, mode='w', encoding='utf-8')
file_handler.setLevel(logging.DEBUG)

# For MCP servers, we should minimize stderr output to avoid [ERROR] logs in clients
# Only log ERROR and CRITICAL to stderr, everything else goes to file only
console_handler = logging.StreamHandler(sys.stderr)
console_handler.setLevel(logging.ERROR)  # Only ERROR and CRITICAL to stderr

# Create formatter
formatter = logging.Formatter(config.log_format)
file_handler.setFormatter(formatter)
console_handler.setFormatter(formatter)

# Configure root logger
logging.basicConfig(
    level=logging.DEBUG,
    handlers=[file_handler, console_handler]
)

# Set all third-party loggers to ERROR level to minimize stderr output for MCP
third_party_loggers = ['fastmcp', 'mcp', 'asyncio', 'urllib3', 'requests']
for logger_name in third_party_loggers:
    logging.getLogger(logger_name).setLevel(logging.ERROR)

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
@mcp.tool(description="Query slides with flexible filtering criteria. Provides powerful slide filtering and search capabilities for PowerPoint presentations with complex search criteria including title matching, content analysis, layout filtering, and slide number selection.")
async def query_slides(
    file_path: Annotated[str, "Path to the PowerPoint file (.pptx). Must be a valid PowerPoint file. Example: 'presentation.pptx' or '/path/to/slides.pptx'"],
    search_criteria: Annotated[Dict[str, Any], "Dictionary containing search and filter criteria. Supports title filtering (contains, starts_with, ends_with, regex, one_of), content filtering (contains_text, has_tables, has_charts, has_images, object_count), layout filtering (type, name), notes filtering (contains, regex, is_empty), slide_numbers (int, List[int], or str slicing), and section filtering"],
    return_fields: Annotated[Optional[List[str]], "List of fields to include in results. Valid values: 'slide_number', 'title', 'subtitle', 'layout_name', 'layout_type', 'object_counts', 'preview_text', 'table_info', 'full_content'. Default: ['slide_number', 'title', 'object_counts']"] = None,
    limit: Annotated[int, "Maximum number of results to return (1-1000, default: 50)"] = 50
) -> str:
    """Query slides with flexible filtering criteria.

    This tool provides powerful slide filtering and search capabilities for PowerPoint presentations.
    It supports complex search criteria including title matching, content analysis, layout filtering,
    and slide number selection with flexible result formatting.

    Args:
        file_path: Path to the PowerPoint file (.pptx). Must be a valid PowerPoint file.
                  Example: "C:¥¥temp¥¥presentation.pptx" or "/path/to/slides.pptx"

        search_criteria: Dictionary containing search and filter criteria. Structure:
            {
                "title": {                    # Title-based filtering
                    "contains": "str",        # Title contains this text
                    "starts_with": "str",     # Title starts with this text
                    "ends_with": "str",       # Title ends with this text
                    "regex": "str",           # Title matches this regex pattern
                    "one_of": ["str1", "str2"] # Title is one of these values
                },
                "content": {                  # Content-based filtering
                    "contains_text": "str",   # Slide text contains this string
                    "has_tables": bool,       # Slide has tables (true/false)
                    "has_charts": bool,       # Slide has charts (true/false)
                    "has_images": bool,       # Slide has images (true/false)
                    "object_count": {         # Object count constraints
                        "min": int,           # Minimum number of objects
                        "max": int            # Maximum number of objects
                    }
                },
                "layout": {                   # Layout-based filtering
                    "type": "str",            # Layout type (e.g., "title", "title_content")
                    "name": "str"             # Layout name (e.g., "Title Slide")
                },
                "notes": {                    # Speaker notes filtering
                    "contains": "str",        # Notes contain this text
                    "regex": "str",           # Notes match this regex
                    "is_empty": bool          # Notes are empty (true/false)
                },
                "slide_numbers": "Specific slide numbers to include (1-based). Supports multiple formats: int (single slide), List[int] (specific slides), or str (Python-style slicing like ':100', '5:20', '25:', '1,5,10')",
                "section": "str"              # Section name to filter by
            }

        return_fields: List of fields to include in results. Valid field names:
            - "slide_number": Slide number (always included)
            - "title": Slide title
            - "subtitle": Slide subtitle
            - "layout_name": Layout name
            - "layout_type": Layout type
            - "object_counts": Object count statistics
            - "preview_text": Preview of slide text content
            - "table_info": Table structure information
            - "full_content": Complete slide content
            Default: ["slide_number", "title", "object_counts"]

        limit: Maximum number of results to return (1-1000, default: 50)

    Returns:
        JSON string with the following structure:
        {
            "file_path": "str",
            "search_criteria": {...},
            "summary": {
                "total_slides_in_presentation": int,
                "slides_matching_criteria": int,
                "results_returned": int,
                "search_time_ms": float
            },
            "results": [
                {
                    "slide_number": int,
                    "title": "str",
                    "subtitle": "str",
                    "layout_name": "str",
                    "layout_type": "str",
                    "object_counts": {
                        "shapes": int,
                        "text_boxes": int,
                        "images": int,
                        "tables": int
                    },
                    "preview_text": "str",
                    "table_info": [
                        {
                            "rows": int,
                            "columns": int,
                            "preview": "str"
                        }
                    ],
                    "full_content": {...}
                }
            ]
        }

        | key | type | description |
        |------|------|-------------|
        | file_path | str | Path to the analyzed file |
        | search_criteria | object | The search criteria that were applied |
        | summary.total_slides_in_presentation | int | Total number of slides in the presentation |
        | summary.slides_matching_criteria | int | Number of slides that matched the search criteria |
        | summary.results_returned | int | Number of results returned (limited by 'limit' parameter) |
        | summary.search_time_ms | float | Time taken to perform the search in milliseconds |
        | results[].slide_number | int | Slide number (1-based) |
        | results[].title | str | Slide title (included if requested in return_fields) |
        | results[].subtitle | str | Slide subtitle (included if requested in return_fields) |
        | results[].layout_name | str | Name of the slide layout (included if requested in return_fields) |
        | results[].layout_type | str | Type of layout (included if requested in return_fields) |
        | results[].object_counts | object | Count of different object types (included if requested in return_fields) |
        | results[].preview_text | str | Preview of slide text content (included if requested in return_fields) |
        | results[].table_info | array | Information about tables in the slide (included if requested in return_fields) |
        | results[].full_content | object | Complete slide content (included if requested in return_fields) |

        If an error occurs, returns:
        {
            "error": str,
            "error_type": "query_slides_error",
            "file_path": str
        }

    Example Usage:
        # Find slides with "Sales" in the title
        query_slides("C:¥¥temp¥¥presentation.pptx", {"title": {"contains": "Sales"}})

        # Find slides with tables and specific layout
        query_slides("C:¥¥temp¥¥presentation.pptx", {
            "content": {"has_tables": true},
            "layout": {"type": "title_content"}
        })

        # Find specific slides with custom return fields
        query_slides("C:¥¥temp¥¥presentation.pptx", 
                    {"slide_numbers": [1, 3, 5]},
                    ["slide_number", "title", "preview_text"])

        # Complex search with multiple criteria
        query_slides("C:¥¥temp¥¥presentation.pptx", {
            "title": {"regex": "Q[1-4].*Results"},
            "content": {"has_tables": true, "has_images": false},
            "notes": {"is_empty": false}
        }, limit=10)
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

@mcp.tool(description="Extract table data with flexible selection and formatting detection. Supports various slide selection methods, table filtering criteria, column selection, and comprehensive formatting detection.")
async def extract_table_data(
    file_path: Annotated[str, "Path to the PowerPoint file (.pptx)"],
    slide_numbers: Annotated[Optional[Union[int, str, List[int]]], "Slide numbers to extract tables from (1-based indexing). Supports: None (all slides), int (single slide), List[int] (specific slides), or str (Python-style slicing like ':100', '5:20', '25:', '1,5,10')"] = None,
    table_criteria: Annotated[Optional[Dict[str, Any]], "Criteria for selecting tables. Keys: min_rows, max_rows, min_columns, max_columns, header_contains (List[str]), header_patterns (List[str])"] = None,
    column_selection: Annotated[Optional[Dict[str, Any]], "Configuration for column selection. Keys: specific_columns (List[str]), column_patterns (List[str]), exclude_columns (List[str]), all_columns (bool)"] = None,
    formatting_detection: Annotated[Optional[Dict[str, Any]], "Configuration for formatting detection. Keys: detect_bold, detect_italic, detect_underline, detect_highlight, detect_colors, detect_hyperlinks, preserve_formatting (all bool)"] = None,
    output_format: Annotated[str, "Output format for extracted data. Valid values: 'structured' (hierarchical with metadata), 'flat' (flattened array), 'grouped_by_slide' (tables grouped by slide)"] = "structured",
    include_metadata: Annotated[bool, "Whether to include table metadata (row_span, col_span, row_col_position, position, size, formatting stats)"] = True
) -> str:
    """Extract table data with flexible selection and formatting detection.

    Args:
        file_path: Path to the PowerPoint file (.pptx)
        slide_numbers: Optional. Slide numbers to extract tables from (1-based indexing).
                       Supports multiple formats: None(=All),int,List[int],Python-style slicing
                       - None: All slides
                       - int: Single slide (e.g., 3)
                       - List[int]: Specific slides (e.g., [1, 5, 10])
                       - str: Python-style slicing:
                         - ":100" or "[:100]": First 100 slides (1-100)
                         - "5:20" or "[5:20]": Slides 5-20
                         - "25:" or "[25:]": Slides 25 to end
                         - "3" or "[3]": Single slide 3
                         - "1,5,10" or "[1,5,10]": Specific slides 1, 5, 10

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

        include_metadata: Whether to include table metadata (row_span,col_span,row_col_position, position, size, formatting stats)

    Returns:
        JSON string containing the extracted table data with structure:
        {
            "summary": {
                "total_tables_found": int,
                "total_tables": int,
                "total_rows": int,
                "slides_with_tables": int,
                "formatting_found": {
                    "bold_cells": int,
                    "italic_cells": int,
                    "highlighted_cells": int,
                    "colored_cells": int
                },
                "slides_processed": int
            },
            "extracted_tables": [
                {
                    "slide_number": int,
                    "table_index": int,
                    "rows": int,
                    "columns": int,
                    "headers": List[str,str,..],
                    "metadata": {
                        "has_formatting": bool,
                        "cell_count": int,
                        "non_empty_cells": int
                    },
                    "position": [int,int],
                    "size": [int,int],
                    "data": [
                        {
                            "header_name": {
                                "value": str,
                                "formatting": {
                                    "bold": bool,
                                    "italic": bool,
                                    "underline": bool,
                                    "highlight": bool,
                                    "strikethrough": bool,
                                    "font_color": str | null,
                                    "background_color": str | null,
                                    "font_size": float,
                                    "hyperlink": str |null
                                },
                                "row_span": int,
                                "col_span": int,
                                "row_col_position": [int,int]
                            }
                        }
                    ]
                }
            ]
        }

    Example Usage:
        # Basic table extraction from all slides
        extract_table_data("C:¥¥temp¥¥presentation.pptx")

        # Extract specific columns from all slides
        extract_table_data("C:¥¥temp¥¥presentation.pptx",
                          column_selection={"specific_columns": ["Name", "Age"]})

        # Extract tables from first 10 slides
        extract_table_data("C:¥¥temp¥¥presentation.pptx", slide_numbers=":10")

        # Extract tables with specific criteria
        extract_table_data("C:¥¥temp¥¥presentation.pptx", slide_numbers=[1, 2],
                          table_criteria={"min_rows": 2, "header_contains": ["Name"]})

        # Extract specific columns with formatting from all slides
        extract_table_data("C:¥¥temp¥¥presentation.pptx",
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

@mcp.tool(description="Extract text with specific formatting attributes from PowerPoint slides. Provides a generalized interface for extracting various types of text formatting with position information.")
async def extract_formatted_text(
    file_path: Annotated[str, "Path to the PowerPoint file (.pptx). Must be a valid PowerPoint file. Example: 'presentation.pptx' or '/path/to/slides.pptx'"],
    formatting_type: Annotated[str, "Type of formatting to extract. Valid values: 'bold', 'italic', 'underlined', 'highlighted', 'strikethrough', 'hyperlinks', 'font_sizes', 'font_colors'"],
    slide_numbers: Annotated[Optional[Union[int, str, List[int]]], "Slide numbers to analyze (1-based indexing). Supports: None (all slides), int (single slide), List[int] (specific slides), or str (Python-style slicing like ':100', '5:20', '25:', '1,5,10')"] = None
) -> str:
    """Extract text with specific formatting attributes from PowerPoint slides.

    This tool provides a generalized interface for extracting various types of text formatting
    from PowerPoint presentations. It analyzes slides and returns both complete text content
    and specific formatted segments with position information.

    Args:
        file_path: Path to the PowerPoint file (.pptx). Must be a valid PowerPoint file.
                  Example: "C:¥¥temp¥¥presentation.pptx" or "/path/to/slides.pptx"

        formatting_type: Type of formatting to extract. Valid values are:
            - "bold": Extract bold text segments and their positions
            - "italic": Extract italic text segments and their positions
            - "underlined": Extract underlined text segments and their positions
            - "highlighted": Extract highlighted text segments and their positions
            - "strikethrough": Extract strikethrough text segments and their positions
            - "hyperlinks": Extract hyperlink text, URLs, and link types (external/internal/email)
            - "font_sizes": Extract text segments with their font size information
            - "font_colors": Extract text segments with their color information (hex format)

        slide_numbers: Optional slide numbers to analyze (1-based indexing).
                      Supports multiple formats:
                      - None: All slides
                      - int: Single slide (e.g., 3)
                      - List[int]: Specific slides (e.g., [1, 5, 10])
                      - str: Python-style slicing:
                        - ":100" or "[:100]": First 100 slides (1-100)
                        - "5:20" or "[5:20]": Slides 5-20
                        - "25:" or "[25:]": Slides 25 to end
                        - "3" or "[3]": Single slide 3
                        - "1,5,10" or "[1,5,10]": Specific slides 1, 5, 10

    Returns:
        JSON string with the following structure:
        {
            "file_path": "str",
            "formatting_type": "str",
            "summary": {
                "total_slides_analyzed": int,
                "slides_with_formatting": int,
                "total_formatted_segments": int
            },
            "results_by_slide": [
                {
                "slide_number": int,
                "title": "str",
                "complete_text": "str",
                "format": "str",
                "formatted_segments": [
                    {
                    "text": "str",
                    "start_position": int
                    }
                ]
                }
            ]
        }

        | key | type | description |
        |------|------|-------------|
        | file_path | str | Path to the analyzed file |
        | formatting_type | str | Type of formatting that was extracted (e.g., bold, italic) |
        | summary.total_slides_analyzed | int | Number of slides that were analyzed |
        | summary.slides_with_formatting | int | Number of slides containing the requested formatting |
        | summary.total_formatted_segments | int | Total number of formatted text segments found |
        | results_by_slide[].slide_number | int | Slide number (1-based) |
        | results_by_slide[].title | str | Slide title (empty string if no title) |
        | results_by_slide[].complete_text | str | Complete text content from all text elements |
        | results_by_slide[].format | str | Formatting type (same as input parameter) |
        | results_by_slide[].formatted_segments[].text | str | The formatted text content |
        | results_by_slide[].formatted_segments[].start_position | int | Character position where formatted text starts |

        If an error occurs, returns:
        {
            "error": str
        }

    Example Usage:
        extract_formatted_text("slides.pptx", "bold")
        # Returns all bold text from all slides

        extract_formatted_text("slides.pptx", "hyperlinks", [1, 2])
        # Returns hyperlinks from slides 1 and 2 only
    """
    logger.info(f"extract_formatted_text called with file_path: {file_path}, formatting_type: {formatting_type}, slide_numbers: {slide_numbers}")

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
        logger.error(f"Error in extract_formatted_text: {e}")
        import traceback
        logger.error(f"Traceback: {traceback.format_exc()}")
        return json.dumps({
            "error": str(e),
            "error_type": "extract_formatted_text_error",
            "file_path": file_path,
            "formatting_type": formatting_type
        }, indent=2)

def main():
    """Main entry point for the FastMCP PowerPoint server."""
    logger.info(f"Starting PowerPoint Analyzer MCP using FastMCP 2.0: {config.server_name} v{config.server_version}")
    logger.info(f"Log file: {log_file}")

    # Set FastMCP logging to ERROR level to reduce stderr output for MCP clients
    fastmcp_logger = logging.getLogger('fastmcp')
    fastmcp_logger.setLevel(logging.ERROR)
    
    # Also set MCP SDK logger to ERROR level
    mcp_logger = logging.getLogger('mcp')
    mcp_logger.setLevel(logging.ERROR)

    logger.info("FastMCP 2.0 server configured with tools")

    try:
        # Run the FastMCP server (banner suppressed by fastmcp.configure(quiet=True))
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