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
@mcp.tool(description="Query slides with flexible filtering criteria. Provides powerful slide filtering and search capabilities for PowerPoint presentations.")
async def query_slides(
    file_path: Annotated[str, "Path to the PowerPoint file (.pptx). Must be a valid PowerPoint file. Example: 'C:\\\\temp\\\\presentation.pptx' or '/path/to/slides.pptx'"],
    search_criteria: Annotated[Dict[str, Any], "Dictionary containing search and filter criteria. Supports title filtering (contains, starts_with, ends_with, regex, one_of), content filtering (contains_text, has_tables, has_charts, has_images), notes filtering (contains, regex, is_empty), and sections (List[str]) filtering"],
    return_fields: Annotated[Optional[List[str]], "List of fields to include in results. Valid values: 'slide_number', 'title', 'subtitle', 'text', 'extracted_tables'. Default: ['slide_number', 'title', 'text']"] = None,
    slide_numbers: Annotated[Optional[Union[int, str, List[int]]], "Slide numbers to query (1-based indexing). Supports: None (all slides), int (single slide), List[int] (specific slides), or str (Python-style slicing like ':100', '5:20', '25:', '1,5,10')"] = None,
    output_type: Annotated[str, "Text output type: 'preview_text_3boxes' (default, shows title + content placeholder + 3 text boxes) or 'full_text' (shows all text elements)"] = "preview_text_3boxes",
    output_format: Annotated[str, "Output format: 'simple' (default, no formatting in text/tables) or 'formatted' (includes formatting)"] = "simple",
    limit: Annotated[int, "Maximum number of slides to return (1-10000, default: 1000)"] = 1000
) -> str:
    """Query slides with flexible filtering criteria.

    This tool provides powerful slide filtering and search capabilities for PowerPoint presentations
    with simplified output optimized for minimal context consumption.

    Args:
        file_path: Path to the PowerPoint file (.pptx). Must be a valid PowerPoint file.
                  Example: "C:\\temp\\presentation.pptx" or "/path/to/slides.pptx"

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
                    "has_images": bool        # Slide has images (true/false)
                },
                "notes": {                    # Speaker notes filtering
                    "contains": "str",        # Notes contain this text
                    "regex": "str",           # Notes match this regex
                    "is_empty": bool          # Notes are empty (true/false)
                },
                "sections": ["str1", "str2"]  # Section names to filter by (List[str])
            }

        return_fields: List of fields to include in results. Valid field names:
            - "slide_number": Slide number (always included)
            - "title": Slide title
            - "subtitle": Slide subtitle
            - "preview_text_3boxes": Preview with title + content placeholder + 3 text boxes
            - "full_text": All text elements without limit
            - "extracted_tables": Table data in simplified format
            Default: ["slide_number", "title", "preview_text_3boxes"]

        slide_numbers: Optional slide numbers to query (1-based indexing).
                      Supports: None (all slides), int, List[int], or str (Python-style slicing)

        output_type: Text output type selection:
            - "preview_text_3boxes": Shows title + content placeholder + up to 3 text boxes (default)
            - "full_text": Shows all text elements without limit

        output_format: Output format selection:
            - "simple": No formatting in text/tables (default)
            - "formatted": Includes formatting information

        limit: Maximum number of slides to return (1-10000, default: 1000)

    Returns:
        JSON string with the following structure:
        {
            "summary": {
                "total_slides_in_presentation": int,
                "slides_matching_criteria": int,
                "results_returned": int,
                "tables_in_slides": {
                    "slide_number": [int, int, ...],
                    "table_count": [int, int, ...]
                }
            },
            "results": [
                {
                    "slide_number": int,
                    "title": "str",
                    "subtitle": "str",
                    "text": "str",  # Content follows output_type parameter
                    "extracted_tables": [
                        {
                            "rows": int,
                            "columns": int,
                            "headers": ["col1", "col2", ...],
                            "data": [[row, col, "value"], ...]
                        }
                    ]
                }
            ]
        }

    Example Usage:
        # Find slides with "Sales" in the title
        query_slides("presentation.pptx", {"title": {"contains": "Sales"}})

        # Find slides with tables
        query_slides("presentation.pptx", {"content": {"has_tables": true}})

        # Find specific slides with custom return fields
        query_slides("presentation.pptx", {}, 
                    return_fields=["slide_number", "title", "text"],
                    slide_numbers=[1, 3, 5])
        
        # Get all text with full_text output type
        query_slides("presentation.pptx", {}, output_type="full_text")
    """
    logger.info(f"query_slides called with file_path: {file_path}, search_criteria: {search_criteria}, output_type: {output_type}")

    try:
        server = get_powerpoint_server()
        
        # Set default return_fields based on output_type parameter
        if return_fields is None:
            return_fields = ["slide_number", "title", "text"]
        
        # Validate output_type parameter
        if output_type not in ["preview_text_3boxes", "full_text"]:
            return json.dumps({
                "error": f"Invalid output_type parameter: {output_type}. Must be 'preview_text_3boxes' or 'full_text'."
            }, ensure_ascii=False)
        
        arguments = {
            "file_path": file_path,
            "search_criteria": search_criteria,
            "return_fields": return_fields,
            "slide_numbers": slide_numbers,
            "output_format": output_format,
            "output_type": output_type,
            "limit": limit
        }

        # Call the async method directly
        result = await server._query_slides_simple(arguments)

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
async def extract_formatted_table_data(
    file_path: Annotated[str, "Path to the PowerPoint file (.pptx)"],
    slide_numbers: Annotated[Optional[Union[int, str, List[int]]], "Slide numbers to extract tables from (1-based indexing). Supports: None (all slides), int (single slide), List[int] (specific slides), or str (Python-style slicing like ':100', '5:20', '25:', '1,5,10')"] = None,
    table_criteria: Annotated[Optional[Dict[str, Any]], "Criteria for selecting tables. Keys: min_rows, max_rows, min_columns, max_columns, header_contains (List[str]), header_patterns (List[str])"] = None,
    column_selection: Annotated[Optional[Dict[str, Any]], "Configuration for column selection. Keys: specific_columns (List[str]), column_patterns (List[str]), exclude_columns (List[str]), all_columns (bool)"] = None,
    formatting_detection: Annotated[Optional[Dict[str, Any]], "Configuration for formatting detection. Keys: detect_bold, detect_italic, detect_underline, detect_highlight, detect_colors, detect_hyperlinks, preserve_formatting (all bool)"] = None,
    output_format: Annotated[str, "Output format for extracted data. Valid values: 'structured' (hierarchical with metadata), 'flat' (flattened array), 'grouped_by_slide' (tables grouped by slide)"] = "structured",
    include_metadata: Annotated[bool, "Whether to include table metadata (row_span, col_span, row_col_position, position, size, formatting stats)"] = True
) -> str:
    """Extract table data with comprehensive formatting detection (legacy tool with full formatting support).

    This tool extracts tables with complete formatting information including bold, italic, colors, etc.
    For simpler output without formatting, use extract_table_data instead.

    Args:
        file_path: Path to the PowerPoint file (.pptx)
        slide_numbers: Optional. Slide numbers to extract tables from (1-based indexing).
                       Supports multiple formats: None(=All),int,List[int],Python-style slicing
        table_criteria: Criteria for selecting tables (optional)
        column_selection: Configuration for column selection (optional)
        formatting_detection: Configuration for formatting detection (optional)
        output_format: Output format for extracted data
        include_metadata: Whether to include table metadata

    Returns:
        JSON string containing the extracted table data with full formatting information

    Example Usage:
        # Extract with formatting detection
        extract_formatted_table_data("C:¥¥temp¥¥presentation.pptx",
                                    formatting_detection={"detect_bold": True, "detect_colors": True})
    """
    logger.info(f"extract_formatted_table_data called with file_path: {file_path}, slide_numbers: {slide_numbers}")

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
        logger.error(f"Error in extract_formatted_table_data: {e}")
        import traceback
        logger.error(f"Traceback: {traceback.format_exc()}")
        return json.dumps({
            "error": str(e),
            "error_type": "extract_formatted_table_data_error",
            "file_path": file_path,
            "slide_numbers": slide_numbers
        }, indent=2)

@mcp.tool(description="Extract table data in simplified format without formatting information. Optimized for minimal context consumption with clean output formats.")
async def extract_table_data(
    file_path: Annotated[str, "Path to the PowerPoint file (.pptx)"],
    slide_numbers: Annotated[Optional[Union[int, str, List[int]]], "Slide numbers to extract tables from (1-based indexing). Supports: None (all slides), int (single slide), List[int] (specific slides), or str (Python-style slicing like ':100', '5:20', '25:', '1,5,10')"] = None,
    column_selection: Annotated[Optional[Dict[str, Any]], "Configuration for column selection. Keys: specific_columns (List[str]), column_patterns (List[str]), exclude_columns (List[str]), all_columns (bool)"] = None,
    output_format: Annotated[str, "Output format: 'row_col_value' (default, values only), 'row_col_formattedvalue' (with formatting), 'html' (HTML table with formatting), 'simple_html' (HTML table without formatting)"] = "row_col_value"
) -> str:
    """Extract table data in simplified format optimized for minimal context consumption.

    This tool provides clean, simplified table extraction without heavy formatting metadata.
    For full formatting details, use extract_formatted_table_data instead.

    Args:
        file_path: Path to the PowerPoint file (.pptx)
        slide_numbers: Optional. Slide numbers to extract tables from (1-based indexing).
                       Supports: None (all slides), int, List[int], or str (Python-style slicing)
        column_selection: Optional column filtering configuration
        output_format: Output format selection:
            - "row_col_value": [row, col, value] format with values only (default)
            - "row_col_formattedvalue": [row, col, value] format with formatting included
            - "html": HTML table with formatting (supports colspan/rowspan)
            - "simple_html": HTML table without formatting (supports colspan/rowspan)

    Returns:
        JSON string with structure:
        For row_col_value/row_col_formattedvalue:
        {
            "extracted_tables": [
                {
                    "slide_number": int,
                    "rows": int,
                    "columns": int,
                    "headers": ["col1", "col2", ...],
                    "data": [[row, col, "value"], [row, col, "value"], ...]
                }
            ]
        }

        For html/simple_html:
        {
            "extracted_html_tables": [
                {
                    "slide_number": int,
                    "rows": int,
                    "columns": int,
                    "headers": ["col1", "col2", ...],
                    "htmldata": "<table style=\"white-space: pre;\">...</table>"
                }
            ]
        }

    Example Usage:
        # Extract tables as simple row/col/value arrays
        extract_table_data("presentation.pptx")

        # Extract as HTML tables
        extract_table_data("presentation.pptx", output_format="html")

        # Extract specific slides only
        extract_table_data("presentation.pptx", slide_numbers=[1, 3, 5])
    """
    logger.info(f"extract_table_data called with file_path: {file_path}, slide_numbers: {slide_numbers}, output_format: {output_format}")

    try:
        server = get_powerpoint_server()
        arguments = {
            "file_path": file_path,
            "slide_numbers": slide_numbers,
            "column_selection": column_selection,
            "output_format": output_format
        }

        # Call the async method directly
        result = await server._extract_table_data_simple(arguments)

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
    file_path: Annotated[str, "Path to the PowerPoint file (.pptx). Must be a valid PowerPoint file. Example: 'C:\\\\temp\\\\presentation.pptx' or '/path/to/slides.pptx'"],
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