#!/usr/bin/env python3
"""Main entry point for the PowerPoint MCP Server using FastMCP 2.0."""

import asyncio
import json
import logging
import sys
import os
from typing import Any, Dict, List, Optional
from pathlib import Path

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

def get_powerpoint_server() -> PowerPointMCPServer:
    """Get or create the PowerPoint server instance."""
    global powerpoint_server
    if powerpoint_server is None:
        logger.info("Initializing PowerPoint MCP Server...")
        powerpoint_server = PowerPointMCPServer()
        logger.info("PowerPoint MCP Server initialized successfully")
    return powerpoint_server

# Create FastMCP instance with correct syntax
mcp = FastMCP(config.server_name)

@mcp.tool
def extract_powerpoint_content(file_path: str) -> str:
    """Extract complete structured content from a PowerPoint file.
    
    Args:
        file_path: Path to the PowerPoint file (.pptx)
        
    Returns:
        JSON string containing the complete structured content of the PowerPoint file
    """
    logger.info(f"extract_powerpoint_content called with file_path: {file_path}")
    
    try:
        server = get_powerpoint_server()
        arguments = {"file_path": file_path}
        logger.debug(f"Calling server._extract_powerpoint_content with arguments: {arguments}")
        
        # Call the async method synchronously
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        try:
            result = loop.run_until_complete(server._extract_powerpoint_content(arguments))
        finally:
            loop.close()
        
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
def get_powerpoint_attributes(file_path: str, attributes: List[str]) -> str:
    """Get specific attributes from PowerPoint slides.
    
    Args:
        file_path: Path to the PowerPoint file (.pptx)
        attributes: List of attributes to extract (title, subtitle, text, tables, images, layout, size, sections, notes, object_counts)
        
    Returns:
        JSON string containing the requested attributes from the PowerPoint file
    """
    logger.info(f"get_powerpoint_attributes called with file_path: {file_path}, attributes: {attributes}")
    
    try:
        server = get_powerpoint_server()
        arguments = {"file_path": file_path, "attributes": attributes}
        logger.debug(f"Calling server._get_powerpoint_attributes with arguments: {arguments}")
        
        # Call the async method synchronously
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        try:
            result = loop.run_until_complete(server._get_powerpoint_attributes(arguments))
        finally:
            loop.close()
        
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
def get_slide_info(file_path: str, slide_number: int) -> str:
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
        
        # Call the async method synchronously
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        try:
            result = loop.run_until_complete(server._get_slide_info(arguments))
        finally:
            loop.close()
        
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
def query_slides(file_path: str, search_criteria: Dict[str, Any], return_fields: Optional[List[str]] = None, limit: int = 50) -> str:
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
        
        # Call the async method synchronously
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        try:
            result = loop.run_until_complete(server._query_slides(arguments))
        finally:
            loop.close()
        
        # Extract text content from CallToolResult
        content_text = ""
        if result.content:
            for content_item in result.content:
                if hasattr(content_item, 'text'):
                    content_text += content_item.text
        
        return content_text
        
    except Exception as e:
        logger.error(f"Error in query_slides: {e}")
        return f"Error: {str(e)}"

@mcp.tool
def extract_table_data(file_path: str, slide_numbers: List[int], table_criteria: Optional[Dict[str, Any]] = None, 
                      column_selection: Optional[Dict[str, Any]] = None, formatting_detection: Optional[Dict[str, Any]] = None,
                      output_format: str = "structured", include_metadata: bool = True) -> str:
    """Extract table data with flexible selection and formatting detection.
    
    Args:
        file_path: Path to the PowerPoint file (.pptx)
        slide_numbers: Slide numbers to extract tables from
        table_criteria: Criteria for selecting tables (optional)
        column_selection: Configuration for column selection (optional)
        formatting_detection: Configuration for formatting detection (optional)
        output_format: Output format (structured, flat, grouped_by_slide)
        include_metadata: Whether to include table metadata
        
    Returns:
        JSON string containing the extracted table data
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
        
        # Call the async method synchronously
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        try:
            result = loop.run_until_complete(server._extract_table_data(arguments))
        finally:
            loop.close()
        
        # Extract text content from CallToolResult
        content_text = ""
        if result.content:
            for content_item in result.content:
                if hasattr(content_item, 'text'):
                    content_text += content_item.text
        
        return content_text
        
    except Exception as e:
        logger.error(f"Error in extract_table_data: {e}")
        return f"Error: {str(e)}"

@mcp.tool
def get_presentation_overview(file_path: str, analysis_depth: str = "basic", include_sample_content: bool = True) -> str:
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
        
        # Call the async method synchronously
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        try:
            result = loop.run_until_complete(server._get_presentation_overview(arguments))
        finally:
            loop.close()
        
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

def main():
    """Main entry point for the FastMCP PowerPoint server."""
    logger.info(f"Starting PowerPoint MCP Server using FastMCP 2.0: {config.server_name} v{config.server_version}")
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