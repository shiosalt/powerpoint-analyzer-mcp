#!/usr/bin/env python3
"""Simplified main entry point for the PowerPoint MCP Server."""

import asyncio
import json
import logging
import sys
from typing import Any, Dict

from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.types import (
    CallToolResult,
    ListToolsResult,
    Tool,
    TextContent,
    ErrorData,
    INTERNAL_ERROR,
    METHOD_NOT_FOUND
)
from mcp.server.models import InitializationOptions
from mcp.types import ServerCapabilities
from mcp import McpError

from powerpoint_mcp_server.core.content_extractor import ContentExtractor
from powerpoint_mcp_server.core.attribute_processor import AttributeProcessor
from powerpoint_mcp_server.utils.file_validator import FileValidator
from powerpoint_mcp_server.utils.zip_extractor import ZipExtractor

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[logging.StreamHandler(sys.stderr)]
)

logger = logging.getLogger(__name__)

# Initialize server
server = Server("powerpoint-mcp-server")

# Initialize components
content_extractor = ContentExtractor()
attribute_processor = AttributeProcessor()
file_validator = FileValidator()


@server.list_tools()
async def list_tools() -> ListToolsResult:
    """List available tools."""
    return ListToolsResult(
        tools=[
            Tool(
                name="extract_powerpoint_content",
                description="Extract complete structured content from a PowerPoint file",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "file_path": {
                            "type": "string",
                            "description": "Path to the PowerPoint file (.pptx)"
                        }
                    },
                    "required": ["file_path"]
                }
            ),
            Tool(
                name="get_powerpoint_attributes",
                description="Get specific attributes from PowerPoint slides",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "file_path": {
                            "type": "string",
                            "description": "Path to the PowerPoint file (.pptx)"
                        },
                        "attributes": {
                            "type": "array",
                            "items": {"type": "string"},
                            "description": "List of attributes to extract (title, subtitle, text, tables, images, layout, size, sections, notes, object_counts)"
                        }
                    },
                    "required": ["file_path", "attributes"]
                }
            ),
            Tool(
                name="get_slide_info",
                description="Get information for a specific slide",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "file_path": {
                            "type": "string",
                            "description": "Path to the PowerPoint file (.pptx)"
                        },
                        "slide_number": {
                            "type": "integer",
                            "description": "Slide number (1-based)"
                        }
                    },
                    "required": ["file_path", "slide_number"]
                }
            )
        ]
    )


@server.call_tool()
async def call_tool(name: str, arguments: Dict[str, Any]) -> CallToolResult:
    """Handle tool calls."""
    try:
        if name == "extract_powerpoint_content":
            return await extract_powerpoint_content(arguments)
        elif name == "get_powerpoint_attributes":
            return await get_powerpoint_attributes(arguments)
        elif name == "get_slide_info":
            return await get_slide_info(arguments)
        else:
            raise McpError(
                ErrorData(
                    code=METHOD_NOT_FOUND,
                    message=f"Unknown tool: {name}"
                )
            )
    except Exception as e:
        logger.error(f"Error in tool call {name}: {str(e)}")
        raise McpError(
            ErrorData(
                code=INTERNAL_ERROR,
                message=f"Tool execution failed: {str(e)}"
            )
        )


async def extract_powerpoint_content(arguments: Dict[str, Any]) -> CallToolResult:
    """Extract complete PowerPoint content."""
    file_path = arguments.get("file_path")
    if not file_path:
        raise ValueError("file_path is required")
    
    # Validate the file
    is_valid, error_message = file_validator.validate_file(file_path)
    if not is_valid:
        raise ValueError(f"File validation failed: {error_message}")
    
    # Extract content from the PowerPoint file
    result = await process_powerpoint_file(file_path)
    
    return CallToolResult(
        content=[
            TextContent(
                type="text",
                text=json.dumps(result, indent=2, ensure_ascii=False)
            )
        ]
    )


async def get_powerpoint_attributes(arguments: Dict[str, Any]) -> CallToolResult:
    """Get specific PowerPoint attributes."""
    file_path = arguments.get("file_path")
    attributes = arguments.get("attributes", [])
    
    if not file_path:
        raise ValueError("file_path is required")
    if not attributes:
        raise ValueError("attributes list is required")
    
    # Validate the file
    is_valid, error_message = file_validator.validate_file(file_path)
    if not is_valid:
        raise ValueError(f"File validation failed: {error_message}")
    
    # Extract content from the PowerPoint file
    full_content = await process_powerpoint_file(file_path)
    
    # Filter to requested attributes
    filtered_content = attribute_processor.filter_attributes(full_content, attributes)
    
    return CallToolResult(
        content=[
            TextContent(
                type="text",
                text=json.dumps(filtered_content, indent=2, ensure_ascii=False)
            )
        ]
    )


async def get_slide_info(arguments: Dict[str, Any]) -> CallToolResult:
    """Get specific slide information."""
    file_path = arguments.get("file_path")
    slide_number = arguments.get("slide_number")
    
    if not file_path:
        raise ValueError("file_path is required")
    if slide_number is None:
        raise ValueError("slide_number is required")
    
    # Validate the file
    is_valid, error_message = file_validator.validate_file(file_path)
    if not is_valid:
        raise ValueError(f"File validation failed: {error_message}")
    
    # Extract specific slide information
    slide_info = await process_single_slide(file_path, slide_number)
    
    return CallToolResult(
        content=[
            TextContent(
                type="text",
                text=json.dumps(slide_info, indent=2, ensure_ascii=False)
            )
        ]
    )


async def process_powerpoint_file(file_path: str) -> Dict[str, Any]:
    """Process a complete PowerPoint file and extract all content."""
    result = {
        'file_path': file_path,
        'slides': [],
        'metadata': {}
    }
    
    # Extract PowerPoint content using ZipExtractor
    with ZipExtractor(file_path) as extractor:
        # Get presentation metadata
        presentation_xml = extractor.read_xml_content('ppt/presentation.xml')
        if presentation_xml:
            result['metadata'] = content_extractor.extract_presentation_metadata(presentation_xml)
            result['slide_size'] = content_extractor.get_slide_size_info(presentation_xml)
            result['sections'] = content_extractor.extract_section_information(presentation_xml)
        
        # Get slide XML files
        slide_files = extractor.get_slide_xml_files()
        
        for i, slide_file in enumerate(slide_files, 1):
            slide_xml = extractor.read_xml_content(slide_file)
            if slide_xml:
                # Extract slide content
                slide_info = content_extractor.extract_slide_content(slide_xml, i)
                
                # Try to get notes for this slide (optional)
                notes_file = f'ppt/notesSlides/notesSlide{i}.xml'
                notes_content = ""
                try:
                    notes_xml = extractor.read_xml_content(notes_file)
                    if notes_xml:
                        notes_content = content_extractor._extract_notes_content(notes_xml)
                except Exception:
                    # Notes file doesn't exist or can't be read - that's okay
                    notes_content = ""
                
                # Create slide data
                slide_data = {
                    'slide_number': i,
                    'title': slide_info.title,
                    'subtitle': slide_info.subtitle,
                    'layout_name': slide_info.layout_name,
                    'layout_type': slide_info.layout_type,
                    'placeholders': slide_info.placeholders,
                    'text_elements': slide_info.text_elements,
                    'tables': slide_info.tables,
                    'notes': notes_content,
                    'object_counts': content_extractor._count_slide_objects(
                        content_extractor.xml_parser.parse_xml_string(slide_xml)
                    )
                }
                
                result['slides'].append(slide_data)
    
    return result


async def process_single_slide(file_path: str, slide_number: int) -> Dict[str, Any]:
    """Process a single slide and extract its information."""
    with ZipExtractor(file_path) as extractor:
        # Get slide XML files
        slide_files = extractor.get_slide_xml_files()
        
        # Get the specific slide (slide_files is a dict, convert to list)
        slide_file_list = list(slide_files.keys())
        if slide_number < 1 or slide_number > len(slide_file_list):
            raise ValueError(f"Slide number {slide_number} is out of range (1-{len(slide_file_list)})")
        slide_file = slide_file_list[slide_number - 1]
        slide_xml = extractor.read_xml_content(slide_file)
        
        if not slide_xml:
            raise ValueError(f"Could not read slide {slide_number}")
        
        # Extract slide content
        slide_info = content_extractor.extract_slide_content(slide_xml, slide_number)
        
        # Try to get notes for this slide (optional)
        notes_file = f'ppt/notesSlides/notesSlide{slide_number}.xml'
        notes_content = ""
        try:
            notes_xml = extractor.read_xml_content(notes_file)
            if notes_xml:
                notes_content = content_extractor._extract_notes_content(notes_xml)
        except Exception:
            # Notes file doesn't exist or can't be read - that's okay
            notes_content = ""
        
        # Get presentation metadata for context
        presentation_xml = extractor.read_xml_content('ppt/presentation.xml')
        slide_size = {}
        if presentation_xml:
            slide_size = content_extractor.get_slide_size_info(presentation_xml)
        
        return {
            'slide_number': slide_number,
            'title': slide_info.title,
            'subtitle': slide_info.subtitle,
            'layout_name': slide_info.layout_name,
            'layout_type': slide_info.layout_type,
            'placeholders': slide_info.placeholders,
            'text_elements': slide_info.text_elements,
            'tables': slide_info.tables,
            'notes': notes_content,
            'object_counts': content_extractor._count_slide_objects(
                content_extractor.xml_parser.parse_xml_string(slide_xml)
            ),
            'slide_size': slide_size
        }


async def main():
    """Main entry point."""
    logger.info("Starting PowerPoint MCP Server...")
    
    async with stdio_server() as (read_stream, write_stream):
        logger.info("MCP server connected to stdio streams")
        await server.run(
            read_stream, 
            write_stream,
            InitializationOptions(
                server_name="powerpoint-mcp-server",
                server_version="1.0.0",
                capabilities=ServerCapabilities()
            )
        )


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        logger.info("Server stopped by user")
        sys.exit(0)
    except Exception as e:
        logger.error(f"Server error: {e}")
        sys.exit(1)