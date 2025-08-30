"""Main MCP server implementation for PowerPoint content extraction."""

import asyncio
import json
import logging
import os
from typing import Any, Dict, List, Optional

from mcp.server import Server
from mcp.server.models import InitializationOptions
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
from mcp import McpError

from .core.content_extractor import ContentExtractor
from .core.attribute_processor import AttributeProcessor
from .utils.file_validator import FileValidator
from .utils.zip_extractor import ZipExtractor
from .config import get_config, get_config_manager

logger = logging.getLogger(__name__)

class PowerPointMCPServer:
    """Main MCP server class for PowerPoint content extraction."""
    
    def __init__(self):
        """Initialize the PowerPoint MCP server."""
        self.config = get_config()
        self.config_manager = get_config_manager()
        
        self.server = Server(self.config.server_name)
        self.content_extractor = ContentExtractor()
        self.attribute_processor = AttributeProcessor()
        self.file_validator = FileValidator()
        self._running = False
        self._setup_handlers()
        
        logger.info(f"PowerPoint MCP Server initialized (version {self.config.server_version})")
        if self.config.debug_mode:
            self.config_manager.log_configuration()
    
    def _setup_handlers(self):
        """Set up MCP request handlers."""
        
        @self.server.list_tools()
        async def list_tools() -> ListToolsResult:
            """List available tools."""
            tools = [
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
            return ListToolsResult(tools=tools)
        
        @self.server.call_tool()
        async def call_tool(name: str, arguments: Dict[str, Any]) -> CallToolResult:
            """Handle tool calls."""
            try:
                if name == "extract_powerpoint_content":
                    return await self._extract_powerpoint_content(arguments)
                elif name == "get_powerpoint_attributes":
                    return await self._get_powerpoint_attributes(arguments)
                elif name == "get_slide_info":
                    return await self._get_slide_info(arguments)
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
    
    async def _extract_powerpoint_content(self, arguments: Dict[str, Any]) -> CallToolResult:
        """Extract complete PowerPoint content."""
        try:
            file_path = arguments.get("file_path")
            if not file_path:
                raise ValueError("file_path is required")
            
            # Validate the file
            is_valid, error_message = self.file_validator.validate_file(file_path)
            if not is_valid:
                raise ValueError(f"File validation failed: {error_message}")
            
            # Extract content from the PowerPoint file
            result = await self._process_powerpoint_file(file_path)
            
            return CallToolResult(
                content=[
                    TextContent(
                        type="text",
                        text=json.dumps(result, indent=2, ensure_ascii=False)
                    )
                ]
            )
            
        except Exception as e:
            logger.error(f"Error extracting PowerPoint content: {e}")
            raise McpError(
                ErrorData(
                    code=INTERNAL_ERROR,
                    message=f"Failed to extract PowerPoint content: {str(e)}"
                )
            )
    
    async def _get_powerpoint_attributes(self, arguments: Dict[str, Any]) -> CallToolResult:
        """Get specific PowerPoint attributes."""
        try:
            file_path = arguments.get("file_path")
            attributes = arguments.get("attributes", [])
            
            if not file_path:
                raise ValueError("file_path is required")
            if not attributes:
                raise ValueError("attributes list is required")
            
            # Validate the file
            is_valid, error_message = self.file_validator.validate_file(file_path)
            if not is_valid:
                raise ValueError(f"File validation failed: {error_message}")
            
            # Extract content from the PowerPoint file
            full_content = await self._process_powerpoint_file(file_path)
            
            # Filter to requested attributes
            filtered_content = self.attribute_processor.filter_attributes(full_content, attributes)
            
            return CallToolResult(
                content=[
                    TextContent(
                        type="text",
                        text=json.dumps(filtered_content, indent=2, ensure_ascii=False)
                    )
                ]
            )
            
        except Exception as e:
            logger.error(f"Error getting PowerPoint attributes: {e}")
            raise McpError(
                ErrorData(
                    code=INTERNAL_ERROR,
                    message=f"Failed to get PowerPoint attributes: {str(e)}"
                )
            )
    
    async def _get_slide_info(self, arguments: Dict[str, Any]) -> CallToolResult:
        """Get specific slide information."""
        try:
            file_path = arguments.get("file_path")
            slide_number = arguments.get("slide_number")
            
            if not file_path:
                raise ValueError("file_path is required")
            if slide_number is None:
                raise ValueError("slide_number is required")
            
            # Validate the file
            is_valid, error_message = self.file_validator.validate_file(file_path)
            if not is_valid:
                raise ValueError(f"File validation failed: {error_message}")
            
            # Extract specific slide information
            slide_info = await self._process_single_slide(file_path, slide_number)
            
            return CallToolResult(
                content=[
                    TextContent(
                        type="text",
                        text=json.dumps(slide_info, indent=2, ensure_ascii=False)
                    )
                ]
            )
            
        except Exception as e:
            logger.error(f"Error getting slide info: {e}")
            raise McpError(
                ErrorData(
                    code=INTERNAL_ERROR,
                    message=f"Failed to get slide info: {str(e)}"
                )
            )
    
    async def _process_powerpoint_file(self, file_path: str) -> Dict[str, Any]:
        """Process a complete PowerPoint file and extract all content."""
        try:
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
                    result['metadata'] = self.content_extractor.extract_presentation_metadata(presentation_xml)
                    result['slide_size'] = self.content_extractor.get_slide_size_info(presentation_xml)
                    result['sections'] = self.content_extractor.extract_section_information(presentation_xml)
                
                # Get slide XML files
                slide_files = extractor.get_slide_xml_files()
                
                for i, slide_file in enumerate(slide_files, 1):
                    slide_xml = extractor.read_xml_content(slide_file)
                    if slide_xml:
                        # Extract slide content
                        slide_info = self.content_extractor.extract_slide_content(slide_xml, i)
                        
                        # Try to get notes for this slide (optional)
                        notes_file = f'ppt/notesSlides/notesSlide{i}.xml'
                        notes_content = ""
                        try:
                            notes_xml = extractor.read_xml_content(notes_file)
                            if notes_xml:
                                notes_content = self.content_extractor._extract_notes_content(notes_xml)
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
                            'object_counts': self.content_extractor._count_slide_objects(
                                self.content_extractor.xml_parser.parse_xml_string(slide_xml)
                            )
                        }
                        
                        result['slides'].append(slide_data)
            
            return result
            
        except Exception as e:
            logger.error(f"Error processing PowerPoint file {file_path}: {e}")
            raise
    
    async def _process_single_slide(self, file_path: str, slide_number: int) -> Dict[str, Any]:
        """Process a single slide and extract its information."""
        try:
            with ZipExtractor(file_path) as extractor:
                # Get slide XML files
                slide_files = extractor.get_slide_xml_files()
                
                # Check will be done below when we convert to list
                
                # Get the specific slide (slide_files is a dict, convert to list)
                slide_file_list = list(slide_files.keys())
                if slide_number < 1 or slide_number > len(slide_file_list):
                    raise ValueError(f"Slide number {slide_number} is out of range (1-{len(slide_file_list)})")
                slide_file = slide_file_list[slide_number - 1]
                slide_xml = extractor.read_xml_content(slide_file)
                
                if not slide_xml:
                    raise ValueError(f"Could not read slide {slide_number}")
                
                # Extract slide content
                slide_info = self.content_extractor.extract_slide_content(slide_xml, slide_number)
                
                # Try to get notes for this slide (optional)
                notes_file = f'ppt/notesSlides/notesSlide{slide_number}.xml'
                notes_content = ""
                try:
                    notes_xml = extractor.read_xml_content(notes_file)
                    if notes_xml:
                        notes_content = self.content_extractor._extract_notes_content(notes_xml)
                except Exception:
                    # Notes file doesn't exist or can't be read - that's okay
                    notes_content = ""
                
                # Get presentation metadata for context
                presentation_xml = extractor.read_xml_content('ppt/presentation.xml')
                slide_size = {}
                if presentation_xml:
                    slide_size = self.content_extractor.get_slide_size_info(presentation_xml)
                
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
                    'object_counts': self.content_extractor._count_slide_objects(
                        self.content_extractor.xml_parser.parse_xml_string(slide_xml)
                    ),
                    'slide_size': slide_size
                }
                
        except Exception as e:
            logger.error(f"Error processing slide {slide_number} from {file_path}: {e}")
            raise
    
    async def run(self):
        """Run the MCP server."""
        try:
            self._running = True
            logger.info("PowerPoint MCP Server starting...")
            
            async with stdio_server() as (read_stream, write_stream):
                logger.info("MCP server connected to stdio streams")
                
                await self.server.run(
                    read_stream,
                    write_stream,
                    InitializationOptions(
                        server_name=self.config.server_name,
                        server_version=self.config.server_version
                    )
                )
        except Exception as e:
            logger.error(f"Error running MCP server: {e}")
            raise
        finally:
            self._running = False
            logger.info("PowerPoint MCP Server stopped")
    
    def is_running(self) -> bool:
        """Check if the server is currently running."""
        return self._running
    
    async def shutdown(self):
        """Shutdown the server gracefully."""
        logger.info("Shutting down PowerPoint MCP Server...")
        self._running = False
        
        # Perform any cleanup operations
        try:
            # Clear any cached data
            if hasattr(self.content_extractor, 'cache_manager'):
                cache_manager = self.content_extractor.cache_manager
                if hasattr(cache_manager, 'clear_cache'):
                    cache_manager.clear_cache()
                    logger.debug("Cache cleared during shutdown")
        except Exception as e:
            logger.warning(f"Error during shutdown cleanup: {e}")
        
        logger.info("PowerPoint MCP Server shutdown complete")


async def main():
    """Main entry point."""
    server = PowerPointMCPServer()
    await server.run()


if __name__ == "__main__":
    asyncio.run(main())