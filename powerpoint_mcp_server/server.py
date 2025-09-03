"""Main PowerPoint Analyzer MCP server implementation for PowerPoint content extraction."""

import asyncio
import json
import logging
import os
import sys
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
from .core.slide_query_engine import SlideQueryEngine, create_filters_from_dict
from .core.enhanced_table_extractor import EnhancedTableExtractor, create_table_criteria_from_dict, create_column_selection_from_dict, create_formatting_detection_from_dict, OutputFormat
from .core.text_formatting_analyzer import TextFormattingAnalyzer, create_formatting_filter_from_dict, GroupingType
from .core.data_filter_engine import DataFilterEngine, create_filter_config_from_dict
from .core.presentation_analyzer import PresentationAnalyzer, AnalysisDepth
from .tools.tool_help import get_tool_help
from .utils.file_validator import FileValidator
from .utils.zip_extractor import ZipExtractor
from .config import get_config, get_config_manager

logger = logging.getLogger(__name__)

class PowerPointMCPServer:
    """Main PowerPoint Analyzer MCP server class for PowerPoint content extraction."""

    def __init__(self):
        """Initialize the PowerPoint Analyzer MCP server."""
        try:
            self.config = get_config()
            self.config_manager = get_config_manager()

            self.server = Server(self.config.server_name)
            self.content_extractor = ContentExtractor()
            self.attribute_processor = AttributeProcessor()
            self.slide_query_engine = SlideQueryEngine(self.content_extractor)
            self.table_extractor = EnhancedTableExtractor(self.content_extractor)
            self.formatting_analyzer = TextFormattingAnalyzer(self.content_extractor)
            self.data_filter_engine = DataFilterEngine()
            self.presentation_analyzer = PresentationAnalyzer(self.content_extractor)
            self.file_validator = FileValidator()
            self._running = False
            self._setup_handlers()

            logger.info(f"PowerPoint Analyzer MCP initialized (version {self.config.server_version})")
            if self.config.debug_mode:
                self.config_manager.log_configuration()

        except Exception as e:
            logger.error(f"Failed to initialize PowerPoint Analyzer MCP: {e}")
            import traceback
            logger.error(f"Initialization traceback: {traceback.format_exc()}")
            raise

    def _setup_handlers(self):
        """Set up MCP request handlers."""
        


        @self.server.list_tools()
        async def list_tools() -> ListToolsResult:
            """List available tools."""
            logger.info("list_tools handler called")
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
                ),
                Tool(
                    name="query_slides",
                    description="Query slides with flexible filtering criteria",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "file_path": {
                                "type": "string",
                                "description": "Path to the PowerPoint file (.pptx)"
                            },
                            "search_criteria": {
                                "type": "object",
                                "description": "Search criteria for filtering slides",
                                "properties": {
                                    "title": {
                                        "type": "object",
                                        "description": "Title-based filters"
                                    },
                                    "content": {
                                        "type": "object",
                                        "description": "Content-based filters"
                                    },
                                    "layout": {
                                        "type": "object",
                                        "description": "Layout-based filters"
                                    },
                                    "slide_numbers": {
                                        "type": "array",
                                        "items": {"type": "integer"},
                                        "description": "Specific slide numbers"
                                    }
                                }
                            },
                            "return_fields": {
                                "type": "array",
                                "items": {"type": "string"},
                                "description": "Fields to include in results",
                                "default": ["slide_number", "title", "object_counts"]
                            },
                            "limit": {
                                "type": "integer",
                                "description": "Maximum number of results",
                                "default": 50
                            }
                        },
                        "required": ["file_path", "search_criteria"]
                    }
                ),
                Tool(
                    name="extract_table_data",
                    description="Extract table data with flexible selection and formatting detection",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "file_path": {
                                "type": "string",
                                "description": "Path to the PowerPoint file (.pptx)"
                            },
                            "slide_numbers": {
                                "type": "array",
                                "items": {"type": "integer"},
                                "description": "Slide numbers to extract tables from"
                            },
                            "table_criteria": {
                                "type": "object",
                                "description": "Criteria for selecting tables"
                            },
                            "column_selection": {
                                "type": "object",
                                "description": "Configuration for column selection"
                            },
                            "formatting_detection": {
                                "type": "object",
                                "description": "Configuration for formatting detection"
                            },
                            "output_format": {
                                "type": "string",
                                "enum": ["structured", "flat", "grouped_by_slide"],
                                "default": "structured"
                            },
                            "include_metadata": {
                                "type": "boolean",
                                "default": True
                            }
                        },
                        "required": ["file_path", "slide_numbers"]
                    }
                ),
                Tool(
                    name="extract_text_formatting",
                    description="Extract text with specific formatting attributes from PowerPoint slides",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "file_path": {
                                "type": "string",
                                "description": "Path to the PowerPoint file (.pptx)"
                            },
                            "formatting_type": {
                                "type": "string",
                                "enum": ["bold", "italic", "underlined", "highlighted", "strikethrough", "hyperlinks", "font_sizes", "font_colors"],
                                "description": "Type of formatting to extract"
                            },
                            "slide_numbers": {
                                "type": "array",
                                "items": {"type": "integer"},
                                "description": "Slide numbers to analyze (optional, analyzes all if not specified)"
                            }
                        },
                        "required": ["file_path", "formatting_type"]
                    }
                ),
                Tool(
                    name="analyze_text_formatting",
                    description="Analyze text formatting patterns across slides",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "file_path": {
                                "type": "string",
                                "description": "Path to the PowerPoint file (.pptx)"
                            },
                            "slide_numbers": {
                                "type": "array",
                                "items": {"type": "integer"},
                                "description": "Slide numbers to analyze (optional)"
                            },
                            "formatting_filter": {
                                "type": "object",
                                "description": "Filter configuration for formatting analysis"
                            },
                            "grouping": {
                                "type": "string",
                                "enum": ["by_slide", "by_formatting_type", "by_content_type", "by_color", "by_font_size", "none"],
                                "default": "none"
                            }
                        },
                        "required": ["file_path"]
                    }
                ),
                Tool(
                    name="filter_and_aggregate",
                    description="Filter and aggregate extracted data",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "data_source": {
                                "type": "object",
                                "description": "Source data to filter and aggregate"
                            },
                            "filter_config": {
                                "type": "object",
                                "description": "Complete filter configuration"
                            }
                        },
                        "required": ["data_source", "filter_config"]
                    }
                ),
                Tool(
                    name="get_presentation_overview",
                    description="Get comprehensive presentation overview and analysis",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "file_path": {
                                "type": "string",
                                "description": "Path to the PowerPoint file (.pptx)"
                            },
                            "analysis_depth": {
                                "type": "string",
                                "enum": ["basic", "detailed", "comprehensive"],
                                "default": "basic"
                            },
                            "include_sample_content": {
                                "type": "boolean",
                                "default": True
                            }
                        },
                        "required": ["file_path"]
                    }
                ),
                Tool(
                    name="analyze_presentation",
                    description="Analyze presentation with flexible options for text and formatting analysis",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "file_path": {
                                "type": "string",
                                "description": "Path to the PowerPoint file (.pptx)"
                            },
                            "analysis_options": {
                                "type": "object",
                                "description": "Analysis configuration options",
                                "properties": {
                                    "include_text": {
                                        "type": "boolean",
                                        "default": True,
                                        "description": "Include text content analysis"
                                    },
                                    "include_formatting": {
                                        "type": "boolean",
                                        "default": True,
                                        "description": "Include formatting analysis"
                                    },
                                    "include_structure": {
                                        "type": "boolean",
                                        "default": True,
                                        "description": "Include structural analysis"
                                    },
                                    "analysis_depth": {
                                        "type": "string",
                                        "enum": ["basic", "detailed", "comprehensive"],
                                        "default": "detailed"
                                    }
                                }
                            }
                        },
                        "required": ["file_path"]
                    }
                ),
                Tool(
                    name="tool_help",
                    description="Get detailed help and documentation for MCP tools",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "tool_name": {
                                "type": "string",
                                "description": "Name of the tool to get help for"
                            }
                        },
                        "required": ["tool_name"]
                    }
                )
            ]
            return ListToolsResult(tools=tools)

        @self.server.call_tool()
        async def call_tool(name: str, arguments: Dict[str, Any]) -> CallToolResult:
            """Handle tool calls."""
            logger.info(f"call_tool handler called: {name}")
            try:
                # Sanitize arguments to prevent boolean parsing issues
                sanitized_arguments = self._sanitize_arguments(arguments)

                if name == "extract_powerpoint_content":
                    return await self._extract_powerpoint_content(sanitized_arguments)
                elif name == "get_powerpoint_attributes":
                    return await self._get_powerpoint_attributes(sanitized_arguments)
                elif name == "get_slide_info":
                    return await self._get_slide_info(sanitized_arguments)
                elif name == "query_slides":
                    return await self._query_slides(sanitized_arguments)
                elif name == "extract_table_data":
                    return await self._extract_table_data(sanitized_arguments)
                elif name == "extract_text_formatting":
                    return await self._extract_text_formatting(sanitized_arguments)
                elif name == "analyze_text_formatting":
                    return await self._analyze_text_formatting(sanitized_arguments)
                elif name == "filter_and_aggregate":
                    return await self._filter_and_aggregate(sanitized_arguments)
                elif name == "get_presentation_overview":
                    return await self._get_presentation_overview(sanitized_arguments)
                elif name == "analyze_presentation":
                    return await self._analyze_presentation(sanitized_arguments)
                elif name == "tool_help":
                    return await self._tool_help(sanitized_arguments)
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
                    sections = self.content_extractor.extract_section_information(presentation_xml)
                    logger.debug(f"Extracted {len(sections)} sections: {sections}")
                    result['sections'] = sections

                # Get slide XML files
                slide_files = extractor.get_slide_xml_files()

                for i, slide_file in enumerate(slide_files, 1):
                    slide_xml = extractor.read_xml_content(slide_file)
                    if slide_xml:
                        # Extract slide content
                        slide_info = self.content_extractor.extract_slide_content(slide_xml, i)

                        # Try to get notes for this slide using proper mapping only
                        notes_content = ""
                        try:
                            # Use the notes mapping to find the correct notes file for this slide
                            notes_to_slide_map = self.content_extractor._build_notes_slide_mapping(extractor)
                            # Find the notes file that corresponds to this slide
                            for notes_file_path, mapped_slide_number in notes_to_slide_map.items():
                                if mapped_slide_number == i:
                                    notes_xml = extractor.read_xml_content(notes_file_path)
                                    if notes_xml:
                                        notes_content = self.content_extractor._extract_notes_content(notes_xml)
                                    break
                            # No fallback - if mapping doesn't find a notes file for this slide, 
                            # it means there are no notes for this slide
                        except Exception:
                            # Notes file doesn't exist or can't be read - that's okay
                            notes_content = ""

                        # Resolve hyperlink relationships
                        logger.info(f"Resolving hyperlinks for slide {i}")
                        self.content_extractor._resolve_hyperlink_relationships(
                            extractor, i, slide_info.text_elements
                        )

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

                # Extract notes
                logger.info("Extracting notes from PowerPoint file")
                notes = self.content_extractor.extract_notes(extractor)
                logger.info(f"Found {len(notes)} notes")
                result['notes'] = notes

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

                # Resolve hyperlink relationships
                logger.info(f"Resolving hyperlinks for slide {slide_number}")
                self.content_extractor._resolve_hyperlink_relationships(
                    extractor, slide_number, slide_info.text_elements
                )

                # Try to get notes for this slide using proper mapping only
                notes_content = ""
                try:
                    # Use the notes mapping to find the correct notes file for this slide
                    notes_to_slide_map = self.content_extractor._build_notes_slide_mapping(extractor)
                    # Find the notes file that corresponds to this slide
                    for notes_file_path, mapped_slide_number in notes_to_slide_map.items():
                        if mapped_slide_number == slide_number:
                            notes_xml = extractor.read_xml_content(notes_file_path)
                            if notes_xml:
                                notes_content = self.content_extractor._extract_notes_content(notes_xml)
                            break
                    # No fallback - if mapping doesn't find a notes file for this slide, 
                    # it means there are no notes for this slide
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

    async def _query_slides(self, arguments: Dict[str, Any]) -> CallToolResult:
        """Query slides with flexible filtering criteria."""
        try:
            file_path = arguments.get("file_path")
            search_criteria = arguments.get("search_criteria", {})
            return_fields = arguments.get("return_fields", ["slide_number", "title", "object_counts"])
            limit = arguments.get("limit", 50)

            if not file_path:
                raise ValueError("file_path is required")

            # Validate the file
            is_valid, error_message = self.file_validator.validate_file(file_path)
            if not is_valid:
                raise ValueError(f"File validation failed: {error_message}")

            # Create filters from dictionary
            filters = create_filters_from_dict(search_criteria)

            # Query slides
            results = self.slide_query_engine.query_slides(
                file_path=file_path,
                filters=filters,
                return_fields=return_fields,
                limit=limit
            )

            # Convert results to serializable format
            serializable_results = []
            for result in results:
                result_dict = {
                    "slide_number": result.slide_number,
                    "title": result.title,
                    "subtitle": result.subtitle,
                    "layout_name": result.layout_name,
                    "layout_type": result.layout_type,
                    "object_counts": result.object_counts,
                    "preview_text": result.preview_text,
                    "table_info": result.table_info,
                    "full_content": result.full_content
                }
                # Filter to only requested fields
                filtered_result = {k: v for k, v in result_dict.items() if k in return_fields or k == "slide_number"}
                serializable_results.append(filtered_result)

            response = {
                "results": serializable_results,
                "total_found": len(results),
                "search_criteria": search_criteria,
                "return_fields": return_fields
            }

            return CallToolResult(
                content=[
                    TextContent(
                        type="text",
                        text=json.dumps(response, indent=2, ensure_ascii=False)
                    )
                ]
            )

        except Exception as e:
            logger.error(f"Error querying slides: {e}")
            raise McpError(
                ErrorData(
                    code=INTERNAL_ERROR,
                    message=f"Failed to query slides: {str(e)}"
                )
            )

    async def _extract_table_data(self, arguments: Dict[str, Any]) -> CallToolResult:
        """Extract table data with flexible selection and formatting detection."""
        try:
            file_path = arguments.get("file_path")
            slide_numbers = arguments.get("slide_numbers", [])
            table_criteria_dict = arguments.get("table_criteria", {})
            column_selection_dict = arguments.get("column_selection", {})
            formatting_detection_dict = arguments.get("formatting_detection", {})
            output_format_str = arguments.get("output_format", "structured")
            include_metadata = arguments.get("include_metadata", True)

            if not file_path:
                raise ValueError("file_path is required")
            if not slide_numbers:
                raise ValueError("slide_numbers is required")

            # Validate the file
            is_valid, error_message = self.file_validator.validate_file(file_path)
            if not is_valid:
                raise ValueError(f"File validation failed: {error_message}")

            # Create configuration objects
            table_criteria = create_table_criteria_from_dict(table_criteria_dict) if table_criteria_dict else None
            column_selection = create_column_selection_from_dict(column_selection_dict) if column_selection_dict else None
            formatting_detection = create_formatting_detection_from_dict(formatting_detection_dict) if formatting_detection_dict else None
            output_format = OutputFormat(output_format_str)

            # Extract table data
            result = self.table_extractor.extract_tables(
                file_path=file_path,
                slide_numbers=slide_numbers,
                table_criteria=table_criteria,
                column_selection=column_selection,
                formatting_detection=formatting_detection,
                output_format=output_format,
                include_metadata=include_metadata
            )

            return CallToolResult(
                content=[
                    TextContent(
                        type="text",
                        text=json.dumps(result, indent=2, ensure_ascii=False)
                    )
                ]
            )

        except Exception as e:
            logger.error(f"Error extracting table data: {e}")
            raise McpError(
                ErrorData(
                    code=INTERNAL_ERROR,
                    message=f"Failed to extract table data: {str(e)}"
                )
            )

    async def _extract_text_formatting(self, arguments: Dict[str, Any]) -> CallToolResult:
        """Extract text with specific formatting attributes from PowerPoint slides."""
        try:
            file_path = arguments.get("file_path")
            formatting_type = arguments.get("formatting_type")
            slide_numbers = arguments.get("slide_numbers")

            if not file_path:
                raise ValueError("file_path is required")
            if not formatting_type:
                raise ValueError("formatting_type is required")

            # Validate formatting_type
            valid_types = ["bold", "italic", "underlined", "highlighted", "strikethrough", "hyperlinks", "font_sizes", "font_colors"]
            if formatting_type not in valid_types:
                raise ValueError(f"Invalid formatting_type: {formatting_type}. Valid options: {valid_types}")

            # Validate the file
            is_valid, error_message = self.file_validator.validate_file(file_path)
            if not is_valid:
                raise ValueError(f"File validation failed: {error_message}")

            # Import the FormattingExtractor
            from .core.formatting_extractor import FormattingExtractor
            
            # Create formatting extractor
            formatting_extractor = FormattingExtractor(self.content_extractor)

            # Extract content from the PowerPoint file
            full_content = await self._process_powerpoint_file(file_path)
            slides = full_content.get('slides', [])

            # Filter slides if specific slide numbers are provided
            if slide_numbers:
                slides = [slide for slide in slides if slide.get('slide_number') in slide_numbers]

            # Process each slide
            results_by_slide = []
            total_formatted_segments = 0
            slides_with_formatting = 0

            for slide in slides:
                slide_number = slide.get('slide_number', 0)
                title = slide.get('title', '')
                text_elements = slide.get('text_elements', [])

                # Combine all text from the slide
                complete_text = ' '.join([elem.get('content_plain', '') for elem in text_elements if elem.get('content_plain')])

                # Extract formatted segments
                formatted_segments = formatting_extractor.extract_formatting_segments(
                    text_elements, formatting_type, slide_number
                )

                # Convert segments to response format
                response_segments = []
                for segment in formatted_segments:
                    response_segments.append({
                        "text": segment.text,
                        "start_position": segment.start_position
                    })

                if response_segments:
                    slides_with_formatting += 1
                    total_formatted_segments += len(response_segments)

                results_by_slide.append({
                    "slide_number": slide_number,
                    "title": title,
                    "complete_text": complete_text,
                    "format": formatting_type,
                    "formatted_segments": response_segments
                })

            # Create response
            result = {
                "file_path": file_path,
                "formatting_type": formatting_type,
                "summary": {
                    "total_slides_analyzed": len(results_by_slide),
                    "slides_with_formatting": slides_with_formatting,
                    "total_formatted_segments": total_formatted_segments
                },
                "results_by_slide": results_by_slide
            }

            return CallToolResult(
                content=[
                    TextContent(
                        type="text",
                        text=json.dumps(result, indent=2, ensure_ascii=False)
                    )
                ]
            )

        except Exception as e:
            logger.error(f"Error extracting text formatting: {e}")
            raise McpError(
                ErrorData(
                    code=INTERNAL_ERROR,
                    message=f"Failed to extract text formatting: {str(e)}"
                )
            )

    async def _analyze_text_formatting(self, arguments: Dict[str, Any]) -> CallToolResult:
        """Analyze text formatting patterns across slides."""
        try:
            file_path = arguments.get("file_path")
            slide_numbers = arguments.get("slide_numbers")
            formatting_filter_dict = arguments.get("formatting_filter", {})
            grouping_str = arguments.get("grouping", "none")

            if not file_path:
                raise ValueError("file_path is required")

            # Validate the file
            is_valid, error_message = self.file_validator.validate_file(file_path)
            if not is_valid:
                raise ValueError(f"File validation failed: {error_message}")

            # Create configuration objects
            formatting_filter = create_formatting_filter_from_dict(formatting_filter_dict) if formatting_filter_dict else None
            grouping = GroupingType(grouping_str)

            # Analyze formatting
            result = self.formatting_analyzer.analyze_formatting(
                file_path=file_path,
                slide_numbers=slide_numbers,
                formatting_filter=formatting_filter,
                grouping=grouping
            )

            # Convert result to serializable format
            serializable_result = {
                "total_elements": result.total_elements,
                "formatted_elements": [
                    {
                        "slide_number": elem.slide_number,
                        "content_type": elem.content_type.value,
                        "element_index": elem.element_index,
                        "text_content": elem.text_content,
                        "formatting": elem.formatting,
                        "position": elem.position,
                        "size": elem.size,
                        "parent_element": elem.parent_element
                    }
                    for elem in result.formatted_elements
                ],
                "formatting_summary": result.formatting_summary,
                "groupings": result.groupings
            }

            return CallToolResult(
                content=[
                    TextContent(
                        type="text",
                        text=json.dumps(serializable_result, indent=2, ensure_ascii=False)
                    )
                ]
            )

        except Exception as e:
            logger.error(f"Error analyzing text formatting: {e}")
            raise McpError(
                ErrorData(
                    code=INTERNAL_ERROR,
                    message=f"Failed to analyze text formatting: {str(e)}"
                )
            )

    async def _filter_and_aggregate(self, arguments: Dict[str, Any]) -> CallToolResult:
        """Filter and aggregate extracted data."""
        try:
            data_source = arguments.get("data_source")
            filter_config_dict = arguments.get("filter_config", {})

            if not data_source:
                raise ValueError("data_source is required")
            if not filter_config_dict:
                raise ValueError("filter_config is required")

            # Create filter configuration
            filter_config = create_filter_config_from_dict(filter_config_dict)

            # Convert data_source to list format if needed
            if isinstance(data_source, dict) and "data" in data_source:
                data_list = data_source["data"]
            elif isinstance(data_source, list):
                data_list = data_source
            else:
                raise ValueError("data_source must be a list or dict with 'data' key")

            # Apply filtering and aggregation
            result = self.data_filter_engine.filter_and_aggregate(data_list, filter_config)

            return CallToolResult(
                content=[
                    TextContent(
                        type="text",
                        text=json.dumps(result, indent=2, ensure_ascii=False)
                    )
                ]
            )

        except Exception as e:
            logger.error(f"Error filtering and aggregating data: {e}")
            raise McpError(
                ErrorData(
                    code=INTERNAL_ERROR,
                    message=f"Failed to filter and aggregate data: {str(e)}"
                )
            )

    async def _get_presentation_overview(self, arguments: Dict[str, Any]) -> CallToolResult:
        """Get comprehensive presentation overview and analysis."""
        try:
            file_path = arguments.get("file_path")
            analysis_depth_str = arguments.get("analysis_depth", "basic")
            include_sample_content = arguments.get("include_sample_content", True)

            if not file_path:
                raise ValueError("file_path is required")

            # Validate the file
            is_valid, error_message = self.file_validator.validate_file(file_path)
            if not is_valid:
                raise ValueError(f"File validation failed: {error_message}")

            # Create configuration objects
            analysis_depth = AnalysisDepth(analysis_depth_str)

            # Analyze presentation
            result = await self.presentation_analyzer.analyze_presentation(
                file_path=file_path,
                analysis_depth=analysis_depth,
                include_sample_content=include_sample_content
            )

            # Convert result to serializable format
            serializable_result = {
                "file_path": result.file_path,
                "metadata": result.metadata,
                "structure": {
                    "total_slides": result.structure.total_slides,
                    "slide_types": result.structure.slide_types,
                    "sections": result.structure.sections,
                    "content_flow": result.structure.content_flow,
                    "structural_issues": result.structure.structural_issues
                },
                "slide_classifications": [
                    {
                        "slide_number": cls.slide_number,
                        "slide_type": cls.slide_type.value,
                        "confidence": cls.confidence,
                        "characteristics": cls.characteristics,
                        "content_summary": cls.content_summary,
                        "object_counts": cls.object_counts
                    }
                    for cls in result.slide_classifications
                ],
                "content_patterns": [
                    {
                        "pattern_type": pattern.pattern_type,
                        "pattern_name": pattern.pattern_name,
                        "occurrences": pattern.occurrences,
                        "slides": pattern.slides,
                        "examples": pattern.examples,
                        "confidence": pattern.confidence
                    }
                    for pattern in result.content_patterns
                ],
                "insights": {
                    "readability_score": result.insights.readability_score,
                    "content_density": result.insights.content_density,
                    "visual_balance": result.insights.visual_balance,
                    "consistency_issues": result.insights.consistency_issues,
                    "recommendations": result.insights.recommendations,
                    "strengths": result.insights.strengths,
                    "areas_for_improvement": result.insights.areas_for_improvement
                },
                "analysis_depth": result.analysis_depth.value,
                "sample_content": result.sample_content
            }

            return CallToolResult(
                content=[
                    TextContent(
                        type="text",
                        text=json.dumps(serializable_result, indent=2, ensure_ascii=False)
                    )
                ]
            )

        except Exception as e:
            logger.error(f"Error getting presentation overview: {e}")
            raise McpError(
                ErrorData(
                    code=INTERNAL_ERROR,
                    message=f"Failed to get presentation overview: {str(e)}"
                )
            )

    async def _analyze_presentation(self, arguments: Dict[str, Any]) -> CallToolResult:
        """Analyze presentation with flexible options for text and formatting analysis."""
        try:
            file_path = arguments.get("file_path")
            analysis_options = arguments.get("analysis_options", {})

            if not file_path:
                raise ValueError("file_path is required")

            # Validate the file
            is_valid, error_message = self.file_validator.validate_file(file_path)
            if not is_valid:
                raise ValueError(f"File validation failed: {error_message}")

            # Parse analysis options with proper boolean handling
            include_text = self._parse_boolean(analysis_options.get("include_text", True))
            include_formatting = self._parse_boolean(analysis_options.get("include_formatting", True))
            include_structure = self._parse_boolean(analysis_options.get("include_structure", True))
            analysis_depth_str = analysis_options.get("analysis_depth", "detailed")

            # Create configuration objects
            analysis_depth = AnalysisDepth(analysis_depth_str)

            # Analyze presentation using the presentation analyzer
            result = await self.presentation_analyzer.analyze_presentation(
                file_path=file_path,
                analysis_depth=analysis_depth,
                include_sample_content=include_text
            )

            # Build response based on requested options
            response = {
                "file_path": result.file_path,
                "analysis_options": {
                    "include_text": include_text,
                    "include_formatting": include_formatting,
                    "include_structure": include_structure,
                    "analysis_depth": analysis_depth_str
                }
            }

            # Add structure information if requested
            if include_structure:
                response["structure"] = {
                    "total_slides": result.structure.total_slides,
                    "slide_types": result.structure.slide_types,
                    "sections": result.structure.sections,
                    "content_flow": result.structure.content_flow,
                    "structural_issues": result.structure.structural_issues
                }

                response["slide_classifications"] = [
                    {
                        "slide_number": cls.slide_number,
                        "slide_type": cls.slide_type.value,
                        "confidence": cls.confidence,
                        "characteristics": cls.characteristics,
                        "content_summary": cls.content_summary,
                        "object_counts": cls.object_counts
                    }
                    for cls in result.slide_classifications
                ]

            # Add text content if requested
            if include_text:
                response["content_patterns"] = [
                    {
                        "pattern_type": pattern.pattern_type,
                        "pattern_name": pattern.pattern_name,
                        "occurrences": pattern.occurrences,
                        "slides": pattern.slides,
                        "examples": pattern.examples,
                        "confidence": pattern.confidence
                    }
                    for pattern in result.content_patterns
                ]

                response["insights"] = {
                    "readability_score": result.insights.readability_score,
                    "content_density": result.insights.content_density,
                    "recommendations": result.insights.recommendations
                }

            # Add formatting analysis if requested
            if include_formatting:
                # Get formatting analysis using the text formatting analyzer
                formatting_result = self.formatting_analyzer.analyze_formatting(
                    file_path=file_path,
                    slide_numbers=None,  # Analyze all slides
                    formatting_filter=None,
                    grouping=GroupingType.BY_FORMATTING_TYPE
                )

                response["formatting_analysis"] = {
                    "total_elements": formatting_result.total_elements,
                    "formatting_summary": formatting_result.formatting_summary,
                    "formatting_patterns": [
                        {
                            "slide_number": item.slide_number,
                            "content_type": item.content_type.value if hasattr(item.content_type, 'value') else str(item.content_type),
                            "element_index": item.element_index,
                            "text_content": item.text_content,
                            "formatting": item.formatting,
                            "position": item.position,
                            "size": item.size,
                            "parent_element": item.parent_element
                        }
                        for item in formatting_result.formatted_elements
                    ],
                    "groupings": formatting_result.groupings
                }

                if include_structure:
                    response["insights"]["visual_balance"] = result.insights.visual_balance
                    response["insights"]["consistency_issues"] = result.insights.consistency_issues

            # Add metadata
            response["metadata"] = result.metadata

            return CallToolResult(
                content=[
                    TextContent(
                        type="text",
                        text=json.dumps(response, indent=2, ensure_ascii=False)
                    )
                ]
            )

        except Exception as e:
            logger.error(f"Error analyzing presentation: {e}")
            raise McpError(
                ErrorData(
                    code=INTERNAL_ERROR,
                    message=f"Failed to analyze presentation: {str(e)}"
                )
            )

    async def _tool_help(self, arguments: Dict[str, Any]) -> CallToolResult:
        """Get detailed help and documentation for MCP tools."""
        try:
            tool_name = arguments.get("tool_name")
            
            if not tool_name:
                raise ValueError("tool_name is required")
            
            # Get help text for the specified tool
            help_text = get_tool_help(tool_name)
            
            if not help_text or "No help available" in help_text:
                raise ValueError(f"No help available for tool: {tool_name}")
            
            return CallToolResult(
                content=[
                    TextContent(
                        type="text",
                        text=help_text
                    )
                ]
            )
            
        except Exception as e:
            logger.error(f"Error getting tool help: {e}")
            raise McpError(
                ErrorData(
                    code=INTERNAL_ERROR,
                    message=f"Failed to get tool help: {str(e)}"
                )
            )

    def _parse_boolean(self, value) -> bool:
        """Parse boolean value from various formats (handles JSON true/false)."""
        if isinstance(value, bool):
            return value
        if isinstance(value, str):
            return value.lower() in ('true', '1', 'yes', 'on')
        if value is None:
            return False
        return bool(value)

    def _sanitize_arguments(self, arguments: Dict[str, Any]) -> Dict[str, Any]:
        """Sanitize arguments to prevent boolean parsing issues."""
        def sanitize_value(value):
            if isinstance(value, str):
                # Handle JSON boolean strings
                if value.lower() == 'true':
                    return True
                elif value.lower() == 'false':
                    return False
                # Handle other string values that might cause issues
                elif value in ('null', 'None'):
                    return None
            elif isinstance(value, dict):
                return {k: sanitize_value(v) for k, v in value.items()}
            elif isinstance(value, list):
                return [sanitize_value(item) for item in value]
            return value

        return {k: sanitize_value(v) for k, v in arguments.items()}

    def is_running(self) -> bool:
        """Check if the server is currently running."""
        return self._running

    async def shutdown(self):
        """Shutdown the server gracefully."""
        logger.info("Shutting down PowerPoint Analyzer MCP...")
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

        logger.info("PowerPoint Analyzer MCP shutdown complete")

    async def run(self):
        """Run the MCP server using direct JSON-RPC implementation."""
        try:
            logger.info("Starting PowerPoint Analyzer MCP...")
            self._running = True
            
            # Use direct JSON-RPC implementation instead of MCP library
            await self._run_direct_jsonrpc()
            
        except asyncio.CancelledError:
            logger.info("Server cancelled, shutting down gracefully...")
        except BrokenPipeError:
            logger.info("Client disconnected, shutting down gracefully...")
        except ConnectionResetError:
            logger.info("Connection reset by client, shutting down gracefully...")
        except EOFError:
            logger.info("End of input stream, shutting down gracefully...")
        except Exception as e:
            logger.error(f"Error running MCP server: {e}")
            import traceback
            logger.error(f"Traceback: {traceback.format_exc()}")
            raise
        finally:
            await self.shutdown()
    
    async def _run_direct_jsonrpc(self):
        """Run server using direct JSON-RPC implementation"""
        logger.info("Starting direct JSON-RPC server...")
        logger.info("Server ready, waiting for requests...")
        
        initialized = False
        
        try:
            while self._running:
                try:
                    # Read line from stdin
                    line = sys.stdin.readline()
                    
                    if not line:
                        logger.info("EOF received, shutting down")
                        break
                    
                    line = line.strip()
                    if not line:
                        continue
                    
                    logger.info(f"Received: {line}")
                    
                    # Parse JSON
                    try:
                        request = json.loads(line)
                        logger.info(f"Parsed request: {request}")
                    except json.JSONDecodeError as e:
                        logger.error(f"JSON parse error: {e}")
                        continue
                    
                    # Handle request
                    method = request.get("method")
                    response = None
                    
                    if method == "initialize":
                        logger.info("Handling initialize request")
                        initialized = True
                        response = {
                            "jsonrpc": "2.0",
                            "id": request.get("id"),
                            "result": {
                                "protocolVersion": "2024-11-05",
                                "capabilities": {
                                    "tools": {},
                                    "resources": {},
                                    "prompts": {}
                                },
                                "serverInfo": {
                                    "name": "powerpoint-mcp-server",
                                    "version": self.config.server_version
                                }
                            }
                        }
                    elif method == "ping":
                        logger.info("Handling ping request")
                        if not initialized:
                            response = {
                                "jsonrpc": "2.0",
                                "id": request.get("id"),
                                "error": {"code": -32002, "message": "Server not initialized"}
                            }
                        else:
                            response = {
                                "jsonrpc": "2.0",
                                "id": request.get("id"),
                                "result": {}
                            }
                    elif method == "tools/list":
                        logger.info("Handling tools/list request")
                        if not initialized:
                            response = {
                                "jsonrpc": "2.0",
                                "id": request.get("id"),
                                "error": {"code": -32002, "message": "Server not initialized"}
                            }
                        else:
                            # Use the existing list_tools handler
                            tools_result = await self._get_tools_list()
                            response = {
                                "jsonrpc": "2.0",
                                "id": request.get("id"),
                                "result": {"tools": tools_result}
                            }
                    elif method == "tools/call":
                        logger.info("Handling tools/call request")
                        if not initialized:
                            response = {
                                "jsonrpc": "2.0",
                                "id": request.get("id"),
                                "error": {"code": -32002, "message": "Server not initialized"}
                            }
                        else:
                            # Handle tool call
                            params = request.get("params", {})
                            tool_name = params.get("name")
                            arguments = params.get("arguments", {})
                            
                            try:
                                result = await self._call_tool(tool_name, arguments)
                                response = {
                                    "jsonrpc": "2.0",
                                    "id": request.get("id"),
                                    "result": result
                                }
                            except Exception as e:
                                response = {
                                    "jsonrpc": "2.0",
                                    "id": request.get("id"),
                                    "error": {
                                        "code": -32603,
                                        "message": f"Tool execution failed: {str(e)}"
                                    }
                                }
                    elif method and method.startswith("notifications/"):
                        logger.info(f"Received notification: {method}")
                        # No response for notifications
                        continue
                    else:
                        response = {
                            "jsonrpc": "2.0",
                            "id": request.get("id"),
                            "error": {"code": -32601, "message": "Method not found"}
                        }
                    
                    # Send response
                    if response is not None:
                        response_json = json.dumps(response)
                        logger.info(f"Sending response: {response_json}")
                        
                        # Write to stdout and flush immediately
                        sys.stdout.write(response_json + "\\n")
                        sys.stdout.flush()
                        logger.info("Response sent and flushed")
                
                except EOFError:
                    logger.info("EOF on stdin, shutting down")
                    break
                except KeyboardInterrupt:
                    logger.info("Keyboard interrupt, shutting down")
                    break
                except Exception as e:
                    logger.error(f"Error processing request: {e}")
                    import traceback
                    logger.error(f"Traceback: {traceback.format_exc()}")
                    continue
        
        except Exception as e:
            logger.error(f"Fatal server error: {e}")
            import traceback
            logger.error(f"Traceback: {traceback.format_exc()}")
            raise
    
    async def _get_tools_list(self):
        """Get tools list for direct JSON-RPC implementation"""
        return [
            {
                "name": "extract_powerpoint_content",
                "description": "Extract complete structured content from a PowerPoint file",
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "file_path": {
                            "type": "string",
                            "description": "Path to the PowerPoint file (.pptx)"
                        }
                    },
                    "required": ["file_path"]
                }
            },
            {
                "name": "get_powerpoint_attributes",
                "description": "Get specific attributes from PowerPoint slides",
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "file_path": {
                            "type": "string",
                            "description": "Path to the PowerPoint file (.pptx)"
                        },
                        "attributes": {
                            "type": "array",
                            "items": {"type": "string"},
                            "description": "List of attributes to extract"
                        }
                    },
                    "required": ["file_path", "attributes"]
                }
            },
            {
                "name": "analyze_presentation",
                "description": "Analyze presentation with flexible options",
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "file_path": {
                            "type": "string",
                            "description": "Path to the PowerPoint file (.pptx)"
                        }
                    },
                    "required": ["file_path"]
                }
            },
            {
                "name": "tool_help",
                "description": "Get detailed help and documentation for MCP tools",
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "tool_name": {
                            "type": "string",
                            "description": "Name of the tool to get help for"
                        }
                    },
                    "required": ["tool_name"]
                }
            }
        ]
    
    async def _call_tool(self, name: str, arguments: dict):
        """Call tool for direct JSON-RPC implementation"""
        logger.info(f"Calling tool: {name} with {arguments}")
        
        # Sanitize arguments
        sanitized_arguments = self._sanitize_arguments(arguments)
        
        if name == "extract_powerpoint_content":
            return await self._extract_powerpoint_content(sanitized_arguments)
        elif name == "get_powerpoint_attributes":
            return await self._get_powerpoint_attributes(sanitized_arguments)
        elif name == "analyze_presentation":
            return await self._analyze_presentation(sanitized_arguments)
        elif name == "tool_help":
            return await self._tool_help(sanitized_arguments)
        else:
            raise ValueError(f"Unknown tool: {name}")


async def main():
    """Main entry point."""
    server = PowerPointMCPServer()
    await server.run()


if __name__ == "__main__":
    asyncio.run(main())
