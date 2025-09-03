"""
Tool help system for providing detailed documentation of MCP tools.
"""

import json
from typing import Dict, Any, List, Optional


class ToolHelpSystem:
    """Provides detailed help and documentation for MCP tools."""
    
    def __init__(self):
        self.tool_docs = self._initialize_tool_docs()
    
    def _initialize_tool_docs(self) -> Dict[str, Dict[str, Any]]:
        """Initialize comprehensive tool documentation."""
        return {
            "query_slides": {
                "description": "Query slides with flexible filtering criteria",
                "parameters": {
                    "file_path": {
                        "type": "string",
                        "required": True,
                        "description": "Path to the PowerPoint file (.pptx)"
                    },
                    "search_criteria": {
                        "type": "object",
                        "required": True,
                        "description": "Search criteria for filtering slides",
                        "schema": {
                            "title": {
                                "type": "object",
                                "description": "Filter slides by title content",
                                "properties": {
                                    "contains": {
                                        "type": "string",
                                        "description": "Title must contain this text (case-insensitive)"
                                    },
                                    "starts_with": {
                                        "type": "string", 
                                        "description": "Title must start with this text (case-insensitive)"
                                    },
                                    "ends_with": {
                                        "type": "string",
                                        "description": "Title must end with this text (case-insensitive)"
                                    },
                                    "regex": {
                                        "type": "string",
                                        "description": "Title must match this regex pattern (case-insensitive)"
                                    },
                                    "one_of": {
                                        "type": "array",
                                        "items": {"type": "string"},
                                        "description": "Title must match at least one of these patterns (regex or string)"
                                    }
                                }
                            },
                            "content": {
                                "type": "object",
                                "description": "Filter slides by content characteristics",
                                "properties": {
                                    "contains_text": {
                                        "type": "string",
                                        "description": "Slide content must contain this text (case-insensitive)"
                                    },
                                    "has_tables": {
                                        "type": "boolean",
                                        "description": "Whether slide must have tables (true) or not have tables (false)"
                                    },
                                    "has_charts": {
                                        "type": "boolean",
                                        "description": "Whether slide must have charts (true) or not have charts (false)"
                                    },
                                    "has_images": {
                                        "type": "boolean",
                                        "description": "Whether slide must have images (true) or not have images (false)"
                                    },
                                    "object_count": {
                                        "type": "object",
                                        "description": "Filter by total number of objects on slide",
                                        "properties": {
                                            "min": {
                                                "type": "integer",
                                                "description": "Minimum number of objects"
                                            },
                                            "max": {
                                                "type": "integer", 
                                                "description": "Maximum number of objects"
                                            }
                                        }
                                    }
                                }
                            },
                            "layout": {
                                "type": "object",
                                "description": "Filter slides by layout properties",
                                "properties": {
                                    "type": {
                                        "type": "string",
                                        "description": "Layout type (e.g., 'Title Slide', 'Content', 'Blank')"
                                    },
                                    "name": {
                                        "type": "string",
                                        "description": "Layout name (case-insensitive substring match)"
                                    }
                                }
                            },
                            "slide_numbers": {
                                "type": "array",
                                "items": {"type": "integer"},
                                "description": "Specific slide numbers to include (1-based)"
                            },
                            "section": {
                                "type": "string",
                                "description": "Filter by presentation section name (if sections are defined)"
                            }
                        }
                    },
                    "return_fields": {
                        "type": "array",
                        "items": {"type": "string"},
                        "required": False,
                        "description": "Fields to include in results",
                        "default": ["slide_number", "title", "object_counts"],
                        "valid_values": [
                            "slide_number",
                            "title", 
                            "subtitle",
                            "layout",
                            "object_counts",
                            "preview_text",
                            "table_info",
                            "full_content"
                        ]
                    },
                    "limit": {
                        "type": "integer",
                        "required": False,
                        "description": "Maximum number of results to return",
                        "default": 50,
                        "minimum": 1
                    }
                },
                "examples": [
                    {
                        "name": "Find slides with 'Introduction' in title",
                        "search_criteria": {
                            "title": {
                                "contains": "Introduction"
                            }
                        },
                        "return_fields": ["slide_number", "title", "preview_text"]
                    },
                    {
                        "name": "Find slides with tables and charts",
                        "search_criteria": {
                            "content": {
                                "has_tables": True,
                                "has_charts": True
                            }
                        },
                        "return_fields": ["slide_number", "title", "table_info", "object_counts"]
                    },
                    {
                        "name": "Find content slides with specific text",
                        "search_criteria": {
                            "layout": {
                                "type": "Content"
                            },
                            "content": {
                                "contains_text": "revenue"
                            }
                        },
                        "return_fields": ["slide_number", "title", "full_content"],
                        "limit": 10
                    },
                    {
                        "name": "Find slides by multiple title patterns",
                        "search_criteria": {
                            "title": {
                                "one_of": ["Summary", "Conclusion", "Next Steps"]
                            }
                        }
                    },
                    {
                        "name": "Find slides with many objects",
                        "search_criteria": {
                            "content": {
                                "object_count": {
                                    "min": 5,
                                    "max": 20
                                }
                            }
                        },
                        "return_fields": ["slide_number", "title", "object_counts"]
                    },
                    {
                        "name": "Get specific slides with full details",
                        "search_criteria": {
                            "slide_numbers": [1, 5, 10]
                        },
                        "return_fields": ["slide_number", "title", "subtitle", "layout", "full_content"]
                    }
                ],
                "notes": [
                    "All text matching is case-insensitive",
                    "Multiple conditions within the same filter category use AND logic",
                    "The 'one_of' condition uses OR logic for its patterns",
                    "Regex patterns are supported in title filters and 'one_of' arrays",
                    "Object counts include text_boxes, tables, images, shapes, and charts",
                    "Layout types are extracted from PowerPoint's built-in layout definitions",
                    "Empty or null values are handled gracefully in all filters"
                ]
            }
        }
    
    def get_tool_help(self, tool_name: str) -> Optional[Dict[str, Any]]:
        """Get comprehensive help for a specific tool."""
        return self.tool_docs.get(tool_name)
    
    def get_parameter_help(self, tool_name: str, parameter_name: str) -> Optional[Dict[str, Any]]:
        """Get detailed help for a specific parameter."""
        tool_doc = self.tool_docs.get(tool_name)
        if not tool_doc:
            return None
        
        parameters = tool_doc.get("parameters", {})
        return parameters.get(parameter_name)
    
    def get_examples(self, tool_name: str) -> List[Dict[str, Any]]:
        """Get usage examples for a tool."""
        tool_doc = self.tool_docs.get(tool_name)
        if not tool_doc:
            return []
        
        return tool_doc.get("examples", [])
    
    def format_help_text(self, tool_name: str) -> str:
        """Format comprehensive help text for a tool."""
        tool_doc = self.get_tool_help(tool_name)
        if not tool_doc:
            return f"No help available for tool: {tool_name}"
        
        help_text = []
        help_text.append(f"# {tool_name}")
        help_text.append(f"\n{tool_doc['description']}\n")
        
        # Parameters section
        help_text.append("## Parameters")
        parameters = tool_doc.get("parameters", {})
        for param_name, param_info in parameters.items():
            help_text.append(f"\n### {param_name}")
            help_text.append(f"- **Type**: {param_info['type']}")
            help_text.append(f"- **Required**: {param_info.get('required', False)}")
            help_text.append(f"- **Description**: {param_info['description']}")
            
            if 'default' in param_info:
                help_text.append(f"- **Default**: {param_info['default']}")
            
            if 'valid_values' in param_info:
                help_text.append(f"- **Valid Values**: {', '.join(param_info['valid_values'])}")
            
            # Schema details for complex objects
            if 'schema' in param_info:
                help_text.append("\n#### Schema:")
                help_text.append(self._format_schema(param_info['schema'], indent=1))
        
        # Examples section
        examples = tool_doc.get("examples", [])
        if examples:
            help_text.append("\n## Examples")
            for i, example in enumerate(examples, 1):
                help_text.append(f"\n### Example {i}: {example['name']}")
                help_text.append("```json")
                example_data = {k: v for k, v in example.items() if k != 'name'}
                help_text.append(json.dumps(example_data, indent=2))
                help_text.append("```")
        
        # Notes section
        notes = tool_doc.get("notes", [])
        if notes:
            help_text.append("\n## Important Notes")
            for note in notes:
                help_text.append(f"- {note}")
        
        return "\n".join(help_text)
    
    def _format_schema(self, schema: Dict[str, Any], indent: int = 0) -> str:
        """Format schema documentation recursively."""
        lines = []
        indent_str = "  " * indent
        
        for key, value in schema.items():
            if isinstance(value, dict):
                lines.append(f"{indent_str}- **{key}**: {value.get('description', 'No description')}")
                if 'type' in value:
                    lines.append(f"{indent_str}  - Type: {value['type']}")
                if 'properties' in value:
                    lines.append(f"{indent_str}  - Properties:")
                    lines.append(self._format_schema(value['properties'], indent + 2))
                if 'items' in value:
                    lines.append(f"{indent_str}  - Items: {value['items']}")
            else:
                lines.append(f"{indent_str}- **{key}**: {value}")
        
        return "\n".join(lines)


# Global instance
tool_help = ToolHelpSystem()


def get_tool_help(tool_name: str) -> str:
    """Get formatted help text for a tool."""
    return tool_help.format_help_text(tool_name)


def get_tool_examples(tool_name: str) -> List[Dict[str, Any]]:
    """Get examples for a tool."""
    return tool_help.get_examples(tool_name)


def get_parameter_help(tool_name: str, parameter_name: str) -> Optional[Dict[str, Any]]:
    """Get help for a specific parameter."""
    return tool_help.get_parameter_help(tool_name, parameter_name)