"""
MCP Resource for PowerPoint extraction capabilities documentation.
"""

import json
from typing import Dict, Any

POWERPOINT_EXTRACTION_CAPABILITIES = {
    "name": "powerpoint_extraction_capabilities",
    "description": "Complete reference for PowerPoint Analyzer MCP extraction capabilities and attributes",
    "uri": "powerpoint://capabilities/extraction",
    "mimeType": "application/json",
    "content": {
        "overview": {
            "description": "The PowerPoint Analyzer MCP provides comprehensive extraction capabilities for PowerPoint presentations (.pptx files only)",
            "supported_formats": [".pptx"],
            "unsupported_formats": [".ppt", ".odp"],
            "key_features": [
                "Slide content extraction with formatting",
                "Table data extraction with cell-level formatting",
                "Text formatting analysis",
                "Flexible slide querying and filtering",
                "Presentation structure analysis",
                "Data aggregation and post-processing"
            ]
        },
        
        "extraction_attributes": {
            "slide_level": {
                "basic_info": {
                    "slide_number": {
                        "type": "integer",
                        "description": "1-based slide number",
                        "example": 1
                    },
                    "layout_name": {
                        "type": "string",
                        "description": "Name of the slide layout",
                        "example": "Title and Content"
                    },
                    "layout_type": {
                        "type": "string",
                        "description": "Type of layout (title, content, etc.)",
                        "example": "titleAndContent"
                    },
                    "title": {
                        "type": "string",
                        "description": "Slide title text",
                        "example": "Project Overview"
                    },
                    "subtitle": {
                        "type": "string",
                        "description": "Slide subtitle text",
                        "example": "Q4 2024 Progress"
                    }
                },
                
                "content_elements": {
                    "text_elements": {
                        "type": "array",
                        "description": "All text elements on the slide with formatting",
                        "structure": {
                            "content_plain": {
                                "type": "string",
                                "description": "Plain text without formatting tags"
                            },
                            "content_formatted": {
                                "type": "string",
                                "description": "Text with HTML-like formatting tags"
                            },
                            "font_sizes": {
                                "type": "array",
                                "description": "List of font sizes used in this element"
                            },
                            "font_colors": {
                                "type": "array",
                                "description": "List of font colors used (hex codes or scheme names)"
                            },
                            "hyperlinks": {
                                "type": "array",
                                "description": "List of hyperlink URLs or relationship IDs"
                            },
                            "formatting_counts": {
                                "bolded": {"type": "integer", "description": "Number of bold text runs"},
                                "italic": {"type": "integer", "description": "Number of italic text runs"},
                                "underlined": {"type": "integer", "description": "Number of underlined text runs"},
                                "highlighted": {"type": "integer", "description": "Number of highlighted text runs"},
                                "strikethrough": {"type": "integer", "description": "Number of strikethrough text runs"}
                            },
                            "position": {
                                "type": "array",
                                "description": "Position as [x, y] coordinates in EMUs"
                            },
                            "size": {
                                "type": "array",
                                "description": "Size as [width, height] in EMUs"
                            }
                        }
                    },
                    
                    "tables": {
                        "type": "array",
                        "description": "Table data with structure and formatting",
                        "structure": {
                            "rows": {"type": "integer", "description": "Number of rows"},
                            "columns": {"type": "integer", "description": "Number of columns"},
                            "cells": {
                                "type": "array",
                                "description": "2D array of cell data",
                                "cell_structure": {
                                    "content": {"type": "string", "description": "Cell text content"},
                                    "row_span": {"type": "integer", "description": "Number of rows spanned"},
                                    "col_span": {"type": "integer", "description": "Number of columns spanned"},
                                    "formatting": {
                                        "type": "object",
                                        "description": "Cell formatting information",
                                        "properties": {
                                            "fill_color": {"type": "string", "description": "Background color"},
                                            "borders": {"type": "object", "description": "Border information"}
                                        }
                                    }
                                }
                            },
                            "position": {"type": "array", "description": "Table position [x, y]"},
                            "size": {"type": "array", "description": "Table size [width, height]"}
                        }
                    },
                    
                    "placeholders": {
                        "type": "array",
                        "description": "Slide placeholders information",
                        "structure": {
                            "type": {"type": "string", "description": "Placeholder type (title, content, etc.)"},
                            "position": {"type": "array", "description": "Position [x, y]"},
                            "size": {"type": "array", "description": "Size [width, height]"},
                            "content": {"type": "string", "description": "Placeholder content if any"}
                        }
                    }
                },
                
                "metadata": {
                    "notes": {
                        "type": "string",
                        "description": "Speaker notes content for the slide"
                    },
                    "object_counts": {
                        "type": "object",
                        "description": "Count of different object types on the slide",
                        "properties": {
                            "shapes": {"type": "integer", "description": "Number of shapes"},
                            "text_boxes": {"type": "integer", "description": "Number of text boxes"},
                            "images": {"type": "integer", "description": "Number of images"},
                            "tables": {"type": "integer", "description": "Number of tables"},
                            "charts": {"type": "integer", "description": "Number of charts"},
                            "media": {"type": "integer", "description": "Number of media objects"},
                            "connectors": {"type": "integer", "description": "Number of connectors"},
                            "groups": {"type": "integer", "description": "Number of grouped objects"}
                        }
                    }
                }
            },
            
            "presentation_level": {
                "metadata": {
                    "slide_count": {"type": "integer", "description": "Total number of slides"},
                    "slide_size": {
                        "type": "object",
                        "description": "Slide dimensions in various units",
                        "properties": {
                            "width_emu": {"type": "integer", "description": "Width in EMUs"},
                            "height_emu": {"type": "integer", "description": "Height in EMUs"},
                            "width_inches": {"type": "number", "description": "Width in inches"},
                            "height_inches": {"type": "number", "description": "Height in inches"},
                            "width_cm": {"type": "number", "description": "Width in centimeters"},
                            "height_cm": {"type": "number", "description": "Height in centimeters"},
                            "aspect_ratio": {"type": "number", "description": "Width/height ratio"}
                        }
                    },
                    "has_sections": {"type": "boolean", "description": "Whether presentation has sections"},
                    "sections": {
                        "type": "array",
                        "description": "Section information",
                        "structure": {
                            "name": {"type": "string", "description": "Section name"},
                            "id": {"type": "string", "description": "Section ID"}
                        }
                    }
                }
            }
        },
        
        "filtering_capabilities": {
            "slide_queries": {
                "title_filters": {
                    "contains": "Filter slides where title contains specific text",
                    "starts_with": "Filter slides where title starts with specific text",
                    "ends_with": "Filter slides where title ends with specific text",
                    "regex": "Filter slides using regular expression patterns",
                    "one_of": "Filter slides where title matches any of the provided patterns"
                },
                "content_filters": {
                    "has_tables": "Filter slides that contain tables",
                    "has_charts": "Filter slides that contain charts",
                    "has_images": "Filter slides that contain images",
                    "object_count_min": "Filter slides with minimum object count",
                    "object_count_max": "Filter slides with maximum object count",
                    "contains_text": "Filter slides containing specific text in any element"
                },
                "layout_filters": {
                    "layout_type": "Filter by layout type",
                    "layout_name": "Filter by layout name"
                },
                "other_filters": {
                    "slide_numbers": "Filter specific slide numbers",
                    "section": "Filter slides in specific section"
                }
            },
            
            "table_extraction": {
                "table_criteria": {
                    "min_rows": "Minimum number of rows",
                    "min_columns": "Minimum number of columns",
                    "max_rows": "Maximum number of rows",
                    "max_columns": "Maximum number of columns",
                    "header_contains": "Headers must contain specific text",
                    "header_patterns": "Headers must match regex patterns"
                },
                "column_selection": {
                    "specific_columns": "Extract only specified columns by name",
                    "column_patterns": "Extract columns matching regex patterns",
                    "exclude_columns": "Exclude specific columns",
                    "all_columns": "Extract all columns (default)"
                },
                "formatting_detection": {
                    "detect_bold": "Detect bold text in cells",
                    "detect_italic": "Detect italic text in cells",
                    "detect_underline": "Detect underlined text in cells",
                    "detect_highlight": "Detect highlighted text in cells",
                    "detect_colors": "Detect text and background colors",
                    "detect_hyperlinks": "Detect hyperlinks in cells"
                }
            },
            
            "text_formatting_analysis": {
                "content_types": {
                    "tables": "Analyze formatting in table text",
                    "text_boxes": "Analyze formatting in text boxes",
                    "titles": "Analyze formatting in slide titles",
                    "bullets": "Analyze formatting in bullet points",
                    "all": "Analyze all text content"
                },
                "formatting_types": {
                    "bold": "Bold text formatting",
                    "italic": "Italic text formatting",
                    "underline": "Underlined text formatting",
                    "highlight": "Highlighted text formatting",
                    "strikethrough": "Strikethrough text formatting",
                    "color": "Text color formatting",
                    "font_size": "Font size information",
                    "hyperlink": "Hyperlink formatting"
                },
                "grouping_options": {
                    "by_slide": "Group results by slide number",
                    "by_formatting_type": "Group results by formatting type",
                    "by_content_type": "Group results by content type",
                    "by_color": "Group results by text color",
                    "by_font_size": "Group results by font size"
                }
            },
            
            "data_filtering": {
                "filter_conditions": {
                    "equals": "Exact match",
                    "not_equals": "Not equal to",
                    "contains": "Contains text",
                    "not_contains": "Does not contain text",
                    "starts_with": "Starts with text",
                    "ends_with": "Ends with text",
                    "regex": "Regular expression match",
                    "not_empty": "Field is not empty",
                    "is_empty": "Field is empty",
                    "has_formatting": "Has specific formatting",
                    "no_formatting": "Has no formatting",
                    "greater_than": "Numeric greater than",
                    "less_than": "Numeric less than",
                    "in_list": "Value in provided list",
                    "not_in_list": "Value not in provided list"
                },
                "aggregation_operations": {
                    "count": "Count of items",
                    "list": "List of all values",
                    "unique": "List of unique values",
                    "concat": "Concatenate values with separator",
                    "sum": "Sum of numeric values",
                    "average": "Average of numeric values",
                    "min": "Minimum value",
                    "max": "Maximum value",
                    "first": "First value",
                    "last": "Last value",
                    "most_common": "Most frequently occurring value",
                    "least_common": "Least frequently occurring value"
                }
            }
        },
        
        "output_formats": {
            "slide_queries": {
                "basic": "Slide number, title, and object counts",
                "detailed": "Basic info plus layout, preview text, and table info",
                "full": "Complete slide content including all elements"
            },
            "table_extraction": {
                "structured": "Hierarchical structure with metadata",
                "flat": "Flattened rows with slide/table identifiers",
                "grouped_by_slide": "Tables grouped by slide number"
            },
            "text_formatting": {
                "elements": "List of formatted text elements",
                "summary": "Statistical summary of formatting usage",
                "grouped": "Elements grouped by specified criteria"
            }
        },
        
        "common_use_cases": {
            "content_extraction": {
                "description": "Extract all text content from slides",
                "tools": ["extract_powerpoint_content", "get_powerpoint_attributes"],
                "attributes": ["text_elements", "title", "subtitle"]
            },
            "table_data_mining": {
                "description": "Extract and analyze table data",
                "tools": ["extract_table_data", "filter_and_aggregate"],
                "workflow": [
                    "Use query_slides to find slides with tables",
                    "Use extract_table_data to get table content",
                    "Use filter_and_aggregate to process data"
                ]
            },
            "formatting_analysis": {
                "description": "Analyze text formatting patterns",
                "tools": ["analyze_text_formatting"],
                "use_cases": [
                    "Find all bold text across presentation",
                    "Identify highlighted important information",
                    "Analyze color usage patterns"
                ]
            },
            "presentation_overview": {
                "description": "Get comprehensive presentation analysis",
                "tools": ["get_presentation_overview"],
                "includes": [
                    "Slide type classification",
                    "Content patterns detection",
                    "Structure analysis",
                    "Recommendations for improvement"
                ]
            }
        },
        
        "best_practices": {
            "performance": [
                "Use specific slide numbers when possible to limit processing",
                "Use return_fields parameter to get only needed data",
                "Cache results when processing the same file multiple times",
                "Use filters to reduce data volume before processing"
            ],
            "accuracy": [
                "Verify file format is .pptx before processing",
                "Handle empty or corrupted slides gracefully",
                "Use multiple filter criteria for precise results",
                "Validate extracted data structure before processing"
            ],
            "workflow": [
                "Start with presentation overview for understanding",
                "Use slide queries to identify relevant slides",
                "Extract specific content with targeted tools",
                "Apply post-processing filters as needed"
            ]
        },
        
        "limitations": {
            "file_formats": [
                "Only .pptx files are supported",
                ".ppt files require conversion first",
                "Password-protected files are not supported"
            ],
            "content_types": [
                "Embedded videos are detected but content not extracted",
                "Audio files are detected but content not extracted",
                "Complex animations are not analyzed",
                "Custom fonts may not be accurately detected"
            ],
            "formatting": [
                "Some advanced formatting may not be captured",
                "Theme-based colors are reported as scheme names",
                "Complex text effects may be simplified",
                "Gradient fills are not fully supported"
            ]
        }
    }
}


def get_powerpoint_extraction_capabilities() -> Dict[str, Any]:
    """Get the PowerPoint extraction capabilities resource."""
    return POWERPOINT_EXTRACTION_CAPABILITIES