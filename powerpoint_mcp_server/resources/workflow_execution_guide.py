"""
MCP Resource for workflow execution guide with decision trees and best practices.
"""

import json
from typing import Dict, Any

WORKFLOW_EXECUTION_GUIDE = {
    "name": "workflow_execution_guide",
    "description": "Comprehensive guide for executing PowerPoint analysis workflows with decision trees and best practices",
    "uri": "powerpoint://guide/workflow",
    "mimeType": "application/json",
    "content": {
        "overview": {
            "description": "This guide provides step-by-step workflows for common PowerPoint analysis tasks",
            "workflow_types": [
                "Content Extraction",
                "Table Data Analysis",
                "Formatting Analysis",
                "Presentation Overview",
                "Complex Data Mining",
                "Quality Assessment"
            ]
        },
        
        "decision_trees": {
            "initial_assessment": {
                "description": "Determine the best approach based on user requirements",
                "decision_flow": {
                    "start": {
                        "question": "What is the primary goal?",
                        "options": {
                            "extract_all_content": "goto: full_extraction",
                            "find_specific_data": "goto: targeted_extraction",
                            "analyze_structure": "goto: structure_analysis",
                            "assess_quality": "goto: quality_assessment",
                            "understand_presentation": "goto: overview_analysis"
                        }
                    },
                    "full_extraction": {
                        "description": "Extract all content from the presentation",
                        "recommended_tools": ["extract_powerpoint_content"],
                        "parameters": {
                            "attributes": ["all"],
                            "include_formatting": true
                        },
                        "next_steps": [
                            "Review extracted content structure",
                            "Apply post-processing filters if needed",
                            "Consider formatting analysis for detailed insights"
                        ]
                    },
                    "targeted_extraction": {
                        "question": "What type of specific data?",
                        "options": {
                            "table_data": "goto: table_extraction_workflow",
                            "formatted_text": "goto: formatting_analysis_workflow",
                            "specific_slides": "goto: slide_query_workflow",
                            "presentation_metadata": "goto: metadata_extraction"
                        }
                    },
                    "structure_analysis": {
                        "description": "Analyze presentation structure and organization",
                        "recommended_tools": ["get_presentation_overview"],
                        "parameters": {
                            "analysis_depth": "detailed",
                            "include_sample_content": true
                        },
                        "focus_areas": [
                            "Slide type distribution",
                            "Content flow analysis",
                            "Structural issues identification"
                        ]
                    },
                    "quality_assessment": {
                        "description": "Assess presentation quality and provide recommendations",
                        "recommended_tools": ["get_presentation_overview"],
                        "parameters": {
                            "analysis_depth": "comprehensive",
                            "include_sample_content": false
                        },
                        "focus_areas": [
                            "Readability score",
                            "Content density analysis",
                            "Visual balance assessment",
                            "Consistency evaluation"
                        ]
                    },
                    "overview_analysis": {
                        "description": "Get comprehensive understanding of the presentation",
                        "workflow": "goto: overview_workflow"
                    }
                }
            },
            
            "table_extraction_decision": {
                "description": "Decide on table extraction approach",
                "decision_flow": {
                    "start": {
                        "question": "Do you know which slides contain tables?",
                        "options": {
                            "yes": "goto: direct_extraction",
                            "no": "goto: find_tables_first",
                            "unsure": "goto: explore_first"
                        }
                    },
                    "find_tables_first": {
                        "description": "Find slides containing tables",
                        "recommended_tools": ["query_slides"],
                        "parameters": {
                            "search_criteria": {
                                "content": {"has_tables": true}
                            },
                            "return_details": "detailed"
                        },
                        "next": "goto: direct_extraction"
                    },
                    "explore_first": {
                        "description": "Get overview to understand table distribution",
                        "recommended_tools": ["get_presentation_overview"],
                        "parameters": {
                            "analysis_depth": "basic"
                        },
                        "next": "Review slide classifications for table slides"
                    },
                    "direct_extraction": {
                        "question": "Do you need specific columns or all data?",
                        "options": {
                            "specific_columns": "goto: selective_extraction",
                            "all_data": "goto: full_table_extraction",
                            "filtered_data": "goto: criteria_based_extraction"
                        }
                    }
                }
            },
            
            "formatting_analysis_decision": {
                "description": "Decide on formatting analysis approach",
                "decision_flow": {
                    "start": {
                        "question": "What type of formatting are you looking for?",
                        "options": {
                            "specific_formatting": "goto: targeted_formatting",
                            "all_formatting": "goto: comprehensive_formatting",
                            "formatting_patterns": "goto: pattern_analysis",
                            "formatting_issues": "goto: consistency_check"
                        }
                    },
                    "targeted_formatting": {
                        "question": "Which formatting types?",
                        "options": {
                            "bold_text": {
                                "tool": "analyze_text_formatting",
                                "parameters": {
                                    "formatting_filter": {
                                        "formatting_types": ["bold"]
                                    }
                                }
                            },
                            "highlighted_text": {
                                "tool": "analyze_text_formatting",
                                "parameters": {
                                    "formatting_filter": {
                                        "formatting_types": ["highlight"]
                                    }
                                }
                            },
                            "colored_text": {
                                "tool": "analyze_text_formatting",
                                "parameters": {
                                    "formatting_filter": {
                                        "formatting_types": ["color"]
                                    }
                                }
                            }
                        }
                    }
                }
            }
        },
        
        "workflows": {
            "table_extraction_workflow": {
                "name": "Complete Table Data Extraction",
                "description": "Extract and analyze table data from PowerPoint presentations",
                "steps": [
                    {
                        "step": 1,
                        "name": "Identify Table-Containing Slides",
                        "tool": "query_slides",
                        "parameters": {
                            "search_criteria": {
                                "content": {"has_tables": true}
                            },
                            "return_details": "detailed"
                        },
                        "purpose": "Find all slides that contain tables",
                        "expected_output": "List of slides with table information"
                    },
                    {
                        "step": 2,
                        "name": "Extract Table Data",
                        "tool": "extract_table_data",
                        "parameters": {
                            "slide_numbers": "from_step_1",
                            "table_criteria": {
                                "min_rows": 2,
                                "min_columns": 2
                            },
                            "formatting_detection": {
                                "detect_bold": true,
                                "detect_highlight": true,
                                "detect_colors": true
                            },
                            "output_format": "structured"
                        },
                        "purpose": "Extract structured table data with formatting",
                        "expected_output": "Structured table data with metadata"
                    },
                    {
                        "step": 3,
                        "name": "Filter and Process Data (Optional)",
                        "tool": "filter_and_aggregate",
                        "parameters": {
                            "data_source": "from_step_2",
                            "filters": [
                                {
                                    "field": "formatting.bold",
                                    "condition": "equals",
                                    "value": true
                                }
                            ]
                        },
                        "purpose": "Apply additional filtering or aggregation",
                        "expected_output": "Processed and filtered data"
                    }
                ],
                "variations": {
                    "specific_columns_only": {
                        "modify_step": 2,
                        "add_parameters": {
                            "column_selection": {
                                "specific_columns": ["Column1", "Column2"],
                                "all_columns": false
                            }
                        }
                    },
                    "tables_with_criteria": {
                        "modify_step": 2,
                        "add_parameters": {
                            "table_criteria": {
                                "header_contains": ["Name", "Value"],
                                "min_rows": 5
                            }
                        }
                    }
                }
            },
            
            "formatting_analysis_workflow": {
                "name": "Text Formatting Analysis",
                "description": "Analyze text formatting patterns across the presentation",
                "steps": [
                    {
                        "step": 1,
                        "name": "Analyze All Text Formatting",
                        "tool": "analyze_text_formatting",
                        "parameters": {
                            "formatting_filter": {
                                "content_types": ["all"],
                                "formatting_types": ["all"]
                            },
                            "grouping": "by_formatting_type"
                        },
                        "purpose": "Get comprehensive formatting overview",
                        "expected_output": "Formatting analysis with groupings"
                    },
                    {
                        "step": 2,
                        "name": "Focus on Specific Formatting (Optional)",
                        "tool": "analyze_text_formatting",
                        "parameters": {
                            "formatting_filter": {
                                "formatting_types": ["bold", "highlight"],
                                "content_types": ["tables", "titles"]
                            },
                            "grouping": "by_slide"
                        },
                        "purpose": "Analyze specific formatting in key content",
                        "expected_output": "Targeted formatting analysis"
                    }
                ],
                "use_cases": [
                    "Find all emphasized text (bold, highlight)",
                    "Identify color-coded information",
                    "Analyze formatting consistency",
                    "Extract hyperlinked content"
                ]
            },
            
            "slide_query_workflow": {
                "name": "Targeted Slide Querying",
                "description": "Find specific slides based on content criteria",
                "steps": [
                    {
                        "step": 1,
                        "name": "Query Slides with Criteria",
                        "tool": "query_slides",
                        "parameters": {
                            "search_criteria": {
                                "title": {
                                    "contains": "target_text"
                                },
                                "content": {
                                    "has_tables": true,
                                    "object_count_min": 3
                                }
                            },
                            "return_details": "detailed",
                            "limit": 10
                        },
                        "purpose": "Find slides matching specific criteria",
                        "expected_output": "List of matching slides with details"
                    },
                    {
                        "step": 2,
                        "name": "Extract Content from Found Slides",
                        "tool": "get_powerpoint_attributes",
                        "parameters": {
                            "slide_numbers": "from_step_1",
                            "attributes": ["text_elements", "tables", "object_counts"]
                        },
                        "purpose": "Get detailed content from matching slides",
                        "expected_output": "Detailed slide content"
                    }
                ],
                "common_queries": {
                    "slides_with_specific_title_pattern": {
                        "search_criteria": {
                            "title": {
                                "regex": "Chapter \\d+.*"
                            }
                        }
                    },
                    "content_heavy_slides": {
                        "search_criteria": {
                            "content": {
                                "object_count_min": 5
                            }
                        }
                    },
                    "slides_with_tables_and_charts": {
                        "search_criteria": {
                            "content": {
                                "has_tables": true,
                                "has_charts": true
                            }
                        }
                    }
                }
            },
            
            "overview_workflow": {
                "name": "Comprehensive Presentation Overview",
                "description": "Get complete understanding of presentation structure and content",
                "steps": [
                    {
                        "step": 1,
                        "name": "Get Basic Overview",
                        "tool": "get_presentation_overview",
                        "parameters": {
                            "analysis_depth": "basic",
                            "include_sample_content": true
                        },
                        "purpose": "Understand basic structure and content",
                        "expected_output": "Basic presentation analysis"
                    },
                    {
                        "step": 2,
                        "name": "Detailed Analysis (If Needed)",
                        "tool": "get_presentation_overview",
                        "parameters": {
                            "analysis_depth": "comprehensive",
                            "include_sample_content": false
                        },
                        "purpose": "Get detailed insights and recommendations",
                        "expected_output": "Comprehensive analysis with recommendations",
                        "condition": "If detailed analysis is needed"
                    }
                ],
                "interpretation_guide": {
                    "slide_classifications": {
                        "title_slide": "Introduction or section dividers",
                        "content_slide": "Main content slides",
                        "table_slide": "Data-heavy slides",
                        "chart_slide": "Visual data presentation",
                        "bullet_slide": "Text-heavy informational slides"
                    },
                    "insights": {
                        "readability_score": {
                            "0-3": "Poor readability, too much content",
                            "4-6": "Moderate readability, some improvements needed",
                            "7-8": "Good readability",
                            "9-10": "Excellent readability"
                        },
                        "content_density": {
                            "low": "Slides may be too sparse",
                            "medium": "Good balance",
                            "high": "Slides may be too crowded"
                        }
                    }
                }
            },
            
            "complex_data_mining_workflow": {
                "name": "Complex Data Mining and Analysis",
                "description": "Advanced workflow for extracting and analyzing complex data patterns",
                "steps": [
                    {
                        "step": 1,
                        "name": "Presentation Overview",
                        "tool": "get_presentation_overview",
                        "parameters": {
                            "analysis_depth": "detailed"
                        },
                        "purpose": "Understand overall structure and identify data-rich areas"
                    },
                    {
                        "step": 2,
                        "name": "Identify Data Sources",
                        "tool": "query_slides",
                        "parameters": {
                            "search_criteria": {
                                "content": {
                                    "has_tables": true
                                }
                            }
                        },
                        "purpose": "Find all slides with tabular data"
                    },
                    {
                        "step": 3,
                        "name": "Extract Table Data",
                        "tool": "extract_table_data",
                        "parameters": {
                            "slide_numbers": "from_step_2",
                            "formatting_detection": {
                                "detect_bold": true,
                                "detect_highlight": true,
                                "detect_colors": true
                            }
                        },
                        "purpose": "Extract all table data with formatting"
                    },
                    {
                        "step": 4,
                        "name": "Analyze Formatting Patterns",
                        "tool": "analyze_text_formatting",
                        "parameters": {
                            "slide_numbers": "from_step_2",
                            "formatting_filter": {
                                "content_types": ["tables"],
                                "formatting_types": ["bold", "highlight", "color"]
                            },
                            "grouping": "by_color"
                        },
                        "purpose": "Identify formatting patterns that might indicate data significance"
                    },
                    {
                        "step": 5,
                        "name": "Filter and Aggregate Data",
                        "tool": "filter_and_aggregate",
                        "parameters": {
                            "data_source": "from_step_3",
                            "filters": [
                                {
                                    "field": "formatting.highlight",
                                    "condition": "equals",
                                    "value": true
                                }
                            ],
                            "grouping": {
                                "fields": ["slide_number"],
                                "aggregations": [
                                    {
                                        "field": "value",
                                        "operation": "list",
                                        "output_field": "highlighted_values"
                                    }
                                ]
                            }
                        },
                        "purpose": "Focus on highlighted/important data points"
                    }
                ]
            }
        },
        
        "error_handling": {
            "common_errors": {
                "file_not_found": {
                    "error": "File not found or inaccessible",
                    "solutions": [
                        "Verify file path is correct",
                        "Check file permissions",
                        "Ensure file exists and is not locked"
                    ]
                },
                "unsupported_format": {
                    "error": "File format not supported",
                    "solutions": [
                        "Convert .ppt files to .pptx format",
                        "Ensure file is a valid PowerPoint presentation",
                        "Check file is not corrupted"
                    ]
                },
                "no_matching_slides": {
                    "error": "No slides match the specified criteria",
                    "solutions": [
                        "Broaden search criteria",
                        "Check slide content manually",
                        "Use get_presentation_overview to understand structure",
                        "Try different filter combinations"
                    ]
                },
                "empty_tables": {
                    "error": "Tables found but no data extracted",
                    "solutions": [
                        "Check table_criteria parameters",
                        "Verify tables have actual content",
                        "Adjust min_rows and min_columns criteria",
                        "Check if tables are actually text boxes"
                    ]
                },
                "formatting_not_detected": {
                    "error": "Expected formatting not found",
                    "solutions": [
                        "Verify formatting_detection parameters are enabled",
                        "Check if formatting exists in the source file",
                        "Try different content_types in the filter",
                        "Use comprehensive analysis first"
                    ]
                }
            },
            
            "debugging_strategies": {
                "start_simple": [
                    "Begin with get_presentation_overview",
                    "Use basic parameters first",
                    "Gradually add complexity"
                ],
                "verify_data": [
                    "Check slide_count in overview",
                    "Verify slide_numbers exist",
                    "Confirm expected content types are present"
                ],
                "incremental_filtering": [
                    "Start with broad criteria",
                    "Add filters one at a time",
                    "Test each filter independently"
                ]
            }
        },
        
        "performance_optimization": {
            "best_practices": [
                "Use specific slide_numbers when possible",
                "Limit return_details to what you need",
                "Use filters to reduce data volume",
                "Cache results for repeated operations",
                "Process large presentations in batches"
            ],
            "parameter_optimization": {
                "query_slides": {
                    "limit": "Set reasonable limit (default 50)",
                    "return_details": "Use 'basic' unless detailed info needed"
                },
                "extract_table_data": {
                    "formatting_detection": "Disable unused formatting detection",
                    "output_format": "Use 'flat' for simple processing"
                },
                "analyze_text_formatting": {
                    "content_types": "Specify only needed content types",
                    "slide_numbers": "Limit to relevant slides"
                }
            }
        },
        
        "integration_patterns": {
            "chaining_tools": {
                "description": "How to chain multiple tools effectively",
                "patterns": [
                    "Overview -> Query -> Extract -> Filter",
                    "Query -> Extract -> Analyze -> Aggregate",
                    "Overview -> Targeted Analysis -> Detailed Extraction"
                ]
            },
            "data_flow": {
                "slide_numbers": "Pass slide numbers between tools",
                "extracted_data": "Use extracted data as input for filtering",
                "analysis_results": "Use analysis to guide further extraction"
            },
            "result_interpretation": {
                "empty_results": "May indicate wrong criteria or no matching content",
                "large_results": "Consider adding filters or pagination",
                "unexpected_structure": "Verify file format and content expectations"
            }
        }
    }
}


def get_workflow_execution_guide() -> Dict[str, Any]:
    """Get the workflow execution guide resource."""
    return WORKFLOW_EXECUTION_GUIDE