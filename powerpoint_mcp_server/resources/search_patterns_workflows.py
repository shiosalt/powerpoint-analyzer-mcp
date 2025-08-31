"""
MCP Resource for search patterns and common workflows with practical examples.
"""

import json
from typing import Dict, Any

SEARCH_PATTERNS_WORKFLOWS = {
    "name": "search_patterns_workflows",
    "description": "Collection of search patterns and common workflows for PowerPoint analysis",
    "uri": "powerpoint://patterns/search",
    "mimeType": "application/json",
    "content": {
        "overview": {
            "description": "Pre-built search patterns and workflows for common PowerPoint analysis tasks",
            "categories": [
                "Title Patterns",
                "Content Patterns", 
                "Formatting Patterns",
                "Data Extraction Patterns",
                "Quality Assessment Patterns"
            ]
        },
        
        "title_patterns": {
            "numbered_sections": {
                "description": "Find slides with numbered titles (1., 2., Chapter 1, etc.)",
                "pattern": "^\\d+[\\.\\.\\s].*",
                "tool": "query_slides",
                "parameters": {
                    "search_criteria": {
                        "title": {
                            "regex": "^\\d+[\\.\\.\\s].*"
                        }
                    }
                },
                "examples": [
                    "1. Introduction",
                    "2. Main Content", 
                    "Chapter 1: Overview",
                    "Section 3.1 Details"
                ]
            },
            
            "question_titles": {
                "description": "Find slides with question titles",
                "pattern": ".*\\?$",
                "tool": "query_slides",
                "parameters": {
                    "search_criteria": {
                        "title": {
                            "regex": ".*\\?$"
                        }
                    }
                },
                "examples": [
                    "What is our goal?",
                    "How do we proceed?",
                    "Why is this important?"
                ]
            },
            
            "agenda_slides": {
                "description": "Find agenda, outline, or overview slides",
                "patterns": ["agenda", "outline", "overview", "contents", "topics"],
                "tool": "query_slides",
                "parameters": {
                    "search_criteria": {
                        "title": {
                            "one_of": [".*agenda.*", ".*outline.*", ".*overview.*", ".*contents.*", ".*topics.*"]
                        }
                    }
                },
                "case_insensitive": true
            },
            
            "conclusion_slides": {
                "description": "Find conclusion, summary, or wrap-up slides",
                "patterns": ["conclusion", "summary", "wrap-up", "takeaways", "next steps"],
                "tool": "query_slides",
                "parameters": {
                    "search_criteria": {
                        "title": {
                            "one_of": [".*conclusion.*", ".*summary.*", ".*wrap.*up.*", ".*takeaways.*", ".*next.*steps.*"]
                        }
                    }
                }
            },
            
            "project_phases": {
                "description": "Find slides related to project phases",
                "patterns": ["phase", "stage", "milestone", "sprint", "iteration"],
                "tool": "query_slides",
                "parameters": {
                    "search_criteria": {
                        "title": {
                            "one_of": [".*phase.*", ".*stage.*", ".*milestone.*", ".*sprint.*", ".*iteration.*"]
                        }
                    }
                }
            }
        },
        
        "content_patterns": {
            "data_heavy_slides": {
                "description": "Find slides with lots of data (tables, charts, high object count)",
                "tool": "query_slides",
                "parameters": {
                    "search_criteria": {
                        "content": {
                            "has_tables": true,
                            "object_count_min": 5
                        }
                    }
                },
                "follow_up": {
                    "tool": "extract_table_data",
                    "purpose": "Extract the actual data from these slides"
                }
            },
            
            "visual_slides": {
                "description": "Find slides with images, charts, or visual content",
                "tool": "query_slides",
                "parameters": {
                    "search_criteria": {
                        "content": {
                            "has_images": true,
                            "has_charts": true
                        }
                    }
                },
                "filter_logic": "OR"
            },
            
            "text_heavy_slides": {
                "description": "Find slides with lots of text content",
                "tool": "query_slides",
                "parameters": {
                    "search_criteria": {
                        "content": {
                            "object_count_min": 3,
                            "has_tables": false,
                            "has_charts": false,
                            "has_images": false
                        }
                    }
                }
            },
            
            "minimal_content_slides": {
                "description": "Find slides with very little content (potential title slides or breaks)",
                "tool": "query_slides",
                "parameters": {
                    "search_criteria": {
                        "content": {
                            "object_count_max": 2
                        }
                    }
                }
            }
        },
        
        "formatting_patterns": {
            "highlighted_important_info": {
                "description": "Find all highlighted text across the presentation",
                "tool": "analyze_text_formatting",
                "parameters": {
                    "formatting_filter": {
                        "formatting_types": ["highlight"],
                        "content_types": ["all"]
                    },
                    "grouping": "by_slide"
                },
                "use_case": "Identify key points or important information"
            },
            
            "bold_headings_and_emphasis": {
                "description": "Find all bold text for headings and emphasis",
                "tool": "analyze_text_formatting",
                "parameters": {
                    "formatting_filter": {
                        "formatting_types": ["bold"],
                        "content_types": ["titles", "text_boxes"]
                    },
                    "grouping": "by_content_type"
                }
            },
            
            "color_coded_information": {
                "description": "Analyze color usage patterns",
                "tool": "analyze_text_formatting",
                "parameters": {
                    "formatting_filter": {
                        "formatting_types": ["color"],
                        "content_types": ["all"]
                    },
                    "grouping": "by_color"
                },
                "interpretation": {
                    "red": "Often used for warnings, errors, or critical information",
                    "green": "Often used for success, positive results, or go signals",
                    "blue": "Often used for information, links, or neutral content",
                    "yellow": "Often used for caution, attention, or highlights"
                }
            },
            
            "hyperlinked_content": {
                "description": "Find all hyperlinks in the presentation",
                "tool": "analyze_text_formatting",
                "parameters": {
                    "formatting_filter": {
                        "formatting_types": ["hyperlink"],
                        "content_types": ["all"]
                    }
                },
                "follow_up": "Extract actual URLs if needed for link analysis"
            }
        },
        
        "data_extraction_patterns": {
            "financial_data_tables": {
                "description": "Extract financial data from tables",
                "workflow": [
                    {
                        "step": 1,
                        "tool": "query_slides",
                        "parameters": {
                            "search_criteria": {
                                "title": {
                                    "one_of": [".*financial.*", ".*budget.*", ".*revenue.*", ".*cost.*", ".*profit.*"]
                                },
                                "content": {
                                    "has_tables": true
                                }
                            }
                        }
                    },
                    {
                        "step": 2,
                        "tool": "extract_table_data",
                        "parameters": {
                            "slide_numbers": "from_step_1",
                            "table_criteria": {
                                "header_contains": ["amount", "value", "cost", "revenue", "$", "€", "£"]
                            },
                            "formatting_detection": {
                                "detect_colors": true,
                                "detect_bold": true
                            }
                        }
                    }
                ]
            },
            
            "project_status_data": {
                "description": "Extract project status information",
                "workflow": [
                    {
                        "step": 1,
                        "tool": "query_slides",
                        "parameters": {
                            "search_criteria": {
                                "title": {
                                    "one_of": [".*status.*", ".*progress.*", ".*update.*", ".*dashboard.*"]
                                },
                                "content": {
                                    "has_tables": true
                                }
                            }
                        }
                    },
                    {
                        "step": 2,
                        "tool": "extract_table_data",
                        "parameters": {
                            "slide_numbers": "from_step_1",
                            "table_criteria": {
                                "header_contains": ["status", "progress", "complete", "%", "task", "milestone"]
                            },
                            "formatting_detection": {
                                "detect_colors": true,
                                "detect_highlight": true
                            }
                        }
                    },
                    {
                        "step": 3,
                        "tool": "filter_and_aggregate",
                        "parameters": {
                            "data_source": "from_step_2",
                            "filters": [
                                {
                                    "field": "formatting.highlight",
                                    "condition": "equals",
                                    "value": true
                                }
                            ]
                        },
                        "purpose": "Focus on highlighted status items (often issues or important updates)"
                    }
                ]
            },
            
            "contact_information": {
                "description": "Extract contact information from slides",
                "workflow": [
                    {
                        "step": 1,
                        "tool": "analyze_text_formatting",
                        "parameters": {
                            "formatting_filter": {
                                "text_patterns": [".*@.*", ".*\\d{3}.*\\d{3}.*\\d{4}.*", ".*www\\..*", ".*http.*"]
                            }
                        }
                    }
                ],
                "patterns": {
                    "email": ".*@.*\\.(com|org|net|edu|gov).*",
                    "phone": ".*\\d{3}[\\-\\s\\.]\\d{3}[\\-\\s\\.]\\d{4}.*",
                    "website": ".*(www\\.|http).*"
                }
            },
            
            "metrics_and_kpis": {
                "description": "Extract metrics and KPI data",
                "workflow": [
                    {
                        "step": 1,
                        "tool": "query_slides",
                        "parameters": {
                            "search_criteria": {
                                "title": {
                                    "one_of": [".*metrics.*", ".*kpi.*", ".*performance.*", ".*results.*"]
                                }
                            }
                        }
                    },
                    {
                        "step": 2,
                        "tool": "extract_table_data",
                        "parameters": {
                            "slide_numbers": "from_step_1",
                            "table_criteria": {
                                "header_patterns": [".*%.*", ".*rate.*", ".*score.*", ".*index.*"]
                            }
                        }
                    }
                ]
            }
        },
        
        "quality_assessment_patterns": {
            "consistency_check": {
                "description": "Check for consistency issues across the presentation",
                "workflow": [
                    {
                        "step": 1,
                        "tool": "get_presentation_overview",
                        "parameters": {
                            "analysis_depth": "comprehensive"
                        }
                    }
                ],
                "focus_areas": [
                    "Title usage consistency",
                    "Content density variation",
                    "Formatting consistency",
                    "Structural issues"
                ]
            },
            
            "readability_assessment": {
                "description": "Assess presentation readability",
                "workflow": [
                    {
                        "step": 1,
                        "tool": "get_presentation_overview",
                        "parameters": {
                            "analysis_depth": "comprehensive"
                        }
                    },
                    {
                        "step": 2,
                        "tool": "analyze_text_formatting",
                        "parameters": {
                            "formatting_filter": {
                                "content_types": ["all"]
                            }
                        }
                    }
                ],
                "metrics": [
                    "Readability score",
                    "Content density",
                    "Text-to-visual ratio",
                    "Formatting complexity"
                ]
            },
            
            "content_balance_analysis": {
                "description": "Analyze balance between different content types",
                "tool": "get_presentation_overview",
                "parameters": {
                    "analysis_depth": "detailed"
                },
                "interpretation": {
                    "slide_type_distribution": "Should have variety in slide types",
                    "visual_balance": "Should balance text and visual content",
                    "content_flow": "Should have logical progression"
                }
            }
        },
        
        "advanced_workflows": {
            "competitive_analysis_extraction": {
                "description": "Extract competitive analysis data",
                "steps": [
                    {
                        "step": 1,
                        "description": "Find competitive analysis slides",
                        "tool": "query_slides",
                        "parameters": {
                            "search_criteria": {
                                "title": {
                                    "one_of": [".*competitor.*", ".*competition.*", ".*vs.*", ".*comparison.*"]
                                }
                            }
                        }
                    },
                    {
                        "step": 2,
                        "description": "Extract comparison tables",
                        "tool": "extract_table_data",
                        "parameters": {
                            "slide_numbers": "from_step_1",
                            "table_criteria": {
                                "min_columns": 3,
                                "header_contains": ["company", "product", "feature", "price"]
                            }
                        }
                    },
                    {
                        "step": 3,
                        "description": "Analyze competitive positioning",
                        "tool": "analyze_text_formatting",
                        "parameters": {
                            "slide_numbers": "from_step_1",
                            "formatting_filter": {
                                "formatting_types": ["bold", "highlight", "color"]
                            }
                        }
                    }
                ]
            },
            
            "risk_assessment_extraction": {
                "description": "Extract risk assessment information",
                "steps": [
                    {
                        "step": 1,
                        "tool": "query_slides",
                        "parameters": {
                            "search_criteria": {
                                "title": {
                                    "one_of": [".*risk.*", ".*issue.*", ".*challenge.*", ".*mitigation.*"]
                                }
                            }
                        }
                    },
                    {
                        "step": 2,
                        "tool": "analyze_text_formatting",
                        "parameters": {
                            "slide_numbers": "from_step_1",
                            "formatting_filter": {
                                "formatting_types": ["color", "highlight"],
                                "colors": ["#FF0000", "#FFA500", "#FFFF00"]  # Red, Orange, Yellow
                            }
                        }
                    },
                    {
                        "step": 3,
                        "tool": "extract_table_data",
                        "parameters": {
                            "slide_numbers": "from_step_1",
                            "table_criteria": {
                                "header_contains": ["risk", "impact", "probability", "mitigation"]
                            }
                        }
                    }
                ]
            },
            
            "timeline_extraction": {
                "description": "Extract timeline and milestone information",
                "steps": [
                    {
                        "step": 1,
                        "tool": "query_slides",
                        "parameters": {
                            "search_criteria": {
                                "title": {
                                    "one_of": [".*timeline.*", ".*schedule.*", ".*roadmap.*", ".*milestone.*"]
                                }
                            }
                        }
                    },
                    {
                        "step": 2,
                        "tool": "extract_table_data",
                        "parameters": {
                            "slide_numbers": "from_step_1",
                            "table_criteria": {
                                "header_patterns": [".*date.*", ".*month.*", ".*quarter.*", ".*year.*", ".*deadline.*"]
                            }
                        }
                    },
                    {
                        "step": 3,
                        "tool": "analyze_text_formatting",
                        "parameters": {
                            "slide_numbers": "from_step_1",
                            "formatting_filter": {
                                "text_patterns": ["\\d{1,2}/\\d{1,2}/\\d{4}", "\\d{4}-\\d{2}-\\d{2}", "Q[1-4]", "Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec"]
                            }
                        }
                    }
                ]
            }
        },
        
        "troubleshooting_patterns": {
            "no_results_debugging": {
                "description": "Steps to take when queries return no results",
                "steps": [
                    {
                        "step": 1,
                        "description": "Get basic presentation info",
                        "tool": "get_presentation_overview",
                        "parameters": {
                            "analysis_depth": "basic"
                        },
                        "purpose": "Understand what content actually exists"
                    },
                    {
                        "step": 2,
                        "description": "Try broader search",
                        "tool": "query_slides",
                        "parameters": {
                            "search_criteria": {},
                            "return_details": "detailed",
                            "limit": 5
                        },
                        "purpose": "See sample of actual slide content"
                    },
                    {
                        "step": 3,
                        "description": "Check specific slide",
                        "tool": "get_powerpoint_attributes",
                        "parameters": {
                            "slide_numbers": [1],
                            "attributes": ["title", "text_elements", "object_counts"]
                        },
                        "purpose": "Verify extraction is working"
                    }
                ]
            },
            
            "unexpected_results_analysis": {
                "description": "Analyze unexpected or confusing results",
                "steps": [
                    {
                        "step": 1,
                        "description": "Examine slide structure",
                        "tool": "get_powerpoint_attributes",
                        "parameters": {
                            "slide_numbers": "problematic_slides",
                            "attributes": ["all"]
                        }
                    },
                    {
                        "step": 2,
                        "description": "Check formatting details",
                        "tool": "analyze_text_formatting",
                        "parameters": {
                            "slide_numbers": "problematic_slides",
                            "formatting_filter": {
                                "content_types": ["all"],
                                "formatting_types": ["all"]
                            }
                        }
                    }
                ]
            }
        },
        
        "performance_patterns": {
            "large_presentation_processing": {
                "description": "Efficiently process large presentations",
                "strategy": [
                    "Start with overview to understand structure",
                    "Use targeted queries instead of full extraction",
                    "Process in batches using slide_numbers",
                    "Use specific return_fields to limit data"
                ],
                "example_workflow": [
                    {
                        "step": 1,
                        "tool": "get_presentation_overview",
                        "parameters": {
                            "analysis_depth": "basic"
                        }
                    },
                    {
                        "step": 2,
                        "tool": "query_slides",
                        "parameters": {
                            "search_criteria": "specific_criteria",
                            "limit": 20
                        }
                    },
                    {
                        "step": 3,
                        "tool": "extract_table_data",
                        "parameters": {
                            "slide_numbers": "from_step_2",
                            "output_format": "flat"
                        }
                    }
                ]
            },
            
            "batch_processing_pattern": {
                "description": "Process slides in batches for better performance",
                "batch_size": 10,
                "example": {
                    "total_slides": 50,
                    "batches": [
                        {"slides": [1, 2, 3, 4, 5, 6, 7, 8, 9, 10]},
                        {"slides": [11, 12, 13, 14, 15, 16, 17, 18, 19, 20]},
                        {"slides": [21, 22, 23, 24, 25, 26, 27, 28, 29, 30]},
                        {"slides": [31, 32, 33, 34, 35, 36, 37, 38, 39, 40]},
                        {"slides": [41, 42, 43, 44, 45, 46, 47, 48, 49, 50]}
                    ]
                }
            }
        }
    }
}


def get_search_patterns_workflows() -> Dict[str, Any]:
    """Get the search patterns and workflows resource."""
    return SEARCH_PATTERNS_WORKFLOWS