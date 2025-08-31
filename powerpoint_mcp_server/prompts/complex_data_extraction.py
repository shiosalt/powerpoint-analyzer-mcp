"""
MCP Prompt for complex data extraction workflows.
"""

from typing import Dict, Any, List, Optional

COMPLEX_DATA_EXTRACTION_PROMPT = {
    "name": "complex_data_extraction",
    "description": "Automated workflow for extracting complex data patterns from PowerPoint presentations",
    "arguments": [
        {
            "name": "file_path",
            "description": "Path to the PowerPoint file to analyze",
            "required": True
        },
        {
            "name": "data_types",
            "description": "Types of data to extract (tables, formatted_text, metrics, financial, project_status)",
            "required": False
        },
        {
            "name": "search_criteria",
            "description": "Specific search criteria for targeting relevant slides",
            "required": False
        },
        {
            "name": "output_format",
            "description": "Desired output format (structured, flat, summary)",
            "required": False
        }
    ],
    "template": """
You are an expert at extracting complex data from PowerPoint presentations. Your task is to analyze the presentation at {file_path} and extract meaningful data patterns.

## Analysis Approach

1. **Initial Assessment**: Start by understanding the presentation structure
2. **Data Discovery**: Identify where the target data is located
3. **Targeted Extraction**: Extract data using appropriate tools and filters
4. **Data Processing**: Clean and structure the extracted data
5. **Quality Validation**: Verify the completeness and accuracy of results

## Step-by-Step Workflow

### Step 1: Presentation Overview
First, get a comprehensive understanding of the presentation:

```
Tool: get_presentation_overview
Parameters:
  analysis_depth: "detailed"
  include_sample_content: true
```

**Analysis Focus:**
- Identify slide types and content distribution
- Look for patterns in slide titles and structure
- Note any sections or organizational structure
- Assess overall complexity and data density

### Step 2: Data Source Identification
Based on the data types requested ({data_types}), identify relevant slides:

**For Table Data:**
```
Tool: query_slides
Parameters:
  search_criteria:
    content:
      has_tables: true
    title:
      one_of: {search_criteria or [".*data.*", ".*table.*", ".*results.*", ".*metrics.*"]}
  return_details: "detailed"
```

**For Financial Data:**
```
Tool: query_slides
Parameters:
  search_criteria:
    title:
      one_of: [".*financial.*", ".*budget.*", ".*revenue.*", ".*cost.*", ".*profit.*", ".*\\$.*"]
    content:
      has_tables: true
```

**For Project Status Data:**
```
Tool: query_slides
Parameters:
  search_criteria:
    title:
      one_of: [".*status.*", ".*progress.*", ".*update.*", ".*dashboard.*", ".*milestone.*"]
    content:
      has_tables: true
```

### Step 3: Comprehensive Data Extraction
Extract data from identified slides using multiple approaches:

**Primary Table Extraction:**
```
Tool: extract_table_data
Parameters:
  slide_numbers: [from previous query results]
  table_criteria:
    min_rows: 2
    min_columns: 2
  column_selection:
    all_columns: true
  formatting_detection:
    detect_bold: true
    detect_highlight: true
    detect_colors: true
    detect_hyperlinks: true
  output_format: "structured"
  include_metadata: true
```

**Formatting Analysis for Context:**
```
Tool: analyze_text_formatting
Parameters:
  slide_numbers: [from previous query results]
  formatting_filter:
    content_types: ["tables", "titles", "text_boxes"]
    formatting_types: ["bold", "highlight", "color"]
  grouping: "by_slide"
```

### Step 4: Advanced Data Processing
Apply intelligent filtering and aggregation:

**Highlight Important Data:**
```
Tool: filter_and_aggregate
Parameters:
  data_source: [from table extraction]
  filters:
    - field: "formatting.highlight"
      condition: "equals"
      value: true
    - field: "formatting.bold"
      condition: "equals"
      value: true
  filter_logic: "OR"
  sorting:
    - field: "slide_number"
      order: "asc"
```

**Color-Coded Analysis:**
```
Tool: filter_and_aggregate
Parameters:
  data_source: [from table extraction]
  filters:
    - field: "formatting.font_color"
      condition: "in_list"
      value: ["#FF0000", "#FFA500", "#008000"]  # Red, Orange, Green
  grouping:
    fields: ["formatting.font_color"]
    aggregations:
      - field: "value"
        operation: "list"
        output_field: "values_by_color"
```

### Step 5: Pattern Recognition and Insights
Analyze the extracted data for patterns:

**Numerical Data Analysis:**
- Identify columns with numerical data
- Look for trends, outliers, or significant values
- Check for percentage values, currency amounts, or metrics

**Status and Progress Indicators:**
- Identify status columns (Complete, In Progress, Not Started)
- Look for progress percentages or completion indicators
- Find risk indicators (Red, Yellow, Green status)

**Temporal Data:**
- Identify date columns or time-based data
- Look for deadlines, milestones, or timeline information

## Data Interpretation Guidelines

### Financial Data Patterns:
- **Revenue/Sales**: Look for currency symbols, "Revenue", "Sales", "Income"
- **Costs/Expenses**: Look for "Cost", "Expense", "Budget", "Spend"
- **Metrics**: Look for percentages, ratios, "ROI", "Margin", "Growth"

### Project Status Patterns:
- **Progress**: Look for percentages, "Complete", "Done", "Finished"
- **Status**: Look for "Green/Yellow/Red", "On Track", "At Risk", "Delayed"
- **Milestones**: Look for dates, "Milestone", "Deadline", "Target"

### Quality Indicators:
- **Highlighted Text**: Often indicates important or exceptional values
- **Bold Text**: Usually indicates headers or key metrics
- **Color Coding**: 
  - Red: Problems, risks, negative values
  - Green: Success, positive values, completed items
  - Yellow/Orange: Warnings, attention needed

## Output Formatting

Structure your final response based on the requested output format ({output_format}):

### Structured Format:
```json
{{
  "presentation_summary": {{
    "file_path": "{file_path}",
    "total_slides": "number",
    "data_slides_found": "number",
    "extraction_timestamp": "ISO timestamp"
  }},
  "extracted_data": {{
    "tables": [
      {{
        "slide_number": "number",
        "table_index": "number", 
        "headers": ["list of headers"],
        "data": [
          {{"header1": {{"value": "text", "formatting": {{}}}}}
        ],
        "metadata": {{}}
      }}
    ],
    "formatted_text": [
      {{
        "slide_number": "number",
        "content_type": "type",
        "text": "content",
        "formatting": {{}}
      }}
    ]
  }},
  "insights": {{
    "key_findings": ["list of important discoveries"],
    "data_quality": "assessment",
    "recommendations": ["list of recommendations"]
  }}
}}
```

### Summary Format:
Provide a concise summary highlighting:
- Total amount of data extracted
- Key patterns or insights discovered
- Data quality assessment
- Most important findings
- Recommendations for further analysis

## Error Handling and Recovery

If any step fails:

1. **No slides found**: Broaden search criteria or check presentation structure
2. **No tables found**: Look for text-based data or different content types
3. **Empty extraction**: Verify table criteria and formatting detection settings
4. **Unexpected results**: Examine individual slides manually for verification

## Quality Validation

Before finalizing results:
1. Verify slide numbers are valid and accessible
2. Check that extracted data makes logical sense
3. Confirm formatting detection captured relevant information
4. Validate that search criteria matched intended content
5. Ensure output format meets requirements

Remember: The goal is to extract meaningful, actionable data while maintaining accuracy and providing context for interpretation.
""",
    "examples": [
        {
            "name": "Financial Data Extraction",
            "description": "Extract financial metrics from quarterly report",
            "arguments": {
                "file_path": "Q4_Financial_Report.pptx",
                "data_types": ["financial", "tables"],
                "search_criteria": [".*revenue.*", ".*profit.*", ".*budget.*"],
                "output_format": "structured"
            },
            "expected_workflow": [
                "Get presentation overview to understand structure",
                "Query slides for financial content",
                "Extract table data with formatting detection",
                "Filter for highlighted/important values",
                "Analyze color-coded status indicators",
                "Structure results with financial context"
            ]
        },
        {
            "name": "Project Status Dashboard",
            "description": "Extract project status and progress data",
            "arguments": {
                "file_path": "Project_Dashboard.pptx", 
                "data_types": ["project_status", "tables"],
                "search_criteria": [".*status.*", ".*progress.*", ".*milestone.*"],
                "output_format": "summary"
            },
            "expected_workflow": [
                "Identify project status slides",
                "Extract progress tables and metrics",
                "Analyze formatting for status indicators",
                "Filter for at-risk or completed items",
                "Summarize overall project health"
            ]
        }
    ]
}


def get_complex_data_extraction_prompt() -> Dict[str, Any]:
    """Get the complex data extraction prompt."""
    return COMPLEX_DATA_EXTRACTION_PROMPT