"""
MCP Prompt for progressive table analysis workflows.
"""

from typing import Dict, Any, List, Optional

PROGRESSIVE_TABLE_ANALYSIS_PROMPT = {
    "name": "progressive_table_analysis",
    "description": "Step-by-step analysis of table data in PowerPoint presentations with progressive refinement",
    "arguments": [
        {
            "name": "file_path",
            "description": "Path to the PowerPoint file to analyze",
            "required": True
        },
        {
            "name": "analysis_focus",
            "description": "Focus area for analysis (overview, specific_columns, formatting_patterns, data_validation)",
            "required": False
        },
        {
            "name": "target_columns",
            "description": "Specific column names or patterns to focus on",
            "required": False
        },
        {
            "name": "refinement_criteria",
            "description": "Criteria for progressive refinement (formatting, values, patterns)",
            "required": False
        }
    ],
    "template": """
You are conducting a progressive table analysis of the PowerPoint presentation at {file_path}. This approach allows you to start broad and progressively narrow down to specific insights through iterative analysis.

## Progressive Analysis Philosophy

This workflow follows a "funnel" approach:
1. **Wide Discovery**: Find all tables and understand their structure
2. **Pattern Recognition**: Identify common patterns and interesting features
3. **Focused Analysis**: Drill down into specific areas of interest
4. **Deep Insights**: Extract detailed insights from targeted data
5. **Validation**: Verify findings and ensure completeness

## Phase 1: Discovery and Overview

### Step 1.1: Presentation Structure Analysis
Start with a broad understanding of the presentation:

```
Tool: get_presentation_overview
Parameters:
  analysis_depth: "basic"
  include_sample_content: true
```

**What to look for:**
- How many slides contain tables?
- What types of slides are most common?
- Are there clear sections or themes?
- What's the overall complexity level?

### Step 1.2: Table Discovery
Find all slides containing tables:

```
Tool: query_slides
Parameters:
  search_criteria:
    content:
      has_tables: true
  return_details: "detailed"
  limit: 50
```

**Analysis Questions:**
- How many slides have tables?
- What are the slide titles suggesting about table content?
- Are tables concentrated in specific sections?
- What's the distribution of table-containing slides?

### Step 1.3: Initial Table Structure Survey
Get a broad view of all table structures:

```
Tool: extract_table_data
Parameters:
  slide_numbers: [all slides with tables from step 1.2]
  table_criteria:
    min_rows: 1
    min_columns: 1
  formatting_detection:
    detect_bold: true
    detect_highlight: true
    detect_colors: true
  output_format: "structured"
  include_metadata: true
```

**Structure Analysis:**
- What are the common table sizes (rows × columns)?
- What headers appear most frequently?
- Which tables have the most formatting?
- Are there consistent patterns across tables?

## Phase 2: Pattern Recognition and Categorization

### Step 2.1: Header Pattern Analysis
Analyze table headers to understand data types:

```
Tool: filter_and_aggregate
Parameters:
  data_source: [from step 1.3]
  grouping:
    fields: ["headers"]
    aggregations:
      - field: "slide_number"
        operation: "list"
        output_field: "slides_with_header"
      - field: "headers"
        operation: "count"
        output_field: "header_frequency"
```

**Pattern Categories:**
- **Temporal**: Date, Month, Quarter, Year, Timeline
- **Financial**: Revenue, Cost, Budget, Profit, $, €, £
- **Status**: Status, Progress, Complete, Done, %
- **Identification**: Name, ID, Project, Task, Item
- **Metrics**: Count, Total, Average, Score, Rating
- **Geographic**: Location, Region, Country, City

### Step 2.2: Formatting Pattern Analysis
Understand how formatting is used across tables:

```
Tool: analyze_text_formatting
Parameters:
  slide_numbers: [all table slides]
  formatting_filter:
    content_types: ["tables"]
    formatting_types: ["bold", "highlight", "color"]
  grouping: "by_formatting_type"
```

**Formatting Insights:**
- Which formatting types are most common?
- Are certain colors used consistently?
- Do bold/highlight patterns suggest importance?
- Is formatting used for categorization?

### Step 2.3: Table Categorization
Based on patterns, categorize tables by purpose:

**Data Tables**: Structured data with clear headers and consistent data types
**Status Tables**: Progress tracking, project status, completion metrics
**Financial Tables**: Revenue, costs, budgets, financial metrics
**Comparison Tables**: Side-by-side comparisons, competitive analysis
**Summary Tables**: Aggregated data, totals, key metrics

## Phase 3: Focused Analysis (Based on {analysis_focus})

### Option A: Overview Focus
If analysis_focus is "overview", provide comprehensive summary:

```
Tool: filter_and_aggregate
Parameters:
  data_source: [from initial extraction]
  grouping:
    fields: ["slide_number"]
    aggregations:
      - field: "rows"
        operation: "sum"
        output_field: "total_rows"
      - field: "columns"
        operation: "average"
        output_field: "avg_columns"
      - field: "headers"
        operation: "unique"
        output_field: "unique_headers"
```

### Option B: Specific Columns Focus
If target_columns are specified ({target_columns}), focus on those:

```
Tool: extract_table_data
Parameters:
  slide_numbers: [relevant slides]
  column_selection:
    specific_columns: {target_columns}
    all_columns: false
  formatting_detection:
    detect_bold: true
    detect_highlight: true
    detect_colors: true
```

### Option C: Formatting Patterns Focus
If analysis_focus is "formatting_patterns":

```
Tool: filter_and_aggregate
Parameters:
  data_source: [from initial extraction]
  filters:
    - field: "formatting.bold"
      condition: "equals"
      value: true
    - field: "formatting.highlight"
      condition: "equals"
      value: true
    - field: "formatting.font_color"
      condition: "not_empty"
  filter_logic: "OR"
  grouping:
    fields: ["formatting.font_color"]
    aggregations:
      - field: "value"
        operation: "list"
        output_field: "formatted_values"
```

## Phase 4: Deep Dive Analysis

### Step 4.1: Value Analysis
Analyze the actual data values for insights:

**Numerical Data Analysis:**
- Identify numerical columns
- Look for ranges, outliers, patterns
- Calculate basic statistics where appropriate

**Categorical Data Analysis:**
- Identify categorical columns
- Count unique values and frequencies
- Look for status patterns (Complete/Incomplete, High/Medium/Low)

**Temporal Data Analysis:**
- Identify date/time columns
- Look for trends over time
- Identify deadlines or milestones

### Step 4.2: Cross-Table Relationships
Look for relationships between tables:

```
Tool: filter_and_aggregate
Parameters:
  data_source: [combined table data]
  grouping:
    fields: ["common_header_name"]
    aggregations:
      - field: "slide_number"
        operation: "list"
        output_field: "slides_with_common_data"
      - field: "value"
        operation: "unique"
        output_field: "unique_values_across_tables"
```

### Step 4.3: Anomaly Detection
Look for unusual patterns or outliers:

**Formatting Anomalies:**
- Cells with unique formatting combinations
- Inconsistent formatting patterns
- Unusual color usage

**Data Anomalies:**
- Extremely high or low values
- Missing data patterns
- Inconsistent data types in columns

## Phase 5: Refinement and Validation

### Step 5.1: Targeted Re-extraction
Based on findings, perform targeted re-extraction:

```
Tool: extract_table_data
Parameters:
  slide_numbers: [slides with interesting patterns]
  table_criteria:
    header_contains: [headers of interest from analysis]
  column_selection:
    specific_columns: [refined column list]
  formatting_detection:
    detect_bold: true
    detect_highlight: true
    detect_colors: true
  output_format: "flat"  # For easier analysis
```

### Step 5.2: Quality Validation
Validate the analysis results:

1. **Completeness Check**: Did we capture all relevant tables?
2. **Accuracy Check**: Do extracted values make sense?
3. **Consistency Check**: Are patterns consistent across similar tables?
4. **Context Check**: Do findings align with slide titles and context?

## Progressive Refinement Strategies

### Refinement by Criteria ({refinement_criteria}):

**By Formatting:**
- Start with all formatted cells
- Narrow to specific formatting types
- Focus on most frequently used formats
- Drill down to specific color patterns

**By Values:**
- Start with all data
- Filter by data types (numerical, categorical, dates)
- Focus on specific value ranges or patterns
- Drill down to outliers or anomalies

**By Patterns:**
- Identify recurring patterns
- Focus on most common patterns
- Analyze pattern variations
- Investigate pattern exceptions

## Output Structure

### Progressive Analysis Report:

```json
{{
  "analysis_summary": {{
    "file_path": "{file_path}",
    "analysis_focus": "{analysis_focus}",
    "total_tables_found": "number",
    "analysis_phases_completed": ["phase1", "phase2", "phase3", "phase4", "phase5"]
  }},
  "phase_1_discovery": {{
    "presentation_overview": {{}},
    "table_distribution": {{}},
    "initial_structure_analysis": {{}}
  }},
  "phase_2_patterns": {{
    "header_patterns": {{}},
    "formatting_patterns": {{}},
    "table_categories": {{}}
  }},
  "phase_3_focused_analysis": {{
    "focus_area": "{analysis_focus}",
    "targeted_findings": {{}}
  }},
  "phase_4_deep_dive": {{
    "value_analysis": {{}},
    "cross_table_relationships": {{}},
    "anomalies_detected": {{}}
  }},
  "phase_5_refinement": {{
    "refined_extraction": {{}},
    "validation_results": {{}},
    "quality_assessment": {{}}
  }},
  "key_insights": [
    "Most important findings from the analysis"
  ],
  "recommendations": [
    "Suggestions for further analysis or action"
  ]
}}
```

## Adaptive Analysis Flow

The analysis should adapt based on what's discovered:

**If many small tables found**: Focus on aggregation and pattern recognition
**If few large tables found**: Focus on detailed structure and content analysis
**If highly formatted tables found**: Emphasize formatting pattern analysis
**If consistent table structures found**: Look for data trends and relationships
**If inconsistent structures found**: Focus on categorization and individual analysis

## Decision Points for Progression

At each phase, decide whether to:
1. **Continue broadly**: If patterns are unclear, cast a wider net
2. **Narrow focus**: If clear patterns emerge, drill down deeper
3. **Pivot analysis**: If unexpected findings suggest different approach
4. **Conclude analysis**: If sufficient insights have been gathered

Remember: The goal is progressive understanding - each phase should build on the previous one and inform the next steps.
""",
    "examples": [
        {
            "name": "Financial Report Analysis",
            "description": "Progressive analysis of quarterly financial tables",
            "arguments": {
                "file_path": "Q4_Report.pptx",
                "analysis_focus": "overview",
                "refinement_criteria": "formatting"
            },
            "expected_progression": [
                "Discover 8 slides with financial tables",
                "Identify revenue, cost, and profit patterns",
                "Focus on highlighted variance cells",
                "Deep dive into quarterly trends",
                "Validate against summary slides"
            ]
        },
        {
            "name": "Project Dashboard Analysis",
            "description": "Progressive analysis focusing on status columns",
            "arguments": {
                "file_path": "Project_Status.pptx",
                "analysis_focus": "specific_columns",
                "target_columns": ["Status", "Progress", "Risk"],
                "refinement_criteria": "values"
            },
            "expected_progression": [
                "Find all project status tables",
                "Focus on Status/Progress/Risk columns",
                "Analyze color-coded status indicators",
                "Identify at-risk projects",
                "Cross-reference with timeline data"
            ]
        }
    ]
}


def get_progressive_table_analysis_prompt() -> Dict[str, Any]:
    """Get the progressive table analysis prompt."""
    return PROGRESSIVE_TABLE_ANALYSIS_PROMPT