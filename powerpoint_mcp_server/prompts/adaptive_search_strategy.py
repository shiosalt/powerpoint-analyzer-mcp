"""
MCP Prompt for adaptive search strategy with intelligent query refinement.
"""

from typing import Dict, Any, List, Optional

ADAPTIVE_SEARCH_STRATEGY_PROMPT = {
    "name": "adaptive_search_strategy",
    "description": "Intelligent search strategy that adapts based on results and refines queries automatically",
    "arguments": [
        {
            "name": "file_path",
            "description": "Path to the PowerPoint file to search",
            "required": True
        },
        {
            "name": "search_objective",
            "description": "What you're looking for (specific_data, content_type, patterns, quality_issues)",
            "required": True
        },
        {
            "name": "initial_keywords",
            "description": "Initial keywords or patterns to search for",
            "required": False
        },
        {
            "name": "search_scope",
            "description": "Scope of search (titles, content, formatting, all)",
            "required": False
        },
        {
            "name": "success_criteria",
            "description": "What constitutes a successful search result",
            "required": False
        }
    ],
    "template": """
You are implementing an adaptive search strategy for the PowerPoint presentation at {file_path}. Your objective is: {search_objective}

This strategy automatically adapts based on search results, refining queries and trying alternative approaches when initial searches don't yield optimal results.

## Adaptive Search Philosophy

The adaptive approach follows these principles:
1. **Start Broad, Then Narrow**: Begin with wide searches, then refine based on results
2. **Learn from Results**: Use search results to inform next search strategies
3. **Multiple Angles**: Try different search approaches if initial ones fail
4. **Context Awareness**: Consider presentation structure and content patterns
5. **Quality Over Quantity**: Focus on finding the most relevant results

## Phase 1: Initial Reconnaissance

### Step 1.1: Presentation Intelligence Gathering
First, understand what you're working with:

```
Tool: get_presentation_overview
Parameters:
  analysis_depth: "basic"
  include_sample_content: true
```

**Intelligence Analysis:**
- Total slides and complexity level
- Slide type distribution (title, content, table, chart slides)
- Sample titles and content patterns
- Overall structure and organization
- Potential search challenges or opportunities

### Step 1.2: Baseline Search
Start with the initial search based on your objective ({search_objective}):

**For Specific Data Search:**
```
Tool: query_slides
Parameters:
  search_criteria:
    title:
      one_of: {initial_keywords or [".*data.*", ".*table.*", ".*results.*"]}
    content:
      has_tables: true
  return_details: "detailed"
  limit: 20
```

**For Content Type Search:**
```
Tool: query_slides
Parameters:
  search_criteria:
    content:
      has_tables: {search_scope includes tables}
      has_charts: {search_scope includes charts}
      has_images: {search_scope includes images}
  return_details: "basic"
  limit: 30
```

**For Pattern Search:**
```
Tool: analyze_text_formatting
Parameters:
  formatting_filter:
    content_types: ["all"]
    formatting_types: ["all"]
  grouping: "by_slide"
```

## Phase 2: Result Analysis and Strategy Adaptation

### Step 2.1: Search Result Evaluation
Analyze the initial search results:

**Result Quality Metrics:**
- Number of results found
- Relevance of slide titles to search objective
- Content richness (object counts, text length)
- Formatting patterns that might indicate importance

**Adaptation Triggers:**
- **Too Many Results (>30)**: Need to narrow search criteria
- **Too Few Results (<3)**: Need to broaden search criteria  
- **Irrelevant Results**: Need to change search approach
- **Mixed Quality**: Need to add filtering criteria

### Step 2.2: Adaptive Strategy Selection

Based on initial results, choose the appropriate adaptation strategy:

#### Strategy A: Refinement (when too many results)
```
Tool: query_slides
Parameters:
  search_criteria:
    title:
      contains: [most relevant keyword from initial results]
    content:
      has_tables: true
      object_count_min: 3  # Focus on content-rich slides
  return_details: "detailed"
  limit: 15
```

#### Strategy B: Expansion (when too few results)
```
Tool: query_slides
Parameters:
  search_criteria:
    title:
      one_of: [expanded keyword list based on sample content]
    content:
      object_count_min: 1  # Lower threshold
  return_details: "detailed"
  limit: 40
```

#### Strategy C: Alternative Approach (when results are irrelevant)
Switch to content-based search instead of title-based:
```
Tool: analyze_text_formatting
Parameters:
  formatting_filter:
    text_contains: [keywords from search objective]
    content_types: ["all"]
  grouping: "by_slide"
```

#### Strategy D: Pattern-Based Search (when looking for specific patterns)
```
Tool: query_slides
Parameters:
  search_criteria:
    title:
      regex: [pattern derived from presentation structure]
  return_details: "detailed"
```

## Phase 3: Intelligent Query Refinement

### Step 3.1: Keyword Evolution
Based on successful searches, evolve your keyword strategy:

**Keyword Expansion Techniques:**
- Use successful slide titles to generate related terms
- Extract common words from relevant content
- Identify domain-specific terminology from the presentation
- Consider synonyms and variations of successful terms

**Example Keyword Evolution:**
```
Initial: ["budget"]
After Analysis: ["budget", "financial", "cost", "expense", "revenue", "profit"]
Refined: ["budget.*plan", "financial.*summary", "cost.*analysis"]
```

### Step 3.2: Multi-Dimensional Search
Combine different search dimensions for better results:

```
Tool: query_slides
Parameters:
  search_criteria:
    title:
      one_of: [evolved keywords]
    content:
      has_tables: true
      object_count_min: [threshold based on presentation complexity]
    slide_numbers: [exclude slides already found irrelevant]
  return_details: "detailed"
```

### Step 3.3: Contextual Filtering
Apply contextual filters based on presentation patterns:

**If presentation has clear sections:**
```
Tool: query_slides
Parameters:
  search_criteria:
    section: [relevant section name]
    title:
      contains: [refined keywords]
```

**If presentation has numbered slides:**
```
Tool: query_slides
Parameters:
  search_criteria:
    title:
      regex: "\\d+.*[keyword pattern]"
```

## Phase 4: Result Validation and Quality Assessment

### Step 4.1: Content Validation
Validate that found slides actually contain what you're looking for:

```
Tool: get_powerpoint_attributes
Parameters:
  slide_numbers: [top candidates from search]
  attributes: ["title", "text_elements", "tables", "object_counts"]
```

**Validation Criteria:**
- Does the content match the search objective?
- Is there sufficient data/information in the slides?
- Are the slides relevant to the user's needs?
- Is the quality of content adequate?

### Step 4.2: Completeness Check
Ensure you haven't missed important content:

```
Tool: query_slides
Parameters:
  search_criteria:
    slide_numbers: [slides NOT in current results]
    content:
      object_count_min: [threshold for potentially relevant slides]
  return_details: "basic"
  limit: 10
```

## Phase 5: Final Optimization and Extraction

### Step 5.1: Optimized Extraction
Based on validated results, perform optimized data extraction:

**For Table Data:**
```
Tool: extract_table_data
Parameters:
  slide_numbers: [validated slide numbers]
  table_criteria:
    min_rows: 2
    header_contains: [relevant headers identified during search]
  formatting_detection:
    detect_bold: true
    detect_highlight: true
    detect_colors: true
```

**For Formatted Content:**
```
Tool: analyze_text_formatting
Parameters:
  slide_numbers: [validated slide numbers]
  formatting_filter:
    formatting_types: [types relevant to search objective]
    text_contains: [refined keywords]
```

### Step 5.2: Result Ranking and Prioritization
Rank results by relevance and quality:

**Ranking Factors:**
1. **Keyword Match Strength**: Exact matches > partial matches > related terms
2. **Content Richness**: More objects/data > sparse content
3. **Formatting Indicators**: Highlighted/bold content suggests importance
4. **Structural Position**: Title slides, summary slides often more important
5. **Context Relevance**: Content that fits the search objective context

## Adaptive Strategies by Search Objective

### For "specific_data" objective:
1. Start with keyword-based title search
2. If insufficient, expand to content search
3. Use table extraction to validate data presence
4. Refine based on actual data found
5. Cross-reference with formatting analysis

### For "content_type" objective:
1. Start with object-type filters (has_tables, has_charts)
2. Analyze distribution and patterns
3. Refine based on content quality
4. Use formatting analysis to identify important content
5. Validate with actual content extraction

### For "patterns" objective:
1. Start with formatting analysis across all slides
2. Identify pattern clusters and anomalies
3. Use pattern insights to guide targeted searches
4. Validate patterns with content extraction
5. Refine pattern definitions based on findings

### For "quality_issues" objective:
1. Start with presentation overview for structural issues
2. Use formatting analysis to find inconsistencies
3. Search for slides with unusual characteristics
4. Validate issues with detailed content analysis
5. Prioritize issues by impact and frequency

## Error Recovery and Fallback Strategies

### When searches return no results:
1. **Broaden Keywords**: Remove restrictive terms, add synonyms
2. **Lower Thresholds**: Reduce object_count_min, expand criteria
3. **Change Approach**: Switch from title-based to content-based search
4. **Manual Sampling**: Extract content from random slides to understand structure

### When searches return too many irrelevant results:
1. **Add Exclusions**: Exclude slides with irrelevant patterns
2. **Increase Specificity**: Use more specific keywords or regex patterns
3. **Add Content Filters**: Require specific object types or characteristics
4. **Use Formatting Filters**: Focus on formatted content that suggests importance

### When results are inconsistent:
1. **Analyze Patterns**: Look for common characteristics in good results
2. **Refine Criteria**: Use successful patterns to improve search criteria
3. **Validate Manually**: Check a sample of results to understand quality
4. **Iterate Strategy**: Adjust approach based on validation findings

## Success Criteria Evaluation

Evaluate success based on your defined criteria ({success_criteria}):

**Quantitative Measures:**
- Number of relevant slides found
- Percentage of search objectives met
- Quality score of extracted content
- Coverage of presentation content

**Qualitative Measures:**
- Relevance of found content to user needs
- Completeness of information extracted
- Actionability of results
- Clarity and usefulness of findings

## Final Output Structure

```json
{{
  "search_summary": {{
    "file_path": "{file_path}",
    "search_objective": "{search_objective}",
    "initial_keywords": {initial_keywords},
    "final_keywords": ["evolved keyword list"],
    "search_phases_completed": ["phase1", "phase2", "phase3", "phase4", "phase5"],
    "adaptations_made": ["list of strategy adaptations"]
  }},
  "search_results": {{
    "total_slides_found": "number",
    "high_relevance_slides": ["slide numbers with high relevance"],
    "medium_relevance_slides": ["slide numbers with medium relevance"],
    "search_coverage": "percentage of presentation searched"
  }},
  "extracted_content": {{
    "primary_findings": [
      {{
        "slide_number": "number",
        "relevance_score": "1-10",
        "content_summary": "brief description",
        "extracted_data": {{}}
      }}
    ],
    "supporting_findings": [],
    "related_content": []
  }},
  "search_intelligence": {{
    "presentation_characteristics": {{}},
    "successful_strategies": ["strategies that worked well"],
    "unsuccessful_strategies": ["strategies that didn't work"],
    "recommendations_for_future_searches": []
  }},
  "quality_assessment": {{
    "success_criteria_met": "boolean",
    "completeness_score": "1-10",
    "relevance_score": "1-10",
    "confidence_level": "high/medium/low"
  }}
}}
```

Remember: The key to adaptive search is learning from each result and continuously refining your approach. Stay flexible and be ready to pivot strategies based on what the data tells you.
""",
    "examples": [
        {
            "name": "Financial Data Hunt",
            "description": "Adaptive search for financial metrics in annual report",
            "arguments": {
                "file_path": "Annual_Report.pptx",
                "search_objective": "specific_data",
                "initial_keywords": ["revenue", "profit"],
                "search_scope": "all",
                "success_criteria": "Find all financial tables with quarterly data"
            },
            "expected_adaptations": [
                "Start with revenue/profit title search",
                "Expand to financial, budget, earnings keywords",
                "Add table requirement for data validation",
                "Refine to quarterly-specific content",
                "Validate with actual table extraction"
            ]
        },
        {
            "name": "Risk Assessment Search",
            "description": "Adaptive search for risk-related content with pattern recognition",
            "arguments": {
                "file_path": "Project_Review.pptx",
                "search_objective": "patterns",
                "initial_keywords": ["risk", "issue"],
                "search_scope": "content",
                "success_criteria": "Identify all risk indicators and their severity levels"
            },
            "expected_adaptations": [
                "Start with risk/issue keyword search",
                "Analyze formatting patterns for severity indicators",
                "Expand to challenge, problem, concern keywords",
                "Focus on color-coded content (red/yellow indicators)",
                "Cross-reference with status tables"
            ]
        }
    ]
}


def get_adaptive_search_strategy_prompt() -> Dict[str, Any]:
    """Get the adaptive search strategy prompt."""
    return ADAPTIVE_SEARCH_STRATEGY_PROMPT