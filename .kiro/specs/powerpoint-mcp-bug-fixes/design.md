# Design Document

## Overview

This design addresses critical bugs in the PowerPoint Analyzer MCP server that affect core functionality across multiple tools. The bugs span formatting detection, table extraction, slide querying, and position calculation systems. The fixes will be implemented systematically with comprehensive integration testing to ensure reliability.

## Architecture

The bug fixes will be implemented across several core modules:

```
powerpoint_mcp_server/
├── core/
│   ├── text_formatting_analyzer.py    # Fix formatting counts and detection
│   ├── enhanced_table_extractor.py    # Fix table extraction and slide numbering
│   ├── slide_query_engine.py          # Fix query validation and section filtering
│   └── formatting_extractor.py        # New: Position-aware formatting extraction
├── server.py                          # Update MCP tool implementations
└── main.py                            # Update FastMCP tool wrappers
```

## Components and Interfaces

### 1. Text Formatting Analyzer Fixes

**Problem**: `analyze_text_formatting` returns all formatting counts as 0

**Root Cause Analysis**:
- The `_analyze_text_formatting_in_element` method has incomplete formatting detection logic
- Bold detection relies on XML attributes but doesn't handle all PowerPoint formatting scenarios
- Italic, underline, and other formatting types have similar detection issues

**Solution**:
- Enhance the formatting detection algorithm to handle multiple XML structures
- Add support for theme-based formatting and inherited styles
- Implement proper namespace handling for formatting elements

### 2. Text Formatting Extraction Fixes

**Problem**: `extract_text_formatting` fails to recognize italic and hyperlinks, and has incorrect position calculations

**Root Cause Analysis**:
- Italic detection uses incorrect XML element matching
- Hyperlink detection doesn't handle relationship references properly
- Position calculation assumes all text starts at position 0
- Formatted segments include entire text instead of only formatted portions

**Solution**:
- Create a new `FormattingExtractor` class with precise text parsing
- Implement character-level position tracking
- Add proper hyperlink relationship resolution
- Extract only the specifically formatted text segments

### 3. Table Extraction Fixes

**Problem**: `extract_table_data` returns zero counts and fails with slide number parameters

**Root Cause Analysis**:
- Summary calculation logic doesn't properly count extracted tables
- Slide number validation has off-by-one errors (0-based vs 1-based indexing)
- Error handling returns cryptic messages instead of proper results

**Solution**:
- Fix the summary calculation in `_format_structured_output`
- Correct slide number validation and conversion
- Improve error handling to return proper JSON responses

### 4. Query Slides Fixes

**Problem**: Invalid search criteria returns all slides instead of zero results, and section filtering doesn't work

**Root Cause Analysis**:
- Search criteria validation happens after slide processing
- Invalid criteria cause exceptions that are caught and ignored
- Section filtering is not implemented (TODO comment in code)

**Solution**:
- Add upfront search criteria validation
- Implement proper section filtering using presentation metadata
- Return empty results for invalid criteria instead of all slides

## Data Models

### Enhanced Formatting Detection

```python
@dataclass
class FormattingSegment:
    """Represents a formatted text segment with precise positioning."""
    text: str
    start_position: int
    end_position: int
    formatting_type: str
    element_index: int
    
@dataclass
class FormattingExtractionResult:
    """Result of formatting extraction with position-aware segments."""
    file_path: str
    formatting_type: str
    summary: Dict[str, int]
    results_by_slide: List[SlideFormattingResult]

@dataclass
class SlideFormattingResult:
    """Formatting results for a single slide."""
    slide_number: int
    title: str
    complete_text: str
    formatted_segments: List[FormattingSegment]
```

### Enhanced Table Extraction

```python
@dataclass
class TableExtractionSummary:
    """Accurate summary of table extraction results."""
    total_tables_found: int
    slides_with_tables: int
    slides_processed: int
    extraction_errors: int
```

### Query Validation

```python
@dataclass
class QueryValidationResult:
    """Result of search criteria validation."""
    is_valid: bool
    errors: List[str]
    warnings: List[str]
```

## Error Handling

### Validation Layer

All MCP tools will implement upfront validation:

1. **File Path Validation**: Ensure file exists and is readable
2. **Parameter Validation**: Validate all input parameters before processing
3. **Search Criteria Validation**: Validate query syntax and field names
4. **Slide Number Validation**: Ensure slide numbers are within valid range

### Error Response Format

Standardized error responses across all tools:

```json
{
    "error": "Descriptive error message",
    "error_code": "VALIDATION_ERROR|PROCESSING_ERROR|FILE_ERROR",
    "details": {
        "invalid_parameters": ["param1", "param2"],
        "suggestions": ["Use slide numbers 1-10", "Check file path"]
    }
}
```

## Testing Strategy

### Integration Test Framework

Create comprehensive integration tests using the MCP protocol:

```python
class MCPBugFixIntegrationTests:
    """Integration tests for all bug fixes using real MCP communication."""
    
    async def test_formatting_analysis_accuracy(self):
        """Test that formatting counts are accurate and non-zero for existing formatting."""
        
    async def test_text_formatting_extraction_precision(self):
        """Test that italic/hyperlinks are recognized and positions are accurate."""
        
    async def test_table_extraction_completeness(self):
        """Test that table extraction returns proper summaries and handles slide numbers."""
        
    async def test_query_validation_strictness(self):
        """Test that invalid queries return zero results, not all slides."""
        
    async def test_section_filtering_accuracy(self):
        """Test that section-based queries work correctly."""
```

### Test Data Requirements

The `test_complex.pptx` file must contain:

1. **Formatting Test Slides**:
   - Slide with bold text in multiple locations
   - Slide with italic text in various elements
   - Slide with underlined text
   - Slide with highlighted text
   - Slide with strikethrough text
   - Slide with colored text (multiple colors)
   - Slide with hyperlinks (internal and external)
   - Slide with font_size variations (multiple sizes)

2. **Table Test Slides**:
   - Slide with simple table (3x3)
   - Slide with complex table (5x7 with formatting)
   - Slide with no tables
   - Multiple slides with tables for batch testing

3. **Section Test Structure**:
   - Multiple sections with descriptive names
   - Slides distributed across sections

4. **Query Test Content**:
   - Slides with specific titles for title filtering
   - Slides with various layouts
   - Slides with different object counts

### Test Execution Strategy

1. **Unit Tests**: Test individual bug fixes in isolation
2. **Integration Tests**: Test complete MCP tool workflows
3. **Regression Tests**: Ensure fixes don't break existing functionality
4. **Performance Tests**: Verify fixes don't impact performance

## Implementation Plan

### Phase 1: Core Formatting Fixes

1. **Fix TextFormattingAnalyzer**:
   - Enhance `_analyze_text_formatting_in_element` method
   - Add comprehensive formatting detection for all types
   - Implement proper XML namespace handling

2. **Create FormattingExtractor**:
   - New class for position-aware text extraction
   - Character-level position tracking
   - Precise formatted segment extraction

### Phase 2: Table and Query Fixes

1. **Fix EnhancedTableExtractor**:
   - Correct summary calculation logic
   - Fix slide number validation and indexing
   - Improve error handling and responses

2. **Fix SlideQueryEngine**:
   - Add upfront search criteria validation
   - Implement section filtering functionality
   - Return empty results for invalid criteria

### Phase 3: Integration and Testing

1. **Update MCP Tool Implementations**:
   - Integrate fixes into server.py and main.py
   - Ensure consistent error handling across tools
   - Add proper logging for debugging

2. **Comprehensive Testing**:
   - Create integration test suite
   - Test with test_complex.pptx file
   - Validate all bug scenarios are resolved

## Risk Mitigation

### Backward Compatibility

- Maintain existing API interfaces
- Ensure response formats remain consistent
- Add new fields without removing existing ones

### Performance Impact

- Cache formatting analysis results
- Optimize XML parsing for repeated operations
- Implement lazy loading for large presentations

### Error Recovery

- Graceful degradation for partial failures
- Detailed logging for debugging
- Fallback mechanisms for edge cases

### 5. Enhanced Search and Filter Capabilities

**Problem**: Search functions lack sections and notes filtering, query_slides has grammar error issues, extract_table_data shows wrong slide numbers

**Root Cause Analysis**:
- Sections and notes are not extracted or used in filtering logic
- Grammar validation in query_slides is insufficient
- Slide number mapping in table extraction has indexing errors
- analyze_text_formatting and get_presentation_overview don't include structural information

**Solution**:
- Add sections and notes extraction to content processing
- Implement comprehensive grammar validation for search criteria
- Fix slide number mapping in table extraction results
- Enhance analysis tools to include sections and notes information

## Success Criteria

1. **Formatting Analysis**: All formatting counts return accurate values (non-zero when formatting exists)
2. **Text Extraction**: Italic and hyperlinks are properly recognized
3. **Position Accuracy**: Start positions are calculated correctly relative to complete text
4. **Segment Precision**: Formatted segments contain only formatted text portions
5. **Table Extraction**: Summary values are accurate and slide numbers work correctly
6. **Query Validation**: Invalid criteria return zero results, not all slides
7. **Section Filtering**: Section-based queries filter correctly
8. **Enhanced Search**: Sections and notes filtering work in all search functions
9. **Grammar Handling**: Invalid search syntax returns zero results with clear errors
10. **Slide Number Accuracy**: Table extraction shows correct slide numbers
11. **Structural Information**: Analysis tools include sections and notes data
12. **Integration Tests**: All tests pass with comprehensive test coverage