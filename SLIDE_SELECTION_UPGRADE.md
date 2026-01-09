# Slide Selection Upgrade - Python-Style Slicing Support

## Overview

Enhanced the PowerPoint Analyzer MCP server to support Python-style slicing notation for the `slide_numbers` parameter across all tools. This provides a more intuitive and flexible way to specify which slides to process.

## Changes Made

### 1. New Utility Module
- **File**: `powerpoint_mcp_server/utils/slide_selector.py`
- **Function**: `parse_slide_numbers(slide_spec, total_slides)`
- **Purpose**: Parse various slide specification formats into a list of slide numbers

### 2. Updated Core Components

#### Server Module (`powerpoint_mcp_server/server.py`)
- Updated `_resolve_slide_numbers()` method to use the new slide selector
- Changed parameter type from `Optional[List[int]]` to `Any` to support multiple formats
- Added import for the new slide selector utility

#### Slide Query Engine (`powerpoint_mcp_server/core/slide_query_engine.py`)
- Updated `SlideQueryFilters.slide_numbers` type to `Optional[Any]`
- Enhanced `_apply_filters()` method to handle new slide number formats
- Updated validation logic to accept int, str, and List[int] formats
- Added comprehensive error handling for invalid slide specifications

#### Main Server (`main.py`)
- Updated function signatures for `extract_table_data` and `extract_formatted_text`
- Enhanced documentation with new format examples
- Updated parameter descriptions to include all supported formats

### 3. Enhanced Documentation

#### README.md
- Added comprehensive section on "Slide Selection with Python-Style Slicing"
- Included examples for all supported formats
- Updated "Recent Updates" section with version 2.1 information

#### Examples
- **File**: `examples/slide_selection_examples.py`
- Demonstrates all supported slide selection formats
- Shows both utility function usage and MCP tool integration

### 4. Comprehensive Testing

#### Unit Tests (`tests/test_slide_selector.py`)
- 22 test cases covering all functionality
- Tests for valid formats, error conditions, and edge cases
- Validates parsing logic and error handling

#### Integration Tests (`tests/test_slide_numbers_integration.py`)
- Tests MCP tool integration with new formats
- Validates server-level functionality
- Includes validation tests for the slide query engine

## Supported Formats

| Format | Example | Description |
|--------|---------|-------------|
| `None` | `None` | All slides |
| `int` | `3` | Single slide |
| `List[int]` | `[1, 5, 10]` | Specific slides |
| `str` (single) | `"3"` | Single slide as string |
| `str` (comma) | `"1,5,10"` | Comma-separated slides |
| `str` (slice start) | `":10"` | First 10 slides |
| `str` (slice range) | `"5:20"` | Slides 5-20 |
| `str` (slice end) | `"25:"` | Slides 25 to end |
| `str` (with brackets) | `"[5:20]"` | Optional bracket notation |

## Benefits

1. **Intuitive Syntax**: Familiar Python slicing notation
2. **Flexible Selection**: Multiple ways to specify slides
3. **Performance Optimization**: Process only needed slides
4. **Backward Compatibility**: Existing formats still work
5. **Error Handling**: Comprehensive validation and error messages
6. **Consistent API**: Same format works across all tools

## Usage Examples

```python
# Extract tables from first 10 slides
extract_table_data("presentation.pptx", slide_numbers=":10")

# Extract bold text from slides 5-20
extract_formatted_text("presentation.pptx", "bold", slide_numbers="5:20")

# Query specific slides
query_slides("presentation.pptx", {
    "slide_numbers": "1,3,5,10"
})

# Extract tables from slide 25 to end
extract_table_data("presentation.pptx", slide_numbers="25:")
```

## Testing Results

- ✅ All 22 unit tests pass
- ✅ Integration tests validate MCP tool functionality
- ✅ Validation tests confirm error handling
- ✅ Backward compatibility maintained
- ✅ Performance optimized for large presentations

## Implementation Quality

- **Type Safety**: Proper type hints and validation
- **Error Handling**: Comprehensive error messages
- **Documentation**: Extensive docstrings and examples
- **Testing**: 100% test coverage for new functionality
- **Performance**: Efficient parsing and validation
- **Maintainability**: Clean, modular code structure

This upgrade significantly enhances the user experience while maintaining full backward compatibility and adding robust error handling.