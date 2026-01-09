# Slide Sorting Bug Fix

## Issue Description

The `extract_table_data` function was returning empty results (`"extracted_tables": []`) when using specific slide numbers (e.g., `[10, 12, 13]`) with `"column_selection": {"all_columns": true}`, even when tables actually existed on those slides.

## Root Cause

The bug was caused by **incorrect slide file sorting** throughout the codebase. PowerPoint slide files are named like:
- `ppt/slides/slide1.xml`
- `ppt/slides/slide2.xml`
- `ppt/slides/slide10.xml`
- `ppt/slides/slide11.xml`

When sorted alphabetically (the original behavior), the order becomes:
```
slide1.xml, slide10.xml, slide11.xml, slide12.xml, ..., slide2.xml, slide20.xml, slide3.xml
```

This caused slide number 10 (index 9) to map to `slide10.xml` in the alphabetical list, but `slide10.xml` was actually at index 1, not index 9. This resulted in:
- Slide 10 → `slide18.xml` (wrong file)
- Slide 12 → `slide2.xml` (wrong file)  
- Slide 13 → `slide20.xml` (wrong file)

## Solution

### 1. Added Numerical Sorting Utility

Added a new method `get_slide_xml_files_sorted()` to `ZipExtractor` that sorts slide files numerically:

```python
def get_slide_xml_files_sorted(self) -> List[str]:
    """Get slide XML file paths sorted numerically by slide number."""
    slide_files_dict = self.get_slide_xml_files()
    
    def extract_slide_number(slide_path):
        """Extract slide number from path like 'ppt/slides/slide1.xml'"""
        import re
        match = re.search(r'slide(\d+)\.xml$', slide_path)
        return int(match.group(1)) if match else 0
    
    return sorted(slide_files_dict.keys(), key=extract_slide_number)
```

### 2. Updated All Affected Files

Fixed slide sorting in the following files:
- `powerpoint_mcp_server/core/enhanced_table_extractor.py`
- `powerpoint_mcp_server/core/formatting_extractor.py`
- `powerpoint_mcp_server/core/text_formatting_analyzer.py`
- `powerpoint_mcp_server/core/presentation_analyzer.py`
- `powerpoint_mcp_server/core/slide_query_engine.py`
- `powerpoint_mcp_server/server.py`

### 3. Before and After Comparison

**Before (Alphabetical - Incorrect):**
```
['ppt/slides/slide1.xml', 'ppt/slides/slide10.xml', 'ppt/slides/slide11.xml', 
 'ppt/slides/slide12.xml', 'ppt/slides/slide2.xml', 'ppt/slides/slide20.xml']
```

**After (Numerical - Correct):**
```
['ppt/slides/slide1.xml', 'ppt/slides/slide2.xml', 'ppt/slides/slide3.xml',
 'ppt/slides/slide4.xml', 'ppt/slides/slide10.xml', 'ppt/slides/slide11.xml']
```

## Impact

This fix resolves the issue where:
- `extract_table_data` with specific slide numbers was accessing wrong slides
- `extract_formatted_text` with slide numbers was analyzing wrong slides  
- `query_slides` with slide number filters was searching wrong slides
- Any other slide-based operations were processing incorrect slides

## Testing

The fix was verified by:
1. Testing `extract_table_data` with specific slide numbers `[10, 12, 13]`
2. Confirming correct slide file mapping (slide 10 → `slide10.xml`)
3. Verifying that existing functionality still works correctly
4. Running existing test suites to ensure no regressions

## Files Modified

- `powerpoint_mcp_server/utils/zip_extractor.py` - Added numerical sorting utility
- `powerpoint_mcp_server/core/enhanced_table_extractor.py` - Fixed slide sorting
- `powerpoint_mcp_server/core/formatting_extractor.py` - Fixed slide sorting
- `powerpoint_mcp_server/core/text_formatting_analyzer.py` - Fixed slide sorting  
- `powerpoint_mcp_server/core/presentation_analyzer.py` - Fixed slide sorting
- `powerpoint_mcp_server/core/slide_query_engine.py` - Fixed slide sorting
- `powerpoint_mcp_server/server.py` - Fixed slide sorting

## Backward Compatibility

This fix maintains full backward compatibility. All existing functionality works exactly the same, but now with correct slide number mapping.