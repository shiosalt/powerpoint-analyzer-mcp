# Test Complex PowerPoint File Documentation

## File: test_complex.pptx

This document describes the content and expected results for the test_complex.pptx file used in integration testing.

## Presentation Structure

### Metadata
- **Title**: N/A (no presentation title set)
- **Author**: N/A
- **Total Slides**: 4
- **Sections**: 3

### Sections
1. **既定のセクション** (Default Section)
2. **表確認** (Table Verification)
3. **複雑な書式** (Complex Formatting)

## Slide-by-Slide Content

### Slide 1: Complex Test Presentation
- **Title**: "Complex Test Presentation"
- **Subtitle**: "サブタイトル"
- **Layout**: Default layout
- **Objects**: 2 shapes, 2 text boxes
- **Tables**: None
- **Formatting**: Basic formatting (to be verified manually)

### Slide 2: Table Slide
- **Title**: "Table Slide"
- **Subtitle**: None
- **Layout**: Default layout
- **Objects**: 1 shape, 1 text box, 2 tables
- **Tables**: 
  - Table 1: 3x2 (Headers: "Header 1", "Header 2")
  - Table 2: 10x2 (Headers: "表２項目", "表示値")
- **Formatting Detected**:
  - Highlight formatting: 2 instances in tables
  - Hyperlinks: 2 instances in tables
  - Colors: 3 different colors in tables

### Slide 3: Text Slide
- **Title**: "Text Slide"
- **Subtitle**: None
- **Layout**: Default layout
- **Objects**: 2 shapes, 2 text boxes
- **Tables**: None
- **Formatting Detected**:
  - Highlight formatting: 5 instances in text boxes
  - Colors: 2 different colors in text boxes
  - Hyperlinks: 2 instances in text boxes

### Slide 4: Text Slide 2nd
- **Title**: "Text Slide 2nd"
- **Subtitle**: None
- **Layout**: Default layout
- **Objects**: 9 shapes, 9 text boxes
- **Tables**: None
- **Formatting Detected**:
  - Colors: 1 instance in title
  - Highlight formatting: 5 instances in text boxes
  - Hyperlinks: 2 instances in text boxes

## Expected Test Results

### Formatting Analysis Tests
Based on the detected formatting, the following should be found:

#### analyze_text_formatting tool:
- **Bold count**: To be verified (not detected in current analysis)
- **Italic count**: To be verified (not detected in current analysis)
- **Underline count**: To be verified (not detected in current analysis)
- **Highlight count**: 12 total (2 + 5 + 5)
- **Strikethrough count**: To be verified (not detected in current analysis)
- **Colored text count**: 6 total (3 + 2 + 1)
- **Hyperlinks count**: 6 total (2 + 2 + 2)

#### extract_text_formatting tool:
- **Bold**: Should extract bold text segments with accurate positions
- **Italic**: Should extract italic text segments with accurate positions
- **Hyperlinks**: Should extract 6 hyperlink segments across slides 2, 3, and 4
- **Highlighted**: Should extract 12 highlighted segments across slides 2, 3, and 4
- **Font colors**: Should extract colored text segments with color information

### Table Extraction Tests

#### extract_table_data tool:
- **Slide 2**: Should extract 2 tables
  - Table 1: 3 rows x 2 columns
  - Table 2: 10 rows x 2 columns
- **Other slides**: Should return 0 tables
- **Summary**: 
  - total_tables_found: 2
  - slides_with_tables: 1
  - slides_processed: depends on slide_numbers parameter

### Query Tests

#### query_slides tool:
- **Valid criteria**: Should return matching slides
- **Invalid criteria**: Should return 0 results
- **Section filtering**:
  - "既定のセクション": Should return slides in default section
  - "表確認": Should return slides in table verification section
  - "複雑な書式": Should return slides in complex formatting section
  - "NonExistentSection": Should return 0 results

## Test Coverage Verification

### Required Formatting Types Status:
- ✅ **Highlighted**: Present (12 instances)
- ✅ **Colored text**: Present (6 instances)
- ✅ **Hyperlinks**: Present (6 instances)
- ❓ **Bold**: Needs manual verification
- ❓ **Italic**: Needs manual verification
- ❓ **Underlined**: Needs manual verification
- ❓ **Strikethrough**: Needs manual verification
- ❓ **Font size variations**: Needs manual verification

### Required Table Content Status:
- ✅ **Tables present**: 2 tables on slide 2
- ✅ **Simple table**: Table 1 (3x2)
- ✅ **Complex table**: Table 2 (10x2)
- ✅ **Slides without tables**: Slides 1, 3, 4

### Required Section Structure Status:
- ✅ **Multiple sections**: 3 sections present
- ✅ **Descriptive names**: Japanese section names
- ❓ **Slides distributed across sections**: Needs verification of slide assignments

## Recommendations for Test Enhancement

1. **Add missing formatting types**: Ensure bold, italic, underlined, strikethrough, and font size variations are present
2. **Verify section assignments**: Confirm which slides belong to which sections
3. **Document exact positions**: Record expected start/end positions for formatted text segments
4. **Add edge cases**: Consider adding slides with no formatting, empty tables, etc.

## Usage in Tests

This file should be used with the following test scenarios:

1. **Formatting accuracy tests**: Verify that formatting counts match expected values
2. **Position accuracy tests**: Verify that formatted segments have correct positions
3. **Table extraction tests**: Verify that tables are extracted with correct structure
4. **Query validation tests**: Verify that invalid queries return zero results
5. **Section filtering tests**: Verify that section-based queries work correctly

## Last Updated
Generated automatically by analyze_test_file.py script.