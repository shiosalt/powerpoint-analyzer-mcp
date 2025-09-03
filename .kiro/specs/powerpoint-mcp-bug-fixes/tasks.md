# Implementation Plan

- [x] 1. Fix TextFormattingAnalyzer formatting count detection


  - Enhance the `_analyze_text_formatting_in_element` method to properly detect all formatting types
  - Fix bold detection to handle XML attributes, child elements, and theme-based formatting
  - Add proper italic, underline, strikethrough, highlight, and color detection
  - Implement comprehensive XML namespace handling for formatting elements
  - _Requirements: 1.1, 1.2, 1.3, 1.4, 1.5, 1.6, 1.7, 1.8_

- [x] 2. Create FormattingExtractor class for position-aware text extraction


  - Create new `powerpoint_mcp_server/core/formatting_extractor.py` file
  - Implement character-level position tracking within complete text content
  - Add precise formatted segment extraction that returns only formatted text portions
  - Implement proper hyperlink relationship resolution using presentation relationships
  - Add support for all formatting types: bold, italic, underlined, highlighted, strikethrough, font_sizes, font_colors
  - _Requirements: 2.1, 2.2, 2.3, 2.4, 2.5, 2.6, 2.7, 2.8, 3.1, 3.2, 3.3, 3.4, 3.5, 3.6_

- [x] 3. Update extract_text_formatting MCP tool implementation


  - Modify the `extract_text_formatting` method in server.py to use the new FormattingExtractor
  - Ensure italic and hyperlinks formatting types are properly recognized
  - Fix position calculations to return accurate start_position and end_position values
  - Ensure formatted_segments contain only the formatted text portions, not complete text
  - Add proper error handling for unsupported formatting types
  - _Requirements: 2.1, 2.2, 2.3, 2.4, 2.5, 2.6, 2.7, 2.8, 2.9, 3.1, 3.2, 3.3, 3.4, 3.5, 3.6, 3.7_

- [x] 4. Fix EnhancedTableExtractor summary calculation and slide numbering


  - Fix the `_format_structured_output` method to properly calculate summary values
  - Ensure total_tables_found, slides_with_tables counts are accurate and non-zero when tables exist
  - Fix slide number validation to handle 1-based indexing correctly
  - Correct off-by-one errors in slide number parameter processing
  - Improve error handling to return proper JSON responses instead of cryptic error messages
  - _Requirements: 4.1, 4.2, 4.3, 4.4, 4.5, 4.6, 4.7, 5.1, 5.2, 5.3, 5.4, 5.5, 5.6, 5.7_

- [x] 5. Add search criteria validation to SlideQueryEngine


  - Implement upfront search criteria validation before slide processing
  - Add validation for field names, operators, and value types in search criteria
  - Return zero results (empty array) for invalid search criteria instead of all slides
  - Add proper error messages for validation failures
  - Ensure syntax errors in search criteria are caught and handled appropriately
  - _Requirements: 6.1, 6.2, 6.3, 6.4, 6.5, 6.6, 6.7_

- [x] 6. Implement section filtering in SlideQueryEngine


  - Add section metadata extraction from presentation.xml
  - Implement section-based filtering logic in `_apply_filters` method
  - Handle non-existent sections by returning zero results
  - Support section names with special characters
  - Combine section filtering correctly with other search criteria
  - _Requirements: 7.1, 7.2, 7.3, 7.4, 7.5, 7.6, 7.7_

- [x] 7. Create comprehensive integration test suite


  - Create `tests/test_bug_fixes_integration.py` file with MCP protocol-based tests
  - Implement test for analyze_text_formatting returning accurate formatting counts
  - Add test for extract_text_formatting recognizing italic and hyperlinks correctly
  - Create test for position accuracy in formatted_segments with correct start_position values
  - Add test for table extraction returning accurate summary values and handling slide numbers
  - Implement test for query_slides returning zero results for invalid search criteria
  - Add test for section-based filtering working correctly
  - _Requirements: 8.1, 8.2, 8.3, 8.4, 8.5, 8.6, 8.7, 8.8, 8.9, 8.10_

- [x] 8. Validate and enhance test_complex.pptx file content


  - Verify test_complex.pptx contains slides with all required formatting types
  - Ensure file has bold, italic, underlined, highlighted, strikethrough, colored text, hyperlinks, and font size variations
  - Add tables on specific slides for table extraction testing
  - Create presentation sections for section-based query testing
  - Document expected test results for each slide and formatting type
  - _Requirements: 9.1, 9.2, 9.3, 9.4, 9.5, 9.6, 9.7_

- [x] 9. Update MCP tool wrappers in main.py


  - Update analyze_text_formatting FastMCP tool wrapper to use fixed implementation
  - Update extract_text_formatting FastMCP tool wrapper to use new FormattingExtractor
  - Update extract_table_data FastMCP tool wrapper to use fixed table extraction
  - Update query_slides FastMCP tool wrapper to use enhanced validation and filtering
  - Ensure consistent error handling and logging across all tool wrappers
  - _Requirements: 8.1, 8.2, 8.3, 8.4, 8.5, 8.6, 8.7, 8.8_

- [x] 10. Run comprehensive integration tests and validate all fixes
  - Execute all integration tests using test_complex.pptx file
  - Verify that formatting_counts return accurate non-zero values for existing formatting
  - Confirm that italic and hyperlinks are correctly recognized in extract_text_formatting
  - Validate that start_position values are correct and formatted_segments contain only formatted text
  - Check that table extraction summary values are accurate and slide number parameters work
  - Ensure invalid search criteria return zero results in query_slides
  - Verify section-based filtering works correctly
  - Document test results and confirm all identified bugs are resolved
  - _Requirements: 8.1, 8.2, 8.3, 8.4, 8.5, 8.6, 8.7, 8.8, 8.9, 8.10_

- [x] 11. Add sections and notes extraction to content processing


  - Enhance ContentExtractor to extract presentation sections from presentation.xml
  - Add speaker notes extraction for each slide from notesSlide relationships
  - Implement sections and notes data structures in slide content models
  - Add sections and notes information to slide query engine data processing
  - Ensure sections and notes are available for filtering operations
  - _Requirements: 10.1, 10.2, 10.3, 10.4, 13.1, 13.2, 13.5, 13.6_

- [x] 12. Implement sections and notes filtering in search functions

  - Add sections field support to SlideQueryEngine search criteria validation
  - Add notes field support to SlideQueryEngine search criteria validation
  - Implement section-based filtering logic with case-insensitive matching
  - Implement notes content filtering with text matching and regex support
  - Add sections and notes filtering to EnhancedTableExtractor table_criteria
  - Ensure combined filtering works correctly with AND logic
  - _Requirements: 10.1, 10.2, 10.3, 10.4, 10.5, 10.6, 10.7_

- [x] 13. Fix query_slides grammar error handling


  - Implement comprehensive search criteria syntax validation before processing
  - Add validation for JSON structure, field names, and operator syntax
  - Return zero results with clear error messages for malformed criteria
  - Add specific error details for different types of syntax issues
  - Ensure grammar validation occurs before any slide processing
  - Test with various invalid syntax scenarios
  - _Requirements: 11.1, 11.2, 11.3, 11.4, 11.5, 11.6, 11.7_

- [x] 14. Fix extract_table_data slide number display issues


  - Identify and fix slide number mapping errors in table extraction results
  - Ensure slide_number field in results matches actual slide position
  - Correct any off-by-one errors in slide number conversion
  - Add logging to show both internal and external slide number values
  - Verify slide number accuracy across all extraction paths
  - Test with slides 10+ to ensure correct numbering
  - _Requirements: 12.1, 12.2, 12.3, 12.4, 12.5, 12.6, 12.7_

- [x] 15. Add sections and notes information to analysis tools


  - Update analyze_text_formatting to include sections information for each slide
  - Update analyze_text_formatting to include notes content for each slide
  - Update get_presentation_overview to include complete sections structure
  - Update get_presentation_overview to include notes summary statistics
  - Handle presentations with no sections or notes gracefully
  - Ensure proper formatting of sections and notes data in responses
  - _Requirements: 13.1, 13.2, 13.3, 13.4, 13.5, 13.6, 13.7, 13.8_

- [x] 16. Create comprehensive integration tests for search and display improvements



  - Add tests for sections and notes filtering in query_slides
  - Add tests for sections and notes filtering in extract_table_data
  - Add tests for grammar error handling returning zero results
  - Add tests for correct slide number display in table extraction
  - Add tests for sections and notes inclusion in analyze_text_formatting
  - Add tests for sections and notes inclusion in get_presentation_overview
  - Ensure test file contains appropriate sections and notes content
  - Verify all new functionality works without breaking existing features
  - _Requirements: 14.1, 14.2, 14.3, 14.4, 14.5, 14.6, 14.7, 14.8, 14.9, 14.10_