# Requirements Document

## Introduction

PowerPoint Analyzer MCPサーバの重大なバグを修正します。現在、複数のMCPツールで機能が正常に動作しておらず、フォーマット検出、テーブル抽出、スライドクエリなどの主要機能に問題があります。これらのバグを体系的に修正し、結合テストで動作を確認します。

## Requirements

### Requirement 1

**User Story:** As a developer, I want the analyze_text_formatting tool to correctly count formatting attributes, so that I can get accurate statistics about text formatting in presentations.

#### Acceptance Criteria

1. WHEN analyze_text_formatting is called THEN formatting_counts SHALL return accurate values for existing formatting
2. WHEN bold text exists in slides THEN bold_count SHALL reflect the actual number of bold text segments
3. WHEN italic text exists in slides THEN italic_count SHALL reflect the actual number of italic text segments
4. WHEN underlined text exists in slides THEN underlined_count SHALL reflect the actual number of underlined text segments
5. WHEN highlighted text exists in slides THEN highlighted_count SHALL reflect the actual number of highlighted text segments
6. WHEN strikethrough text exists in slides THEN strikethrough_count SHALL reflect the actual number of strikethrough text segments
7. WHEN colored text exists in slides THEN colored_text_count SHALL reflect the actual number of colored text segments
8. WHEN hyperlinks exist in slides THEN hyperlinks_count SHALL reflect the actual number of hyperlink segments
9. IF no formatting of a specific type exists THEN the count for that type SHALL be 0
10. WHEN multiple formatting types exist on the same slide THEN each type SHALL be counted independently and accurately

### Requirement 2

**User Story:** As a developer, I want the extract_text_formatting tool to recognize all supported formatting types, so that I can extract any type of text formatting from presentations.

#### Acceptance Criteria

1. WHEN formatting_type is "bold" THEN the system SHALL correctly identify and extract bold text segments
2. WHEN formatting_type is "italic" THEN the system SHALL correctly identify and extract italic text segments
3. WHEN formatting_type is "underlined" THEN the system SHALL correctly identify and extract underlined text segments
4. WHEN formatting_type is "highlighted" THEN the system SHALL correctly identify and extract highlighted text segments
5. WHEN formatting_type is "strikethrough" THEN the system SHALL correctly identify and extract strikethrough text segments
6. WHEN formatting_type is "hyperlinks" THEN the system SHALL correctly identify and extract hyperlink text and URLs
7. WHEN formatting_type is "font_sizes" THEN the system SHALL correctly identify and extract text with font size information
8. WHEN formatting_type is "font_colors" THEN the system SHALL correctly identify and extract text with color information
9. IF a formatting type is not supported THEN the system SHALL return an appropriate error message

### Requirement 3

**User Story:** As a developer, I want accurate position information for formatted text segments, so that I can precisely locate formatted text within the complete text content.

#### Acceptance Criteria

1. WHEN formatted_segments are returned THEN start_position SHALL indicate the correct character position where the formatted text begins
2. WHEN formatted_segments are returned THEN end_position SHALL indicate the correct character position where the formatted text ends
3. WHEN multiple formatted segments exist THEN each segment SHALL have accurate and unique position information
4. WHEN formatted_segments are extracted THEN they SHALL contain ONLY the text that has the specified formatting applied
5. WHEN formatted_segments are returned THEN each segment SHALL be a separate array element containing only the formatted text portion
6. WHEN multiple formatted segments exist in the same text element THEN each formatted portion SHALL be extracted as a separate segment with accurate positions
7. WHEN no formatting of the specified type exists THEN formatted_segments SHALL be an empty array
8. WHEN position calculations are performed THEN they SHALL be relative to the complete_text field
9. IF overlapping formatting exists THEN position information SHALL handle overlaps correctly

### Requirement 4

**User Story:** As a developer, I want the extract_table_data tool to correctly identify and extract tables from slides, so that I can process tabular data from presentations.

#### Acceptance Criteria

1. WHEN extract_table_data is called on slides containing tables THEN summary.total_tables_found SHALL reflect the actual number of tables
2. WHEN tables exist on specified slides THEN summary.slides_with_tables SHALL count slides that actually contain tables
3. WHEN tables are extracted THEN extracted_tables array SHALL contain table data for each found table
4. WHEN table data is extracted THEN each table SHALL include accurate row and column counts
5. WHEN table cells contain text THEN cell content SHALL be extracted correctly
6. IF no tables exist on specified slides THEN summary values SHALL be 0 and extracted_tables SHALL be empty
7. WHEN table extraction succeeds THEN the response SHALL include complete table structure and content

### Requirement 5

**User Story:** As a developer, I want the extract_table_data tool to handle slide number parameters correctly, so that I can extract tables from specific slides without errors.

#### Acceptance Criteria

1. WHEN slide_numbers parameter contains valid slide numbers THEN the tool SHALL process those slides successfully
2. WHEN slide_numbers contains [10, 11] and slide 10 exists THEN the tool SHALL NOT return "Failed to extract table data: 9"
3. WHEN slide_numbers contains [11] and slide 11 exists THEN the tool SHALL return proper results, not empty Tool Execution Result
4. WHEN slide_numbers contains non-existent slide numbers THEN the tool SHALL return appropriate error messages
5. WHEN slide_numbers is empty or null THEN the tool SHALL process all slides in the presentation
6. IF slide numbering is 1-based in the API THEN internal processing SHALL handle the conversion correctly
7. WHEN slide number validation fails THEN error messages SHALL clearly indicate which slide numbers are invalid

### Requirement 6

**User Story:** As a developer, I want the query_slides tool to return zero results for invalid search criteria, so that I can get accurate filtered results.

#### Acceptance Criteria

1. WHEN search_criteria contains syntax errors THEN total_found SHALL be 0 and results SHALL be empty
2. WHEN search_criteria contains invalid field names THEN total_found SHALL be 0 and results SHALL be empty
3. WHEN search_criteria contains invalid operators THEN total_found SHALL be 0 and results SHALL be empty
4. WHEN search_criteria contains invalid values THEN total_found SHALL be 0 and results SHALL be empty
5. IF search_criteria validation fails THEN the system SHALL NOT return all slides as results
6. WHEN no slides match valid search criteria THEN total_found SHALL be 0 and results SHALL be empty
7. WHEN search criteria validation occurs THEN it SHALL happen before slide processing

### Requirement 7

**User Story:** As a developer, I want the query_slides tool to correctly filter slides by section, so that I can find slides within specific presentation sections.

#### Acceptance Criteria

1. WHEN search_criteria specifies a section name THEN only slides in that section SHALL be returned
2. WHEN search_criteria specifies a non-existent section THEN total_found SHALL be 0 and results SHALL be empty
3. WHEN section filtering is applied THEN slides outside the specified section SHALL be excluded
4. WHEN multiple sections exist THEN section filtering SHALL work correctly for each section
5. IF presentation has no sections THEN section-based queries SHALL handle this gracefully
6. WHEN section names contain special characters THEN filtering SHALL work correctly
7. WHEN section filtering is combined with other criteria THEN all criteria SHALL be applied correctly

### Requirement 8

**User Story:** As a developer, I want comprehensive integration tests to verify all bug fixes, so that I can ensure the fixes work correctly and prevent regressions.

#### Acceptance Criteria

1. WHEN integration tests are run THEN they SHALL test all identified bug scenarios using the test_complex.pptx file
2. WHEN formatting analysis tests run THEN they SHALL verify that formatting_counts return non-zero values for existing formatting
3. WHEN text formatting extraction tests run THEN they SHALL verify that italic and hyperlinks are correctly recognized
4. WHEN position accuracy tests run THEN they SHALL verify that start_position values are correct and formatted_segments contain only formatted text
5. WHEN table extraction tests run THEN they SHALL verify that summary values are accurate and extracted_tables contain data
6. WHEN slide number parameter tests run THEN they SHALL verify that valid slide numbers work correctly
7. WHEN query validation tests run THEN they SHALL verify that invalid search criteria return zero results
8. WHEN section filtering tests run THEN they SHALL verify that section-based queries work correctly
9. IF any test fails THEN the test SHALL provide detailed information about the failure and expected vs actual results
10. WHEN all tests pass THEN the bug fixes SHALL be considered complete and verified

### Requirement 9

**User Story:** As a developer, I want test data that covers all bug scenarios, so that I can reliably reproduce and verify the fixes.

#### Acceptance Criteria

1. WHEN test_complex.pptx is used THEN it SHALL contain slides with various formatting types (bold, italic, underlined, highlighted, strikethrough, colored text, hyperlinks)
2. WHEN test file is analyzed THEN it SHALL contain tables on specific slides for table extraction testing
3. WHEN test file is structured THEN it SHALL have sections for section-based query testing
4. WHEN test file is created THEN it SHALL have sufficient slides to test slide number parameter handling
5. IF test file needs specific content THEN manual creation or modification SHALL be performed to ensure proper test coverage
6. WHEN test scenarios are designed THEN they SHALL cover both positive and negative test cases
7. WHEN test data is documented THEN expected results SHALL be clearly specified for verification

### Requirement 10

**User Story:** As a developer, I want search and filter functions to support sections and notes criteria, so that I can filter slides based on presentation structure and speaker notes content.

#### Acceptance Criteria

1. WHEN query_slides search_criteria includes "sections" field THEN slides SHALL be filtered by section membership
2. WHEN query_slides search_criteria includes "notes" field THEN slides SHALL be filtered by speaker notes content
3. WHEN extract_table_data includes section filtering THEN only tables from slides in specified sections SHALL be extracted
4. WHEN extract_table_data includes notes filtering THEN only tables from slides with matching notes SHALL be extracted
5. WHEN section filtering is applied THEN section names SHALL be matched case-insensitively
6. WHEN notes filtering is applied THEN notes content SHALL support text matching and regular expressions
7. WHEN multiple filtering criteria are combined THEN all criteria SHALL be applied with AND logic

### Requirement 11

**User Story:** As a developer, I want query_slides to handle grammar errors properly, so that invalid search criteria return zero results instead of all slides.

#### Acceptance Criteria

1. WHEN search_criteria contains malformed JSON syntax THEN total_found SHALL be 0 and results SHALL be empty array
2. WHEN search_criteria contains invalid field references THEN total_found SHALL be 0 and results SHALL be empty array
3. WHEN search_criteria contains invalid operator syntax THEN total_found SHALL be 0 and results SHALL be empty array
4. WHEN search_criteria validation fails THEN error details SHALL be included in response
5. IF search_criteria is syntactically correct but semantically invalid THEN zero results SHALL be returned
6. WHEN grammar validation occurs THEN it SHALL happen before any slide processing
7. WHEN validation errors are detected THEN clear error messages SHALL indicate the specific syntax issues

### Requirement 12

**User Story:** As a developer, I want extract_table_data to display correct slide numbers in results, so that I can accurately identify which slides contain the extracted tables.

#### Acceptance Criteria

1. WHEN extract_table_data processes slide 10 THEN result slide_number SHALL be 10, not 2
2. WHEN extract_table_data processes multiple slides THEN each result SHALL show the correct source slide number
3. WHEN slide numbering is 1-based in API THEN internal processing SHALL maintain correct mapping
4. WHEN table extraction results are returned THEN slide_number field SHALL match the actual slide position in presentation
5. IF slide number mapping has off-by-one errors THEN they SHALL be corrected in all extraction paths
6. WHEN debugging slide number issues THEN logging SHALL show both internal and external slide number values
7. WHEN slide number validation occurs THEN it SHALL verify correct mapping between API and internal representations

### Requirement 13

**User Story:** As a developer, I want analyze_text_formatting and get_presentation_overview to include sections and notes information, so that I can get complete presentation structure analysis.

#### Acceptance Criteria

1. WHEN analyze_text_formatting is called THEN response SHALL include sections information for each slide
2. WHEN analyze_text_formatting is called THEN response SHALL include notes content for each slide
3. WHEN get_presentation_overview is called THEN response SHALL include complete sections structure
4. WHEN get_presentation_overview is called THEN response SHALL include notes summary statistics
5. WHEN sections information is included THEN section names and slide ranges SHALL be accurate
6. WHEN notes information is included THEN notes content SHALL be properly extracted and formatted
7. WHEN presentation has no sections THEN sections field SHALL be empty array or null
8. WHEN slides have no notes THEN notes field SHALL be empty string or null

### Requirement 14

**User Story:** As a developer, I want comprehensive integration tests for the new search and display improvements, so that I can verify all enhancements work correctly.

#### Acceptance Criteria

1. WHEN integration tests run THEN they SHALL test sections and notes filtering in query_slides
2. WHEN integration tests run THEN they SHALL test sections and notes filtering in extract_table_data
3. WHEN integration tests run THEN they SHALL test grammar error handling in query_slides
4. WHEN integration tests run THEN they SHALL test correct slide number display in extract_table_data
5. WHEN integration tests run THEN they SHALL test sections and notes inclusion in analyze_text_formatting
6. WHEN integration tests run THEN they SHALL test sections and notes inclusion in get_presentation_overview
7. WHEN test file is used THEN it SHALL contain appropriate sections and notes content for testing
8. IF any new functionality test fails THEN detailed error information SHALL be provided
9. WHEN all tests pass THEN the search and display improvements SHALL be considered complete
10. WHEN regression tests run THEN existing functionality SHALL remain unaffected by new features