# Requirements Document

## Introduction

PowerPoint MCP サーバの機能を拡張し、テキストフォーマット抽出機能の改善と包括的なテスト体制の構築を行います。現在の `extract_bold_text` ツールを一般化して複数のフォーマット属性に対応し、より詳細な抽出結果を提供します。また、MCP プロトコルを使用した結合テストを実装し、全ツールの動作確認を行える体制を整備します。

## Requirements

### Requirement 1

**User Story:** As a developer, I want to extract specific text formatting attributes from PowerPoint slides, so that I can analyze text styling patterns and preserve formatting information.

#### Acceptance Criteria

1. WHEN bold text is extracted THEN the system SHALL return both the complete text content and an array of bold text segments with their positions
2. WHEN a formatting attribute is specified THEN the system SHALL support bold, italic, underlined, highlighted, strikethrough, hyperlinks, font_sizes, and font_colors
3. WHEN multiple bold segments exist in a text area THEN the system SHALL return all segments as separate array elements
4. IF no formatting of the specified type is found THEN the system SHALL return an empty array for formatted segments
5. WHEN text content is extracted THEN the system SHALL preserve the original text structure and line breaks

### Requirement 2

**User Story:** As a developer, I want to use a generalized text formatting extraction tool, so that I can extract any supported formatting attribute with a single consistent interface.

#### Acceptance Criteria

1. WHEN the formatting extraction tool is called THEN it SHALL accept a formatting_type parameter specifying the desired attribute
2. WHEN formatting_type is "bold" THEN the system SHALL extract bold text segments
3. WHEN formatting_type is "italic" THEN the system SHALL extract italic text segments
4. WHEN formatting_type is "underlined" THEN the system SHALL extract underlined text segments
5. WHEN formatting_type is "highlighted" THEN the system SHALL extract highlighted text segments
6. WHEN formatting_type is "strikethrough" THEN the system SHALL extract strikethrough text segments
7. WHEN formatting_type is "hyperlinks" THEN the system SHALL extract hyperlink text and URLs
8. WHEN formatting_type is "font_sizes" THEN the system SHALL extract text segments with their font sizes
9. WHEN formatting_type is "font_colors" THEN the system SHALL extract text segments with their color information
10. IF an invalid formatting_type is provided THEN the system SHALL return an error with valid options listed

### Requirement 3

**User Story:** As a developer, I want clear and comprehensive tool documentation, so that I can understand exactly what parameters are available and what the response format will be.

#### Acceptance Criteria

1. WHEN tool descriptions are provided THEN they SHALL include specific examples of valid parameter values
2. WHEN parameter documentation is written THEN it SHALL list all acceptable values for enumerated parameters
3. WHEN return value documentation is created THEN it SHALL specify the exact structure and data types of response objects
4. WHEN tool summaries are written THEN they SHALL include representative usage examples
5. IF a tool has optional parameters THEN the documentation SHALL clearly indicate default values and behavior

### Requirement 4

**User Story:** As a developer, I want to clean up obsolete test code, so that the test suite remains maintainable and only contains relevant tests.

#### Acceptance Criteria

1. WHEN test files are reviewed THEN obsolete tests that no longer match current specifications SHALL be identified
2. WHEN obsolete tests are found THEN they SHALL be removed from the test suite
3. WHEN test cleanup is performed THEN remaining tests SHALL be verified to work with current implementation
4. IF test files contain useful test data or patterns THEN those elements SHALL be preserved or migrated to current tests
5. WHEN cleanup is complete THEN the test directory SHALL contain only functional and relevant test files

### Requirement 5

**User Story:** As a developer, I want comprehensive integration tests using the MCP protocol, so that I can verify the server works correctly with real MCP clients.

#### Acceptance Criteria

1. WHEN integration tests are implemented THEN they SHALL use fastmcp.client.transports for MCP communication
2. WHEN MCP protocol tests are run THEN they SHALL establish actual client-server communication
3. WHEN test PowerPoint files are used THEN they SHALL contain known, predefined content for verification
4. WHEN all tools are tested THEN each tool SHALL be tested with all available parameter combinations
5. WHEN integration tests run THEN they SHALL verify both successful responses and error handling
6. IF MCP communication fails THEN tests SHALL provide clear diagnostic information
7. WHEN test results are generated THEN they SHALL include coverage reports for all MCP tools and options

### Requirement 6

**User Story:** As a developer, I want standardized test PowerPoint files with known content, so that I can create reliable and repeatable tests.

#### Acceptance Criteria

1. WHEN test PowerPoint files are created THEN they SHALL be generated using python-pptx library to ensure consistency
2. WHEN test files are designed THEN they SHALL include edge cases like empty slides, complex layouts, and mixed formatting
3. WHEN test content is defined THEN it SHALL be documented with expected extraction results
4. IF test files are modified THEN the documentation SHALL be updated to reflect changes
5. WHEN tests use these files THEN they SHALL verify exact matches against documented expected results
6. IF python-pptx cannot set certain formatting attributes THEN human assistance SHALL be requested to manually add those attributes

### Requirement 7

**User Story:** As a developer, I want automated verification of all MCP tool functionality, so that I can ensure comprehensive coverage and catch regressions.

#### Acceptance Criteria

1. WHEN automated tests are run THEN they SHALL test every available MCP tool
2. WHEN tool parameters are tested THEN all valid parameters SHALL be exercised
3. WHEN error conditions are tested THEN invalid parameters and edge cases SHALL be verified
4. WHEN test coverage is measured THEN it SHALL achieve at least 90% coverage of MCP tool code paths
5. IF any tool fails testing THEN the test suite SHALL provide detailed failure information and reproduction steps

### Requirement 8

**User Story:** As a developer, I want detailed formatting position information, so that I can reconstruct the original text layout and styling.

#### Acceptance Criteria

1. WHEN formatted text segments are extracted THEN each segment SHALL include start and end character positions
2. WHEN position information is provided THEN it SHALL be relative to the complete text content
3. WHEN overlapping formatting exists THEN the system SHALL handle multiple attributes on the same text correctly
4. WHEN character positions are calculated THEN they SHALL be consistent across different text encoding scenarios