# Design Document

## Overview

This design document outlines the enhancement of the PowerPoint MCP server to provide improved text formatting extraction capabilities and comprehensive testing infrastructure. The enhancements focus on generalizing the current `extract_bold_text` tool to support multiple formatting attributes and implementing robust MCP protocol-based integration testing.

## Architecture

### Current Architecture Analysis

The PowerPoint MCP server follows a layered architecture:

1. **MCP Server Layer** (`server.py`): Handles MCP protocol communication and tool registration
2. **Core Processing Layer**: Contains specialized processors for different content types
   - `ContentExtractor`: Main content extraction engine
   - `TextFormattingAnalyzer`: Current formatting analysis capabilities
   - `AttributeProcessor`: Filters and processes extracted attributes
3. **Utility Layer**: File validation, ZIP extraction, XML parsing
4. **FastMCP Integration** (`main.py`): Alternative FastMCP-based server implementation

### Enhanced Architecture

The enhanced architecture will extend the existing structure with:

1. **Generalized Formatting Extractor**: Replace `extract_bold_text` with `extract_text_formatting`
2. **Enhanced Response Format**: Include both complete text and formatted segments with positions
3. **Comprehensive Test Suite**: MCP protocol-based integration tests
4. **Test Data Generation**: Python-pptx based test file creation

## Components and Interfaces

### 1. Enhanced Text Formatting Extractor

#### New Tool: `extract_text_formatting`

```python
@mcp.tool
async def extract_text_formatting(
    file_path: str, 
    formatting_type: str,
    slide_numbers: Optional[List[int]] = None
) -> str:
    """Extract text with specific formatting attributes from PowerPoint slides.
    
    Args:
        file_path: Path to the PowerPoint file (.pptx)
        formatting_type: Type of formatting to extract. Valid values:
            - "bold": Bold text segments
            - "italic": Italic text segments  
            - "underlined": Underlined text segments
            - "highlighted": Highlighted text segments
            - "strikethrough": Strikethrough text segments
            - "hyperlinks": Hyperlink text and URLs
            - "font_sizes": Text segments with font size information
            - "font_colors": Text segments with color information
        slide_numbers: Slide numbers to analyze (optional, analyzes all if not specified)
        
    Returns:
        JSON string containing:
        {
            "file_path": str,
            "formatting_type": str,
            "summary": {
                "total_slides_analyzed": int,
                "slides_with_formatting": int,
                "total_formatted_segments": int
            },
            "results_by_slide": [
                {
                    "slide_number": int,
                    "title": str,
                    "complete_text": str,
                    "formatted_segments": [
                        {
                            "text": str,
                            "start_position": int,
                            "end_position": int,
                            "formatting_details": dict
                        }
                    ]
                }
            ]
        }
    """
```

#### Enhanced FormattingExtractor Class

```python
class FormattingExtractor:
    """Enhanced formatting extraction with position tracking."""
    
    def extract_formatting_segments(
        self, 
        text_elements: List[Dict], 
        formatting_type: str
    ) -> List[FormattedSegment]:
        """Extract formatted segments with position information."""
        
    def calculate_positions(
        self, 
        complete_text: str, 
        segments: List[str]
    ) -> List[Tuple[int, int]]:
        """Calculate start/end positions for formatted segments."""
        
    def extract_hyperlink_details(
        self, 
        text_element: Dict
    ) -> List[HyperlinkSegment]:
        """Extract hyperlink text and URLs."""
```

### 2. Enhanced Response Format

#### Current Format (extract_bold_text)
```json
{
    "bold_elements": [
        {
            "content": "Full text content",
            "bold_count": 3
        }
    ]
}
```

#### New Enhanced Format
```json
{
    "results_by_slide": [
        {
            "slide_number": 1,
            "complete_text": "This is bold text and this is normal text",
            "format": "bold",
            "formatted_segments": [
                {
                    "text": "bold text",
                    "start_position": 8
                }
            ]
        }
    ]
}
```

### 3. MCP Integration Testing Framework

#### Test Architecture

```python
class MCPIntegrationTestSuite:
    """Comprehensive MCP protocol testing."""
    
    def __init__(self):
        self.client = MCPClient()
        self.test_files = TestFileManager()
        
    async def test_all_tools(self):
        """Test all MCP tools with various parameter combinations."""
        
    async def test_tool_with_parameters(
        self, 
        tool_name: str, 
        parameter_sets: List[Dict]
    ):
        """Test a specific tool with multiple parameter sets."""
        
    def generate_test_report(self) -> Dict:
        """Generate comprehensive test coverage report."""
```

#### Test Client Implementation

```python
from fastmcp.client.transports import StdioClientTransport

class MCPTestClient:
    """MCP client for integration testing."""
    
    async def connect_to_server(self, server_command: List[str]):
        """Establish MCP connection to server."""
        
    async def call_tool(self, tool_name: str, arguments: Dict) -> Dict:
        """Call MCP tool and return response."""
        
    async def list_available_tools(self) -> List[str]:
        """Get list of available tools from server."""
```

### 4. Test Data Generation System

#### Python-pptx Test File Generator

```python
class TestPresentationGenerator:
    """Generate test PowerPoint files with known formatting."""
    
    def create_formatting_test_file(self) -> str:
        """Create PowerPoint with all supported formatting types."""
        
    def add_bold_text_slide(self, presentation):
        """Add slide with various bold text patterns."""
        
    def add_hyperlink_slide(self, presentation):
        """Add slide with hyperlinks."""
        
    def add_mixed_formatting_slide(self, presentation):
        """Add slide with overlapping formatting."""
        
    def document_expected_results(self, file_path: str) -> Dict:
        """Document expected extraction results for test file."""
```

## Data Models

### FormattedSegment

```python
@dataclass
class FormattedSegment:
    """Represents a formatted text segment with position."""
    text: str
    start_position: int
    end_position: int
    formatting_type: str
    formatting_details: Dict[str, Any]
    slide_number: int
    element_index: int
```

### HyperlinkSegment

```python
@dataclass
class HyperlinkSegment(FormattedSegment):
    """Hyperlink-specific formatted segment."""
    url: str
    display_text: str
    link_type: str  # "external", "internal", "email"
```

### TestResult

```python
@dataclass
class TestResult:
    """Result of MCP tool test execution."""
    tool_name: str
    parameters: Dict[str, Any]
    success: bool
    response_time: float
    response_data: Optional[Dict]
    error_message: Optional[str]
```

## Error Handling

### Validation Errors

1. **Invalid formatting_type**: Return error with list of valid options
2. **File not found**: Standard file validation error
3. **Slide number out of range**: Clear error message with valid range
4. **Malformed PowerPoint**: Graceful degradation with partial results

### MCP Protocol Errors

1. **Connection failures**: Retry logic with exponential backoff
2. **Tool not found**: Clear error reporting in test results
3. **Parameter validation**: Detailed parameter error messages
4. **Timeout handling**: Configurable timeout with clear error reporting

## Testing Strategy

### Unit Tests

1. **FormattingExtractor**: Test each formatting type extraction
2. **Position Calculation**: Verify accurate position tracking
3. **Response Format**: Validate JSON structure and content
4. **Error Handling**: Test all error conditions

### Integration Tests

1. **MCP Protocol**: Full client-server communication testing
2. **Tool Coverage**: Test every tool with valid parameters
3. **Error Scenarios**: Test invalid parameters and edge cases
4. **Performance**: Measure response times and resource usage

### Test File Strategy

1. **Generated Files**: Use python-pptx for consistent test data
2. **Known Content**: Document expected results for each test file
3. **Edge Cases**: Empty slides, complex layouts, mixed formatting
4. **Manual Additions**: Request human assistance for unsupported attributes

### Test Execution Framework

```python
class TestExecutor:
    """Orchestrates comprehensive test execution."""
    
    async def run_full_test_suite(self) -> TestReport:
        """Execute all tests and generate report."""
        
    def cleanup_obsolete_tests(self):
        """Remove outdated test files and code."""
        
    def validate_test_files(self) -> List[str]:
        """Verify test files match current specifications."""
```

## Implementation Plan Integration

The implementation will be structured to maintain compatibility with the existing codebase while adding new capabilities:

1. **Extend existing TextFormattingAnalyzer** rather than replacing it
2. **Add new MCP tool** alongside existing tools
3. **Enhance response formats** while maintaining backward compatibility initially
4. **Implement comprehensive testing** as a separate test suite
5. **Generate test data** using automated python-pptx scripts

This design ensures minimal disruption to existing functionality while providing the enhanced capabilities requested.