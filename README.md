# PowerPoint Analyzer MCP Server
An MCP server that enables search and extraction using PowerPoint structure and text formatting attributes.

## Background
Most AI Agent searches that claim PowerPoint support typically ignore PowerPoint file structure and only extract text for searching.
This tool enables outputting text written in bold and other structured information.

## Features

- **Text formatting detection**: Bold, italic, underline, strikethrough, highlighting, hyperlinks
- **Font analysis**: Font sizes, colors, and styling information
- **Slide querying**: Query slides with flexible filtering criteria
- **Table data extraction**: Extract table data with formatting detection
- **Testing suite** for formatting detection validation
- Implementation using Python standard libraries (no external PowerPoint dependencies)
- Direct XML parsing for processing
- Built with FastMCP 2.0

## Project Structure

```
powerpoint-analyzer/
├── main.py                     # Main FastMCP server entry point
├── powerpoint_mcp_server/      # Core server implementation
│   ├── server.py              # Main MCP server implementation
│   ├── config.py              # Configuration management
│   ├── core/                  # Core functionality
│   └── utils/                 # Utility modules
├── tests/                      # Test files
│   ├── test_powerpoint_fastmcp.py  # Main server tests
│   ├── test_formatting_detection.py # Formatting detection tests
│   └── ...                         # Other test files
├── scripts/                    # Utility scripts
│   ├── health_check.py        # Server health check
│   └── start_server.py        # Alternative server startup
├── requirements.txt            # Python dependencies
├── pytest.ini                 # Test configuration
└── README.md                   # Documentation
```

## Installation

1. Clone the repository:
```bash
git clone <repository-url>
cd powerpoint-analyzer
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

3. Configure your AI agent (Claude Desktop, etc.) by adding the following to your configuration file:

**Location of mcp_settings.json:**
- **macOS**: `~/Library/Application Support/Claude/mcp_settings.json`
- **Windows**: `%APPDATA%\Claude\mcp_settings.json`

```json
{
  "mcpServers": {
    "powerpoint-analyzer-mcp": {
      "command": "python",
      "args": ["/path/to/your/powerpoint-analyzer/main.py"],
      "env": {
        "POWERPOINT_MCP_LOG_LEVEL": "INFO"
      }
    }
  }
}
```

**Example with actual paths:**

macOS/Linux:
```json
{
  "mcpServers": {
    "powerpoint-analyzer-mcp": {
      "command": "python",
      "args": ["/Users/username/powerpoint-analyzer/main.py"],
      "env": {
        "POWERPOINT_MCP_LOG_LEVEL": "INFO"
      }
    }
  }
}
```

Windows:
```json
{
  "mcpServers": {
    "powerpoint-analyzer-mcp": {
      "command": "python",
      "args": ["C:\\Users\\username\\powerpoint-analyzer\\main.py"],
      "env": {
        "POWERPOINT_MCP_LOG_LEVEL": "INFO"
      }
    }
  }
}
```

## Technical Approach

This server processes PowerPoint files using the following approach:

- **Direct ZIP handling**: .pptx files are processed as ZIP archives using Python's `zipfile` module
- **XML parsing**: Internal PowerPoint XML structure is parsed using `xml.etree.ElementTree` with namespace support
- **Dual formatting detection**: Supports both XML attribute and child element formats for text formatting properties
- **No external dependencies**: Uses only Python standard library modules for PowerPoint processing
- **Processing**: Extracts only the required information without loading entire presentations into memory
- **Caching**: Caching system for performance on repeated operations

## Text Formatting Detection

The server provides text formatting detection capabilities:

### Supported Formatting Types
- **Bold text**: Detects bold formatting in text elements
- **Italic text**: Identifies italic styling across slides
- **Underlined text**: Finds underlined text with underline styles
- **Strikethrough text**: Detects strikethrough formatting
- **Highlighted text**: Identifies highlighted/background colored text
- **Hyperlinks**: Extracts hyperlink information and relationship IDs
- **Font properties**: Analyzes font sizes, colors (RGB and scheme colors)

### Technical Implementation
- **Dual detection method**: Checks both XML attributes (`b="1"`) and child elements (`<a:b val="1"/>`)
- **Namespace-aware parsing**: Handling of Office Open XML namespaces
- **Validation**: Test suite for detection across different PowerPoint versions
- **Debug capabilities**: Debugging tools for troubleshooting formatting detection issues

### Validation and Testing
- **Test suite**: `tests/test_formatting_detection.py` validates formatting types
- **Debug tools**: `tests/debug_formatting_detection.py` provides XML analysis
- **Validation**: Tested with PowerPoint files containing mixed formatting

## Usage

### Running the Server

```bash
python main.py
```

### Available Tools

This MCP server provides three core tools:

1. **extract_formatted_text**: Extract text with specific formatting types (bold, italic, underline, strikethrough, highlight, hyperlinks, font sizes, font colors)
2. **query_slides**: Query slides with flexible filtering criteria
3. **extract_table_data**: Extract table data with flexible selection and formatting detection

### Slide Selection with Python-Style Slicing

All tools now support flexible slide selection using Python-style slicing notation for the `slide_numbers` parameter:

#### Supported Formats

- **All slides**: `None` or omit the parameter
- **Single slide**: `3` or `"3"`
- **Specific slides**: `[1, 5, 10]` or `"1,5,10"`
- **First N slides**: `":10"` (slides 1-10)
- **Slide range**: `"5:20"` (slides 5-20)
- **From slide to end**: `"25:"` (slides 25 to end)
- **With brackets**: `"[:10]"`, `"[5:20]"`, `"[25:]"` (optional)

#### Examples

```python
# Extract tables from first 10 slides
extract_table_data("presentation.pptx", slide_numbers=":10")

# Extract bold text from slides 5-20
extract_formatted_text("presentation.pptx", "bold", slide_numbers="5:20")

# Query specific slides
query_slides("presentation.pptx", {
    "slide_numbers": "1,3,5,10",
    "title": {"contains": "Summary"}
})

# Extract tables from slide 25 to end
extract_table_data("presentation.pptx", slide_numbers="25:")
```

#### Benefits

- **Efficient processing**: Process only the slides you need
- **Intuitive syntax**: Familiar Python slicing notation
- **Flexible selection**: Mix and match different selection methods
- **Performance optimization**: Reduce processing time for large presentations



## Development

### Requirements

- Python 3.8+
- MCP (Model Context Protocol)
- FastMCP 2.0
- Standard Python libraries (zipfile, xml.etree.ElementTree)

## Recent Updates

### Version 2.1 - Python-Style Slide Selection
- **Enhanced slide selection**: Added support for Python-style slicing notation (`:10`, `5:20`, `25:`, etc.)
- **Flexible slide specification**: Support for single slides, ranges, comma-separated lists, and slicing
- **Improved performance**: Process only the slides you need with efficient selection
- **Backward compatibility**: Existing `[1, 5, 10]` format still supported
- **Comprehensive validation**: Robust error handling for invalid slide specifications

### Version 2.0 - Advanced Text Formatting Detection
- **Fixed critical formatting detection bug**: Bold, italic, underline, and strikethrough attributes now correctly detected
- **Dual detection support**: Handles both XML attribute and child element formats
- **Comprehensive test suite**: Added extensive tests for all formatting types
- **Debug tools**: New debugging utilities for troubleshooting formatting issues
- **Enhanced MCP tools**: New tools for formatted text extraction and analysis

### Formatting Detection Validation
All formatting types have been thoroughly tested and validated:
- ✅ **Bold text**: 8 elements detected in test files
- ✅ **Italic text**: 6 elements detected in test files
- ✅ **Underlined text**: 6 elements detected in test files
- ✅ **Strikethrough text**: 6 elements detected in test files
- ✅ **Highlighted text**: 5 elements detected in test files
- ✅ **Hyperlinks**: 2 elements detected in test files
- ✅ **Font colors**: Multiple colors detected and analyzed

## License


This project is licensed under the Apache License 2.0 - see the [LICENSE](LICENSE) file for details.
