# PowerPoint Analyze MCP Server
An MCP server that enables search and extraction using PowerPoint structure and text formatting attributes.

## Background
Most AI Agent searches that claim PowerPoint support typically ignore PowerPoint file structure and only extract text for searching.
This tool enables outputting text written in bold and other structured information.

## Features

- Extract structured content from PowerPoint (.pptx) files
- Get specific attributes from slides (title, subtitle, text, tables, images, etc.)
- Retrieve information for individual slides
- Query slides with filtering criteria
- Extract table data with formatting detection
- **Text formatting detection**: Bold, italic, underline, strikethrough, highlighting, hyperlinks
- **Font analysis**: Font sizes, colors, and styling information
- Get presentation overview and analysis
- Support for slide layouts, placeholders, and formatting information
- **Testing suite** for formatting detection validation
- **Debug tools** for troubleshooting formatting detection issues
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
    "powerpoint-mcp-server": {
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
    "powerpoint-mcp-server": {
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
    "powerpoint-mcp-server": {
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

#### Core Content Extraction
1. **extract_powerpoint_content**: Extract complete structured content from a PowerPoint file
2. **get_powerpoint_attributes**: Get specific attributes from PowerPoint slides
3. **get_slide_info**: Get information for a specific slide
4. **query_slides**: Query slides with flexible filtering criteria
5. **extract_table_data**: Extract table data with flexible selection and formatting detection

#### Text Formatting Analysis
6. **extract_bold_text**: Extract bold text from slides with location information
7. **extract_formatted_text**: Extract text with specific formatting types (bold, italic, underline, strikethrough, highlight, hyperlinks)
8. **get_formatting_summary**: Get summary of text formatting in the presentation
9. **analyze_text_formatting**: Analyze text formatting patterns across slides with formatting detection

#### Presentation Analysis
10. **get_presentation_overview**: Get presentation overview and analysis
11. **clear_cache**: Clear the cache
12. **reload_file_content**: Reload file content by clearing cache and re-extracting



## Development

### Detailed Structure

```
powerpoint_mcp_server/
├── __init__.py
├── server.py              # Main MCP server implementation
├── config.py              # Configuration management
├── core/
│   ├── __init__.py
│   ├── content_extractor.py    # PowerPoint content extraction with formatting detection
│   ├── attribute_processor.py  # Attribute filtering and processing
│   ├── presentation_analyzer.py # Presentation analysis
│   └── xml_parser.py           # XML parsing utilities
└── utils/
    ├── __init__.py
    ├── file_validator.py       # File validation
    ├── zip_extractor.py        # ZIP archive handling
    └── cache_manager.py        # Caching utilities

tests/
├── test_formatting_detection.py  # Formatting detection tests
├── debug_formatting_detection.py # Debug tools for formatting issues
├── test_powerpoint_fastmcp.py    # FastMCP server tests
└── ...                           # Other test files
```

### Requirements

- Python 3.8+
- MCP (Model Context Protocol)
- FastMCP 2.0
- Standard Python libraries (zipfile, xml.etree.ElementTree)

## Recent Updates

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

TBD