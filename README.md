# PowerPoint MCP Server

A Model Context Protocol (MCP) server for extracting structured information from PowerPoint files using FastMCP 2.0.

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
├── tests/                      # Test files
│   ├── test_powerpoint_fastmcp.py  # Main server tests
│   ├── test_simple_fastmcp.py      # Simple FastMCP tests
│   └── ...                         # Other test files
├── tools/                      # Utility tools
│   ├── fastmcp_example.py          # FastMCP usage examples
│   └── start_mcp_server.py         # Server startup utility
├── temp/                       # Temporary/archived files
└── README.md
```

## Installation

### For Standalone Use

1. Clone the repository:
```bash
git clone <repository-url>
cd powerpoint-analyzer
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

3. Install FastMCP:
```bash
pip install fastmcp
```

4. Run the server:
```bash
python main.py
pip install -e .
```

### For AI Agent Integration

1. Follow the standalone installation steps above

2. Note the absolute path to your installation directory:
```bash
pwd
# Example output: /Users/username/powerpoint-analyzer
```

3. Configure your AI agent using the path from step 2 (see [AI Agent Integration](#ai-agent-integration) section below)

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

Or using the installed console script:
```bash
powerpoint-mcp-server
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

## AI Agent Integration

This MCP server can be integrated with AI agents that support the Model Context Protocol, such as Claude Desktop, Claude Code, and other MCP-compatible applications.

### Claude Desktop Configuration

Add the following configuration to your Claude Desktop `mcp_settings.json` file:

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
        "POWERPOINT_MCP_LOG_LEVEL": "INFO",
        "POWERPOINT_MCP_MAX_FILE_SIZE": "100",
        "POWERPOINT_MCP_CACHE_ENABLED": "true"
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

### Claude Code Configuration

For Claude Code, create or update your `mcp_settings.json`:

```json
{
  "mcpServers": {
    "powerpoint-analyzer": {
      "command": "python",
      "args": ["/absolute/path/to/powerpoint-analyzer/main.py"],
      "env": {
        "POWERPOINT_MCP_LOG_LEVEL": "DEBUG",
        "POWERPOINT_MCP_DEBUG": "true"
      }
    }
  }
}
```

### Alternative: Using the Startup Script

You can also use the startup script for configuration options:

```json
{
  "mcpServers": {
    "powerpoint-mcp-server": {
      "command": "python",
      "args": [
        "/path/to/your/powerpoint-analyzer/scripts/start_server.py",
        "--log-level", "INFO",
        "--max-file-size", "150"
      ]
    }
  }
}
```

### Configuration Options

When integrating with AI agents, you can customize the server behavior using environment variables:

| Environment Variable | Default | Description |
|---------------------|---------|-------------|
| `POWERPOINT_MCP_LOG_LEVEL` | `INFO` | Logging level (DEBUG, INFO, WARNING, ERROR, CRITICAL) |
| `POWERPOINT_MCP_MAX_FILE_SIZE` | `100` | Maximum file size in MB |
| `POWERPOINT_MCP_TIMEOUT` | `300` | Processing timeout in seconds |
| `POWERPOINT_MCP_CACHE_ENABLED` | `true` | Enable/disable caching |
| `POWERPOINT_MCP_DEBUG` | `false` | Enable debug mode |

### Usage Examples with AI Agents

Once configured, you can ask your AI agent to:

1. **Analyze a presentation structure**:
   ```
   Please extract the complete content from my presentation at /Users/username/Documents/quarterly-report.pptx
   ```

2. **Get specific slide information**:
   ```
   Can you get the title and text content from slide 3 of /Users/username/Documents/quarterly-report.pptx?
   ```

3. **Extract specific attributes**:
   ```
   Please get only the titles and object counts from all slides in /Users/username/Documents/quarterly-report.pptx
   ```

4. **Analyze presentation metadata**:
   ```
   What are the slide dimensions and total number of slides in /Users/username/Documents/quarterly-report.pptx?
   ```

5. **Extract tables and structured data**:
   ```
   Please extract all tables from /Users/username/Documents/quarterly-report.pptx and show me their content
   ```

6. **Analyze text formatting**:
   ```
   Please extract all bold text from /Users/username/Documents/quarterly-report.pptx and show me which slides they appear on
   ```

7. **Get formatting summary**:
   ```
   Can you analyze all text formatting (bold, italic, underline, etc.) in /Users/username/Documents/quarterly-report.pptx?
   ```

8. **Extract specific formatting types**:
   ```
   Please extract all text with bold and italic formatting from /Users/username/Documents/quarterly-report.pptx
   ```

### Troubleshooting AI Agent Integration

1. **Server not starting**: Check the path in your `mcp_settings.json` is absolute and correct
2. **Permission errors**: Ensure the Python executable and script files have proper permissions
3. **File access issues**: Make sure the AI agent has access to the PowerPoint files you want to analyze
4. **Debug information**: Set `POWERPOINT_MCP_DEBUG=true` for detailed logging

### Health Check

Before configuring with an AI agent, verify the server works correctly:

```bash
python scripts/health_check.py
```

This will validate all dependencies and configuration settings.

### Verifying AI Agent Integration

After configuring your AI agent:

1. **Restart your AI agent** (Claude Desktop, Claude Code, etc.)

2. **Check if the server is recognized**: Ask your AI agent:
   ```
   What MCP tools do you have available?
   ```
   You should see the PowerPoint MCP server tools listed.

3. **Test with a sample file**: Try extracting content from a PowerPoint file to ensure everything works correctly.

4. **Check logs**: If issues occur, check the AI agent's logs or enable debug mode:
   ```json
   "env": {
     "POWERPOINT_MCP_LOG_LEVEL": "DEBUG",
     "POWERPOINT_MCP_DEBUG": "true"
   }
   ```

## Development

### Project Structure

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
├── utils/
│   ├── __init__.py
│   ├── file_validator.py       # File validation
│   ├── zip_extractor.py        # ZIP archive handling
│   └── cache_manager.py        # Caching utilities
└── tests/
    ├── test_formatting_detection.py  # Formatting detection tests
    ├── debug_formatting_detection.py # Debug tools for formatting issues
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