# PowerPoint MCP Server

A Model Context Protocol (MCP) server for extracting structured information from PowerPoint files using FastMCP 2.0.

## Features

- Extract complete structured content from PowerPoint (.pptx) files
- Get specific attributes from slides (title, subtitle, text, tables, images, etc.)
- Retrieve information for individual slides
- Query slides with flexible filtering criteria
- Extract table data with formatting detection
- Get comprehensive presentation overview and analysis
- Support for slide layouts, placeholders, and formatting information
- Lightweight implementation using Python standard libraries (no external PowerPoint dependencies)
- Direct XML parsing for fast and efficient processing
- Built with FastMCP 2.0 for optimal performance and compatibility

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

This server uses a lightweight approach to PowerPoint processing:

- **Direct ZIP handling**: .pptx files are processed as ZIP archives using Python's `zipfile` module
- **XML parsing**: Internal PowerPoint XML structure is parsed using `xml.etree.ElementTree`
- **No external dependencies**: Uses only Python standard library modules for PowerPoint processing
- **Efficient processing**: Extracts only the required information without loading entire presentations into memory

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

1. **extract_powerpoint_content**: Extract complete structured content from a PowerPoint file
2. **get_powerpoint_attributes**: Get specific attributes from PowerPoint slides
3. **get_slide_info**: Get information for a specific slide

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

You can also use the enhanced startup script for more configuration options:

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
│   ├── content_extractor.py    # PowerPoint content extraction
│   ├── attribute_processor.py  # Attribute filtering and processing
│   └── xml_parser.py           # XML parsing utilities
└── utils/
    ├── __init__.py
    ├── file_validator.py       # File validation
    ├── zip_extractor.py        # ZIP archive handling
    └── cache_manager.py        # Caching utilities
```

### Requirements

- Python 3.8+
- MCP (Model Context Protocol)
- Standard Python libraries (zipfile, xml.etree.ElementTree)

## License

TBD