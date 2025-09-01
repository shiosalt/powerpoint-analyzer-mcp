# PowerPoint MCP Server Configuration

This document describes the configuration options and startup procedures for the PowerPoint MCP Server using FastMCP 2.0.

## Quick Start

### Basic Usage
```bash
# Start the server with default configuration
python main.py
```

### With Custom Configuration
```bash
# Set environment variables before starting
export POWERPOINT_MCP_LOG_LEVEL=DEBUG
export POWERPOINT_MCP_MAX_FILE_SIZE=200
export POWERPOINT_MCP_TIMEOUT=600
python main.py

# Or set them inline
POWERPOINT_MCP_DEBUG=true python main.py
```

## Configuration Options

### Environment Variables

The server can be configured using the following environment variables:

| Variable | Default | Description |
|----------|---------|-------------|
| `POWERPOINT_MCP_LOG_LEVEL` | `INFO` | Logging level (DEBUG, INFO, WARNING, ERROR, CRITICAL) |
| `POWERPOINT_MCP_MAX_FILE_SIZE` | `100` | Maximum file size in MB |
| `POWERPOINT_MCP_TIMEOUT` | `300` | Processing timeout in seconds |
| `POWERPOINT_MCP_CACHE_ENABLED` | `true` | Enable/disable caching |
| `POWERPOINT_MCP_CACHE_TTL` | `3600` | Cache TTL in seconds |
| `POWERPOINT_MCP_DEBUG` | `false` | Enable debug mode |

### FastMCP 2.0 Integration

This server uses FastMCP 2.0 framework for MCP protocol handling:

- **Automatic tool registration**: Tools are automatically registered using decorators
- **Lifespan management**: Server lifecycle is managed through async context managers
- **Stdio communication**: Server communicates via stdin/stdout for MCP protocol
- **Error handling**: Built-in error handling and response formatting

## Server Lifecycle

### Startup Sequence

1. **Configuration Loading**: Load configuration from environment variables using `ServerConfig`
2. **Logging Setup**: Configure file and console logging with specified log level
3. **FastMCP Initialization**: Create FastMCP instance with lifespan management
4. **Tool Registration**: Register all MCP tools using `@mcp.tool` decorators
5. **Server Components**: Initialize PowerPointMCPServer and related components
6. **Server Start**: Start FastMCP server and begin accepting MCP requests via stdio

### Shutdown Sequence

1. **Signal Reception**: Receive shutdown signal (SIGINT or SIGTERM)
2. **Lifespan Cleanup**: Execute lifespan context manager cleanup
3. **Resource Cleanup**: Clear caches and close file handles
4. **Server Stop**: Terminate FastMCP server
5. **Exit**: Clean exit with appropriate exit code

## Health Checks

Use the health check script to verify server readiness:

```bash
python scripts/health_check.py
```

The health check verifies:
- Required Python dependencies are available
- Configuration loads successfully
- Server components (FileValidator, ContentExtractor) can be initialized
- Test PowerPoint files can be validated (if available)

Example output:
```
PowerPoint MCP Server Health Check
========================================

Dependencies:
✓ All required Python modules are available

Configuration:
✓ Server configuration loaded successfully
  - Server name: powerpoint-mcp-server
  - Server version: 0.1.0
  - Log level: INFO
  - Max file size: 100 MB
  - Cache enabled: True
  - Debug mode: False

Components:
✓ FileValidator initialized successfully
✓ ContentExtractor initialized successfully

Test File Processing:
ℹ No test PowerPoint files found (this is optional)

========================================
✓ All health checks passed! Server should be ready to run.
```

## Logging

### Log Levels

- **DEBUG**: Detailed information for debugging
- **INFO**: General information about server operation
- **WARNING**: Warning messages for potential issues
- **ERROR**: Error messages for failed operations
- **CRITICAL**: Critical errors that may cause server shutdown

### Log Format

Default log format: `%(asctime)s - %(name)s - %(levelname)s - %(message)s`

### Log Files

- **File logging**: All logs are written to `powerpoint_mcp_server.log` (DEBUG level)
- **Console logging**: Logs are also output to stderr (configurable level)

Example log output:
```
2024-01-15 10:30:45,123 - __main__ - INFO - Starting PowerPoint MCP Server using FastMCP 2.0: powerpoint-mcp-server v0.1.0
2024-01-15 10:30:45,124 - __main__ - INFO - Log file: powerpoint_mcp_server.log
2024-01-15 10:30:45,125 - __main__ - INFO - FastMCP 2.0 server configured with tools
2024-01-15 10:30:45,126 - __main__ - INFO - Starting FastMCP 2.0 server...
```

## Error Handling

The server implements comprehensive error handling:

### File Access Errors
- File not found
- Permission denied
- File in use

### File Format Errors
- Unsupported file format
- Corrupted PowerPoint file
- Encrypted file

### Processing Errors
- Memory exhaustion
- Processing timeout
- Unexpected file structure

### MCP Protocol Errors
- Invalid tool parameters
- Unsupported tool calls
- JSON serialization errors
- FastMCP framework errors

## Performance Tuning

### File Size Limits
- Default: 100 MB
- Recommended: 50-200 MB depending on available memory
- Large files may require increased timeout values

### Timeout Settings
- Default: 300 seconds (5 minutes)
- Recommended: 60-600 seconds depending on file complexity
- Complex presentations with many slides may require longer timeouts

### Caching
- Default: Enabled with 1-hour TTL
- Improves performance for repeated requests
- Can be disabled for development or testing

## Troubleshooting

### Common Issues

1. **Server won't start**
   - Check dependencies with health check script
   - Verify configuration values
   - Check log output for specific errors

2. **File processing fails**
   - Verify file format (.pptx only)
   - Check file size limits
   - Ensure file is not corrupted or encrypted

3. **Performance issues**
   - Increase timeout values
   - Enable caching
   - Reduce file size limits
   - Check available memory

### Debug Mode

Enable debug mode for detailed troubleshooting:
```bash
POWERPOINT_MCP_DEBUG=true POWERPOINT_MCP_LOG_LEVEL=DEBUG python main.py
```

Debug mode provides:
- Verbose logging (DEBUG level)
- Configuration details
- Detailed error traces
- FastMCP framework debug information

## Available MCP Tools

The server provides the following MCP tools:

### Core Extraction Tools
- `extract_powerpoint_content`: Extract complete structured content from PowerPoint files
- `get_powerpoint_attributes`: Get specific attributes from PowerPoint slides
- `get_slide_info`: Get information for a specific slide
- `query_slides`: Query slides with flexible filtering criteria

### Table Processing Tools
- `extract_table_data`: Extract table data with flexible selection and formatting detection

### Analysis Tools
- `get_presentation_overview`: Get comprehensive presentation overview and analysis
- `analyze_text_formatting`: Analyze text formatting patterns across slides
- `extract_bold_text`: Extract all bold text from slides with location information

### Utility Tools
- `clear_cache`: Clear the analysis cache for specific files or all files
- `reload_file_content`: Reload file content by clearing cache and re-extracting

## MCP Client Integration

### Using with Claude Desktop

Add to your MCP configuration file:
```json
{
  "mcpServers": {
    "powerpoint-analyzer": {
      "command": "python",
      "args": ["path/to/main.py"],
      "cwd": "path/to/powerpoint-analyzer"
    }
  }
}
```

### Using with Other MCP Clients

The server communicates via stdin/stdout using the MCP protocol. Any MCP-compatible client can connect to it.

## Security Considerations

- File path validation prevents directory traversal attacks
- File size limits prevent resource exhaustion
- Processing timeouts prevent infinite loops
- Temporary files are properly cleaned up
- No sensitive information is logged in production mode
- FastMCP framework provides built-in security features