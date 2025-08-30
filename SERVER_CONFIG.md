# PowerPoint MCP Server Configuration

This document describes the configuration options and startup procedures for the PowerPoint MCP Server.

## Quick Start

### Basic Usage
```bash
# Start the server with default configuration
python main.py

# Or use the enhanced startup script
python scripts/start_server.py
```

### With Custom Configuration
```bash
# Start with debug logging
python scripts/start_server.py --debug

# Start with custom file size limit and timeout
python scripts/start_server.py --max-file-size 200 --timeout 600

# Start with no caching
python scripts/start_server.py --no-cache
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

### Command Line Options

The startup script supports the following command line options:

- `--log-level`: Set logging level
- `--max-file-size`: Maximum file size in MB
- `--timeout`: Processing timeout in seconds
- `--no-cache`: Disable caching
- `--debug`: Enable debug mode
- `--version`: Show version information

## Server Lifecycle

### Startup Sequence

1. **Configuration Loading**: Load configuration from environment variables and command line arguments
2. **Logging Setup**: Configure logging based on the specified log level
3. **Signal Handlers**: Set up signal handlers for graceful shutdown
4. **Server Initialization**: Initialize the MCP server and all components
5. **Validation**: Validate configuration and dependencies
6. **Server Start**: Start the MCP server and begin accepting requests

### Shutdown Sequence

1. **Signal Reception**: Receive shutdown signal (SIGINT, SIGTERM, or SIGBREAK on Windows)
2. **Graceful Shutdown**: Stop accepting new requests
3. **Resource Cleanup**: Clear caches and close file handles
4. **Server Stop**: Terminate the MCP server
5. **Exit**: Clean exit with appropriate exit code

## Health Checks

Use the health check script to verify server readiness:

```bash
python scripts/health_check.py
```

The health check verifies:
- Required dependencies are available
- Configuration is valid
- Server components can be initialized
- Test files can be processed (if available)

## Logging

### Log Levels

- **DEBUG**: Detailed information for debugging
- **INFO**: General information about server operation
- **WARNING**: Warning messages for potential issues
- **ERROR**: Error messages for failed operations
- **CRITICAL**: Critical errors that may cause server shutdown

### Log Format

Default log format: `%(asctime)s - %(name)s - %(levelname)s - %(message)s`

Example log output:
```
2024-01-15 10:30:45,123 - powerpoint_mcp_server - INFO - PowerPoint MCP Server initialized (version 0.1.0)
2024-01-15 10:30:45,124 - powerpoint_mcp_server.server - INFO - PowerPoint MCP Server starting...
2024-01-15 10:30:45,125 - powerpoint_mcp_server.server - INFO - MCP server connected to stdio streams
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
- Invalid parameters
- Unsupported tool calls
- Response format errors

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
python scripts/start_server.py --debug
```

Debug mode provides:
- Verbose logging
- Configuration details
- Detailed error traces
- Performance metrics

## Security Considerations

- File path validation prevents directory traversal attacks
- File size limits prevent resource exhaustion
- Processing timeouts prevent infinite loops
- Temporary files are properly cleaned up
- No sensitive information is logged in production mode