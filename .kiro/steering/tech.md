# Technology Stack

## Core Technologies

- **Python 3.8+**: Primary language
- **FastMCP 2.0**: MCP server framework for tool registration and lifecycle management
- **MCP (Model Context Protocol) 1.13.0+**: Protocol for AI agent communication
- **Python Standard Library**: Core processing (no external PowerPoint dependencies)
  - `zipfile`: Handle .pptx files as ZIP archives
  - `xml.etree.ElementTree`: Parse PowerPoint XML with namespace support
  - `json`: Data serialization
  - `logging`: Comprehensive logging system

## Architecture Pattern

- **Modular design**: Separate concerns into core, utils, tools, and resources
- **Async/await**: All MCP tool methods are async
- **Caching layer**: Performance optimization for repeated file operations
- **Direct XML processing**: No external PowerPoint libraries required
- **Dual formatting detection**: Supports both XML attributes and child elements

## Project Structure

```
powerpoint_mcp_server/
├── core/                    # Core business logic
│   ├── content_extractor.py      # Main content extraction
│   ├── text_formatting_analyzer.py # Text formatting detection
│   ├── enhanced_table_extractor.py # Table processing
│   ├── slide_query_engine.py     # Slide filtering/querying
│   └── presentation_analyzer.py  # High-level analysis
├── utils/                   # Utility modules
│   ├── file_validator.py         # File validation
│   ├── zip_extractor.py          # ZIP/PPTX handling
│   └── cache_manager.py          # Caching system
├── tools/                   # MCP tool helpers
└── resources/              # MCP resources (if any)
```

## Common Commands

### Development
```bash
# Install dependencies
pip install -r requirements.txt

# Run tests
pytest
pytest -v                    # Verbose output
pytest tests/test_specific.py # Run specific test

# Run server directly
python main.py

# Health check
python scripts/health_check.py
```

### Testing
```bash
# Run comprehensive test suite
python tests/run_comprehensive_test_suite.py

# Run integration tests
python tests/test_integration.py

# Performance testing
python tests/test_performance.py
```

## Configuration

- Environment variables for configuration (POWERPOINT_MCP_LOG_LEVEL, etc.)
- Logging to both file (`powerpoint_mcp_server.log`) and stderr
- Configurable cache TTL, file size limits, processing timeouts