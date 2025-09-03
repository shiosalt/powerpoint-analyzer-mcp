# Project Structure & Organization

## Root Level Files

- `main.py`: FastMCP 2.0 server entry point with all MCP tool definitions
- `setup.py`: Package configuration and dependencies
- `requirements.txt`: Python dependencies (MCP, FastMCP, pytest)
- `pytest.ini`: Test configuration with asyncio support
- `mcp_config.json`: MCP server configuration for client integration
- `README.md`: Primary documentation
- `LICENSE`: Apache 2.0 license

## Core Package Structure

```
powerpoint_mcp_server/
├── __init__.py
├── server.py              # Main MCP server class (legacy, FastMCP used in main.py)
├── config.py              # Configuration management with environment variables
├── core/                  # Business logic modules
│   ├── content_extractor.py      # Primary content extraction engine
│   ├── attribute_processor.py    # Attribute filtering and processing
│   ├── text_formatting_analyzer.py # Text formatting detection (bold, italic, etc.)
│   ├── enhanced_table_extractor.py # Advanced table extraction
│   ├── slide_query_engine.py     # Slide filtering and querying
│   ├── presentation_analyzer.py  # High-level presentation analysis
│   ├── data_filter_engine.py     # Data filtering utilities
│   ├── formatting_extractor.py   # Formatting extraction utilities
│   ├── workflow_assistant.py     # Workflow helpers
│   └── xml_parser.py             # XML parsing with namespace support
├── utils/                 # Utility modules
│   ├── file_validator.py         # File validation and security
│   ├── zip_extractor.py          # ZIP/PPTX file handling
│   ├── cache_manager.py          # Caching system
│   └── logger.py                 # Logging utilities
├── tools/                 # MCP tool helpers
│   └── tool_help.py              # Tool documentation and help system
├── resources/             # MCP resources (currently empty)
└── prompts/              # MCP prompts (currently empty)
```

## Testing Structure

```
tests/
├── test_*.py              # Unit tests for each module
├── test_files/            # Sample PowerPoint files for testing
├── reports/               # Test reports and outputs
├── run_*.py              # Test runners and suites
└── *_test_framework.py   # Testing frameworks and utilities
```

## Temporary & Development Files

```
temp/                      # Development and debugging files
├── sample*.pptx          # Test PowerPoint files
├── test_*.py             # Ad-hoc test scripts
├── debug_*.py            # Debugging utilities
└── *.md                  # Development notes and summaries
```

## Scripts & Utilities

```
scripts/
├── start_server.py       # Alternative server startup
└── health_check.py       # Server health monitoring
```

## Naming Conventions

- **Files**: snake_case (e.g., `content_extractor.py`)
- **Classes**: PascalCase (e.g., `ContentExtractor`)
- **Methods**: snake_case (e.g., `extract_content`)
- **Constants**: UPPER_SNAKE_CASE (e.g., `MAX_FILE_SIZE`)
- **MCP Tools**: snake_case matching function names (e.g., `extract_powerpoint_content`)

## Key Architectural Principles

- **Separation of concerns**: Core logic, utilities, and MCP interface are separate
- **Async-first**: All MCP tools are async functions
- **Error handling**: Comprehensive error handling with JSON error responses
- **Logging**: Extensive logging for debugging and monitoring
- **Caching**: Built-in caching for performance optimization
- **Standard library focus**: Minimal external dependencies for PowerPoint processing