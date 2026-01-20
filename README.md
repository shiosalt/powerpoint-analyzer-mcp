# PowerPoint Analyzer MCP Server
An MCP server that enables AI agents to extract structured content and text formatting from PowerPoint (.pptx) files using the Model Context Protocol (MCP).

## Background
Most AI tools that claim PowerPoint support only extract plain text, losing structure and formatting information. This MCP server preserves PowerPoint structure, formatting attributes, and enables querying of presentation content.

## Features

- **Text formatting detection**: Bold, italic, underline, strikethrough, highlighting, hyperlinks
- **Font analysis**: Font colors
- **Slide querying**: Filter by title, content, layout, and speaker notes
- **Table data extraction**: Multiple output formats (row/col/value, HTML) with optional formatting
- **Python-style slide selection**: Slicing notation (`:10`, `5:20`, `25:`)
- **Reduced context consumption**: Structured data output
- **No external dependencies**: Uses only Python standard library for PowerPoint processing
- **Built with FastMCP 2.0**: MCP server framework

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

3. Configure your MCP client (Claude Desktop, Cline, etc.) by adding the server configuration:

**For Claude Desktop:**

Location: `~/Library/Application Support/Claude/claude_desktop_config.json` (macOS) or `%APPDATA%\Claude\claude_desktop_config.json` (Windows)

```json
{
  "mcpServers": {
    "powerpoint-analyzer": {
      "command": "python",
      "args": ["/absolute/path/to/powerpoint-analyzer/main.py"],
      "env": {
        "POWERPOINT_MCP_LOG_LEVEL": "INFO"
      }
    }
  }
}
```

**For Cline/Other MCP Clients:**

Add to your MCP configuration file (typically `.kiro/settings/mcp.json` or similar):

```json
{
  "mcpServers": {
    "powerpoint-analyzer": {
      "command": "python",
      "args": ["C:\\path\\to\\powerpoint-analyzer\\main.py"],
      "env": {
        "POWERPOINT_MCP_LOG_LEVEL": "INFO"
      }
    }
  }
}
```

**Note**: Use absolute paths for the `args` parameter. On Windows, use double backslashes (`\\`) or forward slashes (`/`).

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
- **Debug capabilities**: Debugging tools for troubleshooting formatting detection issues

### Testing
- **Test suite**: `tests/test_formatting_detection.py` validates formatting types
- **Debug tools**: `tests/debug_formatting_detection.py` provides XML analysis

## Usage

### Running the Server

```bash
python main.py
```

## Available MCP Tools

This server provides four core tools for PowerPoint analysis:

**Recommended Tools (For AI Agents):**
- `query_slides` - Slide querying with reduced context consumption
- `extract_table_data` - Table extraction with reduced context consumption

**Legacy Tools (Not Recommended - Large Context Consumption):**
- `extract_formatted_table_data` - Full formatting metadata (use only when necessary)
- `extract_formatted_text` - Detailed formatting analysis (use only when necessary)

### 1. query_slides

Query and filter slides with specified criteria. Returns structured slide information with reduced context consumption.

**Parameters:**

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `file_path` | string | Yes | Full path to PowerPoint file (.pptx) |
| `search_criteria` | object | Yes | Filtering criteria (see below) |
| `return_fields` | array | No | Fields to include in results (default: `["slide_number", "title", "text"]`) |
| `slide_numbers` | int/string/array | No | Slides to query (default: all slides) |
| `output_type` | string | No | Text output type: `"preview_text_3boxes"` (default) or `"full_text"` |
| `output_format` | string | No | Output format: `"simple"` (default) or `"formatted"` |
| `limit` | integer | No | Max results to return (default: 1000, max: 10000) |

**search_criteria Structure:**

```json
{
  "title": {
    "contains": "Sales",           // Title contains text
    "starts_with": "Chapter",      // Title starts with text
    "ends_with": "Summary",        // Title ends with text
    "regex": "^Q[1-4] 202[0-9]$", // Title matches regex
    "one_of": ["Intro", "Outro"]  // Title is one of these values
  },
  "content": {
    "contains_text": "revenue",    // Slide text contains string
    "has_tables": true,            // Slide has tables
    "has_charts": true,            // Slide has charts
    "has_images": true             // Slide has images
  },
  "notes": {
    "contains": "important",       // Speaker notes contain text
    "regex": "TODO.*",             // Notes match regex
    "is_empty": false              // Notes are not empty
  },
  "sections": ["Introduction", "Conclusion"]  // Filter by section names
}
```

**return_fields Options:**

- `"slide_number"`: Slide number (always included)
- `"title"`: Slide title
- `"subtitle"`: Slide subtitle
- `"text"`: Text content (format controlled by `output_type`)
- `"extracted_tables"`: Table data in simplified format

**output_type Options:**

- `"preview_text_3boxes"` (default): Shows title + content placeholder + up to 3 text boxes
- `"full_text"`: Shows all text elements without limit

**output_format Options:**

- `"simple"` (default): No formatting information in text/tables
- `"formatted"`: Includes formatting information (bold, italic, colors, etc.)

**Example Usage:**

```python
# Find slides with "Sales" in title
query_slides("C:\\temp\\presentation.pptx", {"title": {"contains": "Sales"}})

# Find slides with tables, return full text
query_slides("C:\\temp\\presentation.pptx", 
            {"content": {"has_tables": true}},
            output_type="full_text")

# Query specific slides with custom fields
query_slides("C:\\temp\\presentation.pptx", {},
            return_fields=["slide_number", "title", "extracted_tables"],
            slide_numbers="1,5,10")
```

---

### 2. extract_table_data

Extract table data in simplified format with reduced context consumption.

**Parameters:**

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `file_path` | string | Yes | Path to PowerPoint file (.pptx) |
| `slide_numbers` | int/string/array | No | Slides to extract from (default: all slides) |
| `column_selection` | object | No | Column filtering configuration |
| `output_format` | string | No | Output format (default: `"row_col_value"`) |

**output_format Options:**

- `"row_col_value"` (default): `[row, col, value]` format with values only
- `"row_col_formattedvalue"`: `[row, col, value]` format with formatting included
- `"html"`: HTML table with formatting (supports colspan/rowspan)
- `"simple_html"`: HTML table without formatting (supports colspan/rowspan)

**column_selection Structure:**

```json
{
  "specific_columns": ["Name", "Price", "Quantity"],  // Select specific columns by name
  "column_patterns": [".*_total$", "^sum_.*"],       // Select columns matching regex patterns
  "exclude_columns": ["Notes", "Internal_ID"],       // Exclude specific columns
  "all_columns": true                                 // Include all columns (default)
}
```

**Output Structure:**

For `row_col_value` / `row_col_formattedvalue`:
```json
{
  "extracted_tables": [
    {
      "slide_number": 3,
      "rows": 5,
      "columns": 3,
      "headers": ["Product", "Price", "Quantity"],
      "data": [
        [1, 0, "Widget A"],
        [1, 1, "$10.00"],
        [1, 2, "100"],
        [2, 0, "Widget B"],
        [2, 1, "$15.00"],
        [2, 2, "50"]
      ]
    }
  ]
}
```

For `html` / `simple_html`:
```json
{
  "extracted_html_tables": [
    {
      "slide_number": 3,
      "rows": 5,
      "columns": 3,
      "headers": ["Product", "Price", "Quantity"],
      "htmldata": "<table style=\"white-space: pre;\">...</table>"
    }
  ]
}
```

**Example Usage:**

```python
# Extract all tables as simple arrays
extract_table_data("C:\\temp\\presentation.pptx")

# Extract as HTML tables with formatting
extract_table_data("C:\\temp\\presentation.pptx", output_format="html")

# Extract from specific slides only
extract_table_data("C:\\temp\\presentation.pptx", slide_numbers=[1, 3, 5])

# Extract specific columns
extract_table_data("C:\\temp\\presentation.pptx",
                  column_selection={"specific_columns": ["Name", "Total"]})
```

---

## Usage Recommendations

### For AI Agents (Recommended)

Use these tools for reduced context consumption:

1. **query_slides**: Primary tool for slide analysis
   - Reduced context consumption
   - Multiple filtering options
   - Structured output
   - Use `output_format="formatted"` when you need formatting information

2. **extract_table_data**: Primary tool for table extraction
   - Reduced context consumption
   - Multiple format options (row/col/value, HTML)
   - Suitable for large presentations

### Legacy Tools (Use Sparingly)

These tools generate large context output and should only be used when absolutely necessary:

- **extract_formatted_table_data**: Only for detailed formatting metadata analysis
- **extract_formatted_text**: Only for comprehensive formatting research

For most formatting needs, use `query_slides` with `output_format="formatted"` instead.

---

### 3. extract_formatted_table_data

⚠️ **NOT RECOMMENDED**: This tool generates very large context output with extensive formatting metadata. Use `extract_table_data` instead for most use cases.

Extract table data with comprehensive formatting detection (legacy tool with full formatting support).

**When to use:**
- Only when you need detailed formatting metadata (bold, italic, colors, etc.)
- For specialized formatting analysis requirements
- When the additional context consumption is acceptable

**For most use cases, use `extract_table_data` instead.**

**Parameters:**

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `file_path` | string | Yes | Path to PowerPoint file (.pptx) |
| `slide_numbers` | int/string/array | No | Slides to extract from (default: all slides) |
| `table_criteria` | object | No | Table filtering criteria |
| `column_selection` | object | No | Column filtering configuration |
| `formatting_detection` | object | No | Formatting detection configuration |
| `output_format` | string | No | Output format: `"structured"`, `"flat"`, `"grouped_by_slide"` |
| `include_metadata` | boolean | No | Include table metadata (default: true) |

**table_criteria Structure:**

```json
{
  "min_rows": 2,                                    // Minimum number of rows
  "max_rows": 100,                                  // Maximum number of rows
  "min_columns": 2,                                 // Minimum number of columns
  "max_columns": 10,                                // Maximum number of columns
  "header_contains": ["Total", "Summary"],         // Header must contain these strings
  "header_patterns": ["^Q[1-4].*", ".*_total$"]   // Header must match these regex patterns
}
```

**formatting_detection Structure:**

```json
{
  "detect_bold": true,           // Detect bold text
  "detect_italic": true,         // Detect italic text
  "detect_underline": true,      // Detect underlined text
  "detect_highlight": true,      // Detect highlighted text
  "detect_colors": true,         // Detect text colors
  "detect_hyperlinks": true,     // Detect hyperlinks
  "preserve_formatting": true    // Preserve formatting in output
}
```

**Example Usage:**

```python
# Extract with formatting detection
extract_formatted_table_data("C:\\temp\\presentation.pptx",
                            formatting_detection={
                              "detect_bold": true,
                              "detect_colors": true
                            })

# Extract tables with specific criteria
extract_formatted_table_data("C:\\temp\\presentation.pptx",
                            table_criteria={
                              "min_rows": 3,
                              "header_contains": ["Total"]
                            })
```

---

### 4. extract_formatted_text

⚠️ **NOT RECOMMENDED**: This tool generates very large context output with detailed formatting analysis. Consider using `query_slides` with `output_format="formatted"` for lighter-weight formatting information.

Extract text with specific formatting attributes from slides.

**When to use:**
- Only when you need comprehensive formatting analysis across all slides
- For specialized text formatting research
- When the additional context consumption is acceptable

**For most use cases, use `query_slides` with appropriate filters instead.**

**Parameters:**

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `file_path` | string | Yes | Path to PowerPoint file (.pptx) |
| `formatting_type` | string | Yes | Type of formatting to extract |
| `slide_numbers` | int/string/array | No | Slides to analyze (default: all slides) |

**formatting_type Options:**

- `"bold"`: Extract bold text segments
- `"italic"`: Extract italic text segments
- `"underlined"`: Extract underlined text segments
- `"highlighted"`: Extract highlighted text segments
- `"strikethrough"`: Extract strikethrough text segments
- `"hyperlinks"`: Extract hyperlinks with URLs and link types
- `"font_sizes"`: Extract text with font size information
- `"font_colors"`: Extract text with color information (hex format)

**Output Structure:**

```json
{
  "file_path": "C:\\temp\\presentation.pptx",
  "formatting_type": "bold",
  "summary": {
    "total_slides_analyzed": 10,
    "slides_with_formatting": 5,
    "total_formatted_segments": 12
  },
  "results_by_slide": [
    {
      "slide_number": 1,
      "title": "Introduction",
      "complete_text": "Welcome to our presentation...",
      "format": "bold",
      "formatted_segments": [
        {
          "text": "Important Note",
          "start_position": 25
        }
      ]
    }
  ]
}
```

**Example Usage:**

```python
# Extract all bold text
extract_formatted_text("C:\\temp\\presentation.pptx", "bold")

# Extract hyperlinks from specific slides
extract_formatted_text("C:\\temp\\presentation.pptx", "hyperlinks", slide_numbers=[1, 2, 3])

# Extract font colors from first 10 slides
extract_formatted_text("C:\\temp\\presentation.pptx", "font_colors", slide_numbers=":10")
```

---

## Slide Selection Syntax

All tools support slide selection using Python-style slicing:

| Format | Example | Description |
|--------|---------|-------------|
| None | (omit parameter) | All slides |
| Integer | `3` | Single slide 3 |
| Array | `[1, 5, 10]` | Specific slides 1, 5, 10 |
| String (comma) | `"1,5,10"` | Specific slides 1, 5, 10 |
| String (slice) | `":10"` | First 10 slides (1-10) |
| String (slice) | `"5:20"` | Slides 5-20 |
| String (slice) | `"25:"` | Slides 25 to end |

**Examples:**

```python
# First 10 slides
query_slides("file.pptx", {}, slide_numbers=":10")

# Slides 5-20
extract_table_data("file.pptx", slide_numbers="5:20")

# Slides 25 to end
extract_formatted_text("file.pptx", "bold", slide_numbers="25:")

# Specific slides
query_slides("file.pptx", {}, slide_numbers="1,3,5,10")
```



## Development

### Requirements

- Python 3.8+
- MCP (Model Context Protocol)
- FastMCP 2.0
- Standard Python libraries (zipfile, xml.etree.ElementTree)

## Recent Updates

### Version 2.2 - Enhanced Documentation & Tool Improvements
- **Comprehensive tool documentation**: Detailed parameter explanations for all MCP tools
- **Output format options**: Multiple formats for tables (row/col/value, HTML, formatted)
- **Flexible filtering**: Advanced search criteria for slides and tables
- **Column selection**: Filter table columns by name or pattern
- **Optimized output**: Minimal context consumption with clean data structures

### Version 2.1 - Python-Style Slide Selection
- **Enhanced slide selection**: Python-style slicing notation (`:10`, `5:20`, `25:`)
- **Flexible specification**: Single slides, ranges, comma-separated lists, and slicing
- **Improved performance**: Process only needed slides with efficient selection
- **Backward compatibility**: Existing `[1, 5, 10]` format still supported

### Version 2.0 - Advanced Text Formatting Detection
- **Fixed formatting detection**: Bold, italic, underline, strikethrough correctly detected
- **Dual detection support**: Handles both XML attribute and child element formats
- **Comprehensive test suite**: Extensive tests for all formatting types
- **Enhanced MCP tools**: New tools for formatted text extraction and analysis

## License

This project is licensed under the Apache License 2.0 - see the [LICENSE](LICENSE) file for details.
