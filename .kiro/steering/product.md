# PowerPoint Analyzer MCP

A Model Context Protocol (MCP) server that enables AI agents to extract structured content and text formatting from PowerPoint (.pptx) files. Unlike typical PowerPoint tools that only extract plain text, this server preserves PowerPoint structure and formatting attributes like bold text, italics, tables, layouts, and metadata.

## Key Features

- **Structured content extraction**: Slides, titles, subtitles, placeholders, layouts
- **Text formatting detection**: Bold, italic, underline, strikethrough, highlighting, hyperlinks
- **Table extraction**: With formatting detection and flexible selection criteria
- **Slide querying**: Filter slides based on content, layout, or formatting criteria
- **Font analysis**: Font sizes, colors, and styling information
- **No external dependencies**: Uses only Python standard library for PowerPoint processing
- **Direct XML parsing**: Processes .pptx files as ZIP archives with XML content
- **Caching system**: Performance optimization for repeated operations

## Target Use Cases

- AI agents analyzing presentation content while preserving formatting context
- Extracting structured data from corporate presentations
- Content analysis that requires understanding of text emphasis (bold, italic)
- Table data extraction from slides with formatting preservation
- Presentation metadata and structure analysis