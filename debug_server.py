#!/usr/bin/env python3
"""Debug version of PowerPoint MCP Server."""

import asyncio
import json
import logging
import sys
import traceback
from typing import Any, Dict

from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.types import (
    CallToolResult,
    ListToolsResult,
    Tool,
    TextContent
)
from mcp.server.models import InitializationOptions
from mcp.types import ServerCapabilities

# Configure detailed logging
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[logging.StreamHandler(sys.stderr)]
)

logger = logging.getLogger(__name__)

# Initialize server
server = Server("powerpoint-mcp-server")

@server.list_tools()
async def list_tools() -> ListToolsResult:
    """List available tools."""
    logger.info("list_tools called")
    try:
        tools = [
            Tool(
                name="test_tool",
                description="A simple test tool",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "message": {
                            "type": "string",
                            "description": "Test message"
                        }
                    },
                    "required": ["message"]
                }
            )
        ]
        logger.info(f"Returning {len(tools)} tools")
        return ListToolsResult(tools=tools)
    except Exception as e:
        logger.error(f"Error in list_tools: {e}")
        logger.error(traceback.format_exc())
        raise

@server.call_tool()
async def call_tool(name: str, arguments: Dict[str, Any]) -> CallToolResult:
    """Handle tool calls."""
    logger.info(f"call_tool called with name={name}, arguments={arguments}")
    try:
        if name == "test_tool":
            message = arguments.get("message", "Hello from PowerPoint MCP Server!")
            return CallToolResult(
                content=[
                    TextContent(
                        type="text",
                        text=f"Test response: {message}"
                    )
                ]
            )
        else:
            return CallToolResult(
                content=[
                    TextContent(
                        type="text",
                        text=f"Unknown tool: {name}"
                    )
                ]
            )
    except Exception as e:
        logger.error(f"Error in call_tool: {e}")
        logger.error(traceback.format_exc())
        raise

async def main():
    """Main entry point."""
    logger.info("Starting Debug PowerPoint MCP Server...")
    
    try:
        async with stdio_server() as (read_stream, write_stream):
            logger.info("MCP server connected to stdio streams")
            await server.run(
                read_stream, 
                write_stream,
                InitializationOptions(
                    server_name="powerpoint-mcp-server",
                    server_version="1.0.0",
                    capabilities=ServerCapabilities()
                )
            )
    except Exception as e:
        logger.error(f"Server error: {e}")
        logger.error(traceback.format_exc())
        raise

if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        logger.info("Server stopped by user")
        sys.exit(0)
    except Exception as e:
        logger.error(f"Main error: {e}")
        logger.error(traceback.format_exc())
        sys.exit(1)