#!/usr/bin/env python3
"""
MCP Test CLI Tool - PowerPoint Analyzer MCP Testing Support

A command-line tool for testing MCP servers via stdio communication.
Supports FastMCP 2.0 servers and provides easy tool discovery and execution.

Usage:
    python mcp_test_cli.py                           # List all available tools
    python mcp_test_cli.py <tool_name>               # Show tool options/parameters
    python mcp_test_cli.py <tool_name> [options]     # Execute tool with options

Examples:
    python mcp_test_cli.py
    python mcp_test_cli.py extract_powerpoint_content
    python mcp_test_cli.py extract_powerpoint_content --file_path "test.pptx"
"""

import asyncio
import json
import sys
import argparse
import subprocess
import os
from typing import Dict, List, Any, Optional
from pathlib import Path


class MCPTestCLI:
    """Command-line interface for testing MCP servers."""
    
    def __init__(self, server_command: List[str] = None):
        """
        Initialize the MCP test CLI.
        
        Args:
            server_command: Command to start the MCP server (default: uses main.py)
        """
        self.server_command = server_command or ["python", "main.py"]
        self.server_process = None
        self.request_id = 1
        
    async def start_server(self):
        """Start the MCP server process."""
        try:
            self.server_process = await asyncio.create_subprocess_exec(
                *self.server_command,
                stdin=asyncio.subprocess.PIPE,
                stdout=asyncio.subprocess.PIPE,
                stderr=asyncio.subprocess.PIPE
            )
            print(f"✅ MCP server started: {' '.join(self.server_command)}")
            return True
        except Exception as e:
            print(f"❌ Failed to start MCP server: {e}")
            return False
    
    async def stop_server(self):
        """Stop the MCP server process."""
        if self.server_process:
            try:
                self.server_process.terminate()
                await self.server_process.wait()
                print("✅ MCP server stopped")
            except Exception as e:
                print(f"⚠️ Error stopping server: {e}")
    
    async def send_request(self, method: str, params: Dict[str, Any] = None) -> Dict[str, Any]:
        """
        Send a JSON-RPC request to the MCP server.
        
        Args:
            method: The method to call
            params: Parameters for the method
            
        Returns:
            Response from the server
        """
        if not self.server_process:
            raise RuntimeError("Server not started")
        
        request = {
            "jsonrpc": "2.0",
            "id": self.request_id,
            "method": method,
            "params": params or {}
        }
        self.request_id += 1
        
        # Send request
        request_json = json.dumps(request) + "\n"
        self.server_process.stdin.write(request_json.encode())
        await self.server_process.stdin.drain()
        
        # Read response
        response_line = await self.server_process.stdout.readline()
        if not response_line:
            raise RuntimeError("No response from server")
        
        try:
            response = json.loads(response_line.decode().strip())
            return response
        except json.JSONDecodeError as e:
            raise RuntimeError(f"Invalid JSON response: {e}")
    
    async def initialize_server(self):
        """Initialize the MCP server."""
        try:
            # Send initialize request
            response = await self.send_request("initialize", {
                "protocolVersion": "2024-11-05",
                "capabilities": {
                    "tools": {}
                },
                "clientInfo": {
                    "name": "mcp-test-cli",
                    "version": "1.0.0"
                }
            })
            
            if "error" in response:
                print(f"❌ Server initialization failed: {response['error']}")
                return False
            
            # Send initialized notification (no response expected)
            initialized_request = {
                "jsonrpc": "2.0",
                "method": "notifications/initialized"
            }
            initialized_json = json.dumps(initialized_request) + "\n"
            self.server_process.stdin.write(initialized_json.encode())
            await self.server_process.stdin.drain()
            
            print("✅ MCP server initialized")
            return True
            
        except Exception as e:
            print(f"❌ Failed to initialize server: {e}")
            return False
    
    async def list_tools(self) -> List[Dict[str, Any]]:
        """Get list of available tools from the server."""
        try:
            response = await self.send_request("tools/list")
            
            if "error" in response:
                print(f"❌ Failed to list tools: {response['error']}")
                return []
            
            return response.get("result", {}).get("tools", [])
            
        except Exception as e:
            print(f"❌ Error listing tools: {e}")
            return []
    
    async def call_tool(self, tool_name: str, arguments: Dict[str, Any]) -> Dict[str, Any]:
        """
        Call a specific tool with arguments.
        
        Args:
            tool_name: Name of the tool to call
            arguments: Arguments to pass to the tool
            
        Returns:
            Tool execution result
        """
        try:
            response = await self.send_request("tools/call", {
                "name": tool_name,
                "arguments": arguments
            })
            
            if "error" in response:
                print(f"❌ Tool execution failed: {response['error']}")
                return response
            
            return response.get("result", {})
            
        except Exception as e:
            print(f"❌ Error calling tool: {e}")
            return {"error": str(e)}
    
    def print_tools_list(self, tools: List[Dict[str, Any]]):
        """Print formatted list of available tools."""
        if not tools:
            print("❌ No tools available")
            return
        
        print(f"\n📋 Available Tools ({len(tools)} total):")
        print("=" * 50)
        
        for i, tool in enumerate(tools, 1):
            name = tool.get("name", "Unknown")
            description = tool.get("description", "No description")
            
            # Truncate long descriptions
            if len(description) > 80:
                description = description[:77] + "..."
            
            print(f"{i:2d}. {name}")
            print(f"    {description}")
            print()
    
    def print_tool_details(self, tool: Dict[str, Any]):
        """Print detailed information about a specific tool."""
        name = tool.get("name", "Unknown")
        description = tool.get("description", "No description")
        input_schema = tool.get("inputSchema", {})
        
        print(f"\n🔧 Tool: {name}")
        print("=" * (len(name) + 8))
        print(f"Description: {description}")
        print()
        
        # Print parameters
        properties = input_schema.get("properties", {})
        required = input_schema.get("required", [])
        
        if properties:
            print("Parameters:")
            for param_name, param_info in properties.items():
                param_type = param_info.get("type", "unknown")
                param_desc = param_info.get("description", "No description")
                is_required = param_name in required
                required_mark = " (required)" if is_required else " (optional)"
                
                print(f"  --{param_name} <{param_type}>{required_mark}")
                print(f"      {param_desc}")
                print()
        else:
            print("No parameters required.")
        
        # Print usage examples
        print("Usage Examples:")
        if properties:
            example_args = []
            for param_name, param_info in properties.items():
                if param_name in required:
                    param_type = param_info.get("type", "string")
                    if param_type == "string":
                        example_args.append(f'--{param_name} "example_value"')
                    elif param_type == "integer":
                        example_args.append(f'--{param_name} 1')
                    elif param_type == "boolean":
                        example_args.append(f'--{param_name} true')
                    elif param_type == "array":
                        example_args.append(f'--{param_name} item1,item2')
                    else:
                        example_args.append(f'--{param_name} <value>')
            
            basic_example = f"python mcp_test_cli.py {name} {' '.join(example_args)}"
            print(f"  {basic_example}")
            
            # Add array-specific examples if there are array parameters
            array_params = [p for p, info in properties.items() if info.get("type") == "array"]
            if array_params:
                print("\n  Array parameter formats:")
                for param in array_params:
                    print(f"    # Comma-separated:")
                    print(f"    --{param} item1,item2,item3")
                    print(f"    # JSON format (Windows CMD):")
                    print(f'    --{param} "[""item1"", ""item2"", ""item3""]"')
                    print(f"    # JSON format (PowerShell):")
                    print(f"    --{param} '[\"item1\", \"item2\", \"item3\"]'")
        else:
            basic_example = f"python mcp_test_cli.py {name}"
            print(f"  {basic_example}")
        
        print()
    
    def print_tool_result(self, result: Dict[str, Any]):
        """Print formatted tool execution result."""
        if "error" in result:
            print(f"❌ Error: {result['error']}")
            return
        
        print("\n✅ Tool Execution Result:")
        print("=" * 30)
        
        # Handle different result formats
        content = result.get("content", [])
        if content:
            for item in content:
                if isinstance(item, dict):
                    if "text" in item:
                        # Try to parse as JSON for pretty printing
                        try:
                            parsed = json.loads(item["text"])
                            print(json.dumps(parsed, indent=2, ensure_ascii=False))
                        except json.JSONDecodeError:
                            print(item["text"])
                    else:
                        print(json.dumps(item, indent=2, ensure_ascii=False))
                else:
                    print(str(item))
        else:
            print(json.dumps(result, indent=2, ensure_ascii=False))
    
    def parse_arguments(self, args: List[str]) -> Dict[str, Any]:
        """
        Parse command line arguments into tool parameters.
        
        Supports multiple formats:
        - JSON arrays: --attributes '["title", "subtitle"]'
        - Comma-separated: --attributes title,subtitle
        - Multiple values: --attributes title --attributes subtitle
        
        Args:
            args: Command line arguments
            
        Returns:
            Dictionary of parsed arguments
        """
        arguments = {}
        i = 0
        
        while i < len(args):
            arg = args[i]
            
            if arg.startswith("--"):
                param_name = arg[2:]
                
                if i + 1 < len(args) and not args[i + 1].startswith("--"):
                    # Has value
                    value = args[i + 1]
                    
                    # Check if this parameter already exists (for array building)
                    if param_name in arguments:
                        # Convert to array if not already
                        if not isinstance(arguments[param_name], list):
                            arguments[param_name] = [arguments[param_name]]
                        # Add new value
                        arguments[param_name].append(self._parse_single_value(value))
                    else:
                        arguments[param_name] = self._parse_single_value(value)
                    
                    i += 2
                else:
                    # Boolean flag
                    arguments[param_name] = True
                    i += 1
            else:
                i += 1
        
        return arguments
    
    def _parse_single_value(self, value: str):
        """Parse a single value with type detection and Windows CMD compatibility."""
        # Boolean values
        if value.lower() in ["true", "false"]:
            return value.lower() == "true"
        
        # Numeric values
        if value.isdigit():
            return int(value)
        if value.replace(".", "", 1).isdigit():
            return float(value)
        
        # Handle Windows CMD quote issues - remove outer quotes if present
        cleaned_value = value
        if (value.startswith('"') and value.endswith('"')) or (value.startswith("'") and value.endswith("'")):
            cleaned_value = value[1:-1]
        
        # Try JSON parsing first (for arrays and objects)
        if cleaned_value.startswith(("[", "{")):
            try:
                # Handle Windows CMD double quote escaping and missing quotes
                json_str = cleaned_value
                
                # If it looks like a JSON array but missing quotes around strings
                if cleaned_value.startswith("[") and not '"' in cleaned_value and not "'" in cleaned_value:
                    # Try to fix missing quotes around array elements
                    # Convert [item1, item2] to ["item1", "item2"]
                    content = cleaned_value[1:-1]  # Remove [ and ]
                    items = [item.strip() for item in content.split(",")]
                    quoted_items = [f'"{item}"' for item in items if item]
                    json_str = f'[{", ".join(quoted_items)}]'
                else:
                    # Handle Windows CMD double quote escaping
                    json_str = cleaned_value.replace('""', '"')
                
                parsed = json.loads(json_str)
                return parsed
            except json.JSONDecodeError:
                # Fall through to other parsing methods
                pass
        
        # Comma-separated values (convert to array)
        if "," in cleaned_value and not (cleaned_value.startswith('"') and cleaned_value.endswith('"')):
            # Split by comma and clean up each item
            items = []
            for item in cleaned_value.split(","):
                item = item.strip().strip('"').strip("'")
                if item:  # Skip empty items
                    items.append(item)
            return items if len(items) > 1 else items[0] if items else cleaned_value
        
        # Regular string
        return cleaned_value


async def main():
    """Main entry point for the CLI tool."""
    if len(sys.argv) == 1:
        # No arguments - list all tools
        cli = MCPTestCLI()
        
        if not await cli.start_server():
            sys.exit(1)
        
        try:
            if not await cli.initialize_server():
                sys.exit(1)
            
            tools = await cli.list_tools()
            cli.print_tools_list(tools)
            
        finally:
            await cli.stop_server()
    
    elif len(sys.argv) == 2:
        # Tool name only - show tool details
        tool_name = sys.argv[1]
        cli = MCPTestCLI()
        
        if not await cli.start_server():
            sys.exit(1)
        
        try:
            if not await cli.initialize_server():
                sys.exit(1)
            
            tools = await cli.list_tools()
            
            # Find the specific tool
            target_tool = None
            for tool in tools:
                if tool.get("name") == tool_name:
                    target_tool = tool
                    break
            
            if target_tool:
                cli.print_tool_details(target_tool)
            else:
                print(f"❌ Tool '{tool_name}' not found")
                print("\nAvailable tools:")
                for tool in tools:
                    print(f"  - {tool.get('name', 'Unknown')}")
                sys.exit(1)
            
        finally:
            await cli.stop_server()
    
    else:
        # Tool name with arguments - execute tool
        tool_name = sys.argv[1]
        arguments = MCPTestCLI().parse_arguments(sys.argv[2:])
        
        cli = MCPTestCLI()
        
        if not await cli.start_server():
            sys.exit(1)
        
        try:
            if not await cli.initialize_server():
                sys.exit(1)
            
            print(f"🚀 Executing tool: {tool_name}")
            print(f"📝 Arguments: {json.dumps(arguments, indent=2)}")
            print()
            
            result = await cli.call_tool(tool_name, arguments)
            cli.print_tool_result(result)
            
        finally:
            await cli.stop_server()


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("\n⚠️ Interrupted by user")
        sys.exit(1)
    except Exception as e:
        print(f"❌ Unexpected error: {e}")
        sys.exit(1)