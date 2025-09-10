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
    python mcp_test_cli.py extract_powerpoint_content --file_path "C:\temp\test.pptx"

*about error: Separator is found, but chunk is longer than limit
The program will fail with this error if the total size of STDIO input and output exceeds 512KB.
This limit can be adjusted by configuring the MCPCLI_MAXLEN environment variable (in kilobytes).
 ex:
   set MCPCLI_MAXLEN=1024
This is a limitation of mcp_test_cli.py,not powerpoint analyzer MCP.
"""

import asyncio
import json
import sys
import argparse
import subprocess
import os
from typing import Dict, List, Any, Optional
from pathlib import Path
maxlen = 512

class MCPTestCLI:
    """Command-line interface for testing MCP servers."""
    global maxlen
    def __init__(self, server_command: List[str] = None):
        """
        Initialize the MCP test CLI.

        Args:
            server_command: Command to start the MCP server (default: uses main.py)
        """
        global maxlen
        maxlen = int(os.getenv('MCPCLI_MAXLEN', maxlen))
        self.server_command = server_command or ["python", "main.py"]
        self.server_process = None
        self.request_id = 1

    async def start_server(self,raw_mode=False):
        """Start the MCP server process."""
        try:
            self.server_process = await asyncio.create_subprocess_exec(
                *self.server_command,
                stdin=asyncio.subprocess.PIPE,
                stdout=asyncio.subprocess.PIPE,
                stderr=asyncio.subprocess.PIPE,
                limit=1024 * maxlen
            )
            if not raw_mode:
                print(f"✅ MCP server started: {' '.join(self.server_command)}")
            return True
        except Exception as e:
            if not raw_mode:
                print(f"❌ Failed to start MCP server: {e}")
            return False

    async def stop_server(self,raw_mode=False):
        """Stop the MCP server process."""
        if self.server_process:
            try:
                self.server_process.terminate()
                await self.server_process.wait()
                if not raw_mode:
                    print("✅ MCP server stopped")
            except Exception as e:
                if not raw_mode:
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
        #self.request_id += 1

        # Send request
        request_json = json.dumps(request) + "\n"
        try:
            self.server_process.stdin.write(request_json.encode())
            await self.server_process.stdin.drain()
        except ValueError as e:
            raise RuntimeError(f"write [{e}].\n you may need set MCPCLI_MAXLEN more than {maxlen}.(default:512)\n")

        # Read response
        try:
            response_line = await self.server_process.stdout.readline()
        except ValueError as e:
            raise RuntimeError(f"read error [{e}].\n you may need set MCPCLI_MAXLEN more than {maxlen}.(default:512)\n")

        if not response_line:
            raise RuntimeError("No response from server")

        try:
            response = json.loads(response_line.decode().strip())
            return response
        except json.JSONDecodeError as e:
            raise RuntimeError(f"Invalid JSON response: {e}")

    async def initialize_server(self,raw_mode=False):
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
            if not raw_mode:
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

    async def call_tool(self, tool_name: str, arguments: Dict[str, Any],raw_mode = False) -> Dict[str, Any]:
        """
        Call a specific tool with arguments.

        Args:
            tool_name: Name of the tool to call
            arguments: Arguments to pass to the tool
            raw_mode: if True , return mcp server output only

        Returns:
            Tool execution result
        """
        try:
            response = await self.send_request("tools/call", {
                "name": tool_name,
                "arguments": arguments
            })
            if not raw_mode:
                if "error" in response:

                    print(f"❌ Tool execution failed: {response['error']}")
                    print(f"work\n",tool_name,arguments,raw_mode,response)
                    return response.get("result", {})
            return response.get("result", {})

        except Exception as e:
            if not raw_mode:
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
            if len(description) > 322:
                description = description[:320] + "..."

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

    def print_tool_result(self, result: Dict[str, Any],raw_mode = False):
        """Print formatted tool execution result."""
        if "error" in result:
            if raw_mode:
                print(result['error'])
            else:
                print(f"❌ Error: {result['error']}")
            return
        if not raw_mode:
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
                        parsed_value = self._parse_single_value(value)
                        # Special handling for parameters that are typically arrays
                        if param_name in ['slide_numbers', 'attributes', 'return_fields'] and not isinstance(parsed_value, list):
                            # Convert single values to arrays for these parameters
                            arguments[param_name] = [parsed_value]
                        else:
                            arguments[param_name] = parsed_value

                    i += 2
                else:
                    # Boolean flag
                    arguments[param_name] = True
                    i += 1
            else:
                i += 1

        return arguments

    def _parse_single_value(self, value: str):
        """
        Parse a single value with type detection and Windows CMD compatibility.
        """
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
                    # Convert [item1, item2] to ["item1", "item2"] or [1, 2] to [1, 2]
                    content = cleaned_value[1:-1]  # Remove [ and ]
                    items = [item.strip() for item in content.split(",")]
                    processed_items = []
                    for item in items:
                        if item:
                            # Check if it's a number
                            if item.isdigit() or (item.replace('.', '', 1).replace('-', '', 1).isdigit() and item.count('.') <= 1 and item.count('-') <= 1):
                                processed_items.append(item)
                            elif item.lower() in ['true', 'false', 'null']:
                                processed_items.append(item.lower())
                            else:
                                processed_items.append(f'"{item}"')
                    json_str = f'[{", ".join(processed_items)}]'

                # If it looks like a JSON object but missing quotes around keys/values
                elif cleaned_value.startswith("{") and not '"' in cleaned_value and not "'" in cleaned_value:
                    # Try to fix missing quotes around object keys and values
                    # Convert {key: value, key2: value2} to {"key": "value", "key2": "value2"}
                    json_str = self._fix_json_object(cleaned_value)

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
                    # Try to convert to appropriate type
                    if item.isdigit():
                        items.append(int(item))
                    elif item.replace(".", "", 1).isdigit():
                        items.append(float(item))
                    elif item.lower() in ["true", "false"]:
                        items.append(item.lower() == "true")
                    else:
                        items.append(item)
            return items if len(items) > 1 else items[0] if items else cleaned_value

        # Regular string
        return cleaned_value

    def _fix_json_object(self, obj_str: str) -> str:
        """Fix malformed JSON object by adding quotes around keys and string values recursively."""
        try:
            # Remove outer braces
            content = obj_str[1:-1].strip()
            if not content:
                return "{}"

            # Parse key-value pairs with proper nesting support
            pairs = self._parse_json_pairs(content)

            # Process each key-value pair
            fixed_pairs = []
            for key, value in pairs:
                # Quote the key
                clean_key = key.strip().strip('"').strip("'")
                quoted_key = f'"{clean_key}"'

                # Handle the value recursively
                quoted_value = self._fix_json_value(value.strip())

                fixed_pairs.append(f'{quoted_key}: {quoted_value}')

            return '{' + ', '.join(fixed_pairs) + '}'

        except Exception:
            # If fixing fails, return original
            return obj_str

    def _parse_json_pairs(self, content: str) -> list:
        """Parse key-value pairs from JSON object content, handling nested structures."""
        pairs = []
        current_pair = ""
        brace_count = 0
        bracket_count = 0
        in_quotes = False
        quote_char = None

        i = 0
        while i < len(content):
            char = content[i]

            # Handle quotes
            if char in ['"', "'"] and (i == 0 or content[i-1] != '\\'):
                if not in_quotes:
                    in_quotes = True
                    quote_char = char
                elif char == quote_char:
                    in_quotes = False
                    quote_char = None

            # Only count braces/brackets outside of quotes
            if not in_quotes:
                if char == '{':
                    brace_count += 1
                elif char == '}':
                    brace_count -= 1
                elif char == '[':
                    bracket_count += 1
                elif char == ']':
                    bracket_count -= 1
                elif char == ',' and brace_count == 0 and bracket_count == 0:
                    # Found a top-level comma separator
                    if current_pair.strip():
                        pairs.append(self._split_key_value(current_pair.strip()))
                    current_pair = ""
                    i += 1
                    continue

            current_pair += char
            i += 1

        # Add the last pair
        if current_pair.strip():
            pairs.append(self._split_key_value(current_pair.strip()))

        return pairs

    def _split_key_value(self, pair: str) -> tuple:
        """Split a key-value pair, handling nested colons properly."""
        # Find the first colon that's not inside quotes or nested structures
        brace_count = 0
        bracket_count = 0
        in_quotes = False
        quote_char = None

        for i, char in enumerate(pair):
            # Handle quotes
            if char in ['"', "'"] and (i == 0 or pair[i-1] != '\\'):
                if not in_quotes:
                    in_quotes = True
                    quote_char = char
                elif char == quote_char:
                    in_quotes = False
                    quote_char = None

            # Only count structures outside of quotes
            if not in_quotes:
                if char == '{':
                    brace_count += 1
                elif char == '}':
                    brace_count -= 1
                elif char == '[':
                    bracket_count += 1
                elif char == ']':
                    bracket_count -= 1
                elif char == ':' and brace_count == 0 and bracket_count == 0:
                    # Found the key-value separator
                    return (pair[:i], pair[i+1:])

        # If no colon found, treat as key with empty value
        return (pair, "")

    def _fix_json_value(self, value: str) -> str:
        """Fix a JSON value, handling different types recursively."""
        value = value.strip()

        # Handle empty values
        if not value:
            return '""'

        # Handle boolean and null values
        if value.lower() in ['true', 'false', 'null']:
            return value.lower()

        # Handle numeric values (including negative numbers)
        if value.isdigit() or (value.replace('.', '', 1).replace('-', '', 1).isdigit() and value.count('.') <= 1 and value.count('-') <= 1):
            return value

        # Handle already quoted strings
        if (value.startswith('"') and value.endswith('"')) or (value.startswith("'") and value.endswith("'")):
            # Normalize to double quotes
            return f'"{value[1:-1]}"'

        # Handle nested objects
        if value.startswith('{') and value.endswith('}'):
            return self._fix_json_object(value)

        # Handle arrays
        if value.startswith('[') and value.endswith(']'):
            return self._fix_json_array(value)

        # Handle unquoted strings
        return f'"{value}"'

    def _fix_json_array(self, arr_str: str) -> str:
        """Fix malformed JSON array by processing each element."""
        try:
            # Remove outer brackets
            content = arr_str[1:-1].strip()
            if not content:
                return "[]"

            # Parse array elements
            elements = self._parse_array_elements(content)

            # Fix each element
            fixed_elements = []
            for element in elements:
                fixed_elements.append(self._fix_json_value(element.strip()))

            return '[' + ', '.join(fixed_elements) + ']'

        except Exception:
            # If fixing fails, return original
            return arr_str

    def _parse_array_elements(self, content: str) -> list:
        """Parse array elements, handling nested structures."""
        elements = []
        current_element = ""
        brace_count = 0
        bracket_count = 0
        in_quotes = False
        quote_char = None

        i = 0
        while i < len(content):
            char = content[i]

            # Handle quotes
            if char in ['"', "'"] and (i == 0 or content[i-1] != '\\'):
                if not in_quotes:
                    in_quotes = True
                    quote_char = char
                elif char == quote_char:
                    in_quotes = False
                    quote_char = None

            # Only count structures outside of quotes
            if not in_quotes:
                if char == '{':
                    brace_count += 1
                elif char == '}':
                    brace_count -= 1
                elif char == '[':
                    bracket_count += 1
                elif char == ']':
                    bracket_count -= 1
                elif char == ',' and brace_count == 0 and bracket_count == 0:
                    # Found a top-level comma separator
                    if current_element.strip():
                        elements.append(current_element.strip())
                    current_element = ""
                    i += 1
                    continue

            current_element += char
            i += 1

        # Add the last element
        if current_element.strip():
            elements.append(current_element.strip())

        return elements


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
            print("\nAvailable tools:")
            for i,tool in enumerate(tools):
                print(f"  -{i+1:2}. {tool.get('name', 'Unknown')}")
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
                for i,tool in enumerate(tools):
                    print(f"  -{i+1:2}. {tool.get('name', 'Unknown')}")
                sys.exit(1)

        finally:
            await cli.stop_server()

    else:
        # Tool name with arguments - execute tool
        tool_name = sys.argv[1]
        raw_output = False
        arguments_list = sys.argv[2:]

        # Check for --raw_output and remove it from arguments_list
        for arg in arguments_list:
            if arg == "--raw_output":
                raw_output = True
                arguments_list.remove(arg)
                break
        arguments = MCPTestCLI().parse_arguments(arguments_list)
        cli = MCPTestCLI()

        if not await cli.start_server(raw_mode=raw_output):
            sys.exit(1)
        try:
            if not await cli.initialize_server(raw_mode=raw_output):
                sys.exit(1)
            if not raw_output:
                print(f"🚀 Executing tool: {tool_name}")
                print(f"📝 Arguments: {json.dumps(arguments, indent=2)}")
                print()

            result = await cli.call_tool(tool_name, arguments,raw_mode=raw_output)
            cli.print_tool_result(result,raw_mode=raw_output)

        finally:
            await cli.stop_server(raw_mode=raw_output)


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("\n⚠️ Interrupted by user")
        sys.exit(1)
    except Exception as e:
        print(f"❌ Unexpected error: {e}")
        sys.exit(1)
