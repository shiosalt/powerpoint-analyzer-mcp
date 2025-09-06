#!/usr/bin/env python3
"""
PowerPoint Analyzer MCP - Quick Test Tool

A simplified wrapper around mcp_test_cli.py for common testing scenarios.
"""

import asyncio
import sys
import subprocess
from pathlib import Path


def run_cli(*args):
    """Run the MCP test CLI with given arguments."""
    cmd = ["python", "mcp_test_cli.py"] + list(args)
    return subprocess.run(cmd, capture_output=False)


def main():
    """Main entry point with common test scenarios."""
    if len(sys.argv) == 1:
        print("🧪 PowerPoint Analyzer MCP - Quick Test Tool")
        print("=" * 50)
        print()
        print("Common test scenarios:")
        print("  python test_tools.py list                    # List all tools")
        print("  python test_tools.py help <tool_name>        # Show tool help")
        print("  python test_tools.py extract <file.pptx>     # Extract content")
        print("  python test_tools.py attrs <file.pptx>       # Get attributes")
        print("  python test_tools.py slide <file.pptx> <n>   # Get slide info")
        print("  python test_tools.py bold <file.pptx>        # Extract bold text")
        print()
        print("Or use mcp_test_cli.py directly for full control.")
        return

    command = sys.argv[1].lower()

    if command == "list":
        # List all tools
        run_cli()

    elif command == "help" and len(sys.argv) == 3:
        # Show tool help
        tool_name = sys.argv[2]
        run_cli(tool_name)

    elif command == "extract" and len(sys.argv) == 3:
        # Extract PowerPoint content
        file_path = sys.argv[2]
        run_cli("extract_powerpoint_content", "--file_path", file_path)

    elif command == "attrs" and len(sys.argv) >= 3:
        # Get PowerPoint attributes
        file_path = sys.argv[2]
        attributes = sys.argv[3:] if len(sys.argv) > 3 else ["title", "subtitle", "text_elements"]
        attrs_json = '["' + '", "'.join(attributes) + '"]'
        run_cli("get_powerpoint_attributes", "--file_path", file_path, "--attributes", attrs_json)

    elif command == "slide" and len(sys.argv) == 4:
        # Get slide info
        file_path = sys.argv[2]
        slide_number = sys.argv[3]
        run_cli("get_slide_info", "--file_path", file_path, "--slide_number", slide_number)

    elif command == "bold" and len(sys.argv) >= 4:
        # Extract bold text
        file_path = sys.argv[2]
        format_type = "bold"
        run_cli("extract_text_formatting", "--file_path", file_path, "--formatting_type", format_type)

    elif command == "format" and len(sys.argv) >= 4:
        # Extract formatted text
        file_path = sys.argv[2]
        format_type = sys.argv[3]
        run_cli("extract_text_formatting", "--file_path", file_path, "--formatting_type", format_type)

    else:
        print("❌ Invalid command or arguments")
        print("Run 'python test_tools.py' for usage help")
        sys.exit(1)


if __name__ == "__main__":
    main()