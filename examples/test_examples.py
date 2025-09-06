#!/usr/bin/env python3
"""
PowerPoint Analyzer MCP - Test Examples

This script demonstrates various ways to test the MCP server using the CLI tools.
"""

import subprocess
import sys
from pathlib import Path


def run_command(description: str, command: list):
    """Run a command and display the result."""
    print(f"\n{'='*60}")
    print(f"🧪 {description}")
    print(f"Command: {' '.join(command)}")
    print('='*60)

    try:
        result = subprocess.run(command, capture_output=False, text=True)
        return result.returncode == 0
    except Exception as e:
        print(f"❌ Error running command: {e}")
        return False


def main():
    """Run example tests."""
    print("🚀 PowerPoint Analyzer MCP - Test Examples")
    print("This script demonstrates various testing scenarios.")

    # Check if test files exist
    test_files = [
        "tests/test_files/test_minimal.pptx",
        "tests/test_files/test_complex.pptx"
    ]

    available_files = []
    for file_path in test_files:
        if Path(file_path).exists():
            available_files.append(file_path)

    if not available_files:
        print("❌ No test files found. Please ensure test files exist in tests/test_files/")
        sys.exit(1)

    test_file = available_files[0]
    print(f"📁 Using test file: {test_file}")

    # Example 1: List all available tools
    run_command(
        "List all available tools",
        ["python", "mcp_test_cli.py"]
    )

    # Example 2: Show help for a specific tool
    run_command(
        "Show help for extract_powerpoint_content tool",
        ["python", "mcp_test_cli.py", "extract_powerpoint_content"]
    )

    # Example 3: Extract complete PowerPoint content
    run_command(
        "Extract complete PowerPoint content",
        ["python", "mcp_test_cli.py", "extract_powerpoint_content", "--file_path", test_file]
    )

    # Example 4: Get specific attributes
    run_command(
        "Get specific attributes (title, subtitle, object_counts)",
        ["python", "mcp_test_cli.py", "get_powerpoint_attributes",
         "--file_path", test_file,
         "--attributes", '["title", "subtitle", "object_counts"]']
    )

    # Example 5: Get slide info
    run_command(
        "Get information for slide 1",
        ["python", "mcp_test_cli.py", "get_slide_info",
         "--file_path", test_file, "--slide_number", "1"]
    )

    # Example 6: Extract bold text
    run_command(
        "Extract bold text formatting",
        ["python", "mcp_test_cli.py", "extract_text_formatting",
         "--file_path", test_file, "--formatting_type", "bold"]
    )


    # Example 7: Extract specific formatting
    run_command(
        "Extract italic text formatting",
        ["python", "mcp_test_cli.py", "extract_text_formatting",
         "--file_path", test_file, "--formatting_type", "italic"]
    )

    # Example 8: Get presentation overview
    run_command(
        "Get presentation overview",
        ["python", "mcp_test_cli.py", "get_presentation_overview",
         "--file_path", test_file, "--analysis_depth", "detailed"]
    )

    # Example 9: Using the simplified test_tools.py
    print(f"\n{'='*60}")
    print("🎯 Simplified Testing with test_tools.py")
    print('='*60)
    print("You can also use the simplified test_tools.py for common scenarios:")
    print(f"  python test_tools.py extract {test_file}")
    print(f"  python test_tools.py attrs {test_file} title subtitle")
    print(f"  python test_tools.py slide {test_file} 1")
    print(f"  python test_tools.py bold {test_file}")

    print(f"\n✅ Test examples completed!")
    print("💡 Try running the individual commands above to test specific functionality.")


if __name__ == "__main__":
    main()