#!/usr/bin/env python3
"""
Examples demonstrating the new Python-style slide selection functionality.

This script shows how to use the enhanced slide_numbers parameter with various
formats for efficient slide processing.
"""

import asyncio
import json
from pathlib import Path

# Add the parent directory to the path so we can import the server
import sys
sys.path.append(str(Path(__file__).parent.parent))

from powerpoint_mcp_server.server import PowerPointMCPServer


async def demonstrate_slide_selection():
    """Demonstrate various slide selection formats."""
    
    # Initialize the server
    server = PowerPointMCPServer()
    
    # Example PowerPoint file (replace with your actual file)
    sample_file = "sample_presentation.pptx"
    
    print("=== PowerPoint Analyzer MCP - Slide Selection Examples ===\n")
    
    # Example 1: Extract tables from first 10 slides
    print("1. Extract tables from first 10 slides using ':10'")
    try:
        arguments = {
            "file_path": sample_file,
            "slide_numbers": ":10",
            "output_format": "structured",
            "include_metadata": True
        }
        result = await server._extract_table_data(arguments)
        print("✓ Success: Processed first 10 slides")
    except Exception as e:
        print(f"✗ Error: {e}")
    
    print()
    
    # Example 2: Extract bold text from slides 5-15
    print("2. Extract bold text from slides 5-15 using '5:15'")
    try:
        arguments = {
            "file_path": sample_file,
            "formatting_type": "bold",
            "slide_numbers": "5:15"
        }
        result = await server._extract_text_formatting(arguments)
        print("✓ Success: Processed slides 5-15")
    except Exception as e:
        print(f"✗ Error: {e}")
    
    print()
    
    # Example 3: Query specific slides
    print("3. Query specific slides using '1,5,10,15'")
    try:
        arguments = {
            "file_path": sample_file,
            "search_criteria": {
                "slide_numbers": "1,5,10,15",
                "title": {"contains": ""}  # Match any title
            },
            "return_fields": ["slide_number", "title"],
            "limit": 50
        }
        result = await server._query_slides(arguments)
        print("✓ Success: Queried specific slides")
    except Exception as e:
        print(f"✗ Error: {e}")
    
    print()
    
    # Example 4: Extract tables from slide 20 to end
    print("4. Extract tables from slide 20 to end using '20:'")
    try:
        arguments = {
            "file_path": sample_file,
            "slide_numbers": "20:",
            "output_format": "structured",
            "include_metadata": True
        }
        result = await server._extract_table_data(arguments)
        print("✓ Success: Processed slides 20 to end")
    except Exception as e:
        print(f"✗ Error: {e}")
    
    print()
    
    # Example 5: Single slide
    print("5. Extract hyperlinks from single slide using '3'")
    try:
        arguments = {
            "file_path": sample_file,
            "formatting_type": "hyperlinks",
            "slide_numbers": 3
        }
        result = await server._extract_text_formatting(arguments)
        print("✓ Success: Processed single slide")
    except Exception as e:
        print(f"✗ Error: {e}")
    
    print()
    
    # Example 6: Traditional list format (still supported)
    print("6. Extract tables using traditional list format [1, 3, 5]")
    try:
        arguments = {
            "file_path": sample_file,
            "slide_numbers": [1, 3, 5],
            "output_format": "structured",
            "include_metadata": True
        }
        result = await server._extract_table_data(arguments)
        print("✓ Success: Processed slides using list format")
    except Exception as e:
        print(f"✗ Error: {e}")
    
    print("\n=== Slide Selection Format Summary ===")
    print("• All slides:        None or omit parameter")
    print("• Single slide:      3 or '3'")
    print("• Specific slides:   [1, 5, 10] or '1,5,10'")
    print("• First N slides:    ':10' (slides 1-10)")
    print("• Slide range:       '5:20' (slides 5-20)")
    print("• From slide to end: '25:' (slides 25 to end)")
    print("• With brackets:     '[:10]', '[5:20]', '[25:]' (optional)")


def demonstrate_slide_selector_utility():
    """Demonstrate the slide selector utility directly."""
    from powerpoint_mcp_server.utils.slide_selector import parse_slide_numbers
    
    print("\n=== Slide Selector Utility Examples ===")
    
    total_slides = 100
    
    examples = [
        (None, "All slides"),
        (5, "Single slide"),
        ([1, 5, 10], "Specific slides (list)"),
        (":10", "First 10 slides"),
        ("5:20", "Slides 5-20"),
        ("25:", "Slides 25 to end"),
        ("1,3,5,7,9", "Comma-separated"),
        ("[:15]", "First 15 slides (with brackets)"),
        ("[10:30]", "Slides 10-30 (with brackets)")
    ]
    
    for slide_spec, description in examples:
        try:
            result = parse_slide_numbers(slide_spec, total_slides)
            if len(result) <= 10:
                result_str = str(result)
            else:
                result_str = f"[{result[0]}, {result[1]}, ..., {result[-2]}, {result[-1]}] ({len(result)} slides)"
            
            print(f"• {description:25} {str(slide_spec):15} → {result_str}")
        except Exception as e:
            print(f"• {description:25} {str(slide_spec):15} → Error: {e}")


if __name__ == "__main__":
    print("PowerPoint Analyzer MCP - Slide Selection Examples")
    print("=" * 50)
    
    # Demonstrate the utility function
    demonstrate_slide_selector_utility()
    
    # Demonstrate with actual server (requires a PowerPoint file)
    print("\nTo test with actual PowerPoint files:")
    print("1. Place a PowerPoint file named 'sample_presentation.pptx' in this directory")
    print("2. Uncomment the line below and run the script")
    print("# asyncio.run(demonstrate_slide_selection())")
    
    # Uncomment the next line to test with actual files
    # asyncio.run(demonstrate_slide_selection())