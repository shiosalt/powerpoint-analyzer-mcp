"""
Simple test script to verify bug fixes without full MCP integration.
"""

import sys
import json
import asyncio
from pathlib import Path

# Add the project root to the path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from powerpoint_mcp_server.server import PowerPointMCPServer

async def test_bug_fixes():
    """Test all bug fixes directly using the server."""
    print("Testing PowerPoint MCP bug fixes...")
    print("=" * 50)
    
    # Initialize server
    server = PowerPointMCPServer()
    test_file = "tests/test_files/test_complex.pptx"
    
    try:
        # Test 1: analyze_text_formatting - should return non-zero counts
        print("Test 1: analyze_text_formatting")
        result1 = await server._analyze_text_formatting({
            "file_path": test_file,
            "slide_numbers": [],
            "formatting_filter": {
                "include_bold": True,
                "include_formatting_details": True
            },
            "grouping": "by_formatting_type"
        })
        
        data1 = json.loads(result1.content[0].text)
        formatting_counts = data1.get("formatting_summary", {}).get("formatting_counts", {})
        
        print(f"  Formatting counts: {formatting_counts}")
        
        total_formatting = sum([
            formatting_counts.get("bold", 0),
            formatting_counts.get("italic", 0),
            formatting_counts.get("underline", 0),
            formatting_counts.get("highlight", 0),
            formatting_counts.get("strikethrough", 0),
            formatting_counts.get("colored_text", 0),
            formatting_counts.get("hyperlinks", 0)
        ])
        
        if total_formatting > 0:
            print("  ✅ PASS: Found formatting (counts are not all zero)")
        else:
            print("  ❌ FAIL: All formatting counts are zero")
        
        print()
        
        # Test 2: extract_formatted_text - test italic and hyperlinks
        print("Test 2: extract_formatted_text - italic")
        result2 = await server._extract_text_formatting({
            "file_path": test_file,
            "formatting_type": "italic",
            "slide_numbers": None
        })
        
        data2 = json.loads(result2.content[0].text)
        print(f"  Italic extraction summary: {data2.get('summary', {})}")
        
        if "error" not in data2:
            print("  ✅ PASS: Italic extraction completed without error")
        else:
            print(f"  ❌ FAIL: Italic extraction error: {data2['error']}")
        
        print()
        
        # Test 3: extract_formatted_text - test hyperlinks
        print("Test 3: extract_formatted_text - hyperlinks")
        result3 = await server._extract_text_formatting({
            "file_path": test_file,
            "formatting_type": "hyperlinks",
            "slide_numbers": None
        })
        
        data3 = json.loads(result3.content[0].text)
        print(f"  Hyperlinks extraction summary: {data3.get('summary', {})}")
        
        if "error" not in data3:
            print("  ✅ PASS: Hyperlinks extraction completed without error")
        else:
            print(f"  ❌ FAIL: Hyperlinks extraction error: {data3['error']}")
        
        print()
        
        # Test 4: extract_formatted_text - test position accuracy
        print("Test 4: extract_formatted_text - position accuracy")
        result4 = await server._extract_text_formatting({
            "file_path": test_file,
            "formatting_type": "bold",
            "slide_numbers": None
        })
        
        data4 = json.loads(result4.content[0].text)
        
        position_errors = 0
        segment_errors = 0
        
        for slide_result in data4.get("results_by_slide", []):
            for segment in slide_result.get("formatted_segments", []):
                # Check position accuracy
                start_pos = segment.get("start_position", -1)
                if start_pos < 0:
                    position_errors += 1
                
                # Check that segment text is not the complete text
                segment_text = segment.get("text", "")
                complete_text = slide_result.get("complete_text", "")
                if segment_text == complete_text and len(complete_text) > 50:
                    segment_errors += 1
        
        if position_errors == 0:
            print("  ✅ PASS: All positions are valid (>= 0)")
        else:
            print(f"  ❌ FAIL: {position_errors} invalid positions found")
        
        if segment_errors == 0:
            print("  ✅ PASS: Formatted segments are not complete text")
        else:
            print(f"  ❌ FAIL: {segment_errors} segments equal complete text")
        
        print()
        
        # Test 5: extract_table_data - test summary accuracy
        print("Test 5: extract_table_data - summary accuracy")
        result5 = await server._extract_table_data({
            "file_path": test_file,
            "slide_numbers": [1, 2, 3, 4],
            "output_format": "structured",
            "include_metadata": True
        })
        
        data5 = json.loads(result5.content[0].text)
        summary5 = data5.get("summary", {})
        extracted_tables = data5.get("extracted_tables", [])
        
        print(f"  Summary: {summary5}")
        print(f"  Extracted tables count: {len(extracted_tables)}")
        
        if summary5.get("total_tables_found", 0) == len(extracted_tables):
            print("  ✅ PASS: Summary count matches extracted tables count")
        else:
            print("  ❌ FAIL: Summary count does not match extracted tables count")
        
        if "error" not in data5:
            print("  ✅ PASS: Table extraction completed without error")
        else:
            print(f"  ❌ FAIL: Table extraction error: {data5['error']}")
        
        print()
        
        # Test 6: extract_table_data - slide number handling
        print("Test 6: extract_table_data - slide number handling")
        try:
            result6 = await server._extract_table_data({
                "file_path": test_file,
                "slide_numbers": [2],  # Valid slide number
                "output_format": "structured",
                "include_metadata": True
            })
            
            data6 = json.loads(result6.content[0].text)
            
            if "error" not in data6:
                print("  ✅ PASS: Valid slide number processed successfully")
            else:
                print(f"  ❌ FAIL: Valid slide number caused error: {data6['error']}")
        
        except Exception as e:
            print(f"  ❌ FAIL: Exception with valid slide number: {e}")
        
        # Test invalid slide number
        try:
            result6b = await server._extract_table_data({
                "file_path": test_file,
                "slide_numbers": [999],  # Invalid slide number
                "output_format": "structured",
                "include_metadata": True
            })
            
            data6b = json.loads(result6b.content[0].text)
            
            if "error" in data6b or "Invalid slide numbers" in str(data6b):
                print("  ✅ PASS: Invalid slide number properly handled")
            else:
                print("  ❌ FAIL: Invalid slide number not properly handled")
        
        except Exception as e:
            if "Invalid slide numbers" in str(e):
                print("  ✅ PASS: Invalid slide number properly rejected")
            else:
                print(f"  ❌ FAIL: Unexpected exception: {e}")
        
        print()
        
        # Test 7: query_slides - validation
        print("Test 7: query_slides - validation")
        
        # Test with invalid return fields
        result7 = await server._query_slides({
            "file_path": test_file,
            "search_criteria": {},
            "return_fields": ["invalid_field"],
            "limit": 50
        })
        
        data7 = json.loads(result7.content[0].text)
        
        if "results" in data7 and len(data7["results"]) == 0:
            print("  ✅ PASS: Invalid return fields returned zero results")
        elif "error" in data7:
            print("  ✅ PASS: Invalid return fields caused validation error")
        else:
            print(f"  ❌ FAIL: Invalid return fields not properly handled: {data7}")
        
        print()
        
        print("Bug fix testing completed!")
        
    except Exception as e:
        print(f"Error during testing: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    asyncio.run(test_bug_fixes())