#!/usr/bin/env python3
"""Test script to verify query_slides validation works correctly."""

import asyncio
import json
import sys
from pathlib import Path

# Add project root to path
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

from powerpoint_mcp_server.core.slide_query_engine import SlideQueryEngine

async def test_validation():
    """Test validation of search criteria."""
    engine = SlideQueryEngine()
    
    # Test case 1: Invalid field name
    test_criteria_1 = {"uso": "bold"}
    result_1 = engine.validate_search_criteria_dict(test_criteria_1)
    print("Test 1 - Invalid field 'uso':")
    print(f"  Is valid: {result_1['is_valid']}")
    print(f"  Errors: {result_1['errors']}")
    print()
    
    # Test case 2: Valid criteria
    test_criteria_2 = {"title": {"contains": "test"}}
    result_2 = engine.validate_search_criteria_dict(test_criteria_2)
    print("Test 2 - Valid criteria:")
    print(f"  Is valid: {result_2['is_valid']}")
    print(f"  Errors: {result_2['errors']}")
    print()
    
    # Test case 3: Multiple invalid fields
    test_criteria_3 = {"uso": "bold", "invalid_field": "test"}
    result_3 = engine.validate_search_criteria_dict(test_criteria_3)
    print("Test 3 - Multiple invalid fields:")
    print(f"  Is valid: {result_3['is_valid']}")
    print(f"  Errors: {result_3['errors']}")
    print()

if __name__ == "__main__":
    asyncio.run(test_validation())