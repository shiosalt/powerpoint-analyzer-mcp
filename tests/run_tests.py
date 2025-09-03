"""
Test runner that properly sets up the Python path for running tests.
"""

import sys
import os
from pathlib import Path

# Add the project root to Python path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

# Now we can import and run tests
if __name__ == "__main__":
    import pytest
    
    # Run specific test files that are known to work
    test_files = [
        "tests/test_data_generator.py",
        "tests/test_comprehensive_mcp_tools.py",
        # Add more test files as they are updated
    ]
    
    print("Running PowerPoint Analyzer MCP tests...")
    print(f"Project root: {project_root}")
    print(f"Python path: {sys.path[0]}")
    
    # Run pytest with the test files
    exit_code = pytest.main([
        "-v",
        "--tb=short",
        *test_files
    ])
    
    sys.exit(exit_code)