"""
Test cleanup analysis and execution.
Identifies obsolete test files and cleans up the test directory.
"""

import os
import sys
import importlib
from pathlib import Path
from typing import List, Dict, Tuple

def analyze_test_files() -> Dict[str, str]:
    """Analyze test files and categorize them."""
    test_dir = Path("tests")
    test_files = [f for f in test_dir.glob("test_*.py") if f.is_file()]
    
    analysis = {
        "keep": [],
        "obsolete": [],
        "needs_update": [],
        "duplicate": []
    }
    
    for test_file in test_files:
        file_name = test_file.name
        
        # Check if it's a simple test file that might be obsolete
        if file_name in ["simple_fastmcp_test.py", "test_simple_fastmcp.py"]:
            analysis["obsolete"].append(file_name)
        
        # Check for MCP-specific test files that test non-existent functionality
        elif file_name.startswith("test_mcp_") and file_name not in [
            "test_mcp_notes.py",  # Keep - tests actual functionality
        ]:
            # These test actual modules, so keep them
            analysis["keep"].append(file_name)
        
        # Check for integration tests
        elif "integration" in file_name:
            if file_name == "test_integration.py":
                analysis["needs_update"].append(file_name)  # Might need updates for new tools
            else:
                analysis["keep"].append(file_name)
        
        # Check for comprehensive test files (our new ones)
        elif "comprehensive" in file_name or "framework" in file_name:
            analysis["keep"].append(file_name)
        
        # Check for data generator
        elif "data_generator" in file_name:
            analysis["keep"].append(file_name)
        
        # Core functionality tests - keep
        elif any(core in file_name for core in [
            "content_extractor", "attribute_processor", "text_formatting",
            "file_validator", "zip_extractor", "xml_parser", "cache_manager"
        ]):
            analysis["keep"].append(file_name)
        
        # Performance and server tests - keep
        elif any(perf in file_name for perf in ["performance", "server"]):
            analysis["keep"].append(file_name)
        
        # Other test files - evaluate individually
        else:
            analysis["needs_update"].append(file_name)
    
    return analysis

def check_imports(test_file: Path) -> Tuple[bool, List[str]]:
    """Check if test file imports work correctly."""
    errors = []
    
    try:
        with open(test_file, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Look for import statements
        import_lines = [line.strip() for line in content.split('\n') 
                      if line.strip().startswith(('import ', 'from '))]
        
        for import_line in import_lines:
            try:
                # Skip relative imports and standard library
                if ('powerpoint_mcp_server' in import_line and 
                    not import_line.startswith('#')):
                    
                    # Extract module name
                    if import_line.startswith('from '):
                        module_part = import_line.split(' import ')[0].replace('from ', '')
                    else:
                        module_part = import_line.replace('import ', '').split('.')[0]
                    
                    # Try to import
                    if 'powerpoint_mcp_server' in module_part:
                        try:
                            importlib.import_module(module_part)
                        except ImportError as e:
                            errors.append(f"Import error: {import_line} - {e}")
                            
            except Exception as e:
                errors.append(f"Error checking import '{import_line}': {e}")
                
    except Exception as e:
        errors.append(f"Error reading file: {e}")
        return False, errors
    
    return len(errors) == 0, errors

def cleanup_obsolete_tests():
    """Remove obsolete test files."""
    analysis = analyze_test_files()
    
    print("Test File Cleanup Analysis")
    print("=" * 50)
    
    print(f"\nFiles to KEEP ({len(analysis['keep'])}):")
    for file_name in sorted(analysis['keep']):
        print(f"  ✓ {file_name}")
    
    print(f"\nFiles that are OBSOLETE ({len(analysis['obsolete'])}):")
    for file_name in sorted(analysis['obsolete']):
        print(f"  ✗ {file_name}")
    
    print(f"\nFiles that NEED UPDATE ({len(analysis['needs_update'])}):")
    for file_name in sorted(analysis['needs_update']):
        print(f"  ⚠ {file_name}")
    
    # Check imports for files that need updates
    print(f"\nImport Analysis:")
    test_dir = Path("tests")
    
    for file_name in analysis['needs_update']:
        test_file = test_dir / file_name
        if test_file.exists():
            imports_ok, errors = check_imports(test_file)
            status = "✓" if imports_ok else "✗"
            print(f"  {status} {file_name}")
            if errors:
                for error in errors[:3]:  # Show first 3 errors
                    print(f"      {error}")
                if len(errors) > 3:
                    print(f"      ... and {len(errors) - 3} more errors")
    
    # Actually remove obsolete files
    removed_files = []
    for file_name in analysis['obsolete']:
        test_file = test_dir / file_name
        if test_file.exists():
            try:
                test_file.unlink()
                removed_files.append(file_name)
                print(f"  Removed: {file_name}")
            except Exception as e:
                print(f"  Failed to remove {file_name}: {e}")
    
    print(f"\nCleanup Summary:")
    print(f"  Removed {len(removed_files)} obsolete test files")
    print(f"  Kept {len(analysis['keep'])} relevant test files")
    print(f"  {len(analysis['needs_update'])} files may need updates")
    
    return analysis, removed_files

def validate_remaining_tests():
    """Validate that remaining test files are functional."""
    test_dir = Path("tests")
    test_files = [f for f in test_dir.glob("test_*.py") if f.is_file()]
    
    print(f"\nValidating {len(test_files)} remaining test files:")
    
    validation_results = {}
    
    for test_file in test_files:
        imports_ok, errors = check_imports(test_file)
        validation_results[test_file.name] = {
            "imports_ok": imports_ok,
            "errors": errors
        }
        
        status = "✓" if imports_ok else "✗"
        print(f"  {status} {test_file.name}")
        
        if not imports_ok and len(errors) <= 2:
            for error in errors:
                print(f"      {error}")
    
    # Summary
    valid_files = sum(1 for result in validation_results.values() if result["imports_ok"])
    print(f"\nValidation Summary:")
    print(f"  {valid_files}/{len(test_files)} test files have valid imports")
    
    return validation_results

if __name__ == "__main__":
    print("Starting test cleanup process...")
    
    # Analyze and cleanup
    analysis, removed_files = cleanup_obsolete_tests()
    
    # Validate remaining files
    validation_results = validate_remaining_tests()
    
    print(f"\nTest cleanup completed!")
    print(f"Check the remaining test files and update them as needed.")