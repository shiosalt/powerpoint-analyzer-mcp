#!/usr/bin/env python3
"""Health check script for PowerPoint Analyzer MCP."""

import sys
import json
from pathlib import Path

# Add the parent directory to the path so we can import the server
sys.path.insert(0, str(Path(__file__).parent.parent))

from powerpoint_mcp_server.config import get_config
from powerpoint_mcp_server.utils.file_validator import FileValidator
from powerpoint_mcp_server.core.content_extractor import ContentExtractor


def check_dependencies():
    """Check if all required dependencies are available."""
    try:
        import xml.etree.ElementTree
        import zipfile
        import json
        import asyncio
        import logging
        print("✓ All required Python modules are available")
        return True
    except ImportError as e:
        print(f"✗ Missing required dependency: {e}")
        return False


def check_configuration():
    """Check server configuration."""
    try:
        config = get_config()
        print("✓ Server configuration loaded successfully")
        print(f"  - Server name: {config.server_name}")
        print(f"  - Server version: {config.server_version}")
        print(f"  - Log level: {config.log_level}")
        print(f"  - Max file size: {config.max_file_size_mb} MB")
        print(f"  - Cache enabled: {config.cache_enabled}")
        print(f"  - Debug mode: {config.debug_mode}")
        return True
    except Exception as e:
        print(f"✗ Configuration error: {e}")
        return False


def check_components():
    """Check if server components can be initialized."""
    try:
        # Test file validator
        validator = FileValidator()
        print("✓ FileValidator initialized successfully")
        
        # Test content extractor
        extractor = ContentExtractor()
        print("✓ ContentExtractor initialized successfully")
        
        return True
    except Exception as e:
        print(f"✗ Component initialization error: {e}")
        return False


def check_test_file_processing():
    """Check if a test PowerPoint file can be processed (if available)."""
    test_files = [
        "tests/test_files/sample.pptx",
        "tests/test_files/test_presentation.pptx",
        "sample.pptx"
    ]
    
    validator = FileValidator()
    
    for test_file in test_files:
        if Path(test_file).exists():
            try:
                is_valid, error_msg = validator.validate_file(test_file)
                if is_valid:
                    print(f"✓ Test file {test_file} is valid and can be processed")
                    return True
                else:
                    print(f"⚠ Test file {test_file} validation failed: {error_msg}")
            except Exception as e:
                print(f"⚠ Error checking test file {test_file}: {e}")
    
    print("ℹ No test PowerPoint files found (this is optional)")
    return True


def main():
    """Run health checks."""
    print("PowerPoint Analyzer MCP Health Check")
    print("=" * 40)
    
    checks = [
        ("Dependencies", check_dependencies),
        ("Configuration", check_configuration),
        ("Components", check_components),
        ("Test File Processing", check_test_file_processing)
    ]
    
    all_passed = True
    
    for check_name, check_func in checks:
        print(f"\n{check_name}:")
        try:
            result = check_func()
            if not result:
                all_passed = False
        except Exception as e:
            print(f"✗ {check_name} check failed with exception: {e}")
            all_passed = False
    
    print("\n" + "=" * 40)
    if all_passed:
        print("✓ All health checks passed! Server should be ready to run.")
        return 0
    else:
        print("✗ Some health checks failed. Please review the issues above.")
        return 1


if __name__ == "__main__":
    sys.exit(main())