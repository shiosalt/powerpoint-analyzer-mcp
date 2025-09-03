#!/usr/bin/env python3
"""Startup script for PowerPoint Analyzer MCP with enhanced configuration options."""

import argparse
import asyncio
import os
import sys
from pathlib import Path

# Add the parent directory to the path so we can import the server
sys.path.insert(0, str(Path(__file__).parent.parent))

import main
from powerpoint_mcp_server.config import get_config_manager


def parse_arguments():
    """Parse command line arguments."""
    parser = argparse.ArgumentParser(
        description="PowerPoint Analyzer MCP - Extract structured content from PowerPoint files"
    )
    
    parser.add_argument(
        "--log-level",
        choices=["DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"],
        help="Set logging level (overrides POWERPOINT_MCP_LOG_LEVEL)"
    )
    
    parser.add_argument(
        "--max-file-size",
        type=int,
        help="Maximum file size in MB (overrides POWERPOINT_MCP_MAX_FILE_SIZE)"
    )
    
    parser.add_argument(
        "--timeout",
        type=int,
        help="Processing timeout in seconds (overrides POWERPOINT_MCP_TIMEOUT)"
    )
    
    parser.add_argument(
        "--no-cache",
        action="store_true",
        help="Disable caching"
    )
    
    parser.add_argument(
        "--debug",
        action="store_true",
        help="Enable debug mode"
    )
    
    parser.add_argument(
        "--version",
        action="version",
        version="PowerPoint Analyzer MCP 0.1.0"
    )
    
    return parser.parse_args()


def apply_cli_config(args):
    """Apply command line configuration overrides."""
    config_manager = get_config_manager()
    
    updates = {}
    
    if args.log_level:
        updates['log_level'] = args.log_level
    
    if args.max_file_size:
        updates['max_file_size_mb'] = args.max_file_size
    
    if args.timeout:
        updates['processing_timeout_seconds'] = args.timeout
    
    if args.no_cache:
        updates['cache_enabled'] = False
    
    if args.debug:
        updates['debug_mode'] = True
        if 'log_level' not in updates:
            updates['log_level'] = 'DEBUG'
    
    if updates:
        config_manager.update_config(**updates)
        print(f"Applied configuration overrides: {updates}")


def main():
    """Main entry point for the startup script."""
    try:
        # Parse command line arguments
        args = parse_arguments()
        
        # Apply CLI configuration overrides
        apply_cli_config(args)
        
        # Run the main server function
        main.main()
        
    except KeyboardInterrupt:
        print("\nServer stopped by user")
        return 0
    except Exception as e:
        print(f"Server error: {e}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nServer stopped by user")
        sys.exit(0)
    except Exception as e:
        print(f"Server error: {e}", file=sys.stderr)
        sys.exit(1)