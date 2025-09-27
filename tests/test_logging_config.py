#!/usr/bin/env python3
"""Test script to verify logging configuration for MCP."""

import logging
import sys
import os

# Add the project root to the path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from powerpoint_mcp_server.config import get_config

def test_logging_config():
    """Test the logging configuration."""
    config = get_config()
    
    # Create log file handler
    log_file = "test_powerpoint_mcp_server.log"
    file_handler = logging.FileHandler(log_file, mode='w', encoding='utf-8')
    file_handler.setLevel(logging.DEBUG)

    # For MCP servers, we should minimize stderr output to avoid [ERROR] logs in clients
    # Only log ERROR and CRITICAL to stderr, everything else goes to file only
    console_handler = logging.StreamHandler(sys.stderr)
    console_handler.setLevel(logging.ERROR)  # Only ERROR and CRITICAL to stderr

    # Create formatter
    formatter = logging.Formatter(config.log_format)
    file_handler.setFormatter(formatter)
    console_handler.setFormatter(formatter)

    # Configure root logger
    logging.basicConfig(
        level=logging.DEBUG,
        handlers=[file_handler, console_handler],
        force=True  # Force reconfiguration
    )

    # Set all third-party loggers to ERROR level to minimize stderr output for MCP
    third_party_loggers = ['fastmcp', 'mcp', 'asyncio', 'urllib3', 'requests']
    for logger_name in third_party_loggers:
        logging.getLogger(logger_name).setLevel(logging.ERROR)

    logger = logging.getLogger(__name__)
    
    print("Testing logging configuration...")
    print("This should appear on stdout")
    
    # Test different log levels
    logger.debug("This DEBUG message should only go to file")
    logger.info("This INFO message should only go to file")
    logger.warning("This WARNING message should only go to file")
    logger.error("This ERROR message should go to both file and stderr")
    logger.critical("This CRITICAL message should go to both file and stderr")
    
    # Test third-party logger
    fastmcp_logger = logging.getLogger('fastmcp')
    fastmcp_logger.info("This FastMCP INFO should only go to file")
    fastmcp_logger.error("This FastMCP ERROR should go to both file and stderr")
    
    print(f"Check the log file: {log_file}")
    print("Only ERROR and CRITICAL messages should have appeared on stderr")

if __name__ == "__main__":
    test_logging_config()