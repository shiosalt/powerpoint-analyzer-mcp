#!/usr/bin/env python3
"""
Simple FastMCP example to understand correct usage
"""

import logging
import sys
from mcp.server.fastmcp import FastMCP

# Configure logging
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("fastmcp_example.log", mode='w', encoding='utf-8'),
        logging.StreamHandler(sys.stderr)
    ]
)

logger = logging.getLogger(__name__)

# Create FastMCP instance
mcp = FastMCP("example-server")

@mcp.tool()
def hello(name: str) -> str:
    """Say hello to someone.
    
    Args:
        name: The name of the person to greet
        
    Returns:
        A greeting message
    """
    logger.info(f"hello tool called with name: {name}")
    return f"Hello, {name}!"

@mcp.tool()
def add_numbers(a: int, b: int) -> int:
    """Add two numbers.
    
    Args:
        a: First number
        b: Second number
        
    Returns:
        The sum of a and b
    """
    logger.info(f"add_numbers tool called with a: {a}, b: {b}")
    return a + b

def main():
    """Main entry point."""
    logger.info("Starting FastMCP example server...")
    
    try:
        # Run the FastMCP server
        mcp.run()
    except Exception as e:
        logger.error(f"Server error: {e}")
        import traceback
        logger.error(f"Traceback: {traceback.format_exc()}")
        raise

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        logger.info("Server stopped by user")
        sys.exit(0)
    except Exception as e:
        logger.error(f"Fatal error: {e}")
        sys.exit(1)