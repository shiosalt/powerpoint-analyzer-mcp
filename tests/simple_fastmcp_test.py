#!/usr/bin/env python3
"""
Simple test to verify FastMCP 2.0 is working
"""

from fastmcp import FastMCP

# Create a simple FastMCP server for testing
mcp = FastMCP("test-server")

@mcp.tool
def hello(name: str) -> str:
    """Say hello to someone.
    
    Args:
        name: The name of the person to greet
        
    Returns:
        A greeting message
    """
    return f"Hello, {name}!"

@mcp.tool
def add(a: int, b: int) -> int:
    """Add two numbers.
    
    Args:
        a: First number
        b: Second number
        
    Returns:
        The sum of a and b
    """
    return a + b

if __name__ == "__main__":
    print("Starting simple FastMCP test server...")
    print("Tools registered:")
    print("- hello: Say hello to someone")
    print("- add: Add two numbers")
    print("Starting server...")
    mcp.run()