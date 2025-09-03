"""Setup script for PowerPoint Analyzer MCP."""

from setuptools import setup, find_packages

with open("requirements.txt", "r", encoding="utf-8") as f:
    requirements = [line.strip() for line in f if line.strip() and not line.startswith("#")]

setup(
    name="powerpoint-analyzer-mcp",
    version="0.1.0",
    description="PowerPoint Analyzer MCP server for extracting structured information from PowerPoint files",
    author="PowerPoint Analyzer MCP",
    packages=find_packages(),
    install_requires=requirements,
    python_requires=">=3.8",
    entry_points={
        "console_scripts": [
            "powerpoint-analyzer-mcp=powerpoint_mcp_server.server:main",
        ],
    },
    classifiers=[
        "Development Status :: 3 - Alpha",
        "Intended Audience :: Developers",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
    ],
)