"""
MCP protocol integration testing framework.
Tests the PowerPoint MCP server using actual MCP client-server communication.
"""

import asyncio
import json
import logging
import subprocess
import sys
import time
from pathlib import Path
from typing import Dict, List, Any, Optional, Tuple
from dataclasses import dataclass
from contextlib import asynccontextmanager

# Import FastMCP client components
try:
    from fastmcp import Client
    from fastmcp.client.transports import PythonStdioTransport
    import subprocess
except ImportError:
    print("FastMCP client not available. Install with: pip install fastmcp")
    sys.exit(1)

logger = logging.getLogger(__name__)


@dataclass
class TestResult:
    """Result of MCP tool test execution."""
    tool_name: str
    parameters: Dict[str, Any]
    success: bool
    response_time: float
    response_data: Optional[Dict]
    error_message: Optional[str]
    test_description: str


class MCPTestClient:
    """MCP client for integration testing."""
    
    def __init__(self):
        """Initialize the MCP test client."""
        self.client = None
        self.session = None
        self.connected = False
        self.client_context = None
    
    async def connect_to_server(self, server_command: List[str]) -> bool:
        """Establish MCP connection to server."""
        try:
            logger.info(f"Starting MCP server with command: {' '.join(server_command)}")
            
            # Extract script path and args from server_command
            # server_command is [python_executable, script_path, ...args]
            if len(server_command) < 2:
                raise ValueError("Server command must have at least python executable and script path")
            
            python_cmd = server_command[0]
            script_path = server_command[1]
            args = server_command[2:] if len(server_command) > 2 else None
            
            # Create transport and client
            transport = PythonStdioTransport(
                script_path=script_path,
                args=args,
                python_cmd=python_cmd
            )
            self.client = Client(transport)
            
            # Enter client context
            self.client_context = await self.client.__aenter__()
            
            # Get session (it's already available)
            self.session = self.client.session
            
            self.connected = True
            logger.info("Successfully connected to MCP server")
            return True
            
        except Exception as e:
            logger.error(f"Failed to connect to MCP server: {e}")
            self.connected = False
            return False
    
    async def disconnect(self):
        """Disconnect from MCP server."""
        try:
            if self.client_context:
                await self.client.__aexit__(None, None, None)
                self.client_context = None
                self.session = None
            
            self.connected = False
            logger.info("Disconnected from MCP server")
            
        except Exception as e:
            logger.error(f"Error during disconnect: {e}")
    
    async def list_available_tools(self) -> List[str]:
        """Get list of available tools from server."""
        try:
            if not self.connected or not self.session:
                raise RuntimeError("Not connected to MCP server")
            
            tools = await self.session.list_tools()
            tool_names = [tool.name for tool in tools.tools]
            logger.info(f"Available tools: {tool_names}")
            return tool_names
            
        except Exception as e:
            logger.error(f"Error listing tools: {e}")
            return []
    
    async def call_tool(self, tool_name: str, arguments: Dict) -> Tuple[bool, Dict, float]:
        """Call MCP tool and return success, response, and response time."""
        start_time = time.time()
        
        try:
            if not self.connected or not self.session:
                raise RuntimeError("Not connected to MCP server")
            
            logger.info(f"Calling tool {tool_name} with arguments: {arguments}")
            
            result = await self.session.call_tool(tool_name, arguments)
            response_time = time.time() - start_time
            
            # Extract response data
            response_data = {}
            if result and hasattr(result, 'content'):
                for content_item in result.content:
                    if hasattr(content_item, 'text'):
                        try:
                            response_data = json.loads(content_item.text)
                        except json.JSONDecodeError:
                            response_data = {"raw_text": content_item.text}
            
            logger.info(f"Tool {tool_name} completed in {response_time:.2f}s")
            return True, response_data, response_time
            
        except Exception as e:
            response_time = time.time() - start_time
            logger.error(f"Error calling tool {tool_name}: {e}")
            return False, {"error": str(e)}, response_time


class MCPIntegrationTestSuite:
    """Comprehensive MCP protocol testing."""
    
    def __init__(self, test_files_dir: Optional[str] = None):
        """Initialize the test suite."""
        self.client = MCPTestClient()
        
        # Get the correct paths relative to the project root
        current_dir = Path(__file__).parent
        project_root = current_dir.parent
        
        # Set test_files_dir
        if test_files_dir is None:
            self.test_files_dir = current_dir / "test_files"
        else:
            self.test_files_dir = Path(test_files_dir)
        
        self.test_results = []
        
        # Get the correct path to main.py (should be in the project root)
        main_py_path = project_root / "main.py"
        
        if not main_py_path.exists():
            raise RuntimeError(f"main.py not found at {main_py_path}")
        
        self.server_command = [sys.executable, str(main_py_path)]
    
    async def setup(self) -> bool:
        """Set up the test environment."""
        try:
            # Connect to MCP server
            success = await self.client.connect_to_server(self.server_command)
            if not success:
                logger.error("Failed to connect to MCP server")
                return False
            
            # Verify test files exist
            if not self.test_files_dir.exists():
                logger.error(f"Test files directory not found: {self.test_files_dir}")
                return False
            
            logger.info("Test environment setup completed")
            return True
            
        except Exception as e:
            logger.error(f"Error setting up test environment: {e}")
            return False
    
    async def teardown(self):
        """Clean up test environment."""
        try:
            await self.client.disconnect()
            logger.info("Test environment cleanup completed")
        except Exception as e:
            logger.error(f"Error during teardown: {e}")
    
    async def test_all_tools(self) -> List[TestResult]:
        """Test all MCP tools with various parameter combinations."""
        try:
            # Get available tools
            available_tools = await self.client.list_available_tools()
            
            if not available_tools:
                logger.error("No tools available from MCP server")
                return []
            
            # Test each tool
            for tool_name in available_tools:
                await self.test_tool_with_parameters(tool_name)
            
            return self.test_results
            
        except Exception as e:
            logger.error(f"Error testing all tools: {e}")
            return []
    
    async def test_tool_with_parameters(self, tool_name: str):
        """Test a specific tool with multiple parameter sets."""
        try:
            parameter_sets = self._get_parameter_sets_for_tool(tool_name)
            
            for params in parameter_sets:
                test_description = f"{tool_name} with {params.get('description', 'default parameters')}"
                
                success, response_data, response_time = await self.client.call_tool(
                    tool_name, params['arguments']
                )
                
                result = TestResult(
                    tool_name=tool_name,
                    parameters=params['arguments'],
                    success=success,
                    response_time=response_time,
                    response_data=response_data if success else None,
                    error_message=response_data.get('error') if not success else None,
                    test_description=test_description
                )
                
                self.test_results.append(result)
                logger.info(f"Test completed: {test_description} - {'PASS' if success else 'FAIL'}")
                
        except Exception as e:
            logger.error(f"Error testing tool {tool_name}: {e}")
    
    def _get_parameter_sets_for_tool(self, tool_name: str) -> List[Dict]:
        """Get parameter sets for testing a specific tool."""
        test_file = self.test_files_dir / "test_formatting_comprehensive.pptx"
        
        if tool_name == "extract_powerpoint_content":
            return [
                {
                    "description": "basic content extraction",
                    "arguments": {"file_path": str(test_file)}
                }
            ]
        
        elif tool_name == "get_powerpoint_attributes":
            return [
                {
                    "description": "title and subtitle attributes",
                    "arguments": {
                        "file_path": str(test_file),
                        "attributes": ["title", "subtitle"]
                    }
                },
                {
                    "description": "all text attributes",
                    "arguments": {
                        "file_path": str(test_file),
                        "attributes": ["title", "text_elements", "tables", "object_counts"]
                    }
                }
            ]
        
        elif tool_name == "get_slide_info":
            return [
                {
                    "description": "first slide info",
                    "arguments": {
                        "file_path": str(test_file),
                        "slide_number": 1
                    }
                },
                {
                    "description": "middle slide info",
                    "arguments": {
                        "file_path": str(test_file),
                        "slide_number": 3
                    }
                }
            ]
        
        elif tool_name == "extract_text_formatting":
            return [
                {
                    "description": "bold text extraction",
                    "arguments": {
                        "file_path": str(test_file),
                        "formatting_type": "bold"
                    }
                },
                {
                    "description": "italic text extraction",
                    "arguments": {
                        "file_path": str(test_file),
                        "formatting_type": "italic"
                    }
                },
                {
                    "description": "underlined text extraction",
                    "arguments": {
                        "file_path": str(test_file),
                        "formatting_type": "underlined"
                    }
                },
                {
                    "description": "font_sizes extraction",
                    "arguments": {
                        "file_path": str(test_file),
                        "formatting_type": "font_sizes"
                    }
                },
                {
                    "description": "font_colors extraction",
                    "arguments": {
                        "file_path": str(test_file),
                        "formatting_type": "font_colors"
                    }
                },
                {
                    "description": "hyperlinks extraction",
                    "arguments": {
                        "file_path": str(test_file),
                        "formatting_type": "hyperlinks"
                    }
                },
                {
                    "description": "bold text from specific slides",
                    "arguments": {
                        "file_path": str(test_file),
                        "formatting_type": "bold",
                        "slide_numbers": [1, 4]
                    }
                }
            ]
        
        elif tool_name == "extract_bold_text":
            return [
                {
                    "description": "all slides bold text",
                    "arguments": {"file_path": str(test_file)}
                },
                {
                    "description": "specific slides bold text",
                    "arguments": {
                        "file_path": str(test_file),
                        "slide_numbers": [1, 4]
                    }
                }
            ]
        
        elif tool_name == "query_slides":
            return [
                {
                    "description": "query by slide numbers",
                    "arguments": {
                        "file_path": str(test_file),
                        "search_criteria": {"slide_numbers": [1, 2, 3]}
                    }
                }
            ]
        
        elif tool_name == "extract_table_data":
            return [
                {
                    "description": "extract tables from all slides",
                    "arguments": {
                        "file_path": str(test_file),
                        "slide_numbers": [1, 2, 3, 4, 5, 6, 7, 8]
                    }
                }
            ]
        
        elif tool_name == "get_presentation_overview":
            return [
                {
                    "description": "basic overview",
                    "arguments": {
                        "file_path": str(test_file),
                        "analysis_depth": "basic"
                    }
                },
                {
                    "description": "detailed overview",
                    "arguments": {
                        "file_path": str(test_file),
                        "analysis_depth": "detailed"
                    }
                }
            ]
        
        elif tool_name == "analyze_text_formatting":
            return [
                {
                    "description": "analyze all slides formatting",
                    "arguments": {"file_path": str(test_file)}
                }
            ]
        
        else:
            # Default parameter set for unknown tools
            return [
                {
                    "description": "default parameters",
                    "arguments": {"file_path": str(test_file)}
                }
            ]
    
    async def test_error_conditions(self):
        """Test error handling with invalid parameters."""
        error_test_cases = [
            {
                "tool_name": "extract_text_formatting",
                "arguments": {
                    "file_path": str(self.test_files_dir / "test_formatting_comprehensive.pptx"),
                    "formatting_type": "invalid_type"
                },
                "description": "invalid formatting type"
            },
            {
                "tool_name": "get_slide_info",
                "arguments": {
                    "file_path": str(self.test_files_dir / "test_formatting_comprehensive.pptx"),
                    "slide_number": 999
                },
                "description": "invalid slide number"
            },
            {
                "tool_name": "extract_powerpoint_content",
                "arguments": {
                    "file_path": "nonexistent_file.pptx"
                },
                "description": "nonexistent file"
            }
        ]
        
        for test_case in error_test_cases:
            success, response_data, response_time = await self.client.call_tool(
                test_case["tool_name"], test_case["arguments"]
            )
            
            # For error conditions, we expect success=False or error in response
            expected_error = not success or "error" in response_data
            
            result = TestResult(
                tool_name=test_case["tool_name"],
                parameters=test_case["arguments"],
                success=expected_error,  # Success means we got expected error
                response_time=response_time,
                response_data=response_data,
                error_message=None if expected_error else "Expected error but got success",
                test_description=f"Error test: {test_case['description']}"
            )
            
            self.test_results.append(result)
    
    def generate_test_report(self) -> Dict:
        """Generate comprehensive test coverage report."""
        total_tests = len(self.test_results)
        passed_tests = sum(1 for result in self.test_results if result.success)
        failed_tests = total_tests - passed_tests
        
        # Group results by tool
        tools_tested = {}
        for result in self.test_results:
            if result.tool_name not in tools_tested:
                tools_tested[result.tool_name] = {"total": 0, "passed": 0, "failed": 0}
            
            tools_tested[result.tool_name]["total"] += 1
            if result.success:
                tools_tested[result.tool_name]["passed"] += 1
            else:
                tools_tested[result.tool_name]["failed"] += 1
        
        # Calculate average response times
        avg_response_time = sum(result.response_time for result in self.test_results) / total_tests if total_tests > 0 else 0
        
        report = {
            "summary": {
                "total_tests": total_tests,
                "passed_tests": passed_tests,
                "failed_tests": failed_tests,
                "success_rate": (passed_tests / total_tests * 100) if total_tests > 0 else 0,
                "average_response_time": avg_response_time
            },
            "tools_coverage": tools_tested,
            "failed_tests": [
                {
                    "tool": result.tool_name,
                    "description": result.test_description,
                    "error": result.error_message,
                    "parameters": result.parameters
                }
                for result in self.test_results if not result.success
            ],
            "performance": {
                "fastest_test": min(self.test_results, key=lambda x: x.response_time, default=None),
                "slowest_test": max(self.test_results, key=lambda x: x.response_time, default=None)
            }
        }
        
        return report


async def run_comprehensive_tests():
    """Run comprehensive MCP integration tests."""
    test_suite = MCPIntegrationTestSuite()
    
    try:
        # Setup
        logger.info("Setting up MCP integration test suite...")
        if not await test_suite.setup():
            logger.error("Failed to setup test environment")
            return
        
        # Run all tool tests
        logger.info("Running comprehensive tool tests...")
        await test_suite.test_all_tools()
        
        # Run error condition tests
        logger.info("Running error condition tests...")
        await test_suite.test_error_conditions()
        
        # Generate report
        report = test_suite.generate_test_report()
        
        # Save report
        report_file = Path("tests/mcp_integration_test_report.json")
        with open(report_file, 'w', encoding='utf-8') as f:
            json.dump(report, f, indent=2, ensure_ascii=False, default=str)
        
        # Print summary
        print("\n" + "="*60)
        print("MCP INTEGRATION TEST RESULTS")
        print("="*60)
        print(f"Total Tests: {report['summary']['total_tests']}")
        print(f"Passed: {report['summary']['passed_tests']}")
        print(f"Failed: {report['summary']['failed_tests']}")
        print(f"Success Rate: {report['summary']['success_rate']:.1f}%")
        print(f"Average Response Time: {report['summary']['average_response_time']:.2f}s")
        
        print(f"\nDetailed report saved to: {report_file}")
        
        if report['summary']['failed_tests'] > 0:
            print("\nFailed Tests:")
            for failed_test in report['failed_tests']:
                print(f"  - {failed_test['tool']}: {failed_test['description']}")
                print(f"    Error: {failed_test['error']}")
        
    except Exception as e:
        logger.error(f"Error running tests: {e}")
        raise
    
    finally:
        # Cleanup
        await test_suite.teardown()


if __name__ == "__main__":
    # Configure logging
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )
    
    # Run tests
    asyncio.run(run_comprehensive_tests())