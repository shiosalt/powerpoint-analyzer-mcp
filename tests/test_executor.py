"""
Test execution and reporting system for PowerPoint MCP server.
Orchestrates test runs and generates comprehensive reports.
"""

import asyncio
import json
import logging
import time
import psutil
import os
from pathlib import Path
from typing import Dict, List, Any, Optional
from dataclasses import dataclass, asdict
from datetime import datetime

from tests.mcp_integration_test_framework import MCPIntegrationTestSuite, TestResult

logger = logging.getLogger(__name__)


@dataclass
class TestExecutionReport:
    """Comprehensive test execution report."""
    execution_id: str
    start_time: datetime
    end_time: datetime
    duration_seconds: float
    total_tests: int
    passed_tests: int
    failed_tests: int
    success_rate: float
    average_response_time: float
    memory_usage_mb: float
    cpu_usage_percent: float
    test_results: List[Dict]
    performance_metrics: Dict[str, Any]
    coverage_analysis: Dict[str, Any]
    error_summary: Dict[str, Any]


class TestExecutor:
    """Orchestrates comprehensive test execution."""
    
    def __init__(self, test_files_dir: Optional[str] = None):
        """Initialize the test executor."""
        # Get the correct paths relative to the project root
        current_dir = Path(__file__).parent
        
        if test_files_dir is None:
            self.test_files_dir = current_dir / "test_files"
        else:
            self.test_files_dir = Path(test_files_dir)
            
        self.reports_dir = current_dir / "reports"
        self.reports_dir.mkdir(exist_ok=True)
        
        # Performance monitoring
        self.process = psutil.Process(os.getpid())
        self.initial_memory = self.process.memory_info().rss / 1024 / 1024  # MB
        
        # Test configuration
        self.test_timeout = 30.0  # seconds per test
        self.max_retries = 2
    
    async def run_full_test_suite(self) -> TestExecutionReport:
        """Execute all tests and generate comprehensive report."""
        execution_id = f"test_run_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        start_time = datetime.now()
        
        logger.info(f"Starting full test suite execution: {execution_id}")
        
        try:
            # Initialize test suite
            test_suite = MCPIntegrationTestSuite(str(self.test_files_dir))
            
            # Setup test environment
            if not await test_suite.setup():
                raise RuntimeError("Failed to setup test environment")
            
            # Run all tests
            test_results = []
            
            # Run tool tests
            logger.info("Running MCP tool tests...")
            tool_results = await test_suite.test_all_tools()
            test_results.extend(tool_results)
            
            # Run error condition tests
            logger.info("Running error condition tests...")
            await test_suite.test_error_conditions()
            test_results.extend(test_suite.test_results)
            
            # Run performance tests
            logger.info("Running performance tests...")
            performance_results = await self._run_performance_tests(test_suite)
            test_results.extend(performance_results)
            
            # Generate comprehensive report
            end_time = datetime.now()
            duration = (end_time - start_time).total_seconds()
            
            report = self._generate_execution_report(
                execution_id, start_time, end_time, duration, test_results
            )
            
            # Save report
            await self._save_report(report)
            
            # Cleanup
            await test_suite.teardown()
            
            logger.info(f"Test suite execution completed: {execution_id}")
            return report
            
        except Exception as e:
            logger.error(f"Error during test execution: {e}")
            raise
    
    async def _run_performance_tests(self, test_suite: MCPIntegrationTestSuite) -> List[TestResult]:
        """Run performance-focused tests."""
        performance_results = []
        
        # Test response time requirements
        test_cases = [
            {
                "tool": "extract_powerpoint_content",
                "args": {"file_path": str(self.test_files_dir / "test_formatting_comprehensive.pptx")},
                "max_time": 10.0,
                "description": "Content extraction performance"
            },
            {
                "tool": "extract_text_formatting",
                "args": {
                    "file_path": str(self.test_files_dir / "test_formatting_comprehensive.pptx"),
                    "formatting_type": "bold"
                },
                "max_time": 5.0,
                "description": "Formatting extraction performance"
            },
            {
                "tool": "get_powerpoint_attributes",
                "args": {
                    "file_path": str(self.test_files_dir / "test_formatting_comprehensive.pptx"),
                    "attributes": ["title", "text_elements"]
                },
                "max_time": 8.0,
                "description": "Attribute extraction performance"
            }
        ]
        
        for test_case in test_cases:
            success, response_data, response_time = await test_suite.client.call_tool(
                test_case["tool"], test_case["args"]
            )
            
            # Check if performance requirement is met
            performance_ok = response_time <= test_case["max_time"]
            
            result = TestResult(
                tool_name=test_case["tool"],
                parameters=test_case["args"],
                success=success and performance_ok,
                response_time=response_time,
                response_data=response_data if success else None,
                error_message=f"Performance requirement not met: {response_time:.2f}s > {test_case['max_time']}s" if not performance_ok else None,
                test_description=f"Performance test: {test_case['description']}"
            )
            
            performance_results.append(result)
        
        return performance_results
    
    def _generate_execution_report(
        self,
        execution_id: str,
        start_time: datetime,
        end_time: datetime,
        duration: float,
        test_results: List[TestResult]
    ) -> TestExecutionReport:
        """Generate comprehensive execution report."""
        
        # Basic statistics
        total_tests = len(test_results)
        passed_tests = sum(1 for result in test_results if result.success)
        failed_tests = total_tests - passed_tests
        success_rate = (passed_tests / total_tests * 100) if total_tests > 0 else 0
        
        # Performance metrics
        response_times = [result.response_time for result in test_results]
        avg_response_time = sum(response_times) / len(response_times) if response_times else 0
        
        # Memory usage
        current_memory = self.process.memory_info().rss / 1024 / 1024  # MB
        memory_usage = current_memory - self.initial_memory
        
        # CPU usage (approximate)
        cpu_usage = self.process.cpu_percent()
        
        # Performance metrics
        performance_metrics = {
            "fastest_test": min(response_times) if response_times else 0,
            "slowest_test": max(response_times) if response_times else 0,
            "median_response_time": sorted(response_times)[len(response_times)//2] if response_times else 0,
            "tests_under_1s": sum(1 for t in response_times if t < 1.0),
            "tests_over_5s": sum(1 for t in response_times if t > 5.0),
            "memory_growth_mb": memory_usage,
            "peak_memory_mb": current_memory
        }
        
        # Coverage analysis
        tools_tested = set(result.tool_name for result in test_results)
        coverage_analysis = {
            "tools_tested": list(tools_tested),
            "total_tools_tested": len(tools_tested),
            "test_distribution": {
                tool: sum(1 for r in test_results if r.tool_name == tool)
                for tool in tools_tested
            }
        }
        
        # Error analysis
        failed_results = [result for result in test_results if not result.success]
        error_summary = {
            "total_errors": len(failed_results),
            "error_types": {},
            "most_common_errors": []
        }
        
        # Categorize errors
        for result in failed_results:
            error_msg = result.error_message or "Unknown error"
            if error_msg not in error_summary["error_types"]:
                error_summary["error_types"][error_msg] = 0
            error_summary["error_types"][error_msg] += 1
        
        # Most common errors
        error_summary["most_common_errors"] = sorted(
            error_summary["error_types"].items(),
            key=lambda x: x[1],
            reverse=True
        )[:5]
        
        return TestExecutionReport(
            execution_id=execution_id,
            start_time=start_time,
            end_time=end_time,
            duration_seconds=duration,
            total_tests=total_tests,
            passed_tests=passed_tests,
            failed_tests=failed_tests,
            success_rate=success_rate,
            average_response_time=avg_response_time,
            memory_usage_mb=memory_usage,
            cpu_usage_percent=cpu_usage,
            test_results=[asdict(result) for result in test_results],
            performance_metrics=performance_metrics,
            coverage_analysis=coverage_analysis,
            error_summary=error_summary
        )
    
    async def _save_report(self, report: TestExecutionReport):
        """Save test execution report to file."""
        report_file = self.reports_dir / f"{report.execution_id}_report.json"
        
        # Convert report to dict for JSON serialization
        report_dict = asdict(report)
        
        # Convert datetime objects to strings
        report_dict['start_time'] = report.start_time.isoformat()
        report_dict['end_time'] = report.end_time.isoformat()
        
        with open(report_file, 'w', encoding='utf-8') as f:
            json.dump(report_dict, f, indent=2, ensure_ascii=False, default=str)
        
        logger.info(f"Test report saved: {report_file}")
        
        # Also save a summary report
        summary_file = self.reports_dir / f"{report.execution_id}_summary.txt"
        await self._save_summary_report(report, summary_file)
    
    async def _save_summary_report(self, report: TestExecutionReport, summary_file: Path):
        """Save human-readable summary report."""
        with open(summary_file, 'w', encoding='utf-8') as f:
            f.write(f"PowerPoint MCP Server Test Execution Report\n")
            f.write(f"=" * 50 + "\n\n")
            
            f.write(f"Execution ID: {report.execution_id}\n")
            f.write(f"Start Time: {report.start_time}\n")
            f.write(f"End Time: {report.end_time}\n")
            f.write(f"Duration: {report.duration_seconds:.2f} seconds\n\n")
            
            f.write(f"Test Results Summary:\n")
            f.write(f"  Total Tests: {report.total_tests}\n")
            f.write(f"  Passed: {report.passed_tests}\n")
            f.write(f"  Failed: {report.failed_tests}\n")
            f.write(f"  Success Rate: {report.success_rate:.1f}%\n\n")
            
            f.write(f"Performance Metrics:\n")
            f.write(f"  Average Response Time: {report.average_response_time:.2f}s\n")
            f.write(f"  Fastest Test: {report.performance_metrics['fastest_test']:.2f}s\n")
            f.write(f"  Slowest Test: {report.performance_metrics['slowest_test']:.2f}s\n")
            f.write(f"  Tests Under 1s: {report.performance_metrics['tests_under_1s']}\n")
            f.write(f"  Tests Over 5s: {report.performance_metrics['tests_over_5s']}\n\n")
            
            f.write(f"Resource Usage:\n")
            f.write(f"  Memory Usage: {report.memory_usage_mb:.1f} MB\n")
            f.write(f"  CPU Usage: {report.cpu_usage_percent:.1f}%\n\n")
            
            f.write(f"Coverage Analysis:\n")
            f.write(f"  Tools Tested: {report.coverage_analysis['total_tools_tested']}\n")
            for tool, count in report.coverage_analysis['test_distribution'].items():
                f.write(f"    {tool}: {count} tests\n")
            
            if report.error_summary['total_errors'] > 0:
                f.write(f"\nError Summary:\n")
                f.write(f"  Total Errors: {report.error_summary['total_errors']}\n")
                f.write(f"  Most Common Errors:\n")
                for error, count in report.error_summary['most_common_errors']:
                    f.write(f"    {error}: {count} occurrences\n")
        
        logger.info(f"Summary report saved: {summary_file}")
    
    def generate_coverage_report(self, test_results: List[TestResult]) -> Dict[str, Any]:
        """Generate detailed test coverage report."""
        coverage = {
            "tools_coverage": {},
            "parameter_coverage": {},
            "error_coverage": {},
            "overall_coverage": 0.0
        }
        
        # Expected tools (based on our implementation)
        expected_tools = [
            "extract_powerpoint_content",
            "get_powerpoint_attributes",
            "get_slide_info",
            "extract_text_formatting",
            "extract_bold_text",
            "query_slides",
            "extract_table_data",
            "get_presentation_overview",
            "analyze_text_formatting"
        ]
        
        # Analyze tool coverage
        for tool in expected_tools:
            tool_tests = [r for r in test_results if r.tool_name == tool]
            coverage["tools_coverage"][tool] = {
                "total_tests": len(tool_tests),
                "passed_tests": sum(1 for t in tool_tests if t.success),
                "failed_tests": sum(1 for t in tool_tests if not t.success),
                "coverage_percentage": (len(tool_tests) / len(test_results) * 100) if test_results else 0
            }
        
        # Calculate overall coverage
        tools_tested = len([tool for tool in expected_tools if coverage["tools_coverage"][tool]["total_tests"] > 0])
        coverage["overall_coverage"] = (tools_tested / len(expected_tools) * 100) if expected_tools else 0
        
        return coverage


async def main():
    """Main entry point for test execution."""
    # Configure logging
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )
    
    executor = TestExecutor()
    
    try:
        print("Starting comprehensive test execution...")
        report = await executor.run_full_test_suite()
        
        print(f"\nTest Execution Completed!")
        print(f"Execution ID: {report.execution_id}")
        print(f"Total Tests: {report.total_tests}")
        print(f"Success Rate: {report.success_rate:.1f}%")
        print(f"Duration: {report.duration_seconds:.2f} seconds")
        print(f"Average Response Time: {report.average_response_time:.2f}s")
        
        if report.failed_tests > 0:
            print(f"\nFailed Tests: {report.failed_tests}")
            print("Check the detailed report for error analysis.")
        
        print(f"\nReports saved in: tests/reports/")
        
    except Exception as e:
        print(f"Test execution failed: {e}")
        logger.error(f"Test execution error: {e}")
        raise


if __name__ == "__main__":
    asyncio.run(main())