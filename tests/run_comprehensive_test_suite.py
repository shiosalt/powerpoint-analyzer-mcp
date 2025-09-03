#!/usr/bin/env python3
"""
Comprehensive test suite runner for PowerPoint MCP server.
Executes all tests, validates against known results, and generates final reports.
"""

import asyncio
import json
import logging
import sys
import time
from pathlib import Path
from typing import Dict, List, Any, Optional
from datetime import datetime

# Add current directory to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent))

# Import test components
from tests.test_executor import TestExecutor, TestExecutionReport
from tests.mcp_integration_test_framework import MCPIntegrationTestSuite
from tests.test_data_generator import TestPresentationGenerator

logger = logging.getLogger(__name__)


class ComprehensiveTestSuiteRunner:
    """Orchestrates the complete test suite execution."""
    
    def __init__(self):
        """Initialize the comprehensive test suite runner."""
        # Get the correct paths relative to the project root
        current_dir = Path(__file__).parent
        
        self.test_files_dir = current_dir / "test_files"
        self.reports_dir = current_dir / "reports"
        self.reports_dir.mkdir(exist_ok=True)
        
        # Test components
        self.test_executor = TestExecutor()
        self.test_generator = TestPresentationGenerator()
        
        # Results storage
        self.validation_results = {}
        self.final_report = {}
    
    async def run_complete_test_suite(self) -> Dict[str, Any]:
        """Run the complete test suite with validation and reporting."""
        suite_start_time = datetime.now()
        logger.info("Starting comprehensive test suite execution...")
        
        try:
            # Step 1: Ensure test data is available
            logger.info("Step 1: Preparing test data...")
            await self._prepare_test_data()
            
            # Step 2: Run comprehensive MCP integration tests
            logger.info("Step 2: Running MCP integration tests...")
            integration_report = await self._run_integration_tests()
            
            # Step 3: Run full test executor suite
            logger.info("Step 3: Running full test executor suite...")
            execution_report = await self.test_executor.run_full_test_suite()
            
            # Step 4: Validate results against known test data
            logger.info("Step 4: Validating results against expected outcomes...")
            validation_report = await self._validate_test_results()
            
            # Step 5: Generate comprehensive final report
            logger.info("Step 5: Generating comprehensive final report...")
            final_report = await self._generate_final_report(
                suite_start_time, integration_report, execution_report, validation_report
            )
            
            # Step 6: Save all reports
            logger.info("Step 6: Saving comprehensive reports...")
            await self._save_comprehensive_reports(final_report)
            
            logger.info("Comprehensive test suite execution completed successfully!")
            return final_report
            
        except Exception as e:
            logger.error(f"Error during comprehensive test suite execution: {e}")
            raise
    
    async def _prepare_test_data(self):
        """Ensure all required test data files are available."""
        try:
            # Check if test files exist
            test_file = self.test_files_dir / "test_formatting_comprehensive.pptx"
            expected_results_file = self.test_files_dir / "test_formatting_comprehensive.json"
            
            if not test_file.exists():
                logger.info("Generating test presentation file...")
                await self.test_generator.generate_comprehensive_test_file()
            
            if not expected_results_file.exists():
                logger.info("Generating expected results file...")
                await self.test_generator.generate_expected_results()
            
            # Verify test files are valid
            if not test_file.exists() or test_file.stat().st_size == 0:
                raise RuntimeError("Test presentation file is missing or empty")
            
            logger.info("Test data preparation completed")
            
        except Exception as e:
            logger.error(f"Error preparing test data: {e}")
            raise
    
    async def _run_integration_tests(self) -> Dict[str, Any]:
        """Run MCP integration tests."""
        try:
            test_suite = MCPIntegrationTestSuite(str(self.test_files_dir))
            
            # Setup test environment
            if not await test_suite.setup():
                raise RuntimeError("Failed to setup MCP integration test environment")
            
            try:
                # Run all tool tests
                await test_suite.test_all_tools()
                
                # Run error condition tests
                await test_suite.test_error_conditions()
                
                # Generate integration report
                integration_report = test_suite.generate_test_report()
                
                # Save integration report
                integration_report_file = self.reports_dir / "mcp_integration_report.json"
                with open(integration_report_file, 'w', encoding='utf-8') as f:
                    json.dump(integration_report, f, indent=2, ensure_ascii=False, default=str)
                
                logger.info(f"Integration test report saved: {integration_report_file}")
                return integration_report
                
            finally:
                await test_suite.teardown()
                
        except Exception as e:
            logger.error(f"Error running integration tests: {e}")
            raise
    
    async def _validate_test_results(self) -> Dict[str, Any]:
        """Validate test results against known expected outcomes."""
        try:
            validation_report = {
                "validation_timestamp": datetime.now().isoformat(),
                "validations_performed": [],
                "validation_summary": {
                    "total_validations": 0,
                    "passed_validations": 0,
                    "failed_validations": 0,
                    "validation_success_rate": 0.0
                },
                "validation_details": []
            }
            
            # Load expected results
            expected_results_file = self.test_files_dir / "test_formatting_comprehensive.json"
            if not expected_results_file.exists():
                logger.warning("Expected results file not found, skipping validation")
                return validation_report
            
            with open(expected_results_file, 'r', encoding='utf-8') as f:
                expected_results = json.load(f)
            
            # Run validation tests
            test_suite = MCPIntegrationTestSuite(str(self.test_files_dir))
            
            if not await test_suite.setup():
                raise RuntimeError("Failed to setup validation test environment")
            
            try:
                # Validate bold text extraction
                await self._validate_bold_text_extraction(test_suite, expected_results, validation_report)
                
                # Validate formatting extraction for each type
                formatting_types = ["bold", "italic", "underlined", "font_sizes", "font_colors"]
                for formatting_type in formatting_types:
                    await self._validate_formatting_extraction(
                        test_suite, expected_results, formatting_type, validation_report
                    )
                
                # Validate content extraction
                await self._validate_content_extraction(test_suite, expected_results, validation_report)
                
                # Calculate validation summary
                total_validations = len(validation_report["validation_details"])
                passed_validations = sum(1 for v in validation_report["validation_details"] if v["passed"])
                failed_validations = total_validations - passed_validations
                
                validation_report["validation_summary"] = {
                    "total_validations": total_validations,
                    "passed_validations": passed_validations,
                    "failed_validations": failed_validations,
                    "validation_success_rate": (passed_validations / total_validations * 100) if total_validations > 0 else 0
                }
                
                logger.info(f"Validation completed: {passed_validations}/{total_validations} passed")
                return validation_report
                
            finally:
                await test_suite.teardown()
                
        except Exception as e:
            logger.error(f"Error during validation: {e}")
            raise
    
    async def _validate_bold_text_extraction(self, test_suite, expected_results, validation_report):
        """Validate bold text extraction against expected results."""
        try:
            success, response_data, response_time = await test_suite.client.call_tool(
                "extract_text_formatting",
                {
                    "file_path": str(self.test_files_dir / "test_formatting_comprehensive.pptx"),
                    "formatting_type": "bold"
                }
            )
            
            validation_detail = {
                "test_name": "bold_text_extraction_validation",
                "passed": False,
                "details": {},
                "errors": []
            }
            
            if not success:
                validation_detail["errors"].append("Tool call failed")
            else:
                # Validate response structure
                if "results_by_slide" in response_data:
                    slide_1_results = None
                    for slide_result in response_data["results_by_slide"]:
                        if slide_result["slide_number"] == 1:
                            slide_1_results = slide_result
                            break
                    
                    if slide_1_results and "slide_1_bold" in expected_results:
                        expected_bold = expected_results["slide_1_bold"]
                        actual_segments = slide_1_results["formatted_segments"]
                        expected_segments = expected_bold.get("bold_segments", [])
                        
                        # Compare segment counts
                        if len(actual_segments) == len(expected_segments):
                            validation_detail["passed"] = True
                            validation_detail["details"]["segments_match"] = True
                        else:
                            validation_detail["errors"].append(
                                f"Segment count mismatch: expected {len(expected_segments)}, got {len(actual_segments)}"
                            )
                    else:
                        validation_detail["errors"].append("Missing slide 1 results or expected data")
                else:
                    validation_detail["errors"].append("Missing results_by_slide in response")
            
            validation_report["validation_details"].append(validation_detail)
            
        except Exception as e:
            validation_report["validation_details"].append({
                "test_name": "bold_text_extraction_validation",
                "passed": False,
                "details": {},
                "errors": [f"Validation error: {str(e)}"]
            })
    
    async def _validate_formatting_extraction(self, test_suite, expected_results, formatting_type, validation_report):
        """Validate formatting extraction for a specific type."""
        try:
            success, response_data, response_time = await test_suite.client.call_tool(
                "extract_text_formatting",
                {
                    "file_path": str(self.test_files_dir / "test_formatting_comprehensive.pptx"),
                    "formatting_type": formatting_type
                }
            )
            
            validation_detail = {
                "test_name": f"{formatting_type}_formatting_extraction_validation",
                "passed": False,
                "details": {},
                "errors": []
            }
            
            if not success:
                validation_detail["errors"].append("Tool call failed")
            else:
                # Basic structure validation
                if "formatting_type" in response_data and response_data["formatting_type"] == formatting_type:
                    if "summary" in response_data and "results_by_slide" in response_data:
                        validation_detail["passed"] = True
                        validation_detail["details"]["structure_valid"] = True
                        validation_detail["details"]["slides_analyzed"] = response_data["summary"]["total_slides_analyzed"]
                    else:
                        validation_detail["errors"].append("Missing summary or results_by_slide")
                else:
                    validation_detail["errors"].append("Invalid formatting_type in response")
            
            validation_report["validation_details"].append(validation_detail)
            
        except Exception as e:
            validation_report["validation_details"].append({
                "test_name": f"{formatting_type}_formatting_extraction_validation",
                "passed": False,
                "details": {},
                "errors": [f"Validation error: {str(e)}"]
            })
    
    async def _validate_content_extraction(self, test_suite, expected_results, validation_report):
        """Validate basic content extraction."""
        try:
            success, response_data, response_time = await test_suite.client.call_tool(
                "extract_powerpoint_content",
                {"file_path": str(self.test_files_dir / "test_formatting_comprehensive.pptx")}
            )
            
            validation_detail = {
                "test_name": "content_extraction_validation",
                "passed": False,
                "details": {},
                "errors": []
            }
            
            if not success:
                validation_detail["errors"].append("Tool call failed")
            else:
                # Validate basic structure
                required_fields = ["slides", "metadata"]
                missing_fields = [field for field in required_fields if field not in response_data]
                
                if not missing_fields:
                    if len(response_data["slides"]) >= 8:  # Expected number of slides
                        validation_detail["passed"] = True
                        validation_detail["details"]["slide_count"] = len(response_data["slides"])
                        validation_detail["details"]["has_metadata"] = True
                    else:
                        validation_detail["errors"].append(f"Expected at least 8 slides, got {len(response_data['slides'])}")
                else:
                    validation_detail["errors"].append(f"Missing required fields: {missing_fields}")
            
            validation_report["validation_details"].append(validation_detail)
            
        except Exception as e:
            validation_report["validation_details"].append({
                "test_name": "content_extraction_validation",
                "passed": False,
                "details": {},
                "errors": [f"Validation error: {str(e)}"]
            })
    
    async def _generate_final_report(self, suite_start_time, integration_report, execution_report, validation_report):
        """Generate comprehensive final report."""
        suite_end_time = datetime.now()
        suite_duration = (suite_end_time - suite_start_time).total_seconds()
        
        final_report = {
            "comprehensive_test_suite_report": {
                "execution_id": f"comprehensive_suite_{suite_start_time.strftime('%Y%m%d_%H%M%S')}",
                "start_time": suite_start_time.isoformat(),
                "end_time": suite_end_time.isoformat(),
                "total_duration_seconds": suite_duration,
                "test_environment": {
                    "python_version": sys.version,
                    "test_files_directory": str(self.test_files_dir),
                    "reports_directory": str(self.reports_dir)
                }
            },
            "integration_test_results": integration_report,
            "execution_test_results": {
                "execution_id": execution_report.execution_id,
                "duration_seconds": execution_report.duration_seconds,
                "total_tests": execution_report.total_tests,
                "passed_tests": execution_report.passed_tests,
                "failed_tests": execution_report.failed_tests,
                "success_rate": execution_report.success_rate,
                "average_response_time": execution_report.average_response_time,
                "performance_metrics": execution_report.performance_metrics,
                "coverage_analysis": execution_report.coverage_analysis
            },
            "validation_results": validation_report,
            "overall_summary": {
                "total_test_categories": 3,
                "integration_success_rate": integration_report.get("summary", {}).get("success_rate", 0),
                "execution_success_rate": execution_report.success_rate,
                "validation_success_rate": validation_report["validation_summary"]["validation_success_rate"],
                "overall_success_rate": 0,
                "recommendations": []
            }
        }
        
        # Calculate overall success rate
        success_rates = [
            final_report["overall_summary"]["integration_success_rate"],
            final_report["overall_summary"]["execution_success_rate"],
            final_report["overall_summary"]["validation_success_rate"]
        ]
        final_report["overall_summary"]["overall_success_rate"] = sum(success_rates) / len(success_rates)
        
        # Generate recommendations
        recommendations = []
        if final_report["overall_summary"]["integration_success_rate"] < 90:
            recommendations.append("Improve MCP integration test coverage and error handling")
        if final_report["overall_summary"]["execution_success_rate"] < 90:
            recommendations.append("Address failing test cases in execution suite")
        if final_report["overall_summary"]["validation_success_rate"] < 90:
            recommendations.append("Review and update expected test results validation")
        if execution_report.average_response_time > 3.0:
            recommendations.append("Optimize tool performance to reduce response times")
        
        final_report["overall_summary"]["recommendations"] = recommendations
        
        return final_report
    
    async def _save_comprehensive_reports(self, final_report):
        """Save all comprehensive reports."""
        try:
            # Save final comprehensive report
            final_report_file = self.reports_dir / f"{final_report['comprehensive_test_suite_report']['execution_id']}_final_report.json"
            with open(final_report_file, 'w', encoding='utf-8') as f:
                json.dump(final_report, f, indent=2, ensure_ascii=False, default=str)
            
            # Save human-readable summary
            summary_file = self.reports_dir / f"{final_report['comprehensive_test_suite_report']['execution_id']}_summary.txt"
            await self._save_human_readable_summary(final_report, summary_file)
            
            logger.info(f"Comprehensive reports saved:")
            logger.info(f"  Final report: {final_report_file}")
            logger.info(f"  Summary: {summary_file}")
            
        except Exception as e:
            logger.error(f"Error saving comprehensive reports: {e}")
            raise
    
    async def _save_human_readable_summary(self, final_report, summary_file):
        """Save human-readable summary report."""
        with open(summary_file, 'w', encoding='utf-8') as f:
            f.write("PowerPoint MCP Server - Comprehensive Test Suite Report\n")
            f.write("=" * 60 + "\n\n")
            
            # Overall summary
            overall = final_report["overall_summary"]
            f.write(f"Overall Success Rate: {overall['overall_success_rate']:.1f}%\n")
            f.write(f"Total Duration: {final_report['comprehensive_test_suite_report']['total_duration_seconds']:.2f} seconds\n\n")
            
            # Integration tests
            integration = final_report["integration_test_results"]["summary"]
            f.write(f"Integration Tests:\n")
            f.write(f"  Total Tests: {integration['total_tests']}\n")
            f.write(f"  Success Rate: {integration['success_rate']:.1f}%\n")
            f.write(f"  Average Response Time: {integration['average_response_time']:.2f}s\n\n")
            
            # Execution tests
            execution = final_report["execution_test_results"]
            f.write(f"Execution Tests:\n")
            f.write(f"  Total Tests: {execution['total_tests']}\n")
            f.write(f"  Success Rate: {execution['success_rate']:.1f}%\n")
            f.write(f"  Average Response Time: {execution['average_response_time']:.2f}s\n\n")
            
            # Validation tests
            validation = final_report["validation_results"]["validation_summary"]
            f.write(f"Validation Tests:\n")
            f.write(f"  Total Validations: {validation['total_validations']}\n")
            f.write(f"  Success Rate: {validation['validation_success_rate']:.1f}%\n\n")
            
            # Recommendations
            if overall["recommendations"]:
                f.write("Recommendations:\n")
                for i, rec in enumerate(overall["recommendations"], 1):
                    f.write(f"  {i}. {rec}\n")
            else:
                f.write("No recommendations - all tests passed successfully!\n")


async def main():
    """Main entry point for comprehensive test suite."""
    # Configure logging
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler('tests/comprehensive_test_suite.log', mode='w')
        ]
    )
    
    runner = ComprehensiveTestSuiteRunner()
    
    try:
        print("Starting PowerPoint MCP Server Comprehensive Test Suite...")
        print("=" * 60)
        
        final_report = await runner.run_complete_test_suite()
        
        print("\nComprehensive Test Suite Completed!")
        print("=" * 60)
        print(f"Overall Success Rate: {final_report['overall_summary']['overall_success_rate']:.1f}%")
        print(f"Total Duration: {final_report['comprehensive_test_suite_report']['total_duration_seconds']:.2f} seconds")
        
        # Print category results
        print(f"\nCategory Results:")
        print(f"  Integration Tests: {final_report['overall_summary']['integration_success_rate']:.1f}%")
        print(f"  Execution Tests: {final_report['overall_summary']['execution_success_rate']:.1f}%")
        print(f"  Validation Tests: {final_report['overall_summary']['validation_success_rate']:.1f}%")
        
        # Print recommendations
        if final_report['overall_summary']['recommendations']:
            print(f"\nRecommendations:")
            for i, rec in enumerate(final_report['overall_summary']['recommendations'], 1):
                print(f"  {i}. {rec}")
        
        print(f"\nDetailed reports saved in: tests/reports/")
        
        # Return appropriate exit code
        if final_report['overall_summary']['overall_success_rate'] >= 80:
            print("\n✅ Test suite PASSED")
            return 0
        else:
            print("\n❌ Test suite FAILED")
            return 1
            
    except Exception as e:
        print(f"\n❌ Test suite execution failed: {e}")
        logger.error(f"Comprehensive test suite error: {e}")
        return 1


if __name__ == "__main__":
    exit_code = asyncio.run(main())
    sys.exit(exit_code)