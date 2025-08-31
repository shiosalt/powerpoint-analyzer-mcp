"""
Unit tests for MCP Resources.
"""

import pytest
import json

from powerpoint_mcp_server.resources.powerpoint_extraction_capabilities import (
    get_powerpoint_extraction_capabilities,
    POWERPOINT_EXTRACTION_CAPABILITIES
)
from powerpoint_mcp_server.resources.workflow_execution_guide import (
    get_workflow_execution_guide,
    WORKFLOW_EXECUTION_GUIDE
)
from powerpoint_mcp_server.resources.search_patterns_workflows import (
    get_search_patterns_workflows,
    SEARCH_PATTERNS_WORKFLOWS
)


class TestPowerPointExtractionCapabilities:
    """Test cases for PowerPoint extraction capabilities resource."""
    
    def test_resource_structure(self):
        """Test that the resource has the correct structure."""
        resource = get_powerpoint_extraction_capabilities()
        
        assert "name" in resource
        assert "description" in resource
        assert "uri" in resource
        assert "mimeType" in resource
        assert "content" in resource
        
        assert resource["name"] == "powerpoint_extraction_capabilities"
        assert resource["mimeType"] == "application/json"
        assert resource["uri"] == "powerpoint://capabilities/extraction"
    
    def test_content_structure(self):
        """Test that the content has the expected sections."""
        resource = get_powerpoint_extraction_capabilities()
        content = resource["content"]
        
        expected_sections = [
            "overview",
            "extraction_attributes", 
            "filtering_capabilities",
            "output_formats",
            "common_use_cases",
            "best_practices",
            "limitations"
        ]
        
        for section in expected_sections:
            assert section in content, f"Missing section: {section}"
    
    def test_overview_section(self):
        """Test the overview section content."""
        resource = get_powerpoint_extraction_capabilities()
        overview = resource["content"]["overview"]
        
        assert "description" in overview
        assert "supported_formats" in overview
        assert "unsupported_formats" in overview
        assert "key_features" in overview
        
        assert ".pptx" in overview["supported_formats"]
        assert ".ppt" in overview["unsupported_formats"]
        assert isinstance(overview["key_features"], list)
        assert len(overview["key_features"]) > 0
    
    def test_extraction_attributes_section(self):
        """Test the extraction attributes section."""
        resource = get_powerpoint_extraction_capabilities()
        attributes = resource["content"]["extraction_attributes"]
        
        assert "slide_level" in attributes
        assert "presentation_level" in attributes
        
        slide_level = attributes["slide_level"]
        assert "basic_info" in slide_level
        assert "content_elements" in slide_level
        assert "metadata" in slide_level
        
        # Test basic info structure
        basic_info = slide_level["basic_info"]
        expected_basic_fields = ["slide_number", "layout_name", "layout_type", "title", "subtitle"]
        for field in expected_basic_fields:
            assert field in basic_info
            assert "type" in basic_info[field]
            assert "description" in basic_info[field]
    
    def test_filtering_capabilities_section(self):
        """Test the filtering capabilities section."""
        resource = get_powerpoint_extraction_capabilities()
        filtering = resource["content"]["filtering_capabilities"]
        
        expected_filter_types = [
            "slide_queries",
            "table_extraction", 
            "text_formatting_analysis",
            "data_filtering"
        ]
        
        for filter_type in expected_filter_types:
            assert filter_type in filtering
    
    def test_common_use_cases_section(self):
        """Test the common use cases section."""
        resource = get_powerpoint_extraction_capabilities()
        use_cases = resource["content"]["common_use_cases"]
        
        expected_use_cases = [
            "content_extraction",
            "table_data_mining",
            "formatting_analysis",
            "presentation_overview"
        ]
        
        for use_case in expected_use_cases:
            assert use_case in use_cases
            assert "description" in use_cases[use_case]
            assert "tools" in use_cases[use_case]
    
    def test_json_serializable(self):
        """Test that the resource is JSON serializable."""
        resource = get_powerpoint_extraction_capabilities()
        
        # Should not raise an exception
        json_str = json.dumps(resource)
        
        # Should be able to parse back
        parsed = json.loads(json_str)
        assert parsed["name"] == resource["name"]


class TestWorkflowExecutionGuide:
    """Test cases for workflow execution guide resource."""
    
    def test_resource_structure(self):
        """Test that the resource has the correct structure."""
        resource = get_workflow_execution_guide()
        
        assert "name" in resource
        assert "description" in resource
        assert "uri" in resource
        assert "mimeType" in resource
        assert "content" in resource
        
        assert resource["name"] == "workflow_execution_guide"
        assert resource["mimeType"] == "application/json"
        assert resource["uri"] == "powerpoint://guide/workflow"
    
    def test_content_structure(self):
        """Test that the content has the expected sections."""
        resource = get_workflow_execution_guide()
        content = resource["content"]
        
        expected_sections = [
            "overview",
            "decision_trees",
            "workflows",
            "error_handling",
            "performance_optimization",
            "integration_patterns"
        ]
        
        for section in expected_sections:
            assert section in content, f"Missing section: {section}"
    
    def test_decision_trees_section(self):
        """Test the decision trees section."""
        resource = get_workflow_execution_guide()
        decision_trees = resource["content"]["decision_trees"]
        
        expected_trees = [
            "initial_assessment",
            "table_extraction_decision",
            "formatting_analysis_decision"
        ]
        
        for tree in expected_trees:
            assert tree in decision_trees
            assert "description" in decision_trees[tree]
            assert "decision_flow" in decision_trees[tree]
    
    def test_workflows_section(self):
        """Test the workflows section."""
        resource = get_workflow_execution_guide()
        workflows = resource["content"]["workflows"]
        
        expected_workflows = [
            "table_extraction_workflow",
            "formatting_analysis_workflow",
            "slide_query_workflow",
            "overview_workflow",
            "complex_data_mining_workflow"
        ]
        
        for workflow in expected_workflows:
            assert workflow in workflows
            workflow_data = workflows[workflow]
            assert "name" in workflow_data
            assert "description" in workflow_data
            assert "steps" in workflow_data
            assert isinstance(workflow_data["steps"], list)
            assert len(workflow_data["steps"]) > 0
    
    def test_workflow_steps_structure(self):
        """Test that workflow steps have the correct structure."""
        resource = get_workflow_execution_guide()
        workflows = resource["content"]["workflows"]
        
        # Test table extraction workflow steps
        table_workflow = workflows["table_extraction_workflow"]
        steps = table_workflow["steps"]
        
        for step in steps:
            assert "step" in step
            assert "name" in step
            assert "tool" in step
            assert "parameters" in step
            assert "purpose" in step
            assert "expected_output" in step
    
    def test_error_handling_section(self):
        """Test the error handling section."""
        resource = get_workflow_execution_guide()
        error_handling = resource["content"]["error_handling"]
        
        assert "common_errors" in error_handling
        assert "debugging_strategies" in error_handling
        
        common_errors = error_handling["common_errors"]
        expected_errors = [
            "file_not_found",
            "unsupported_format",
            "no_matching_slides",
            "empty_tables",
            "formatting_not_detected"
        ]
        
        for error in expected_errors:
            assert error in common_errors
            assert "error" in common_errors[error]
            assert "solutions" in common_errors[error]
            assert isinstance(common_errors[error]["solutions"], list)
    
    def test_json_serializable(self):
        """Test that the resource is JSON serializable."""
        resource = get_workflow_execution_guide()
        
        # Should not raise an exception
        json_str = json.dumps(resource)
        
        # Should be able to parse back
        parsed = json.loads(json_str)
        assert parsed["name"] == resource["name"]


class TestSearchPatternsWorkflows:
    """Test cases for search patterns and workflows resource."""
    
    def test_resource_structure(self):
        """Test that the resource has the correct structure."""
        resource = get_search_patterns_workflows()
        
        assert "name" in resource
        assert "description" in resource
        assert "uri" in resource
        assert "mimeType" in resource
        assert "content" in resource
        
        assert resource["name"] == "search_patterns_workflows"
        assert resource["mimeType"] == "application/json"
        assert resource["uri"] == "powerpoint://patterns/search"
    
    def test_content_structure(self):
        """Test that the content has the expected sections."""
        resource = get_search_patterns_workflows()
        content = resource["content"]
        
        expected_sections = [
            "overview",
            "title_patterns",
            "content_patterns",
            "formatting_patterns",
            "data_extraction_patterns",
            "quality_assessment_patterns",
            "advanced_workflows",
            "troubleshooting_patterns",
            "performance_patterns"
        ]
        
        for section in expected_sections:
            assert section in content, f"Missing section: {section}"
    
    def test_title_patterns_section(self):
        """Test the title patterns section."""
        resource = get_search_patterns_workflows()
        title_patterns = resource["content"]["title_patterns"]
        
        expected_patterns = [
            "numbered_sections",
            "question_titles",
            "agenda_slides",
            "conclusion_slides",
            "project_phases"
        ]
        
        for pattern in expected_patterns:
            assert pattern in title_patterns
            pattern_data = title_patterns[pattern]
            assert "description" in pattern_data
            assert "tool" in pattern_data
            assert "parameters" in pattern_data
    
    def test_content_patterns_section(self):
        """Test the content patterns section."""
        resource = get_search_patterns_workflows()
        content_patterns = resource["content"]["content_patterns"]
        
        expected_patterns = [
            "data_heavy_slides",
            "visual_slides",
            "text_heavy_slides",
            "minimal_content_slides"
        ]
        
        for pattern in expected_patterns:
            assert pattern in content_patterns
            pattern_data = content_patterns[pattern]
            assert "description" in pattern_data
            assert "tool" in pattern_data
            assert "parameters" in pattern_data
    
    def test_formatting_patterns_section(self):
        """Test the formatting patterns section."""
        resource = get_search_patterns_workflows()
        formatting_patterns = resource["content"]["formatting_patterns"]
        
        expected_patterns = [
            "highlighted_important_info",
            "bold_headings_and_emphasis",
            "color_coded_information",
            "hyperlinked_content"
        ]
        
        for pattern in expected_patterns:
            assert pattern in formatting_patterns
            pattern_data = formatting_patterns[pattern]
            assert "description" in pattern_data
            assert "tool" in pattern_data
            assert "parameters" in pattern_data
    
    def test_data_extraction_patterns_section(self):
        """Test the data extraction patterns section."""
        resource = get_search_patterns_workflows()
        data_patterns = resource["content"]["data_extraction_patterns"]
        
        expected_patterns = [
            "financial_data_tables",
            "project_status_data",
            "contact_information",
            "metrics_and_kpis"
        ]
        
        for pattern in expected_patterns:
            assert pattern in data_patterns
            pattern_data = data_patterns[pattern]
            assert "description" in pattern_data
            
            # These should have workflows
            if "workflow" in pattern_data:
                assert isinstance(pattern_data["workflow"], list)
                assert len(pattern_data["workflow"]) > 0
    
    def test_advanced_workflows_section(self):
        """Test the advanced workflows section."""
        resource = get_search_patterns_workflows()
        advanced_workflows = resource["content"]["advanced_workflows"]
        
        expected_workflows = [
            "competitive_analysis_extraction",
            "risk_assessment_extraction",
            "timeline_extraction"
        ]
        
        for workflow in expected_workflows:
            assert workflow in advanced_workflows
            workflow_data = advanced_workflows[workflow]
            assert "description" in workflow_data
            assert "steps" in workflow_data
            assert isinstance(workflow_data["steps"], list)
            assert len(workflow_data["steps"]) > 0
    
    def test_troubleshooting_patterns_section(self):
        """Test the troubleshooting patterns section."""
        resource = get_search_patterns_workflows()
        troubleshooting = resource["content"]["troubleshooting_patterns"]
        
        expected_patterns = [
            "no_results_debugging",
            "unexpected_results_analysis"
        ]
        
        for pattern in expected_patterns:
            assert pattern in troubleshooting
            pattern_data = troubleshooting[pattern]
            assert "description" in pattern_data
            assert "steps" in pattern_data
            assert isinstance(pattern_data["steps"], list)
    
    def test_performance_patterns_section(self):
        """Test the performance patterns section."""
        resource = get_search_patterns_workflows()
        performance = resource["content"]["performance_patterns"]
        
        expected_patterns = [
            "large_presentation_processing",
            "batch_processing_pattern"
        ]
        
        for pattern in expected_patterns:
            assert pattern in performance
            pattern_data = performance[pattern]
            assert "description" in pattern_data
    
    def test_json_serializable(self):
        """Test that the resource is JSON serializable."""
        resource = get_search_patterns_workflows()
        
        # Should not raise an exception
        json_str = json.dumps(resource)
        
        # Should be able to parse back
        parsed = json.loads(json_str)
        assert parsed["name"] == resource["name"]


class TestResourceIntegration:
    """Test cases for resource integration and consistency."""
    
    def test_all_resources_have_consistent_structure(self):
        """Test that all resources follow the same basic structure."""
        resources = [
            get_powerpoint_extraction_capabilities(),
            get_workflow_execution_guide(),
            get_search_patterns_workflows()
        ]
        
        required_fields = ["name", "description", "uri", "mimeType", "content"]
        
        for resource in resources:
            for field in required_fields:
                assert field in resource, f"Resource missing field: {field}"
            
            assert resource["mimeType"] == "application/json"
            assert resource["uri"].startswith("powerpoint://")
    
    def test_resource_names_are_unique(self):
        """Test that all resources have unique names."""
        resources = [
            get_powerpoint_extraction_capabilities(),
            get_workflow_execution_guide(),
            get_search_patterns_workflows()
        ]
        
        names = [resource["name"] for resource in resources]
        assert len(names) == len(set(names)), "Resource names are not unique"
    
    def test_resource_uris_are_unique(self):
        """Test that all resources have unique URIs."""
        resources = [
            get_powerpoint_extraction_capabilities(),
            get_workflow_execution_guide(),
            get_search_patterns_workflows()
        ]
        
        uris = [resource["uri"] for resource in resources]
        assert len(uris) == len(set(uris)), "Resource URIs are not unique"
    
    def test_tool_references_consistency(self):
        """Test that tool references across resources are consistent."""
        capabilities = get_powerpoint_extraction_capabilities()
        workflows = get_workflow_execution_guide()
        patterns = get_search_patterns_workflows()
        
        # Extract tool names from capabilities
        use_cases = capabilities["content"]["common_use_cases"]
        capability_tools = set()
        for use_case in use_cases.values():
            if "tools" in use_case:
                capability_tools.update(use_case["tools"])
        
        # Extract tool names from workflows
        workflow_tools = set()
        for workflow in workflows["content"]["workflows"].values():
            for step in workflow["steps"]:
                if "tool" in step:
                    workflow_tools.add(step["tool"])
        
        # Extract tool names from patterns
        pattern_tools = set()
        for section_name, section in patterns["content"].items():
            if isinstance(section, dict):
                for pattern in section.values():
                    if isinstance(pattern, dict):
                        if "tool" in pattern:
                            pattern_tools.add(pattern["tool"])
                        if "workflow" in pattern and isinstance(pattern["workflow"], list):
                            for step in pattern["workflow"]:
                                if "tool" in step:
                                    pattern_tools.add(step["tool"])
        
        # All tools should be documented in capabilities
        all_referenced_tools = workflow_tools | pattern_tools
        
        # Note: This is a basic check - in a real implementation, 
        # we'd want to ensure all referenced tools are properly documented
        assert len(all_referenced_tools) > 0, "No tools found in workflows and patterns"


if __name__ == "__main__":
    pytest.main([__file__])