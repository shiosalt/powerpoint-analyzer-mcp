"""
Unit tests for MCP Prompts.
"""

import pytest
import json

from powerpoint_mcp_server.prompts.complex_data_extraction import (
    get_complex_data_extraction_prompt,
    COMPLEX_DATA_EXTRACTION_PROMPT
)
from powerpoint_mcp_server.prompts.progressive_table_analysis import (
    get_progressive_table_analysis_prompt,
    PROGRESSIVE_TABLE_ANALYSIS_PROMPT
)
from powerpoint_mcp_server.prompts.adaptive_search_strategy import (
    get_adaptive_search_strategy_prompt,
    ADAPTIVE_SEARCH_STRATEGY_PROMPT
)


class TestComplexDataExtractionPrompt:
    """Test cases for complex data extraction prompt."""
    
    def test_prompt_structure(self):
        """Test that the prompt has the correct structure."""
        prompt = get_complex_data_extraction_prompt()
        
        assert "name" in prompt
        assert "description" in prompt
        assert "arguments" in prompt
        assert "template" in prompt
        assert "examples" in prompt
        
        assert prompt["name"] == "complex_data_extraction"
        assert isinstance(prompt["arguments"], list)
        assert isinstance(prompt["template"], str)
        assert isinstance(prompt["examples"], list)
    
    def test_arguments_structure(self):
        """Test that arguments have the correct structure."""
        prompt = get_complex_data_extraction_prompt()
        arguments = prompt["arguments"]
        
        required_args = []
        optional_args = []
        
        for arg in arguments:
            assert "name" in arg
            assert "description" in arg
            assert "required" in arg
            
            if arg["required"]:
                required_args.append(arg["name"])
            else:
                optional_args.append(arg["name"])
        
        # Should have at least file_path as required
        assert "file_path" in required_args
        
        # Should have optional parameters for customization
        assert len(optional_args) > 0
    
    def test_template_content(self):
        """Test that the template contains expected sections."""
        prompt = get_complex_data_extraction_prompt()
        template = prompt["template"]
        
        expected_sections = [
            "Analysis Approach",
            "Step-by-Step Workflow",
            "Data Interpretation Guidelines",
            "Output Formatting",
            "Error Handling",
            "Quality Validation"
        ]
        
        for section in expected_sections:
            assert section in template, f"Missing section: {section}"
    
    def test_template_tool_references(self):
        """Test that the template references appropriate tools."""
        prompt = get_complex_data_extraction_prompt()
        template = prompt["template"]
        
        expected_tools = [
            "get_presentation_overview",
            "query_slides",
            "extract_table_data",
            "analyze_text_formatting",
            "filter_and_aggregate"
        ]
        
        for tool in expected_tools:
            assert tool in template, f"Missing tool reference: {tool}"
    
    def test_examples_structure(self):
        """Test that examples have the correct structure."""
        prompt = get_complex_data_extraction_prompt()
        examples = prompt["examples"]
        
        assert len(examples) > 0
        
        for example in examples:
            assert "name" in example
            assert "description" in example
            assert "arguments" in example
            assert "expected_workflow" in example
            
            # Arguments should match prompt arguments
            example_args = example["arguments"]
            assert "file_path" in example_args
    
    def test_template_parameter_substitution(self):
        """Test that template uses parameter substitution correctly."""
        prompt = get_complex_data_extraction_prompt()
        template = prompt["template"]
        
        # Should contain parameter placeholders
        expected_placeholders = [
            "{file_path}",
            "{data_types}",
            "{search_criteria}",
            "{output_format}"
        ]
        
        for placeholder in expected_placeholders:
            assert placeholder in template, f"Missing placeholder: {placeholder}"


class TestProgressiveTableAnalysisPrompt:
    """Test cases for progressive table analysis prompt."""
    
    def test_prompt_structure(self):
        """Test that the prompt has the correct structure."""
        prompt = get_progressive_table_analysis_prompt()
        
        assert "name" in prompt
        assert "description" in prompt
        assert "arguments" in prompt
        assert "template" in prompt
        assert "examples" in prompt
        
        assert prompt["name"] == "progressive_table_analysis"
    
    def test_progressive_phases(self):
        """Test that the template includes progressive analysis phases."""
        prompt = get_progressive_table_analysis_prompt()
        template = prompt["template"]
        
        expected_phases = [
            "Phase 1: Discovery and Overview",
            "Phase 2: Pattern Recognition",
            "Phase 3: Focused Analysis",
            "Phase 4: Deep Dive Analysis",
            "Phase 5: Refinement and Validation"
        ]
        
        for phase in expected_phases:
            assert phase in template, f"Missing phase: {phase}"
    
    def test_analysis_focus_options(self):
        """Test that template handles different analysis focus options."""
        prompt = get_progressive_table_analysis_prompt()
        template = prompt["template"]
        
        focus_options = [
            "Option A: Overview Focus",
            "Option B: Specific Columns Focus", 
            "Option C: Formatting Patterns Focus"
        ]
        
        for option in focus_options:
            assert option in template, f"Missing focus option: {option}"
    
    def test_refinement_strategies(self):
        """Test that template includes refinement strategies."""
        prompt = get_progressive_table_analysis_prompt()
        template = prompt["template"]
        
        refinement_sections = [
            "Progressive Refinement Strategies",
            "Refinement by Criteria",
            "Adaptive Analysis Flow"
        ]
        
        for section in refinement_sections:
            assert section in template, f"Missing refinement section: {section}"
    
    def test_examples_progression(self):
        """Test that examples show clear progression."""
        prompt = get_progressive_table_analysis_prompt()
        examples = prompt["examples"]
        
        for example in examples:
            assert "expected_progression" in example
            progression = example["expected_progression"]
            assert isinstance(progression, list)
            assert len(progression) >= 3  # Should have multiple steps


class TestAdaptiveSearchStrategyPrompt:
    """Test cases for adaptive search strategy prompt."""
    
    def test_prompt_structure(self):
        """Test that the prompt has the correct structure."""
        prompt = get_adaptive_search_strategy_prompt()
        
        assert "name" in prompt
        assert "description" in prompt
        assert "arguments" in prompt
        assert "template" in prompt
        assert "examples" in prompt
        
        assert prompt["name"] == "adaptive_search_strategy"
    
    def test_adaptive_phases(self):
        """Test that the template includes adaptive search phases."""
        prompt = get_adaptive_search_strategy_prompt()
        template = prompt["template"]
        
        expected_phases = [
            "Phase 1: Initial Reconnaissance",
            "Phase 2: Result Analysis and Strategy Adaptation",
            "Phase 3: Intelligent Query Refinement",
            "Phase 4: Result Validation",
            "Phase 5: Final Optimization"
        ]
        
        for phase in expected_phases:
            assert phase in template, f"Missing phase: {phase}"
    
    def test_adaptation_strategies(self):
        """Test that template includes different adaptation strategies."""
        prompt = get_adaptive_search_strategy_prompt()
        template = prompt["template"]
        
        strategies = [
            "Strategy A: Refinement",
            "Strategy B: Expansion", 
            "Strategy C: Alternative Approach",
            "Strategy D: Pattern-Based Search"
        ]
        
        for strategy in strategies:
            assert strategy in template, f"Missing strategy: {strategy}"
    
    def test_search_objectives(self):
        """Test that template handles different search objectives."""
        prompt = get_adaptive_search_strategy_prompt()
        template = prompt["template"]
        
        objectives = [
            'For "specific_data" objective',
            'For "content_type" objective',
            'For "patterns" objective',
            'For "quality_issues" objective'
        ]
        
        for objective in objectives:
            assert objective in template, f"Missing objective: {objective}"
    
    def test_error_recovery_strategies(self):
        """Test that template includes error recovery strategies."""
        prompt = get_adaptive_search_strategy_prompt()
        template = prompt["template"]
        
        recovery_sections = [
            "Error Recovery and Fallback Strategies",
            "When searches return no results",
            "When searches return too many irrelevant results",
            "When results are inconsistent"
        ]
        
        for section in recovery_sections:
            assert section in template, f"Missing recovery section: {section}"
    
    def test_examples_adaptations(self):
        """Test that examples show expected adaptations."""
        prompt = get_adaptive_search_strategy_prompt()
        examples = prompt["examples"]
        
        for example in examples:
            assert "expected_adaptations" in example
            adaptations = example["expected_adaptations"]
            assert isinstance(adaptations, list)
            assert len(adaptations) >= 3  # Should have multiple adaptation steps


class TestPromptIntegration:
    """Test cases for prompt integration and consistency."""
    
    def test_all_prompts_have_consistent_structure(self):
        """Test that all prompts follow the same basic structure."""
        prompts = [
            get_complex_data_extraction_prompt(),
            get_progressive_table_analysis_prompt(),
            get_adaptive_search_strategy_prompt()
        ]
        
        required_fields = ["name", "description", "arguments", "template", "examples"]
        
        for prompt in prompts:
            for field in required_fields:
                assert field in prompt, f"Prompt missing field: {field}"
    
    def test_prompt_names_are_unique(self):
        """Test that all prompts have unique names."""
        prompts = [
            get_complex_data_extraction_prompt(),
            get_progressive_table_analysis_prompt(),
            get_adaptive_search_strategy_prompt()
        ]
        
        names = [prompt["name"] for prompt in prompts]
        assert len(names) == len(set(names)), "Prompt names are not unique"
    
    def test_all_prompts_reference_file_path(self):
        """Test that all prompts require file_path argument."""
        prompts = [
            get_complex_data_extraction_prompt(),
            get_progressive_table_analysis_prompt(),
            get_adaptive_search_strategy_prompt()
        ]
        
        for prompt in prompts:
            arguments = prompt["arguments"]
            file_path_args = [arg for arg in arguments if arg["name"] == "file_path"]
            assert len(file_path_args) == 1, f"Prompt {prompt['name']} missing file_path argument"
            assert file_path_args[0]["required"] is True, f"file_path should be required in {prompt['name']}"
    
    def test_tool_references_consistency(self):
        """Test that tool references across prompts are consistent."""
        prompts = [
            get_complex_data_extraction_prompt(),
            get_progressive_table_analysis_prompt(),
            get_adaptive_search_strategy_prompt()
        ]
        
        # Extract tool names from all prompts
        all_tools = set()
        for prompt in prompts:
            template = prompt["template"]
            
            # Look for tool references in the format "Tool: tool_name"
            import re
            tool_matches = re.findall(r'Tool:\s*(\w+)', template)
            all_tools.update(tool_matches)
        
        # Should have found some tools
        assert len(all_tools) > 0, "No tool references found in prompts"
        
        # Common tools should appear across multiple prompts
        expected_common_tools = [
            "get_presentation_overview",
            "query_slides",
            "extract_table_data"
        ]
        
        for tool in expected_common_tools:
            assert tool in all_tools, f"Expected common tool {tool} not found"
    
    def test_examples_have_realistic_arguments(self):
        """Test that examples use realistic argument values."""
        prompts = [
            get_complex_data_extraction_prompt(),
            get_progressive_table_analysis_prompt(),
            get_adaptive_search_strategy_prompt()
        ]
        
        for prompt in prompts:
            examples = prompt["examples"]
            for example in examples:
                arguments = example["arguments"]
                
                # Should have file_path
                assert "file_path" in arguments
                assert arguments["file_path"].endswith(".pptx")
                
                # Arguments should be realistic
                for key, value in arguments.items():
                    assert value is not None
                    assert value != ""
    
    def test_templates_are_comprehensive(self):
        """Test that templates are comprehensive and detailed."""
        prompts = [
            get_complex_data_extraction_prompt(),
            get_progressive_table_analysis_prompt(),
            get_adaptive_search_strategy_prompt()
        ]
        
        for prompt in prompts:
            template = prompt["template"]
            
            # Should be substantial content
            assert len(template) > 5000, f"Template for {prompt['name']} seems too short"
            
            # Should contain code blocks for tool usage
            assert "```" in template, f"Template for {prompt['name']} missing code examples"
            
            # Should contain parameter references
            assert "{" in template and "}" in template, f"Template for {prompt['name']} missing parameter substitution"
    
    def test_json_serializable(self):
        """Test that all prompts are JSON serializable."""
        prompts = [
            get_complex_data_extraction_prompt(),
            get_progressive_table_analysis_prompt(),
            get_adaptive_search_strategy_prompt()
        ]
        
        for prompt in prompts:
            # Should not raise an exception
            json_str = json.dumps(prompt)
            
            # Should be able to parse back
            parsed = json.loads(json_str)
            assert parsed["name"] == prompt["name"]


if __name__ == "__main__":
    pytest.main([__file__])