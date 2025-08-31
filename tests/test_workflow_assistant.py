"""
Unit tests for WorkflowAssistant.
"""

import pytest
import time
from unittest.mock import Mock, patch

from powerpoint_mcp_server.core.workflow_assistant import (
    WorkflowAssistant,
    WorkflowDetector,
    WorkflowSession,
    WorkflowStep,
    WorkflowSuggestion,
    ErrorRecoveryStrategy,
    WorkflowPattern,
    ExecutionContext,
    get_workflow_assistant
)


class TestWorkflowDetector:
    """Test cases for WorkflowDetector."""
    
    @pytest.fixture
    def detector(self):
        """Create a WorkflowDetector instance."""
        return WorkflowDetector()
    
    @pytest.fixture
    def sample_session(self):
        """Create a sample workflow session."""
        session = WorkflowSession(
            session_id="test_session",
            file_path="test.pptx",
            start_time=time.time()
        )
        return session
    
    def test_detect_table_extraction_pattern(self, detector, sample_session):
        """Test detection of table extraction pattern."""
        # Add steps that match table extraction pattern
        sample_session.steps = [
            WorkflowStep(
                tool_name="query_slides",
                parameters={"search_criteria": {"content": {"has_tables": True}}},
                timestamp=time.time()
            ),
            WorkflowStep(
                tool_name="extract_table_data",
                parameters={"slide_numbers": [1, 2, 3]},
                timestamp=time.time()
            )
        ]
        
        pattern = detector.detect_pattern(sample_session)
        assert pattern == WorkflowPattern.TABLE_EXTRACTION
    
    def test_detect_formatting_analysis_pattern(self, detector, sample_session):
        """Test detection of formatting analysis pattern."""
        sample_session.steps = [
            WorkflowStep(
                tool_name="analyze_text_formatting",
                parameters={"formatting_filter": {"formatting_types": ["bold"]}},
                timestamp=time.time()
            )
        ]
        
        pattern = detector.detect_pattern(sample_session)
        assert pattern == WorkflowPattern.FORMATTING_ANALYSIS
    
    def test_detect_content_search_pattern(self, detector, sample_session):
        """Test detection of content search pattern."""
        sample_session.steps = [
            WorkflowStep(
                tool_name="query_slides",
                parameters={"search_criteria": {"title": {"contains": "budget"}}},
                timestamp=time.time()
            )
        ]
        
        pattern = detector.detect_pattern(sample_session)
        assert pattern == WorkflowPattern.CONTENT_SEARCH
    
    def test_detect_presentation_overview_pattern(self, detector, sample_session):
        """Test detection of presentation overview pattern."""
        sample_session.steps = [
            WorkflowStep(
                tool_name="get_presentation_overview",
                parameters={"analysis_depth": "detailed"},
                timestamp=time.time()
            )
        ]
        
        pattern = detector.detect_pattern(sample_session)
        assert pattern == WorkflowPattern.PRESENTATION_OVERVIEW
    
    def test_detect_data_mining_pattern(self, detector, sample_session):
        """Test detection of data mining pattern."""
        sample_session.steps = [
            WorkflowStep(
                tool_name="query_slides",
                parameters={"search_criteria": {"content": {"has_tables": True}}},
                timestamp=time.time()
            ),
            WorkflowStep(
                tool_name="extract_table_data",
                parameters={"slide_numbers": [1, 2]},
                timestamp=time.time()
            ),
            WorkflowStep(
                tool_name="filter_and_aggregate",
                parameters={"filters": [{"field": "value", "condition": "not_empty"}]},
                timestamp=time.time()
            )
        ]
        
        pattern = detector.detect_pattern(sample_session)
        assert pattern == WorkflowPattern.DATA_MINING
    
    def test_detect_unknown_pattern(self, detector, sample_session):
        """Test detection when pattern is unknown."""
        sample_session.steps = [
            WorkflowStep(
                tool_name="unknown_tool",
                parameters={},
                timestamp=time.time()
            )
        ]
        
        pattern = detector.detect_pattern(sample_session)
        assert pattern == WorkflowPattern.UNKNOWN
    
    def test_detect_exploration_context(self, detector, sample_session):
        """Test detection of exploration context."""
        sample_session.steps = [
            WorkflowStep(
                tool_name="get_presentation_overview",
                parameters={},
                timestamp=time.time()
            )
        ]
        
        context = detector.detect_context(sample_session)
        assert context == ExecutionContext.EXPLORATION
    
    def test_detect_targeted_extraction_context(self, detector, sample_session):
        """Test detection of targeted extraction context."""
        sample_session.steps = [
            WorkflowStep(
                tool_name="extract_table_data",
                parameters={},
                timestamp=time.time()
            )
        ]
        
        context = detector.detect_context(sample_session)
        assert context == ExecutionContext.TARGETED_EXTRACTION
    
    def test_detect_analysis_context(self, detector, sample_session):
        """Test detection of analysis context."""
        sample_session.steps = [
            WorkflowStep(
                tool_name="filter_and_aggregate",
                parameters={},
                timestamp=time.time()
            )
        ]
        
        context = detector.detect_context(sample_session)
        assert context == ExecutionContext.ANALYSIS
    
    def test_detect_optimization_context(self, detector, sample_session):
        """Test detection of optimization context."""
        # Add repeated tool usage
        sample_session.steps = [
            WorkflowStep(
                tool_name="query_slides",
                parameters={},
                timestamp=time.time()
            ),
            WorkflowStep(
                tool_name="query_slides",
                parameters={},
                timestamp=time.time()
            ),
            WorkflowStep(
                tool_name="query_slides",
                parameters={},
                timestamp=time.time()
            )
        ]
        
        context = detector.detect_context(sample_session)
        assert context == ExecutionContext.OPTIMIZATION
    
    def test_matches_sequence(self, detector):
        """Test sequence matching."""
        tool_sequence = ["get_presentation_overview", "query_slides", "extract_table_data"]
        pattern = ["query_slides", "extract_table_data"]
        
        assert detector._matches_sequence(tool_sequence, pattern) is True
        
        pattern = ["extract_table_data", "query_slides"]
        assert detector._matches_sequence(tool_sequence, pattern) is False
    
    def test_partial_sequence_match(self, detector):
        """Test partial sequence matching."""
        tool_sequence = ["get_presentation_overview", "query_slides"]
        pattern = ["query_slides", "extract_table_data"]
        
        score = detector._partial_sequence_match(tool_sequence, pattern)
        assert 0.0 < score < 1.0  # Should be partial match


class TestWorkflowAssistant:
    """Test cases for WorkflowAssistant."""
    
    @pytest.fixture
    def assistant(self):
        """Create a WorkflowAssistant instance."""
        return WorkflowAssistant()
    
    def test_start_session(self, assistant):
        """Test starting a new workflow session."""
        session = assistant.start_session("test_session", "test.pptx")
        
        assert session.session_id == "test_session"
        assert session.file_path == "test.pptx"
        assert session.start_time > 0
        assert len(session.steps) == 0
        assert "test_session" in assistant.sessions
    
    def test_record_step(self, assistant):
        """Test recording a workflow step."""
        session = assistant.start_session("test_session", "test.pptx")
        
        assistant.record_step(
            session_id="test_session",
            tool_name="query_slides",
            parameters={"search_criteria": {"content": {"has_tables": True}}},
            result_summary={"slides_found": 5},
            execution_time=1.5,
            success=True
        )
        
        assert len(session.steps) == 1
        step = session.steps[0]
        assert step.tool_name == "query_slides"
        assert step.success is True
        assert step.execution_time == 1.5
        assert step.result_summary == {"slides_found": 5}
    
    def test_record_failed_step(self, assistant):
        """Test recording a failed workflow step."""
        session = assistant.start_session("test_session", "test.pptx")
        
        assistant.record_step(
            session_id="test_session",
            tool_name="extract_table_data",
            parameters={"slide_numbers": [999]},
            success=False,
            error_message="Slide not found"
        )
        
        assert len(session.steps) == 1
        step = session.steps[0]
        assert step.success is False
        assert step.error_message == "Slide not found"
    
    def test_get_next_suggestions_table_extraction(self, assistant):
        """Test getting suggestions for table extraction workflow."""
        session = assistant.start_session("test_session", "test.pptx")
        
        # Record a query_slides step
        assistant.record_step(
            session_id="test_session",
            tool_name="query_slides",
            parameters={"search_criteria": {"content": {"has_tables": True}}},
            success=True
        )
        
        suggestions = assistant.get_next_suggestions("test_session")
        
        assert len(suggestions) > 0
        # Should suggest extract_table_data after query_slides
        extract_suggestion = next(
            (s for s in suggestions if s.tool_name == "extract_table_data"), 
            None
        )
        assert extract_suggestion is not None
        assert extract_suggestion.confidence > 0.5
    
    def test_get_next_suggestions_first_step(self, assistant):
        """Test getting suggestions for first step."""
        session = assistant.start_session("test_session", "test.pptx")
        
        suggestions = assistant.get_next_suggestions("test_session")
        
        assert len(suggestions) > 0
        # Should suggest presentation overview as first step
        overview_suggestion = next(
            (s for s in suggestions if s.tool_name == "get_presentation_overview"), 
            None
        )
        assert overview_suggestion is not None
        assert overview_suggestion.confidence > 0.8
    
    def test_get_error_recovery_strategies_no_results(self, assistant):
        """Test getting error recovery strategies for no results."""
        session = assistant.start_session("test_session", "test.pptx")
        
        strategies = assistant.get_error_recovery_strategies(
            session_id="test_session",
            error_message="No results found",
            failed_tool="query_slides",
            failed_parameters={"search_criteria": {"title": {"contains": "nonexistent"}}}
        )
        
        assert len(strategies) > 0
        # Should suggest broadening search criteria
        broaden_strategy = next(
            (s for s in strategies if "broaden" in s.strategy_name.lower()), 
            None
        )
        assert broaden_strategy is not None
        assert broaden_strategy.confidence > 0.5
    
    def test_get_error_recovery_strategies_file_access(self, assistant):
        """Test getting error recovery strategies for file access errors."""
        session = assistant.start_session("test_session", "test.pptx")
        
        strategies = assistant.get_error_recovery_strategies(
            session_id="test_session",
            error_message="File not found",
            failed_tool="get_presentation_overview",
            failed_parameters={"file_path": "nonexistent.pptx"}
        )
        
        assert len(strategies) > 0
        # Should suggest verifying file path
        verify_strategy = next(
            (s for s in strategies if "verify" in s.strategy_name.lower()), 
            None
        )
        assert verify_strategy is not None
    
    def test_get_session_summary(self, assistant):
        """Test getting session summary."""
        session = assistant.start_session("test_session", "test.pptx")
        
        # Add some steps
        assistant.record_step(
            session_id="test_session",
            tool_name="query_slides",
            parameters={},
            execution_time=1.0,
            success=True
        )
        assistant.record_step(
            session_id="test_session",
            tool_name="extract_table_data",
            parameters={},
            execution_time=2.0,
            success=False,
            error_message="No tables found"
        )
        
        summary = assistant.get_session_summary("test_session")
        
        assert summary["session_id"] == "test_session"
        assert summary["file_path"] == "test.pptx"
        assert summary["total_steps"] == 2
        assert summary["successful_steps"] == 1
        assert summary["failed_steps"] == 1
        assert summary["average_execution_time"] == 1.5
        assert "query_slides" in summary["tools_used"]
        assert "extract_table_data" in summary["tools_used"]
    
    def test_cleanup_session(self, assistant):
        """Test cleaning up a session."""
        session = assistant.start_session("test_session", "test.pptx")
        assert "test_session" in assistant.sessions
        
        assistant.cleanup_session("test_session")
        assert "test_session" not in assistant.sessions
    
    def test_learning_from_successful_steps(self, assistant):
        """Test that assistant learns from successful steps."""
        session = assistant.start_session("test_session", "test.pptx")
        
        # Record successful step
        assistant.record_step(
            session_id="test_session",
            tool_name="query_slides",
            parameters={"search_criteria": {"content": {"has_tables": True}}},
            success=True,
            execution_time=1.0
        )
        
        # Check that learning data was recorded
        assert len(assistant.learning_data) > 0
    
    def test_learning_from_failed_steps(self, assistant):
        """Test that assistant learns from failed steps."""
        session = assistant.start_session("test_session", "test.pptx")
        
        # Record failed step
        assistant.record_step(
            session_id="test_session",
            tool_name="extract_table_data",
            parameters={"slide_numbers": [999]},
            success=False,
            error_message="Slide not found"
        )
        
        # Check that error learning data was recorded
        error_keys = [key for key in assistant.learning_data.keys() if key.startswith("error_")]
        assert len(error_keys) > 0
    
    def test_get_learning_insights(self, assistant):
        """Test getting learning insights."""
        session = assistant.start_session("test_session", "test.pptx")
        
        # Add some learning data
        assistant.record_step(
            session_id="test_session",
            tool_name="query_slides",
            parameters={},
            success=True
        )
        assistant.record_step(
            session_id="test_session",
            tool_name="extract_table_data",
            parameters={},
            success=False,
            error_message="Error"
        )
        
        insights = assistant.get_learning_insights()
        
        assert "most_common_patterns" in insights
        assert "most_successful_tools" in insights
        assert "common_errors" in insights
        assert "performance_insights" in insights
    
    def test_simplify_parameters(self, assistant):
        """Test parameter simplification for error recovery."""
        complex_params = {
            "file_path": "test.pptx",
            "slide_numbers": [1, 2, 3],
            "search_criteria": {"complex": "criteria"},
            "formatting_detection": {"detect_bold": True},
            "other_param": "value"
        }
        
        simplified = assistant._simplify_parameters(complex_params)
        
        # Should keep only essential parameters
        assert "file_path" in simplified
        assert "slide_numbers" in simplified
        assert "search_criteria" not in simplified
        assert "formatting_detection" not in simplified


class TestWorkflowStep:
    """Test cases for WorkflowStep."""
    
    def test_workflow_step_creation(self):
        """Test creating a WorkflowStep."""
        step = WorkflowStep(
            tool_name="query_slides",
            parameters={"search_criteria": {"title": {"contains": "test"}}},
            timestamp=time.time(),
            execution_time=1.5,
            result_summary={"slides_found": 3},
            success=True
        )
        
        assert step.tool_name == "query_slides"
        assert step.parameters == {"search_criteria": {"title": {"contains": "test"}}}
        assert step.execution_time == 1.5
        assert step.result_summary == {"slides_found": 3}
        assert step.success is True
        assert step.error_message is None


class TestWorkflowSuggestion:
    """Test cases for WorkflowSuggestion."""
    
    def test_workflow_suggestion_creation(self):
        """Test creating a WorkflowSuggestion."""
        suggestion = WorkflowSuggestion(
            tool_name="extract_table_data",
            parameters={"slide_numbers": [1, 2, 3]},
            reasoning="Extract data from found slides",
            confidence=0.8,
            expected_outcome="Table data with formatting",
            alternatives=[{"tool": "get_powerpoint_attributes"}]
        )
        
        assert suggestion.tool_name == "extract_table_data"
        assert suggestion.parameters == {"slide_numbers": [1, 2, 3]}
        assert suggestion.reasoning == "Extract data from found slides"
        assert suggestion.confidence == 0.8
        assert suggestion.expected_outcome == "Table data with formatting"
        assert len(suggestion.alternatives) == 1


class TestErrorRecoveryStrategy:
    """Test cases for ErrorRecoveryStrategy."""
    
    def test_error_recovery_strategy_creation(self):
        """Test creating an ErrorRecoveryStrategy."""
        strategy = ErrorRecoveryStrategy(
            strategy_name="broaden_search",
            description="Broaden search criteria",
            recovery_steps=[
                {
                    "tool": "query_slides",
                    "parameters": {"search_criteria": {}},
                    "reasoning": "Use minimal criteria"
                }
            ],
            confidence=0.7,
            applicable_errors=["no results", "empty results"]
        )
        
        assert strategy.strategy_name == "broaden_search"
        assert strategy.description == "Broaden search criteria"
        assert len(strategy.recovery_steps) == 1
        assert strategy.confidence == 0.7
        assert "no results" in strategy.applicable_errors


class TestGlobalWorkflowAssistant:
    """Test cases for global workflow assistant."""
    
    def test_get_workflow_assistant_singleton(self):
        """Test that get_workflow_assistant returns singleton."""
        assistant1 = get_workflow_assistant()
        assistant2 = get_workflow_assistant()
        
        assert assistant1 is assistant2
        assert isinstance(assistant1, WorkflowAssistant)


if __name__ == "__main__":
    pytest.main([__file__])