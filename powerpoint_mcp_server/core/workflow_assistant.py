"""
Intelligent workflow assistance for PowerPoint Analyzer MCP.
"""

import logging
from typing import Dict, List, Any, Optional, Tuple
from dataclasses import dataclass, field
from enum import Enum
from collections import defaultdict, Counter
import json
import time

logger = logging.getLogger(__name__)


class WorkflowPattern(Enum):
    """Enumeration of detected workflow patterns."""
    TABLE_EXTRACTION = "table_extraction"
    FORMATTING_ANALYSIS = "formatting_analysis"
    CONTENT_SEARCH = "content_search"
    PRESENTATION_OVERVIEW = "presentation_overview"
    DATA_MINING = "data_mining"
    QUALITY_ASSESSMENT = "quality_assessment"
    UNKNOWN = "unknown"


class ExecutionContext(Enum):
    """Enumeration of execution contexts."""
    EXPLORATION = "exploration"
    TARGETED_EXTRACTION = "targeted_extraction"
    ANALYSIS = "analysis"
    VALIDATION = "validation"
    OPTIMIZATION = "optimization"


@dataclass
class WorkflowStep:
    """A single step in a workflow."""
    tool_name: str
    parameters: Dict[str, Any]
    timestamp: float
    execution_time: Optional[float] = None
    result_summary: Optional[Dict[str, Any]] = None
    success: bool = True
    error_message: Optional[str] = None


@dataclass
class WorkflowSession:
    """A workflow session tracking multiple related operations."""
    session_id: str
    file_path: str
    start_time: float
    steps: List[WorkflowStep] = field(default_factory=list)
    context: ExecutionContext = ExecutionContext.EXPLORATION
    detected_pattern: WorkflowPattern = WorkflowPattern.UNKNOWN
    user_objectives: List[str] = field(default_factory=list)
    current_strategy: Optional[str] = None
    performance_metrics: Dict[str, Any] = field(default_factory=dict)


@dataclass
class WorkflowSuggestion:
    """A suggested next step in the workflow."""
    tool_name: str
    parameters: Dict[str, Any]
    reasoning: str
    confidence: float
    expected_outcome: str
    alternatives: List[Dict[str, Any]] = field(default_factory=list)


@dataclass
class ErrorRecoveryStrategy:
    """A strategy for recovering from errors."""
    strategy_name: str
    description: str
    recovery_steps: List[Dict[str, Any]]
    confidence: float
    applicable_errors: List[str]


class WorkflowDetector:
    """
    Detects workflow patterns from user interactions.
    """
    
    def __init__(self):
        """Initialize the workflow detector."""
        self.pattern_signatures = {
            WorkflowPattern.TABLE_EXTRACTION: {
                'tools': ['query_slides', 'extract_table_data'],
                'parameters': ['has_tables', 'slide_numbers', 'table_criteria'],
                'sequence_patterns': [
                    ['query_slides', 'extract_table_data'],
                    ['get_presentation_overview', 'query_slides', 'extract_table_data']
                ]
            },
            WorkflowPattern.FORMATTING_ANALYSIS: {
                'tools': ['analyze_text_formatting'],
                'parameters': ['formatting_filter', 'content_types', 'formatting_types'],
                'sequence_patterns': [
                    ['analyze_text_formatting'],
                    ['query_slides', 'analyze_text_formatting']
                ]
            },
            WorkflowPattern.CONTENT_SEARCH: {
                'tools': ['query_slides'],
                'parameters': ['search_criteria', 'title', 'content'],
                'sequence_patterns': [
                    ['query_slides'],
                    ['get_presentation_overview', 'query_slides']
                ]
            },
            WorkflowPattern.PRESENTATION_OVERVIEW: {
                'tools': ['get_presentation_overview'],
                'parameters': ['analysis_depth'],
                'sequence_patterns': [
                    ['get_presentation_overview']
                ]
            },
            WorkflowPattern.DATA_MINING: {
                'tools': ['query_slides', 'extract_table_data', 'filter_and_aggregate'],
                'parameters': ['has_tables', 'filters', 'grouping'],
                'sequence_patterns': [
                    ['query_slides', 'extract_table_data', 'filter_and_aggregate'],
                    ['get_presentation_overview', 'query_slides', 'extract_table_data', 'analyze_text_formatting']
                ]
            }
        }
    
    def detect_pattern(self, session: WorkflowSession) -> WorkflowPattern:
        """Detect the workflow pattern from session steps."""
        try:
            if len(session.steps) == 0:
                return WorkflowPattern.UNKNOWN
            
            # Extract tool sequence
            tool_sequence = [step.tool_name for step in session.steps]
            
            # Extract used parameters
            used_parameters = set()
            for step in session.steps:
                used_parameters.update(step.parameters.keys())
            
            # Score each pattern
            pattern_scores = {}
            
            for pattern, signature in self.pattern_signatures.items():
                score = 0.0
                
                # Tool usage score
                pattern_tools = set(signature['tools'])
                used_tools = set(tool_sequence)
                tool_overlap = len(pattern_tools & used_tools) / len(pattern_tools)
                score += tool_overlap * 0.4
                
                # Parameter usage score
                pattern_params = set(signature['parameters'])
                param_overlap = len(pattern_params & used_parameters) / len(pattern_params) if pattern_params else 0
                score += param_overlap * 0.3
                
                # Sequence pattern score
                sequence_score = 0.0
                for seq_pattern in signature['sequence_patterns']:
                    if self._matches_sequence(tool_sequence, seq_pattern):
                        sequence_score = max(sequence_score, 1.0)
                    else:
                        # Partial sequence match
                        partial_score = self._partial_sequence_match(tool_sequence, seq_pattern)
                        sequence_score = max(sequence_score, partial_score)
                
                score += sequence_score * 0.3
                
                pattern_scores[pattern] = score
            
            # Return the highest scoring pattern
            best_pattern = max(pattern_scores, key=pattern_scores.get)
            if pattern_scores[best_pattern] > 0.5:
                return best_pattern
            else:
                return WorkflowPattern.UNKNOWN
                
        except Exception as e:
            logger.warning(f"Failed to detect workflow pattern: {e}")
            return WorkflowPattern.UNKNOWN
    
    def _matches_sequence(self, tool_sequence: List[str], pattern: List[str]) -> bool:
        """Check if tool sequence matches a pattern exactly."""
        if len(tool_sequence) < len(pattern):
            return False
        
        # Check if pattern appears as subsequence
        for i in range(len(tool_sequence) - len(pattern) + 1):
            if tool_sequence[i:i+len(pattern)] == pattern:
                return True
        
        return False
    
    def _partial_sequence_match(self, tool_sequence: List[str], pattern: List[str]) -> float:
        """Calculate partial sequence match score."""
        if not pattern:
            return 0.0
        
        # Find longest common subsequence
        max_match = 0
        for i in range(len(tool_sequence)):
            match_count = 0
            pattern_idx = 0
            
            for j in range(i, len(tool_sequence)):
                if pattern_idx < len(pattern) and tool_sequence[j] == pattern[pattern_idx]:
                    match_count += 1
                    pattern_idx += 1
            
            max_match = max(max_match, match_count)
        
        return max_match / len(pattern)
    
    def detect_context(self, session: WorkflowSession) -> ExecutionContext:
        """Detect the current execution context."""
        try:
            if len(session.steps) == 0:
                return ExecutionContext.EXPLORATION
            
            recent_steps = session.steps[-3:]  # Look at last 3 steps
            recent_tools = [step.tool_name for step in recent_steps]
            
            # Context detection rules
            if 'get_presentation_overview' in recent_tools:
                return ExecutionContext.EXPLORATION
            
            if any(tool in recent_tools for tool in ['extract_table_data', 'analyze_text_formatting']):
                return ExecutionContext.TARGETED_EXTRACTION
            
            if 'filter_and_aggregate' in recent_tools:
                return ExecutionContext.ANALYSIS
            
            # Check for repeated similar operations (optimization)
            if len(session.steps) >= 3:
                last_three_tools = [step.tool_name for step in session.steps[-3:]]
                if len(set(last_three_tools)) == 1:  # Same tool repeated
                    return ExecutionContext.OPTIMIZATION
            
            # Check for validation patterns
            if len(session.steps) >= 2:
                last_two = recent_tools[-2:]
                if last_two == ['extract_table_data', 'get_powerpoint_attributes']:
                    return ExecutionContext.VALIDATION
            
            return ExecutionContext.TARGETED_EXTRACTION
            
        except Exception as e:
            logger.warning(f"Failed to detect execution context: {e}")
            return ExecutionContext.EXPLORATION


class WorkflowAssistant:
    """
    Provides intelligent workflow assistance and suggestions.
    """
    
    def __init__(self):
        """Initialize the workflow assistant."""
        self.detector = WorkflowDetector()
        self.sessions = {}
        self.learning_data = defaultdict(list)
    
    def start_session(self, session_id: str, file_path: str) -> WorkflowSession:
        """Start a new workflow session."""
        session = WorkflowSession(
            session_id=session_id,
            file_path=file_path,
            start_time=time.time()
        )
        self.sessions[session_id] = session
        return session
    
    def record_step(
        self,
        session_id: str,
        tool_name: str,
        parameters: Dict[str, Any],
        result_summary: Optional[Dict[str, Any]] = None,
        execution_time: Optional[float] = None,
        success: bool = True,
        error_message: Optional[str] = None
    ) -> None:
        """Record a workflow step."""
        try:
            if session_id not in self.sessions:
                logger.warning(f"Session {session_id} not found, creating new session")
                self.start_session(session_id, "unknown")
            
            session = self.sessions[session_id]
            
            step = WorkflowStep(
                tool_name=tool_name,
                parameters=parameters,
                timestamp=time.time(),
                execution_time=execution_time,
                result_summary=result_summary,
                success=success,
                error_message=error_message
            )
            
            session.steps.append(step)
            
            # Update session context and pattern
            session.context = self.detector.detect_context(session)
            session.detected_pattern = self.detector.detect_pattern(session)
            
            # Learn from this step
            self._learn_from_step(session, step)
            
        except Exception as e:
            logger.error(f"Failed to record workflow step: {e}")
    
    def get_next_suggestions(
        self,
        session_id: str,
        max_suggestions: int = 3
    ) -> List[WorkflowSuggestion]:
        """Get suggestions for the next workflow step."""
        try:
            if session_id not in self.sessions:
                return []
            
            session = self.sessions[session_id]
            suggestions = []
            
            # Generate suggestions based on detected pattern and context
            if session.detected_pattern == WorkflowPattern.TABLE_EXTRACTION:
                suggestions.extend(self._suggest_table_extraction_next_steps(session))
            elif session.detected_pattern == WorkflowPattern.FORMATTING_ANALYSIS:
                suggestions.extend(self._suggest_formatting_analysis_next_steps(session))
            elif session.detected_pattern == WorkflowPattern.CONTENT_SEARCH:
                suggestions.extend(self._suggest_content_search_next_steps(session))
            elif session.detected_pattern == WorkflowPattern.DATA_MINING:
                suggestions.extend(self._suggest_data_mining_next_steps(session))
            else:
                suggestions.extend(self._suggest_general_next_steps(session))
            
            # Add context-specific suggestions
            suggestions.extend(self._suggest_context_specific_steps(session))
            
            # Sort by confidence and return top suggestions
            suggestions.sort(key=lambda x: x.confidence, reverse=True)
            return suggestions[:max_suggestions]
            
        except Exception as e:
            logger.error(f"Failed to generate suggestions: {e}")
            return []
    
    def get_error_recovery_strategies(
        self,
        session_id: str,
        error_message: str,
        failed_tool: str,
        failed_parameters: Dict[str, Any]
    ) -> List[ErrorRecoveryStrategy]:
        """Get error recovery strategies."""
        try:
            if session_id not in self.sessions:
                return []
            
            session = self.sessions[session_id]
            strategies = []
            
            # Analyze the error and suggest recovery strategies
            if "no results" in error_message.lower() or "empty" in error_message.lower():
                strategies.extend(self._suggest_no_results_recovery(session, failed_tool, failed_parameters))
            
            if "file not found" in error_message.lower() or "access" in error_message.lower():
                strategies.extend(self._suggest_file_access_recovery(session, failed_tool, failed_parameters))
            
            if "invalid" in error_message.lower() or "format" in error_message.lower():
                strategies.extend(self._suggest_format_error_recovery(session, failed_tool, failed_parameters))
            
            # Add general recovery strategies
            strategies.extend(self._suggest_general_recovery_strategies(session, failed_tool, failed_parameters))
            
            # Sort by confidence
            strategies.sort(key=lambda x: x.confidence, reverse=True)
            return strategies[:3]  # Return top 3 strategies
            
        except Exception as e:
            logger.error(f"Failed to generate error recovery strategies: {e}")
            return []
    
    def _suggest_table_extraction_next_steps(self, session: WorkflowSession) -> List[WorkflowSuggestion]:
        """Suggest next steps for table extraction workflow."""
        suggestions = []
        recent_tools = [step.tool_name for step in session.steps[-3:]]
        
        if 'query_slides' in recent_tools and 'extract_table_data' not in recent_tools:
            # Suggest table extraction after slide query
            suggestions.append(WorkflowSuggestion(
                tool_name='extract_table_data',
                parameters={
                    'slide_numbers': 'from_previous_query',
                    'formatting_detection': {
                        'detect_bold': True,
                        'detect_highlight': True,
                        'detect_colors': True
                    }
                },
                reasoning='Extract table data from slides found in previous query',
                confidence=0.9,
                expected_outcome='Structured table data with formatting information'
            ))
        
        elif 'extract_table_data' in recent_tools and 'filter_and_aggregate' not in recent_tools:
            # Suggest filtering after extraction
            suggestions.append(WorkflowSuggestion(
                tool_name='filter_and_aggregate',
                parameters={
                    'data_source': 'from_table_extraction',
                    'filters': [
                        {
                            'field': 'formatting.highlight',
                            'condition': 'equals',
                            'value': True
                        }
                    ]
                },
                reasoning='Filter extracted data to focus on highlighted important information',
                confidence=0.8,
                expected_outcome='Filtered data focusing on key information'
            ))
        
        return suggestions
    
    def _suggest_formatting_analysis_next_steps(self, session: WorkflowSession) -> List[WorkflowSuggestion]:
        """Suggest next steps for formatting analysis workflow."""
        suggestions = []
        recent_tools = [step.tool_name for step in session.steps[-3:]]
        
        if 'analyze_text_formatting' in recent_tools:
            # Suggest targeted extraction based on formatting findings
            suggestions.append(WorkflowSuggestion(
                tool_name='extract_table_data',
                parameters={
                    'slide_numbers': 'slides_with_formatting',
                    'formatting_detection': {
                        'detect_bold': True,
                        'detect_highlight': True,
                        'detect_colors': True
                    }
                },
                reasoning='Extract data from slides with interesting formatting patterns',
                confidence=0.7,
                expected_outcome='Table data from slides with significant formatting'
            ))
        
        return suggestions
    
    def _suggest_content_search_next_steps(self, session: WorkflowSession) -> List[WorkflowSuggestion]:
        """Suggest next steps for content search workflow."""
        suggestions = []
        recent_tools = [step.tool_name for step in session.steps[-3:]]
        
        if 'query_slides' in recent_tools:
            # Suggest content extraction from found slides
            suggestions.append(WorkflowSuggestion(
                tool_name='get_powerpoint_attributes',
                parameters={
                    'slide_numbers': 'from_search_results',
                    'attributes': ['text_elements', 'tables', 'object_counts']
                },
                reasoning='Extract detailed content from slides found in search',
                confidence=0.8,
                expected_outcome='Detailed content from matching slides'
            ))
        
        return suggestions
    
    def _suggest_data_mining_next_steps(self, session: WorkflowSession) -> List[WorkflowSuggestion]:
        """Suggest next steps for data mining workflow."""
        suggestions = []
        recent_tools = [step.tool_name for step in session.steps[-3:]]
        
        if 'filter_and_aggregate' not in recent_tools and 'extract_table_data' in recent_tools:
            # Suggest aggregation after extraction
            suggestions.append(WorkflowSuggestion(
                tool_name='filter_and_aggregate',
                parameters={
                    'data_source': 'from_extraction',
                    'grouping': {
                        'fields': ['slide_number'],
                        'aggregations': [
                            {
                                'field': 'value',
                                'operation': 'count',
                                'output_field': 'value_count'
                            }
                        ]
                    }
                },
                reasoning='Aggregate extracted data to identify patterns and trends',
                confidence=0.8,
                expected_outcome='Aggregated data showing patterns across slides'
            ))
        
        return suggestions
    
    def _suggest_general_next_steps(self, session: WorkflowSession) -> List[WorkflowSuggestion]:
        """Suggest general next steps when pattern is unknown."""
        suggestions = []
        
        if len(session.steps) == 0:
            # First step should be overview
            suggestions.append(WorkflowSuggestion(
                tool_name='get_presentation_overview',
                parameters={
                    'analysis_depth': 'basic',
                    'include_sample_content': True
                },
                reasoning='Start with presentation overview to understand structure and content',
                confidence=0.9,
                expected_outcome='Understanding of presentation structure and content types'
            ))
        
        elif 'get_presentation_overview' not in [step.tool_name for step in session.steps]:
            # Suggest overview if not done yet
            suggestions.append(WorkflowSuggestion(
                tool_name='get_presentation_overview',
                parameters={
                    'analysis_depth': 'detailed'
                },
                reasoning='Get comprehensive overview to guide further analysis',
                confidence=0.7,
                expected_outcome='Detailed presentation analysis and recommendations'
            ))
        
        return suggestions
    
    def _suggest_context_specific_steps(self, session: WorkflowSession) -> List[WorkflowSuggestion]:
        """Suggest steps based on execution context."""
        suggestions = []
        
        if session.context == ExecutionContext.EXPLORATION:
            suggestions.append(WorkflowSuggestion(
                tool_name='query_slides',
                parameters={
                    'search_criteria': {
                        'content': {
                            'has_tables': True
                        }
                    },
                    'return_details': 'detailed'
                },
                reasoning='Explore slides with tables for potential data extraction',
                confidence=0.6,
                expected_outcome='List of slides containing tabular data'
            ))
        
        elif session.context == ExecutionContext.OPTIMIZATION:
            # Suggest performance improvements
            last_step = session.steps[-1] if session.steps else None
            if last_step and last_step.execution_time and last_step.execution_time > 5.0:
                suggestions.append(WorkflowSuggestion(
                    tool_name=last_step.tool_name,
                    parameters={**last_step.parameters, 'limit': 10},
                    reasoning='Optimize performance by limiting results',
                    confidence=0.7,
                    expected_outcome='Faster execution with focused results'
                ))
        
        return suggestions
    
    def _suggest_no_results_recovery(
        self,
        session: WorkflowSession,
        failed_tool: str,
        failed_parameters: Dict[str, Any]
    ) -> List[ErrorRecoveryStrategy]:
        """Suggest recovery strategies for no results errors."""
        strategies = []
        
        if failed_tool == 'query_slides':
            strategies.append(ErrorRecoveryStrategy(
                strategy_name='broaden_search_criteria',
                description='Broaden search criteria to find more results',
                recovery_steps=[
                    {
                        'tool': 'query_slides',
                        'parameters': {
                            **failed_parameters,
                            'search_criteria': {
                                'content': {'object_count_min': 1}  # Very broad criteria
                            }
                        },
                        'reasoning': 'Use minimal criteria to find any content'
                    }
                ],
                confidence=0.8,
                applicable_errors=['no results', 'empty results']
            ))
        
        elif failed_tool == 'extract_table_data':
            strategies.append(ErrorRecoveryStrategy(
                strategy_name='lower_table_criteria',
                description='Lower table criteria to find smaller tables',
                recovery_steps=[
                    {
                        'tool': 'extract_table_data',
                        'parameters': {
                            **failed_parameters,
                            'table_criteria': {
                                'min_rows': 1,
                                'min_columns': 1
                            }
                        },
                        'reasoning': 'Accept any table structure'
                    }
                ],
                confidence=0.7,
                applicable_errors=['no tables found', 'empty extraction']
            ))
        
        return strategies
    
    def _suggest_file_access_recovery(
        self,
        session: WorkflowSession,
        failed_tool: str,
        failed_parameters: Dict[str, Any]
    ) -> List[ErrorRecoveryStrategy]:
        """Suggest recovery strategies for file access errors."""
        strategies = []
        
        strategies.append(ErrorRecoveryStrategy(
            strategy_name='verify_file_path',
            description='Verify file path and accessibility',
            recovery_steps=[
                {
                    'action': 'check_file_exists',
                    'parameters': {'file_path': session.file_path},
                    'reasoning': 'Ensure file exists and is accessible'
                }
            ],
            confidence=0.9,
            applicable_errors=['file not found', 'access denied', 'permission error']
        ))
        
        return strategies
    
    def _suggest_format_error_recovery(
        self,
        session: WorkflowSession,
        failed_tool: str,
        failed_parameters: Dict[str, Any]
    ) -> List[ErrorRecoveryStrategy]:
        """Suggest recovery strategies for format errors."""
        strategies = []
        
        strategies.append(ErrorRecoveryStrategy(
            strategy_name='validate_file_format',
            description='Validate PowerPoint file format',
            recovery_steps=[
                {
                    'action': 'check_file_format',
                    'parameters': {'file_path': session.file_path},
                    'reasoning': 'Ensure file is valid .pptx format'
                }
            ],
            confidence=0.8,
            applicable_errors=['invalid format', 'corrupted file', 'unsupported format']
        ))
        
        return strategies
    
    def _suggest_general_recovery_strategies(
        self,
        session: WorkflowSession,
        failed_tool: str,
        failed_parameters: Dict[str, Any]
    ) -> List[ErrorRecoveryStrategy]:
        """Suggest general recovery strategies."""
        strategies = []
        
        strategies.append(ErrorRecoveryStrategy(
            strategy_name='retry_with_basic_parameters',
            description='Retry with simplified parameters',
            recovery_steps=[
                {
                    'tool': failed_tool,
                    'parameters': self._simplify_parameters(failed_parameters),
                    'reasoning': 'Use basic parameters to avoid complex failures'
                }
            ],
            confidence=0.6,
            applicable_errors=['general error', 'timeout', 'processing error']
        ))
        
        return strategies
    
    def _simplify_parameters(self, parameters: Dict[str, Any]) -> Dict[str, Any]:
        """Simplify parameters for error recovery."""
        simplified = {}
        
        # Keep only essential parameters
        essential_keys = ['file_path', 'slide_numbers']
        for key in essential_keys:
            if key in parameters:
                simplified[key] = parameters[key]
        
        return simplified
    
    def _learn_from_step(self, session: WorkflowSession, step: WorkflowStep) -> None:
        """Learn from workflow step for future suggestions."""
        try:
            # Record successful patterns
            if step.success:
                pattern_key = f"{session.detected_pattern.value}_{session.context.value}"
                self.learning_data[pattern_key].append({
                    'tool': step.tool_name,
                    'parameters': step.parameters,
                    'execution_time': step.execution_time,
                    'result_summary': step.result_summary
                })
            
            # Record error patterns
            else:
                error_key = f"error_{step.tool_name}"
                self.learning_data[error_key].append({
                    'parameters': step.parameters,
                    'error': step.error_message,
                    'context': session.context.value
                })
                
        except Exception as e:
            logger.warning(f"Failed to learn from step: {e}")
    
    def get_session_summary(self, session_id: str) -> Dict[str, Any]:
        """Get a summary of the workflow session."""
        try:
            if session_id not in self.sessions:
                return {}
            
            session = self.sessions[session_id]
            
            total_time = time.time() - session.start_time
            successful_steps = sum(1 for step in session.steps if step.success)
            failed_steps = len(session.steps) - successful_steps
            
            tools_used = Counter(step.tool_name for step in session.steps)
            avg_execution_time = sum(
                step.execution_time for step in session.steps 
                if step.execution_time is not None
            ) / len(session.steps) if session.steps else 0
            
            return {
                'session_id': session_id,
                'file_path': session.file_path,
                'total_time': total_time,
                'total_steps': len(session.steps),
                'successful_steps': successful_steps,
                'failed_steps': failed_steps,
                'detected_pattern': session.detected_pattern.value,
                'current_context': session.context.value,
                'tools_used': dict(tools_used),
                'average_execution_time': avg_execution_time,
                'performance_metrics': session.performance_metrics
            }
            
        except Exception as e:
            logger.error(f"Failed to generate session summary: {e}")
            return {}
    
    def cleanup_session(self, session_id: str) -> None:
        """Clean up a workflow session."""
        if session_id in self.sessions:
            del self.sessions[session_id]
    
    def get_learning_insights(self) -> Dict[str, Any]:
        """Get insights from learning data."""
        try:
            insights = {
                'most_common_patterns': {},
                'most_successful_tools': {},
                'common_errors': {},
                'performance_insights': {}
            }
            
            # Analyze patterns
            pattern_counts = Counter()
            for key in self.learning_data.keys():
                if not key.startswith('error_'):
                    pattern_counts[key] += len(self.learning_data[key])
            
            insights['most_common_patterns'] = dict(pattern_counts.most_common(5))
            
            # Analyze tool success
            tool_success = Counter()
            for entries in self.learning_data.values():
                for entry in entries:
                    if 'tool' in entry:
                        tool_success[entry['tool']] += 1
            
            insights['most_successful_tools'] = dict(tool_success.most_common(5))
            
            # Analyze errors
            error_counts = Counter()
            for key, entries in self.learning_data.items():
                if key.startswith('error_'):
                    tool_name = key.replace('error_', '')
                    error_counts[tool_name] += len(entries)
            
            insights['common_errors'] = dict(error_counts.most_common(5))
            
            return insights
            
        except Exception as e:
            logger.error(f"Failed to generate learning insights: {e}")
            return {}


# Global workflow assistant instance
_workflow_assistant = None


def get_workflow_assistant() -> WorkflowAssistant:
    """Get the global workflow assistant instance."""
    global _workflow_assistant
    if _workflow_assistant is None:
        _workflow_assistant = WorkflowAssistant()
    return _workflow_assistant