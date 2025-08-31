"""
Data filtering and aggregation engine for post-processing extracted data.
"""

import re
import logging
from typing import Dict, List, Any, Optional, Union, Tuple, Callable
from dataclasses import dataclass, field
from enum import Enum
from collections import defaultdict
import statistics

logger = logging.getLogger(__name__)


class FilterCondition(Enum):
    """Enumeration of available filter conditions."""
    EQUALS = "equals"
    NOT_EQUALS = "not_equals"
    CONTAINS = "contains"
    NOT_CONTAINS = "not_contains"
    STARTS_WITH = "starts_with"
    ENDS_WITH = "ends_with"
    REGEX = "regex"
    NOT_EMPTY = "not_empty"
    IS_EMPTY = "is_empty"
    HAS_FORMATTING = "has_formatting"
    NO_FORMATTING = "no_formatting"
    GREATER_THAN = "greater_than"
    LESS_THAN = "less_than"
    GREATER_EQUAL = "greater_equal"
    LESS_EQUAL = "less_equal"
    IN_LIST = "in_list"
    NOT_IN_LIST = "not_in_list"


class AggregationOperation(Enum):
    """Enumeration of available aggregation operations."""
    COUNT = "count"
    LIST = "list"
    UNIQUE = "unique"
    CONCAT = "concat"
    SUM = "sum"
    AVERAGE = "average"
    MIN = "min"
    MAX = "max"
    FIRST = "first"
    LAST = "last"
    MOST_COMMON = "most_common"
    LEAST_COMMON = "least_common"


class SortOrder(Enum):
    """Enumeration of sort orders."""
    ASC = "asc"
    DESC = "desc"


@dataclass
class FilterRule:
    """A single filter rule."""
    field: str
    condition: FilterCondition
    value: Any = None
    formatting_criteria: Optional[Dict[str, Any]] = None
    case_sensitive: bool = False


@dataclass
class AggregationRule:
    """A single aggregation rule."""
    field: str
    operation: AggregationOperation
    output_field: str
    separator: str = ", "  # For concat operations


@dataclass
class GroupingRule:
    """A grouping rule for data aggregation."""
    fields: List[str]
    aggregations: List[AggregationRule]


@dataclass
class SortRule:
    """A sorting rule."""
    field: str
    order: SortOrder = SortOrder.ASC


@dataclass
class FilterConfig:
    """Complete filter configuration."""
    filters: List[FilterRule] = field(default_factory=list)
    filter_logic: str = "AND"  # "AND" or "OR"
    grouping: Optional[GroupingRule] = None
    sorting: List[SortRule] = field(default_factory=list)
    limit: Optional[int] = None
    offset: int = 0


class DataFilterEngine:
    """
    Engine for filtering and aggregating extracted data.
    """
    
    def __init__(self):
        """Initialize the data filter engine."""
        self._filter_cache = {}
    
    def filter_and_aggregate(
        self,
        data: List[Dict[str, Any]],
        filter_config: FilterConfig
    ) -> Dict[str, Any]:
        """
        Apply filters and aggregations to data.
        
        Args:
            data: List of data records to filter and aggregate
            filter_config: Configuration for filtering and aggregation
            
        Returns:
            Dictionary containing filtered and aggregated results
        """
        logger.info(f"Filtering and aggregating {len(data)} records")
        
        try:
            # Apply filters
            filtered_data = self._apply_filters(data, filter_config.filters, filter_config.filter_logic)
            
            # Apply grouping and aggregation
            if filter_config.grouping:
                aggregated_data = self._apply_grouping_and_aggregation(filtered_data, filter_config.grouping)
            else:
                aggregated_data = filtered_data
            
            # Apply sorting
            if filter_config.sorting:
                aggregated_data = self._apply_sorting(aggregated_data, filter_config.sorting)
            
            # Apply limit and offset
            if filter_config.offset > 0 or filter_config.limit is not None:
                aggregated_data = self._apply_pagination(
                    aggregated_data, filter_config.offset, filter_config.limit
                )
            
            result = {
                "data": aggregated_data,
                "summary": {
                    "original_count": len(data),
                    "filtered_count": len(filtered_data),
                    "final_count": len(aggregated_data),
                    "filters_applied": len(filter_config.filters),
                    "grouping_applied": filter_config.grouping is not None,
                    "sorting_applied": len(filter_config.sorting) > 0
                }
            }
            
            logger.info(f"Filtering complete: {len(data)} -> {len(aggregated_data)} records")
            return result
            
        except Exception as e:
            logger.error(f"Error filtering and aggregating data: {e}")
            raise
    
    def _apply_filters(
        self,
        data: List[Dict[str, Any]],
        filters: List[FilterRule],
        filter_logic: str
    ) -> List[Dict[str, Any]]:
        """Apply filter rules to data."""
        if not filters:
            return data
        
        try:
            filtered_data = []
            
            for record in data:
                if filter_logic.upper() == "AND":
                    # All filters must pass
                    if all(self._evaluate_filter(record, filter_rule) for filter_rule in filters):
                        filtered_data.append(record)
                else:  # OR logic
                    # At least one filter must pass
                    if any(self._evaluate_filter(record, filter_rule) for filter_rule in filters):
                        filtered_data.append(record)
            
            return filtered_data
            
        except Exception as e:
            logger.warning(f"Failed to apply filters: {e}")
            return data
    
    def _evaluate_filter(self, record: Dict[str, Any], filter_rule: FilterRule) -> bool:
        """Evaluate a single filter rule against a record."""
        try:
            field_value = self._get_nested_field_value(record, filter_rule.field)
            
            if filter_rule.condition == FilterCondition.EQUALS:
                return self._compare_values(field_value, filter_rule.value, filter_rule.case_sensitive, "equals")
            
            elif filter_rule.condition == FilterCondition.NOT_EQUALS:
                return not self._compare_values(field_value, filter_rule.value, filter_rule.case_sensitive, "equals")
            
            elif filter_rule.condition == FilterCondition.CONTAINS:
                return self._compare_values(field_value, filter_rule.value, filter_rule.case_sensitive, "contains")
            
            elif filter_rule.condition == FilterCondition.NOT_CONTAINS:
                return not self._compare_values(field_value, filter_rule.value, filter_rule.case_sensitive, "contains")
            
            elif filter_rule.condition == FilterCondition.STARTS_WITH:
                return self._compare_values(field_value, filter_rule.value, filter_rule.case_sensitive, "starts_with")
            
            elif filter_rule.condition == FilterCondition.ENDS_WITH:
                return self._compare_values(field_value, filter_rule.value, filter_rule.case_sensitive, "ends_with")
            
            elif filter_rule.condition == FilterCondition.REGEX:
                return self._evaluate_regex(field_value, filter_rule.value, filter_rule.case_sensitive)
            
            elif filter_rule.condition == FilterCondition.NOT_EMPTY:
                return self._is_not_empty(field_value)
            
            elif filter_rule.condition == FilterCondition.IS_EMPTY:
                return not self._is_not_empty(field_value)
            
            elif filter_rule.condition == FilterCondition.HAS_FORMATTING:
                return self._has_formatting(field_value, filter_rule.formatting_criteria)
            
            elif filter_rule.condition == FilterCondition.NO_FORMATTING:
                return not self._has_formatting(field_value, filter_rule.formatting_criteria)
            
            elif filter_rule.condition == FilterCondition.GREATER_THAN:
                return self._compare_numeric(field_value, filter_rule.value, ">")
            
            elif filter_rule.condition == FilterCondition.LESS_THAN:
                return self._compare_numeric(field_value, filter_rule.value, "<")
            
            elif filter_rule.condition == FilterCondition.GREATER_EQUAL:
                return self._compare_numeric(field_value, filter_rule.value, ">=")
            
            elif filter_rule.condition == FilterCondition.LESS_EQUAL:
                return self._compare_numeric(field_value, filter_rule.value, "<=")
            
            elif filter_rule.condition == FilterCondition.IN_LIST:
                return self._in_list(field_value, filter_rule.value, filter_rule.case_sensitive)
            
            elif filter_rule.condition == FilterCondition.NOT_IN_LIST:
                return not self._in_list(field_value, filter_rule.value, filter_rule.case_sensitive)
            
            else:
                logger.warning(f"Unknown filter condition: {filter_rule.condition}")
                return True
                
        except Exception as e:
            logger.warning(f"Failed to evaluate filter rule: {e}")
            return True  # Default to including the record
    
    def _get_nested_field_value(self, record: Dict[str, Any], field_path: str) -> Any:
        """Get value from nested field path (e.g., 'formatting.bold')."""
        try:
            value = record
            for field_part in field_path.split('.'):
                if isinstance(value, dict):
                    value = value.get(field_part)
                elif isinstance(value, list) and field_part.isdigit():
                    index = int(field_part)
                    value = value[index] if 0 <= index < len(value) else None
                else:
                    return None
            return value
        except Exception:
            return None
    
    def _compare_values(self, field_value: Any, filter_value: Any, case_sensitive: bool, operation: str) -> bool:
        """Compare two values based on the operation."""
        try:
            # Convert to strings for text operations
            field_str = str(field_value) if field_value is not None else ""
            filter_str = str(filter_value) if filter_value is not None else ""
            
            if not case_sensitive:
                field_str = field_str.lower()
                filter_str = filter_str.lower()
            
            if operation == "equals":
                return field_str == filter_str
            elif operation == "contains":
                return filter_str in field_str
            elif operation == "starts_with":
                return field_str.startswith(filter_str)
            elif operation == "ends_with":
                return field_str.endswith(filter_str)
            
            return False
            
        except Exception:
            return False
    
    def _evaluate_regex(self, field_value: Any, pattern: str, case_sensitive: bool) -> bool:
        """Evaluate regex pattern against field value."""
        try:
            field_str = str(field_value) if field_value is not None else ""
            flags = 0 if case_sensitive else re.IGNORECASE
            return bool(re.search(pattern, field_str, flags))
        except re.error:
            logger.warning(f"Invalid regex pattern: {pattern}")
            return False
        except Exception:
            return False
    
    def _is_not_empty(self, value: Any) -> bool:
        """Check if value is not empty."""
        if value is None:
            return False
        if isinstance(value, str):
            return bool(value.strip())
        if isinstance(value, (list, dict)):
            return len(value) > 0
        return True
    
    def _has_formatting(self, field_value: Any, formatting_criteria: Optional[Dict[str, Any]]) -> bool:
        """Check if field has formatting based on criteria."""
        try:
            if not isinstance(field_value, dict):
                return False
            
            # If no specific criteria, check for any formatting
            if not formatting_criteria:
                formatting_fields = ['bold', 'italic', 'underline', 'highlight', 'strikethrough', 
                                   'font_color', 'background_color', 'hyperlink']
                return any(field_value.get(field) for field in formatting_fields)
            
            # Check specific formatting criteria
            for criteria_key, criteria_value in formatting_criteria.items():
                field_formatting_value = field_value.get(criteria_key)
                
                if criteria_value is True:
                    # Must have this formatting
                    if not field_formatting_value:
                        return False
                elif criteria_value is False:
                    # Must not have this formatting
                    if field_formatting_value:
                        return False
                else:
                    # Must match specific value
                    if field_formatting_value != criteria_value:
                        return False
            
            return True
            
        except Exception:
            return False
    
    def _compare_numeric(self, field_value: Any, filter_value: Any, operator: str) -> bool:
        """Compare numeric values."""
        try:
            field_num = float(field_value) if field_value is not None else 0
            filter_num = float(filter_value) if filter_value is not None else 0
            
            if operator == ">":
                return field_num > filter_num
            elif operator == "<":
                return field_num < filter_num
            elif operator == ">=":
                return field_num >= filter_num
            elif operator == "<=":
                return field_num <= filter_num
            
            return False
            
        except (ValueError, TypeError):
            return False
    
    def _in_list(self, field_value: Any, value_list: List[Any], case_sensitive: bool) -> bool:
        """Check if field value is in the provided list."""
        try:
            if not isinstance(value_list, list):
                return False
            
            field_str = str(field_value) if field_value is not None else ""
            
            for list_item in value_list:
                list_str = str(list_item) if list_item is not None else ""
                
                if case_sensitive:
                    if field_str == list_str:
                        return True
                else:
                    if field_str.lower() == list_str.lower():
                        return True
            
            return False
            
        except Exception:
            return False
    
    def _apply_grouping_and_aggregation(
        self,
        data: List[Dict[str, Any]],
        grouping_rule: GroupingRule
    ) -> List[Dict[str, Any]]:
        """Apply grouping and aggregation to data."""
        try:
            # Group data by specified fields
            groups = defaultdict(list)
            
            for record in data:
                # Create group key from grouping fields
                group_key_parts = []
                for field in grouping_rule.fields:
                    value = self._get_nested_field_value(record, field)
                    group_key_parts.append(str(value) if value is not None else "")
                
                group_key = "|".join(group_key_parts)
                groups[group_key].append(record)
            
            # Apply aggregations to each group
            aggregated_data = []
            
            for group_key, group_records in groups.items():
                aggregated_record = {}
                
                # Add grouping fields to the result
                group_key_parts = group_key.split("|")
                for i, field in enumerate(grouping_rule.fields):
                    if i < len(group_key_parts):
                        aggregated_record[field] = group_key_parts[i] if group_key_parts[i] else None
                
                # Apply aggregation operations
                for agg_rule in grouping_rule.aggregations:
                    aggregated_value = self._apply_aggregation(group_records, agg_rule)
                    aggregated_record[agg_rule.output_field] = aggregated_value
                
                aggregated_data.append(aggregated_record)
            
            return aggregated_data
            
        except Exception as e:
            logger.warning(f"Failed to apply grouping and aggregation: {e}")
            return data
    
    def _apply_aggregation(self, records: List[Dict[str, Any]], agg_rule: AggregationRule) -> Any:
        """Apply a single aggregation operation to a group of records."""
        try:
            # Extract field values from all records
            values = []
            for record in records:
                value = self._get_nested_field_value(record, agg_rule.field)
                if value is not None:
                    values.append(value)
            
            if not values:
                return None
            
            if agg_rule.operation == AggregationOperation.COUNT:
                return len(values)
            
            elif agg_rule.operation == AggregationOperation.LIST:
                return values
            
            elif agg_rule.operation == AggregationOperation.UNIQUE:
                return list(set(str(v) for v in values))
            
            elif agg_rule.operation == AggregationOperation.CONCAT:
                str_values = [str(v) for v in values if str(v).strip()]
                return agg_rule.separator.join(str_values)
            
            elif agg_rule.operation == AggregationOperation.SUM:
                numeric_values = [float(v) for v in values if self._is_numeric(v)]
                return sum(numeric_values) if numeric_values else 0
            
            elif agg_rule.operation == AggregationOperation.AVERAGE:
                numeric_values = [float(v) for v in values if self._is_numeric(v)]
                return statistics.mean(numeric_values) if numeric_values else 0
            
            elif agg_rule.operation == AggregationOperation.MIN:
                numeric_values = [float(v) for v in values if self._is_numeric(v)]
                return min(numeric_values) if numeric_values else None
            
            elif agg_rule.operation == AggregationOperation.MAX:
                numeric_values = [float(v) for v in values if self._is_numeric(v)]
                return max(numeric_values) if numeric_values else None
            
            elif agg_rule.operation == AggregationOperation.FIRST:
                return values[0]
            
            elif agg_rule.operation == AggregationOperation.LAST:
                return values[-1]
            
            elif agg_rule.operation == AggregationOperation.MOST_COMMON:
                from collections import Counter
                counter = Counter(str(v) for v in values)
                return counter.most_common(1)[0][0] if counter else None
            
            elif agg_rule.operation == AggregationOperation.LEAST_COMMON:
                from collections import Counter
                counter = Counter(str(v) for v in values)
                return counter.most_common()[-1][0] if counter else None
            
            else:
                logger.warning(f"Unknown aggregation operation: {agg_rule.operation}")
                return None
                
        except Exception as e:
            logger.warning(f"Failed to apply aggregation {agg_rule.operation}: {e}")
            return None
    
    def _is_numeric(self, value: Any) -> bool:
        """Check if value can be converted to a number."""
        try:
            float(value)
            return True
        except (ValueError, TypeError):
            return False
    
    def _apply_sorting(self, data: List[Dict[str, Any]], sort_rules: List[SortRule]) -> List[Dict[str, Any]]:
        """Apply sorting rules to data."""
        try:
            if not sort_rules:
                return data
            
            # Create a sorting key function
            def sort_key(record):
                key_values = []
                for sort_rule in sort_rules:
                    value = self._get_nested_field_value(record, sort_rule.field)
                    
                    # Handle None values
                    if value is None:
                        value = ""
                    
                    # Convert to comparable type
                    if self._is_numeric(value):
                        value = float(value)
                    else:
                        value = str(value).lower()
                    
                    # Apply reverse for descending order
                    if sort_rule.order == SortOrder.DESC:
                        if isinstance(value, (int, float)):
                            value = -value
                        else:
                            # For strings, we'll handle DESC in the sorted() call
                            pass
                    
                    key_values.append(value)
                
                return key_values
            
            # Sort the data
            sorted_data = sorted(data, key=sort_key)
            
            # Handle descending order for string fields
            # (This is a simplified approach; a more robust solution would handle mixed types better)
            if any(rule.order == SortOrder.DESC for rule in sort_rules):
                # For simplicity, if any field is DESC, we'll use a more complex sorting approach
                def complex_sort_key(record):
                    key_values = []
                    for sort_rule in sort_rules:
                        value = self._get_nested_field_value(record, sort_rule.field)
                        
                        if value is None:
                            value = ""
                        
                        if self._is_numeric(value):
                            value = float(value)
                            if sort_rule.order == SortOrder.DESC:
                                value = -value
                        else:
                            value = str(value).lower()
                        
                        key_values.append((value, sort_rule.order == SortOrder.DESC))
                    
                    return key_values
                
                # Custom sorting with mixed ASC/DESC
                def compare_keys(key1, key2):
                    for (val1, desc1), (val2, desc2) in zip(key1, key2):
                        if val1 == val2:
                            continue
                        
                        if isinstance(val1, str) and isinstance(val2, str):
                            result = -1 if val1 < val2 else 1
                        else:
                            result = -1 if val1 < val2 else 1
                        
                        return -result if desc1 else result
                    
                    return 0
                
                # Use the complex key for sorting
                from functools import cmp_to_key
                sorted_data = sorted(data, key=lambda x: [
                    (self._get_nested_field_value(x, rule.field), rule.order == SortOrder.DESC)
                    for rule in sort_rules
                ])
            
            return sorted_data
            
        except Exception as e:
            logger.warning(f"Failed to apply sorting: {e}")
            return data
    
    def _apply_pagination(self, data: List[Dict[str, Any]], offset: int, limit: Optional[int]) -> List[Dict[str, Any]]:
        """Apply pagination (offset and limit) to data."""
        try:
            start_index = max(0, offset)
            
            if limit is not None:
                end_index = start_index + limit
                return data[start_index:end_index]
            else:
                return data[start_index:]
                
        except Exception as e:
            logger.warning(f"Failed to apply pagination: {e}")
            return data
    
    def create_simple_filter(
        self,
        field: str,
        condition: str,
        value: Any,
        case_sensitive: bool = False
    ) -> FilterRule:
        """Create a simple filter rule."""
        try:
            condition_enum = FilterCondition(condition)
            return FilterRule(
                field=field,
                condition=condition_enum,
                value=value,
                case_sensitive=case_sensitive
            )
        except ValueError:
            raise ValueError(f"Invalid filter condition: {condition}")
    
    def create_formatting_filter(
        self,
        field: str,
        formatting_criteria: Dict[str, Any]
    ) -> FilterRule:
        """Create a formatting-based filter rule."""
        return FilterRule(
            field=field,
            condition=FilterCondition.HAS_FORMATTING,
            formatting_criteria=formatting_criteria
        )
    
    def create_aggregation_rule(
        self,
        field: str,
        operation: str,
        output_field: str,
        separator: str = ", "
    ) -> AggregationRule:
        """Create an aggregation rule."""
        try:
            operation_enum = AggregationOperation(operation)
            return AggregationRule(
                field=field,
                operation=operation_enum,
                output_field=output_field,
                separator=separator
            )
        except ValueError:
            raise ValueError(f"Invalid aggregation operation: {operation}")
    
    def create_sort_rule(self, field: str, order: str = "asc") -> SortRule:
        """Create a sort rule."""
        try:
            order_enum = SortOrder(order)
            return SortRule(field=field, order=order_enum)
        except ValueError:
            raise ValueError(f"Invalid sort order: {order}")
    
    def clear_cache(self):
        """Clear the filter cache."""
        self._filter_cache.clear()
        logger.debug("Data filter cache cleared")


def create_filter_config_from_dict(config_dict: Dict[str, Any]) -> FilterConfig:
    """Create FilterConfig from a dictionary representation."""
    config = FilterConfig()
    
    # Parse filters
    if 'filters' in config_dict:
        filters = []
        for filter_dict in config_dict['filters']:
            filter_rule = FilterRule(
                field=filter_dict['field'],
                condition=FilterCondition(filter_dict['condition']),
                value=filter_dict.get('value'),
                formatting_criteria=filter_dict.get('formatting_criteria'),
                case_sensitive=filter_dict.get('case_sensitive', False)
            )
            filters.append(filter_rule)
        config.filters = filters
    
    # Parse filter logic
    config.filter_logic = config_dict.get('filter_logic', 'AND')
    
    # Parse grouping
    if 'grouping' in config_dict:
        grouping_dict = config_dict['grouping']
        aggregations = []
        
        for agg_dict in grouping_dict.get('aggregations', []):
            agg_rule = AggregationRule(
                field=agg_dict['field'],
                operation=AggregationOperation(agg_dict['operation']),
                output_field=agg_dict['output_field'],
                separator=agg_dict.get('separator', ', ')
            )
            aggregations.append(agg_rule)
        
        config.grouping = GroupingRule(
            fields=grouping_dict['fields'],
            aggregations=aggregations
        )
    
    # Parse sorting
    if 'sorting' in config_dict:
        sort_rules = []
        for sort_dict in config_dict['sorting']:
            sort_rule = SortRule(
                field=sort_dict['field'],
                order=SortOrder(sort_dict.get('order', 'asc'))
            )
            sort_rules.append(sort_rule)
        config.sorting = sort_rules
    
    # Parse pagination
    config.limit = config_dict.get('limit')
    config.offset = config_dict.get('offset', 0)
    
    return config