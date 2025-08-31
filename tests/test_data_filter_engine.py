"""
Unit tests for DataFilterEngine.
"""

import pytest
from unittest.mock import Mock

from powerpoint_mcp_server.core.data_filter_engine import (
    DataFilterEngine,
    FilterRule,
    FilterCondition,
    AggregationRule,
    AggregationOperation,
    GroupingRule,
    SortRule,
    SortOrder,
    FilterConfig,
    create_filter_config_from_dict
)


class TestDataFilterEngine:
    """Test cases for DataFilterEngine."""
    
    @pytest.fixture
    def filter_engine(self):
        """Create a DataFilterEngine instance."""
        return DataFilterEngine()
    
    @pytest.fixture
    def sample_data(self):
        """Create sample data for testing."""
        return [
            {
                "slide_number": 1,
                "title": "Introduction",
                "content": "Welcome to our presentation",
                "formatting": {"bold": True, "font_color": "#000000"},
                "word_count": 25,
                "category": "intro"
            },
            {
                "slide_number": 2,
                "title": "Project Overview",
                "content": "This project aims to improve efficiency",
                "formatting": {"italic": True, "font_color": "#FF0000"},
                "word_count": 35,
                "category": "overview"
            },
            {
                "slide_number": 3,
                "title": "Technical Details",
                "content": "Here are the technical specifications",
                "formatting": {"bold": True, "highlight": True},
                "word_count": 30,
                "category": "technical"
            },
            {
                "slide_number": 4,
                "title": "Conclusion",
                "content": "Thank you for your attention",
                "formatting": {},
                "word_count": 20,
                "category": "conclusion"
            }
        ]
    
    def test_create_simple_filter(self, filter_engine):
        """Test creating a simple filter rule."""
        filter_rule = filter_engine.create_simple_filter(
            field="title",
            condition="contains",
            value="Project",
            case_sensitive=False
        )
        
        assert filter_rule.field == "title"
        assert filter_rule.condition == FilterCondition.CONTAINS
        assert filter_rule.value == "Project"
        assert filter_rule.case_sensitive is False
    
    def test_create_formatting_filter(self, filter_engine):
        """Test creating a formatting-based filter rule."""
        formatting_criteria = {"bold": True, "font_color": "#000000"}
        filter_rule = filter_engine.create_formatting_filter(
            field="formatting",
            formatting_criteria=formatting_criteria
        )
        
        assert filter_rule.field == "formatting"
        assert filter_rule.condition == FilterCondition.HAS_FORMATTING
        assert filter_rule.formatting_criteria == formatting_criteria
    
    def test_create_aggregation_rule(self, filter_engine):
        """Test creating an aggregation rule."""
        agg_rule = filter_engine.create_aggregation_rule(
            field="word_count",
            operation="sum",
            output_field="total_words"
        )
        
        assert agg_rule.field == "word_count"
        assert agg_rule.operation == AggregationOperation.SUM
        assert agg_rule.output_field == "total_words"
    
    def test_create_sort_rule(self, filter_engine):
        """Test creating a sort rule."""
        sort_rule = filter_engine.create_sort_rule(field="slide_number", order="desc")
        
        assert sort_rule.field == "slide_number"
        assert sort_rule.order == SortOrder.DESC
    
    def test_get_nested_field_value(self, filter_engine, sample_data):
        """Test getting nested field values."""
        record = sample_data[0]
        
        # Simple field
        assert filter_engine._get_nested_field_value(record, "title") == "Introduction"
        
        # Nested field
        assert filter_engine._get_nested_field_value(record, "formatting.bold") is True
        
        # Non-existent field
        assert filter_engine._get_nested_field_value(record, "nonexistent") is None
        
        # Non-existent nested field
        assert filter_engine._get_nested_field_value(record, "formatting.nonexistent") is None
    
    def test_compare_values_equals(self, filter_engine):
        """Test value comparison for equals operation."""
        assert filter_engine._compare_values("Hello", "Hello", True, "equals") is True
        assert filter_engine._compare_values("Hello", "hello", False, "equals") is True
        assert filter_engine._compare_values("Hello", "hello", True, "equals") is False
        assert filter_engine._compare_values("Hello", "World", False, "equals") is False
    
    def test_compare_values_contains(self, filter_engine):
        """Test value comparison for contains operation."""
        assert filter_engine._compare_values("Hello World", "World", True, "contains") is True
        assert filter_engine._compare_values("Hello World", "world", False, "contains") is True
        assert filter_engine._compare_values("Hello World", "world", True, "contains") is False
        assert filter_engine._compare_values("Hello World", "xyz", False, "contains") is False
    
    def test_compare_values_starts_with(self, filter_engine):
        """Test value comparison for starts_with operation."""
        assert filter_engine._compare_values("Hello World", "Hello", True, "starts_with") is True
        assert filter_engine._compare_values("Hello World", "hello", False, "starts_with") is True
        assert filter_engine._compare_values("Hello World", "World", True, "starts_with") is False
    
    def test_compare_values_ends_with(self, filter_engine):
        """Test value comparison for ends_with operation."""
        assert filter_engine._compare_values("Hello World", "World", True, "ends_with") is True
        assert filter_engine._compare_values("Hello World", "world", False, "ends_with") is True
        assert filter_engine._compare_values("Hello World", "Hello", True, "ends_with") is False
    
    def test_evaluate_regex(self, filter_engine):
        """Test regex evaluation."""
        assert filter_engine._evaluate_regex("Hello World", r"H\w+", True) is True
        assert filter_engine._evaluate_regex("Hello World", r"h\w+", False) is True
        assert filter_engine._evaluate_regex("Hello World", r"h\w+", True) is False
        assert filter_engine._evaluate_regex("Hello World", r"\d+", False) is False
    
    def test_is_not_empty(self, filter_engine):
        """Test empty value checking."""
        assert filter_engine._is_not_empty("Hello") is True
        assert filter_engine._is_not_empty("") is False
        assert filter_engine._is_not_empty("   ") is False
        assert filter_engine._is_not_empty(None) is False
        assert filter_engine._is_not_empty([1, 2, 3]) is True
        assert filter_engine._is_not_empty([]) is False
        assert filter_engine._is_not_empty({"key": "value"}) is True
        assert filter_engine._is_not_empty({}) is False
    
    def test_has_formatting(self, filter_engine):
        """Test formatting detection."""
        # Test with any formatting
        formatting_dict = {"bold": True, "italic": False}
        assert filter_engine._has_formatting(formatting_dict, None) is True
        
        # Test with no formatting
        no_formatting_dict = {"bold": False, "italic": False}
        assert filter_engine._has_formatting(no_formatting_dict, None) is False
        
        # Test with specific criteria
        criteria = {"bold": True}
        assert filter_engine._has_formatting(formatting_dict, criteria) is True
        
        criteria = {"italic": True}
        assert filter_engine._has_formatting(formatting_dict, criteria) is False
        
        # Test with non-dict value
        assert filter_engine._has_formatting("not a dict", None) is False
    
    def test_compare_numeric(self, filter_engine):
        """Test numeric comparison."""
        assert filter_engine._compare_numeric(10, 5, ">") is True
        assert filter_engine._compare_numeric(5, 10, ">") is False
        assert filter_engine._compare_numeric(5, 10, "<") is True
        assert filter_engine._compare_numeric(10, 5, "<") is False
        assert filter_engine._compare_numeric(10, 10, ">=") is True
        assert filter_engine._compare_numeric(10, 10, "<=") is True
        
        # Test with string numbers
        assert filter_engine._compare_numeric("10", "5", ">") is True
        
        # Test with non-numeric values
        assert filter_engine._compare_numeric("abc", "5", ">") is False
    
    def test_in_list(self, filter_engine):
        """Test list membership checking."""
        test_list = ["apple", "banana", "cherry"]
        
        assert filter_engine._in_list("apple", test_list, True) is True
        assert filter_engine._in_list("Apple", test_list, False) is True
        assert filter_engine._in_list("Apple", test_list, True) is False
        assert filter_engine._in_list("grape", test_list, False) is False
        
        # Test with non-list value
        assert filter_engine._in_list("apple", "not a list", False) is False
    
    def test_evaluate_filter_equals(self, filter_engine, sample_data):
        """Test filter evaluation for equals condition."""
        filter_rule = FilterRule(
            field="category",
            condition=FilterCondition.EQUALS,
            value="intro"
        )
        
        assert filter_engine._evaluate_filter(sample_data[0], filter_rule) is True
        assert filter_engine._evaluate_filter(sample_data[1], filter_rule) is False
    
    def test_evaluate_filter_contains(self, filter_engine, sample_data):
        """Test filter evaluation for contains condition."""
        filter_rule = FilterRule(
            field="title",
            condition=FilterCondition.CONTAINS,
            value="Project"
        )
        
        assert filter_engine._evaluate_filter(sample_data[0], filter_rule) is False
        assert filter_engine._evaluate_filter(sample_data[1], filter_rule) is True
    
    def test_evaluate_filter_not_empty(self, filter_engine, sample_data):
        """Test filter evaluation for not_empty condition."""
        filter_rule = FilterRule(
            field="content",
            condition=FilterCondition.NOT_EMPTY
        )
        
        # All sample data should have non-empty content
        for record in sample_data:
            assert filter_engine._evaluate_filter(record, filter_rule) is True
    
    def test_evaluate_filter_has_formatting(self, filter_engine, sample_data):
        """Test filter evaluation for has_formatting condition."""
        filter_rule = FilterRule(
            field="formatting",
            condition=FilterCondition.HAS_FORMATTING,
            formatting_criteria={"bold": True}
        )
        
        assert filter_engine._evaluate_filter(sample_data[0], filter_rule) is True  # Has bold
        assert filter_engine._evaluate_filter(sample_data[1], filter_rule) is False  # No bold
        assert filter_engine._evaluate_filter(sample_data[2], filter_rule) is True  # Has bold
        assert filter_engine._evaluate_filter(sample_data[3], filter_rule) is False  # No formatting
    
    def test_evaluate_filter_greater_than(self, filter_engine, sample_data):
        """Test filter evaluation for greater_than condition."""
        filter_rule = FilterRule(
            field="word_count",
            condition=FilterCondition.GREATER_THAN,
            value=25
        )
        
        assert filter_engine._evaluate_filter(sample_data[0], filter_rule) is False  # 25 not > 25
        assert filter_engine._evaluate_filter(sample_data[1], filter_rule) is True   # 35 > 25
        assert filter_engine._evaluate_filter(sample_data[2], filter_rule) is True   # 30 > 25
        assert filter_engine._evaluate_filter(sample_data[3], filter_rule) is False  # 20 not > 25
    
    def test_evaluate_filter_in_list(self, filter_engine, sample_data):
        """Test filter evaluation for in_list condition."""
        filter_rule = FilterRule(
            field="category",
            condition=FilterCondition.IN_LIST,
            value=["intro", "conclusion"]
        )
        
        assert filter_engine._evaluate_filter(sample_data[0], filter_rule) is True   # intro
        assert filter_engine._evaluate_filter(sample_data[1], filter_rule) is False  # overview
        assert filter_engine._evaluate_filter(sample_data[2], filter_rule) is False  # technical
        assert filter_engine._evaluate_filter(sample_data[3], filter_rule) is True   # conclusion
    
    def test_apply_filters_and_logic(self, filter_engine, sample_data):
        """Test applying filters with AND logic."""
        filters = [
            FilterRule(field="word_count", condition=FilterCondition.GREATER_THAN, value=25),
            FilterRule(field="category", condition=FilterCondition.NOT_EQUALS, value="conclusion")
        ]
        
        filtered_data = filter_engine._apply_filters(sample_data, filters, "AND")
        
        # Should match records with word_count > 25 AND category != "conclusion"
        # That's records 1 and 2 (indices 1 and 2)
        assert len(filtered_data) == 2
        assert filtered_data[0]["slide_number"] == 2
        assert filtered_data[1]["slide_number"] == 3
    
    def test_apply_filters_or_logic(self, filter_engine, sample_data):
        """Test applying filters with OR logic."""
        filters = [
            FilterRule(field="category", condition=FilterCondition.EQUALS, value="intro"),
            FilterRule(field="category", condition=FilterCondition.EQUALS, value="conclusion")
        ]
        
        filtered_data = filter_engine._apply_filters(sample_data, filters, "OR")
        
        # Should match records with category = "intro" OR category = "conclusion"
        # That's records 0 and 3
        assert len(filtered_data) == 2
        assert filtered_data[0]["category"] == "intro"
        assert filtered_data[1]["category"] == "conclusion"
    
    def test_is_numeric(self, filter_engine):
        """Test numeric value detection."""
        assert filter_engine._is_numeric(42) is True
        assert filter_engine._is_numeric(3.14) is True
        assert filter_engine._is_numeric("42") is True
        assert filter_engine._is_numeric("3.14") is True
        assert filter_engine._is_numeric("abc") is False
        assert filter_engine._is_numeric(None) is False
    
    def test_apply_aggregation_count(self, filter_engine, sample_data):
        """Test count aggregation."""
        agg_rule = AggregationRule(
            field="slide_number",
            operation=AggregationOperation.COUNT,
            output_field="count"
        )
        
        result = filter_engine._apply_aggregation(sample_data, agg_rule)
        assert result == 4
    
    def test_apply_aggregation_sum(self, filter_engine, sample_data):
        """Test sum aggregation."""
        agg_rule = AggregationRule(
            field="word_count",
            operation=AggregationOperation.SUM,
            output_field="total_words"
        )
        
        result = filter_engine._apply_aggregation(sample_data, agg_rule)
        assert result == 110  # 25 + 35 + 30 + 20
    
    def test_apply_aggregation_average(self, filter_engine, sample_data):
        """Test average aggregation."""
        agg_rule = AggregationRule(
            field="word_count",
            operation=AggregationOperation.AVERAGE,
            output_field="avg_words"
        )
        
        result = filter_engine._apply_aggregation(sample_data, agg_rule)
        assert result == 27.5  # (25 + 35 + 30 + 20) / 4
    
    def test_apply_aggregation_min_max(self, filter_engine, sample_data):
        """Test min and max aggregation."""
        min_rule = AggregationRule(
            field="word_count",
            operation=AggregationOperation.MIN,
            output_field="min_words"
        )
        
        max_rule = AggregationRule(
            field="word_count",
            operation=AggregationOperation.MAX,
            output_field="max_words"
        )
        
        min_result = filter_engine._apply_aggregation(sample_data, min_rule)
        max_result = filter_engine._apply_aggregation(sample_data, max_rule)
        
        assert min_result == 20
        assert max_result == 35
    
    def test_apply_aggregation_concat(self, filter_engine, sample_data):
        """Test concat aggregation."""
        agg_rule = AggregationRule(
            field="category",
            operation=AggregationOperation.CONCAT,
            output_field="categories",
            separator=" | "
        )
        
        result = filter_engine._apply_aggregation(sample_data, agg_rule)
        assert result == "intro | overview | technical | conclusion"
    
    def test_apply_aggregation_unique(self, filter_engine):
        """Test unique aggregation."""
        data_with_duplicates = [
            {"category": "intro"},
            {"category": "overview"},
            {"category": "intro"},
            {"category": "technical"}
        ]
        
        agg_rule = AggregationRule(
            field="category",
            operation=AggregationOperation.UNIQUE,
            output_field="unique_categories"
        )
        
        result = filter_engine._apply_aggregation(data_with_duplicates, agg_rule)
        assert len(result) == 3
        assert "intro" in result
        assert "overview" in result
        assert "technical" in result
    
    def test_apply_aggregation_first_last(self, filter_engine, sample_data):
        """Test first and last aggregation."""
        first_rule = AggregationRule(
            field="title",
            operation=AggregationOperation.FIRST,
            output_field="first_title"
        )
        
        last_rule = AggregationRule(
            field="title",
            operation=AggregationOperation.LAST,
            output_field="last_title"
        )
        
        first_result = filter_engine._apply_aggregation(sample_data, first_rule)
        last_result = filter_engine._apply_aggregation(sample_data, last_rule)
        
        assert first_result == "Introduction"
        assert last_result == "Conclusion"
    
    def test_apply_sorting_single_field_asc(self, filter_engine, sample_data):
        """Test sorting by single field ascending."""
        sort_rules = [SortRule(field="word_count", order=SortOrder.ASC)]
        
        sorted_data = filter_engine._apply_sorting(sample_data, sort_rules)
        
        word_counts = [record["word_count"] for record in sorted_data]
        assert word_counts == [20, 25, 30, 35]
    
    def test_apply_sorting_single_field_desc(self, filter_engine, sample_data):
        """Test sorting by single field descending."""
        sort_rules = [SortRule(field="word_count", order=SortOrder.DESC)]
        
        sorted_data = filter_engine._apply_sorting(sample_data, sort_rules)
        
        word_counts = [record["word_count"] for record in sorted_data]
        assert word_counts == [35, 30, 25, 20]
    
    def test_apply_sorting_multiple_fields(self, filter_engine):
        """Test sorting by multiple fields."""
        # Create data with same category but different word counts
        data = [
            {"category": "intro", "word_count": 30},
            {"category": "intro", "word_count": 20},
            {"category": "overview", "word_count": 25},
            {"category": "overview", "word_count": 35}
        ]
        
        sort_rules = [
            SortRule(field="category", order=SortOrder.ASC),
            SortRule(field="word_count", order=SortOrder.DESC)
        ]
        
        sorted_data = filter_engine._apply_sorting(data, sort_rules)
        
        # Should be sorted by category ASC, then word_count DESC within each category
        expected_order = [
            ("intro", 30),
            ("intro", 20),
            ("overview", 35),
            ("overview", 25)
        ]
        
        actual_order = [(record["category"], record["word_count"]) for record in sorted_data]
        assert actual_order == expected_order
    
    def test_apply_pagination(self, filter_engine, sample_data):
        """Test pagination (offset and limit)."""
        # Test with offset only
        paginated = filter_engine._apply_pagination(sample_data, offset=1, limit=None)
        assert len(paginated) == 3
        assert paginated[0]["slide_number"] == 2
        
        # Test with limit only
        paginated = filter_engine._apply_pagination(sample_data, offset=0, limit=2)
        assert len(paginated) == 2
        assert paginated[0]["slide_number"] == 1
        assert paginated[1]["slide_number"] == 2
        
        # Test with both offset and limit
        paginated = filter_engine._apply_pagination(sample_data, offset=1, limit=2)
        assert len(paginated) == 2
        assert paginated[0]["slide_number"] == 2
        assert paginated[1]["slide_number"] == 3
    
    def test_apply_grouping_and_aggregation(self, filter_engine):
        """Test grouping and aggregation."""
        data = [
            {"category": "intro", "word_count": 25, "slide_number": 1},
            {"category": "intro", "word_count": 30, "slide_number": 2},
            {"category": "overview", "word_count": 35, "slide_number": 3},
            {"category": "overview", "word_count": 20, "slide_number": 4}
        ]
        
        grouping_rule = GroupingRule(
            fields=["category"],
            aggregations=[
                AggregationRule(
                    field="word_count",
                    operation=AggregationOperation.SUM,
                    output_field="total_words"
                ),
                AggregationRule(
                    field="slide_number",
                    operation=AggregationOperation.COUNT,
                    output_field="slide_count"
                )
            ]
        )
        
        result = filter_engine._apply_grouping_and_aggregation(data, grouping_rule)
        
        assert len(result) == 2
        
        # Find intro and overview groups
        intro_group = next(r for r in result if r["category"] == "intro")
        overview_group = next(r for r in result if r["category"] == "overview")
        
        assert intro_group["total_words"] == 55  # 25 + 30
        assert intro_group["slide_count"] == 2
        
        assert overview_group["total_words"] == 55  # 35 + 20
        assert overview_group["slide_count"] == 2
    
    def test_filter_and_aggregate_complete(self, filter_engine, sample_data):
        """Test complete filter and aggregate workflow."""
        filter_config = FilterConfig(
            filters=[
                FilterRule(
                    field="word_count",
                    condition=FilterCondition.GREATER_THAN,
                    value=20
                )
            ],
            filter_logic="AND",
            sorting=[
                SortRule(field="word_count", order=SortOrder.DESC)
            ],
            limit=2
        )
        
        result = filter_engine.filter_and_aggregate(sample_data, filter_config)
        
        assert "data" in result
        assert "summary" in result
        
        # Should have filtered out the record with word_count=20, sorted by word_count DESC, limited to 2
        data = result["data"]
        assert len(data) == 2
        assert data[0]["word_count"] == 35
        assert data[1]["word_count"] == 30
        
        # Check summary
        summary = result["summary"]
        assert summary["original_count"] == 4
        assert summary["filtered_count"] == 3  # Filtered out word_count=20
        assert summary["final_count"] == 2    # Limited to 2
        assert summary["filters_applied"] == 1
        assert summary["sorting_applied"] is True
    
    def test_cache_operations(self, filter_engine):
        """Test cache operations."""
        # Add something to cache
        filter_engine._filter_cache["test_key"] = "test_value"
        assert len(filter_engine._filter_cache) == 1
        
        # Clear cache
        filter_engine.clear_cache()
        assert len(filter_engine._filter_cache) == 0


class TestFilterConfigCreation:
    """Test cases for filter configuration creation."""
    
    def test_create_filter_config_from_dict_complete(self):
        """Test creating complete filter config from dictionary."""
        config_dict = {
            "filters": [
                {
                    "field": "title",
                    "condition": "contains",
                    "value": "Project",
                    "case_sensitive": False
                },
                {
                    "field": "formatting",
                    "condition": "has_formatting",
                    "formatting_criteria": {"bold": True}
                }
            ],
            "filter_logic": "OR",
            "grouping": {
                "fields": ["category"],
                "aggregations": [
                    {
                        "field": "word_count",
                        "operation": "sum",
                        "output_field": "total_words"
                    }
                ]
            },
            "sorting": [
                {
                    "field": "word_count",
                    "order": "desc"
                }
            ],
            "limit": 10,
            "offset": 5
        }
        
        config = create_filter_config_from_dict(config_dict)
        
        # Check filters
        assert len(config.filters) == 2
        assert config.filters[0].field == "title"
        assert config.filters[0].condition == FilterCondition.CONTAINS
        assert config.filters[0].value == "Project"
        assert config.filters[0].case_sensitive is False
        
        assert config.filters[1].field == "formatting"
        assert config.filters[1].condition == FilterCondition.HAS_FORMATTING
        assert config.filters[1].formatting_criteria == {"bold": True}
        
        # Check filter logic
        assert config.filter_logic == "OR"
        
        # Check grouping
        assert config.grouping is not None
        assert config.grouping.fields == ["category"]
        assert len(config.grouping.aggregations) == 1
        assert config.grouping.aggregations[0].field == "word_count"
        assert config.grouping.aggregations[0].operation == AggregationOperation.SUM
        assert config.grouping.aggregations[0].output_field == "total_words"
        
        # Check sorting
        assert len(config.sorting) == 1
        assert config.sorting[0].field == "word_count"
        assert config.sorting[0].order == SortOrder.DESC
        
        # Check pagination
        assert config.limit == 10
        assert config.offset == 5
    
    def test_create_filter_config_from_dict_minimal(self):
        """Test creating minimal filter config from dictionary."""
        config_dict = {}
        
        config = create_filter_config_from_dict(config_dict)
        
        assert len(config.filters) == 0
        assert config.filter_logic == "AND"
        assert config.grouping is None
        assert len(config.sorting) == 0
        assert config.limit is None
        assert config.offset == 0


if __name__ == "__main__":
    pytest.main([__file__])