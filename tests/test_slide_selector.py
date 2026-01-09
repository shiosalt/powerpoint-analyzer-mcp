"""Tests for the slide selector utility with Python-style slicing support."""

import pytest
from powerpoint_mcp_server.utils.slide_selector import parse_slide_numbers, validate_slide_numbers


class TestSlideSelector:
    """Test cases for slide selector utility."""

    def test_parse_none_returns_all_slides(self):
        """Test that None returns all slides."""
        result = parse_slide_numbers(None, 100)
        expected = list(range(1, 101))
        assert result == expected

    def test_parse_single_int(self):
        """Test parsing single integer."""
        result = parse_slide_numbers(3, 100)
        assert result == [3]

    def test_parse_list_of_ints(self):
        """Test parsing list of integers."""
        result = parse_slide_numbers([1, 5, 10], 100)
        assert result == [1, 5, 10]

    def test_parse_list_removes_duplicates(self):
        """Test that duplicate slide numbers are removed."""
        result = parse_slide_numbers([1, 5, 1, 10, 5], 100)
        assert result == [1, 5, 10]

    def test_parse_slice_from_start(self):
        """Test parsing slice from start (:10)."""
        result = parse_slide_numbers(":10", 100)
        expected = list(range(1, 11))
        assert result == expected

    def test_parse_slice_with_brackets(self):
        """Test parsing slice with brackets ([:10])."""
        result = parse_slide_numbers("[:10]", 100)
        expected = list(range(1, 11))
        assert result == expected

    def test_parse_slice_range(self):
        """Test parsing slice range (5:20)."""
        result = parse_slide_numbers("5:20", 100)
        expected = list(range(5, 21))
        assert result == expected

    def test_parse_slice_to_end(self):
        """Test parsing slice to end (25:)."""
        result = parse_slide_numbers("25:", 100)
        expected = list(range(25, 101))
        assert result == expected

    def test_parse_single_string_number(self):
        """Test parsing single number as string."""
        result = parse_slide_numbers("3", 100)
        assert result == [3]

    def test_parse_comma_separated(self):
        """Test parsing comma-separated numbers."""
        result = parse_slide_numbers("1,5,10", 100)
        assert result == [1, 5, 10]

    def test_parse_comma_separated_with_brackets(self):
        """Test parsing comma-separated numbers with brackets."""
        result = parse_slide_numbers("[1,5,10]", 100)
        assert result == [1, 5, 10]

    def test_parse_comma_separated_removes_duplicates(self):
        """Test that comma-separated duplicates are removed."""
        result = parse_slide_numbers("1,5,1,10,5", 100)
        assert result == [1, 5, 10]

    def test_invalid_slide_number_raises_error(self):
        """Test that invalid slide numbers raise ValueError."""
        with pytest.raises(ValueError, match="out of range"):
            parse_slide_numbers(101, 100)

    def test_invalid_list_slide_number_raises_error(self):
        """Test that invalid slide numbers in list raise ValueError."""
        with pytest.raises(ValueError, match="out of range"):
            parse_slide_numbers([1, 101], 100)

    def test_invalid_slice_start_raises_error(self):
        """Test that invalid slice start raises ValueError."""
        with pytest.raises(ValueError, match="beyond total slides"):
            parse_slide_numbers("101:110", 100)

    def test_slice_end_beyond_total_is_capped(self):
        """Test that slice end beyond total slides is capped."""
        result = parse_slide_numbers("95:110", 100)
        expected = list(range(95, 101))
        assert result == expected

    def test_invalid_slice_format_raises_error(self):
        """Test that invalid slice format raises ValueError."""
        with pytest.raises(ValueError, match="Invalid slice notation"):
            parse_slide_numbers("1:2:3", 100)

    def test_start_greater_than_end_raises_error(self):
        """Test that start > end raises ValueError."""
        with pytest.raises(ValueError, match="cannot be greater than"):
            parse_slide_numbers("20:10", 100)

    def test_invalid_type_raises_error(self):
        """Test that invalid type raises ValueError."""
        with pytest.raises(ValueError, match="Invalid slide specification type"):
            parse_slide_numbers({"invalid": "type"}, 100)

    def test_empty_string_raises_error(self):
        """Test that empty string raises ValueError."""
        with pytest.raises(ValueError, match="Invalid slide specification"):
            parse_slide_numbers("", 100)

    def test_validate_slide_numbers_filters_invalid(self):
        """Test that validate_slide_numbers filters out invalid numbers."""
        result = validate_slide_numbers([1, 5, 101, 200], 100)
        assert result == [1, 5]

    def test_validate_slide_numbers_empty_raises_error(self):
        """Test that validate_slide_numbers raises error for no valid slides."""
        with pytest.raises(ValueError, match="No valid slide numbers found"):
            validate_slide_numbers([101, 200], 100)


if __name__ == "__main__":
    pytest.main([__file__])