"""Integration tests for the new slide_numbers parameter formats."""

import pytest
import asyncio
from powerpoint_mcp_server.server import PowerPointMCPServer


class TestSlideNumbersIntegration:
    """Integration tests for slide_numbers parameter with various formats."""

    @pytest.fixture
    def server(self):
        """Create a PowerPoint MCP server instance."""
        return PowerPointMCPServer()

    @pytest.fixture
    def sample_file(self):
        """Path to a sample PowerPoint file for testing."""
        # This should be a real test file - for now using a placeholder
        return "tests/test_files/sample_presentation.pptx"

    @pytest.mark.asyncio
    async def test_extract_table_data_with_slice_notation(self, server, sample_file):
        """Test extract_table_data with slice notation."""
        try:
            # Test with slice notation ":5" (first 5 slides)
            arguments = {
                "file_path": sample_file,
                "slide_numbers": ":5",
                "output_format": "structured",
                "include_metadata": True
            }
            
            result = await server._extract_table_data(arguments)
            
            # Should return a result without errors
            assert result is not None
            assert hasattr(result, 'content')
            
        except FileNotFoundError:
            # Skip test if sample file doesn't exist
            pytest.skip(f"Sample file {sample_file} not found")
        except Exception as e:
            # For other errors, we want to see what happened
            pytest.fail(f"Unexpected error: {e}")

    @pytest.mark.asyncio
    async def test_extract_table_data_with_range_notation(self, server, sample_file):
        """Test extract_table_data with range notation."""
        try:
            # Test with range notation "2:4" (slides 2-4)
            arguments = {
                "file_path": sample_file,
                "slide_numbers": "2:4",
                "output_format": "structured",
                "include_metadata": True
            }
            
            result = await server._extract_table_data(arguments)
            
            # Should return a result without errors
            assert result is not None
            assert hasattr(result, 'content')
            
        except FileNotFoundError:
            # Skip test if sample file doesn't exist
            pytest.skip(f"Sample file {sample_file} not found")
        except Exception as e:
            # For other errors, we want to see what happened
            pytest.fail(f"Unexpected error: {e}")

    @pytest.mark.asyncio
    async def test_extract_table_data_with_comma_separated(self, server, sample_file):
        """Test extract_table_data with comma-separated notation."""
        try:
            # Test with comma-separated notation "1,3,5"
            arguments = {
                "file_path": sample_file,
                "slide_numbers": "1,3,5",
                "output_format": "structured",
                "include_metadata": True
            }
            
            result = await server._extract_table_data(arguments)
            
            # Should return a result without errors
            assert result is not None
            assert hasattr(result, 'content')
            
        except FileNotFoundError:
            # Skip test if sample file doesn't exist
            pytest.skip(f"Sample file {sample_file} not found")
        except Exception as e:
            # For other errors, we want to see what happened
            pytest.fail(f"Unexpected error: {e}")

    @pytest.mark.asyncio
    async def test_extract_formatted_text_with_slice_notation(self, server, sample_file):
        """Test extract_formatted_text with slice notation."""
        try:
            # Test with slice notation ":3" (first 3 slides)
            arguments = {
                "file_path": sample_file,
                "formatting_type": "bold",
                "slide_numbers": ":3"
            }
            
            result = await server._extract_text_formatting(arguments)
            
            # Should return a result without errors
            assert result is not None
            assert hasattr(result, 'content')
            
        except FileNotFoundError:
            # Skip test if sample file doesn't exist
            pytest.skip(f"Sample file {sample_file} not found")
        except Exception as e:
            # For other errors, we want to see what happened
            pytest.fail(f"Unexpected error: {e}")

    @pytest.mark.asyncio
    async def test_query_slides_with_slice_notation(self, server, sample_file):
        """Test query_slides with slice notation in search criteria."""
        try:
            # Test with slice notation in search criteria
            arguments = {
                "file_path": sample_file,
                "search_criteria": {
                    "slide_numbers": ":5"
                },
                "return_fields": ["slide_number", "title"],
                "limit": 50
            }
            
            result = await server._query_slides(arguments)
            
            # Should return a result without errors
            assert result is not None
            assert hasattr(result, 'content')
            
        except FileNotFoundError:
            # Skip test if sample file doesn't exist
            pytest.skip(f"Sample file {sample_file} not found")
        except Exception as e:
            # For other errors, we want to see what happened
            pytest.fail(f"Unexpected error: {e}")

    def test_slide_numbers_validation_accepts_new_formats(self):
        """Test that slide_numbers validation accepts new formats."""
        from powerpoint_mcp_server.core.slide_query_engine import SlideQueryFilters, SlideQueryEngine
        
        engine = SlideQueryEngine()
        
        # Test with integer
        filters = SlideQueryFilters(slide_numbers=3)
        result = engine._validate_search_criteria(filters, ["slide_number"])
        assert result['is_valid'] == True
        
        # Test with string slice
        filters = SlideQueryFilters(slide_numbers=":10")
        result = engine._validate_search_criteria(filters, ["slide_number"])
        assert result['is_valid'] == True
        
        # Test with comma-separated string
        filters = SlideQueryFilters(slide_numbers="1,3,5")
        result = engine._validate_search_criteria(filters, ["slide_number"])
        assert result['is_valid'] == True
        
        # Test with list (existing format)
        filters = SlideQueryFilters(slide_numbers=[1, 3, 5])
        result = engine._validate_search_criteria(filters, ["slide_number"])
        assert result['is_valid'] == True

    def test_slide_numbers_validation_rejects_invalid_formats(self):
        """Test that slide_numbers validation rejects invalid formats."""
        from powerpoint_mcp_server.core.slide_query_engine import SlideQueryFilters, SlideQueryEngine
        
        engine = SlideQueryEngine()
        
        # Test with invalid integer
        filters = SlideQueryFilters(slide_numbers=0)
        result = engine._validate_search_criteria(filters, ["slide_number"])
        assert result['is_valid'] == False
        
        # Test with empty string
        filters = SlideQueryFilters(slide_numbers="")
        result = engine._validate_search_criteria(filters, ["slide_number"])
        assert result['is_valid'] == False
        
        # Test with invalid type
        filters = SlideQueryFilters(slide_numbers={"invalid": "type"})
        result = engine._validate_search_criteria(filters, ["slide_number"])
        assert result['is_valid'] == False


if __name__ == "__main__":
    pytest.main([__file__, "-v"])