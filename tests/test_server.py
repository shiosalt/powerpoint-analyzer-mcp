"""
Unit tests for PowerPoint MCP Server.

Tests the main server functionality and MCP protocol compliance.
"""

import pytest
import json
from unittest.mock import Mock, patch, AsyncMock, mock_open
from mcp import McpError

from powerpoint_mcp_server.server import PowerPointMCPServer


class TestPowerPointMCPServer:
    """Test cases for PowerPointMCPServer class."""
    
    def setup_method(self):
        """Set up test fixtures."""
        self.server = PowerPointMCPServer()
    
    def test_server_initialization(self):
        """Test that server initializes correctly."""
        assert self.server is not None
        assert hasattr(self.server, 'server')
        assert self.server.server.name == "powerpoint-mcp-server"
        assert hasattr(self.server, 'content_extractor')
        assert hasattr(self.server, 'attribute_processor')
        assert hasattr(self.server, 'file_validator')
    
    @pytest.mark.asyncio
    async def test_extract_powerpoint_content_missing_file_path(self):
        """Test extract_powerpoint_content with missing file_path."""
        arguments = {}
        
        with pytest.raises(McpError) as exc_info:
            await self.server._extract_powerpoint_content(arguments)
        
        assert "file_path is required" in str(exc_info.value)
    
    @pytest.mark.asyncio
    async def test_extract_powerpoint_content_invalid_file(self):
        """Test extract_powerpoint_content with invalid file."""
        arguments = {"file_path": "nonexistent.pptx"}
        
        # Mock file validator to return invalid
        with patch.object(self.server.file_validator, 'validate_file') as mock_validate:
            mock_validate.return_value = {
                'is_valid': False,
                'error': 'File not found'
            }
            
            with pytest.raises(McpError) as exc_info:
                await self.server._extract_powerpoint_content(arguments)
            
            assert "File validation failed" in str(exc_info.value)
    
    @pytest.mark.asyncio
    async def test_extract_powerpoint_content_success(self):
        """Test successful extract_powerpoint_content."""
        arguments = {"file_path": "test.pptx"}
        
        # Mock file validator
        with patch.object(self.server.file_validator, 'validate_file') as mock_validate:
            mock_validate.return_value = {'is_valid': True}
            
            # Mock the process_powerpoint_file method
            with patch.object(self.server, '_process_powerpoint_file') as mock_process:
                mock_process.return_value = {
                    'file_path': 'test.pptx',
                    'slides': [{'slide_number': 1, 'title': 'Test Slide'}],
                    'metadata': {'slide_count': 1}
                }
                
                result = await self.server._extract_powerpoint_content(arguments)
                
                assert result is not None
                assert hasattr(result, 'content')
                assert len(result.content) > 0
                
                # Parse the JSON response
                response_data = json.loads(result.content[0].text)
                assert response_data['file_path'] == 'test.pptx'
                assert len(response_data['slides']) == 1
                assert response_data['slides'][0]['title'] == 'Test Slide'
    
    @pytest.mark.asyncio
    async def test_get_powerpoint_attributes_missing_params(self):
        """Test get_powerpoint_attributes with missing parameters."""
        # Missing file_path
        arguments = {"attributes": ["title"]}
        
        with pytest.raises(McpError) as exc_info:
            await self.server._get_powerpoint_attributes(arguments)
        
        assert "file_path is required" in str(exc_info.value)
        
        # Missing attributes
        arguments = {"file_path": "test.pptx"}
        
        with pytest.raises(McpError) as exc_info:
            await self.server._get_powerpoint_attributes(arguments)
        
        assert "attributes list is required" in str(exc_info.value)
    
    @pytest.mark.asyncio
    async def test_get_powerpoint_attributes_success(self):
        """Test successful get_powerpoint_attributes."""
        arguments = {
            "file_path": "test.pptx",
            "attributes": ["title", "subtitle"]
        }
        
        # Mock file validator
        with patch.object(self.server.file_validator, 'validate_file') as mock_validate:
            mock_validate.return_value = {'is_valid': True}
            
            # Mock the process_powerpoint_file method
            with patch.object(self.server, '_process_powerpoint_file') as mock_process:
                mock_process.return_value = {
                    'slides': [
                        {'slide_number': 1, 'title': 'Test Title', 'subtitle': 'Test Subtitle', 'text_elements': []}
                    ]
                }
                
                # Mock attribute processor
                with patch.object(self.server.attribute_processor, 'filter_attributes') as mock_filter:
                    mock_filter.return_value = {
                        'slides': [
                            {'slide_number': 1, 'title': 'Test Title', 'subtitle': 'Test Subtitle'}
                        ]
                    }
                    
                    result = await self.server._get_powerpoint_attributes(arguments)
                    
                    assert result is not None
                    assert hasattr(result, 'content')
                    
                    # Parse the JSON response
                    response_data = json.loads(result.content[0].text)
                    assert 'slides' in response_data
                    assert response_data['slides'][0]['title'] == 'Test Title'
                    assert response_data['slides'][0]['subtitle'] == 'Test Subtitle'
    
    @pytest.mark.asyncio
    async def test_get_slide_info_missing_params(self):
        """Test get_slide_info with missing parameters."""
        # Missing file_path
        arguments = {"slide_number": 1}
        
        with pytest.raises(McpError) as exc_info:
            await self.server._get_slide_info(arguments)
        
        assert "file_path is required" in str(exc_info.value)
        
        # Missing slide_number
        arguments = {"file_path": "test.pptx"}
        
        with pytest.raises(McpError) as exc_info:
            await self.server._get_slide_info(arguments)
        
        assert "slide_number is required" in str(exc_info.value)
    
    @pytest.mark.asyncio
    async def test_get_slide_info_success(self):
        """Test successful get_slide_info."""
        arguments = {
            "file_path": "test.pptx",
            "slide_number": 1
        }
        
        # Mock file validator
        with patch.object(self.server.file_validator, 'validate_file') as mock_validate:
            mock_validate.return_value = {'is_valid': True}
            
            # Mock the process_single_slide method
            with patch.object(self.server, '_process_single_slide') as mock_process:
                mock_process.return_value = {
                    'slide_number': 1,
                    'title': 'Test Slide',
                    'subtitle': None,
                    'layout_name': 'Title Slide',
                    'text_elements': [],
                    'tables': [],
                    'object_counts': {'shapes': 1}
                }
                
                result = await self.server._get_slide_info(arguments)
                
                assert result is not None
                assert hasattr(result, 'content')
                
                # Parse the JSON response
                response_data = json.loads(result.content[0].text)
                assert response_data['slide_number'] == 1
                assert response_data['title'] == 'Test Slide'
                assert response_data['layout_name'] == 'Title Slide'
    
    @pytest.mark.asyncio
    async def test_process_single_slide_invalid_slide_number(self):
        """Test _process_single_slide with invalid slide number."""
        # Mock ZipExtractor
        with patch('powerpoint_mcp_server.server.ZipExtractor') as mock_extractor_class:
            mock_extractor = Mock()
            mock_extractor_class.return_value.__enter__.return_value = mock_extractor
            mock_extractor.get_slide_xml_files.return_value = ['slide1.xml', 'slide2.xml']
            
            with pytest.raises(ValueError) as exc_info:
                await self.server._process_single_slide("test.pptx", 5)
            
            assert "Slide number 5 is out of range (1-2)" in str(exc_info.value)
    
    def test_server_has_required_components(self):
        """Test that server has all required components."""
        assert hasattr(self.server, 'content_extractor')
        assert hasattr(self.server, 'attribute_processor')
        assert hasattr(self.server, 'file_validator')
        
        # Test that components are properly initialized
        assert self.server.content_extractor is not None
        assert self.server.attribute_processor is not None
        assert self.server.file_validator is not None