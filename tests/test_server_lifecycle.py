"""Integration tests for server startup and shutdown procedures."""

import asyncio
import logging
import os
import pytest
import signal
import sys
import tempfile
from unittest.mock import Mock, patch, AsyncMock
from pathlib import Path

# Add the parent directory to the path so we can import the server
sys.path.insert(0, str(Path(__file__).parent.parent))

from powerpoint_mcp_server.server import PowerPointMCPServer
from powerpoint_mcp_server.config import ServerConfig, ConfigManager
from main import ServerApplication


class TestServerLifecycle:
    """Test server startup and shutdown procedures."""
    
    @pytest.fixture
    def server_config(self):
        """Create a test server configuration."""
        return ServerConfig(
            log_level='DEBUG',
            server_name='test-powerpoint-mcp-server',
            server_version='0.1.0-test',
            max_file_size_mb=50,
            processing_timeout_seconds=60,
            cache_enabled=True,
            cache_ttl_seconds=1800,
            debug_mode=True
        )
    
    @pytest.fixture
    def config_manager(self, server_config):
        """Create a test configuration manager."""
        return ConfigManager(server_config)
    
    def test_server_config_creation(self, server_config):
        """Test server configuration creation and validation."""
        assert server_config.server_name == 'test-powerpoint-mcp-server'
        assert server_config.server_version == '0.1.0-test'
        assert server_config.log_level == 'DEBUG'
        assert server_config.max_file_size_mb == 50
        assert server_config.debug_mode is True
    
    def test_server_config_validation(self):
        """Test server configuration validation."""
        # Test invalid log level
        config = ServerConfig(log_level='INVALID')
        assert config.log_level == 'INFO'  # Should default to INFO
        
        # Test invalid file size
        config = ServerConfig(max_file_size_mb=-10)
        assert config.max_file_size_mb == 100  # Should default to 100
        
        # Test invalid timeout
        config = ServerConfig(processing_timeout_seconds=0)
        assert config.processing_timeout_seconds == 300  # Should default to 300
    
    def test_server_config_from_env(self):
        """Test server configuration from environment variables."""
        with patch.dict(os.environ, {
            'POWERPOINT_MCP_LOG_LEVEL': 'WARNING',
            'POWERPOINT_MCP_MAX_FILE_SIZE': '200',
            'POWERPOINT_MCP_TIMEOUT': '600',
            'POWERPOINT_MCP_CACHE_ENABLED': 'false',
            'POWERPOINT_MCP_DEBUG': 'true'
        }):
            config = ServerConfig.from_env()
            assert config.log_level == 'WARNING'
            assert config.max_file_size_mb == 200
            assert config.processing_timeout_seconds == 600
            assert config.cache_enabled is False
            assert config.debug_mode is True
    
    def test_config_manager_update(self, config_manager):
        """Test configuration manager update functionality."""
        config_manager.update_config(log_level='ERROR', max_file_size_mb=75)
        
        config = config_manager.get_config()
        assert config.log_level == 'ERROR'
        assert config.max_file_size_mb == 75
    
    def test_server_initialization(self, server_config):
        """Test PowerPoint MCP server initialization."""
        with patch('powerpoint_mcp_server.config.get_config', return_value=server_config), \
             patch('powerpoint_mcp_server.config.get_config_manager') as mock_config_manager:
            
            mock_config_manager.return_value = ConfigManager(server_config)
            server = PowerPointMCPServer()
            
            assert server.config.server_name == server_config.server_name
            assert server.config.server_version == server_config.server_version
            assert server.server is not None
            assert server.content_extractor is not None
            assert server.attribute_processor is not None
            assert server.file_validator is not None
            assert server._running is False
    
    @pytest.mark.asyncio
    async def test_server_startup_shutdown(self, server_config):
        """Test server startup and shutdown procedures."""
        with patch('powerpoint_mcp_server.config.get_config', return_value=server_config):
            server = PowerPointMCPServer()
            
            # Test initial state
            assert not server.is_running()
            
            # Test shutdown when not running
            await server.shutdown()
            assert not server.is_running()
    
    @pytest.mark.asyncio
    async def test_server_application_startup(self):
        """Test server application startup procedures."""
        app = ServerApplication()
        
        # Mock the server run method to avoid actual MCP connection
        with patch.object(app, '_validate_configuration') as mock_validate:
            mock_validate.return_value = None
            
            await app.startup()
            
            assert app.server is not None
            mock_validate.assert_called_once()
    
    @pytest.mark.asyncio
    async def test_server_application_shutdown(self):
        """Test server application shutdown procedures."""
        app = ServerApplication()
        
        # Initialize server
        await app.startup()
        
        # Mock cleanup methods
        with patch.object(app, '_cleanup_server_resources') as mock_cleanup:
            mock_cleanup.return_value = None
            
            await app.shutdown()
            
            mock_cleanup.assert_called_once()
    
    @pytest.mark.asyncio
    async def test_server_application_configuration_validation(self):
        """Test server application configuration validation."""
        app = ServerApplication()
        
        # Initialize server first
        await app.startup()
        
        # Test successful validation
        await app._validate_configuration()
        
        # Test validation with missing server
        app.server = None
        with pytest.raises(RuntimeError, match="Server instance not initialized"):
            await app._validate_configuration()
    
    @pytest.mark.asyncio
    async def test_server_application_resource_cleanup(self):
        """Test server application resource cleanup."""
        app = ServerApplication()
        await app.startup()
        
        # Mock cache manager
        mock_cache_manager = Mock()
        mock_cache_manager.clear_cache = Mock()
        
        if hasattr(app.server.content_extractor, 'cache_manager'):
            app.server.content_extractor.cache_manager = mock_cache_manager
        
        await app._cleanup_server_resources()
        
        # Verify cleanup was attempted (may not be called if cache manager doesn't exist)
        # This test ensures the cleanup method runs without errors
    
    def test_server_application_logging_setup(self):
        """Test server application logging setup."""
        app = ServerApplication()
        
        # Verify logging is configured
        logger = logging.getLogger('powerpoint_mcp_server')
        assert logger.level <= logging.INFO  # Should be INFO or more verbose
    
    def test_server_application_signal_handlers(self):
        """Test server application signal handler setup."""
        app = ServerApplication()
        
        # Verify signal handlers are set up
        assert signal.getsignal(signal.SIGINT) is not signal.SIG_DFL
        assert signal.getsignal(signal.SIGTERM) is not signal.SIG_DFL
        
        # Test shutdown event is not set initially
        assert not app.shutdown_event.is_set()
    
    @pytest.mark.asyncio
    async def test_server_application_run_with_shutdown(self):
        """Test server application run with shutdown signal."""
        app = ServerApplication()
        
        # Mock the server run method to complete quickly
        async def mock_server_run():
            await asyncio.sleep(0.1)
        
        with patch.object(app, 'startup') as mock_startup, \
             patch.object(app, 'shutdown') as mock_shutdown:
            
            mock_startup.return_value = None
            mock_shutdown.return_value = None
            
            # Create a mock server that completes quickly
            app.server = Mock()
            app.server.run = mock_server_run
            
            # Set shutdown event after a short delay
            async def trigger_shutdown():
                await asyncio.sleep(0.05)
                app.shutdown_event.set()
            
            # Run both the app and the shutdown trigger
            await asyncio.gather(
                app.run(),
                trigger_shutdown()
            )
            
            mock_startup.assert_called_once()
            mock_shutdown.assert_called_once()
    
    @pytest.mark.asyncio
    async def test_server_application_run_with_server_error(self):
        """Test server application run with server error."""
        app = ServerApplication()
        
        # Mock the server run method to raise an exception
        async def mock_server_run():
            raise RuntimeError("Server error")
        
        with patch.object(app, 'startup') as mock_startup, \
             patch.object(app, 'shutdown') as mock_shutdown:
            
            mock_startup.return_value = None
            mock_shutdown.return_value = None
            
            # Create a mock server that raises an error
            app.server = Mock()
            app.server.run = mock_server_run
            
            # The run method should raise the server error
            with pytest.raises(RuntimeError, match="Server error"):
                await app.run()
            
            mock_startup.assert_called_once()
            mock_shutdown.assert_called_once()
    
    def test_config_to_dict(self, server_config):
        """Test configuration conversion to dictionary."""
        config_dict = server_config.to_dict()
        
        expected_keys = {
            'log_level', 'log_format', 'server_name', 'server_version',
            'max_file_size_mb', 'processing_timeout_seconds', 'cache_enabled',
            'cache_ttl_seconds', 'debug_mode'
        }
        
        assert set(config_dict.keys()) == expected_keys
        assert config_dict['server_name'] == 'test-powerpoint-mcp-server'
        assert config_dict['debug_mode'] is True
    
    def test_config_max_file_size_bytes(self, server_config):
        """Test configuration max file size in bytes calculation."""
        bytes_size = server_config.get_max_file_size_bytes()
        expected_bytes = 50 * 1024 * 1024  # 50 MB in bytes
        assert bytes_size == expected_bytes


if __name__ == "__main__":
    pytest.main([__file__, "-v"])