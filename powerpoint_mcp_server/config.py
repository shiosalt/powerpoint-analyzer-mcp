"""Configuration management for PowerPoint MCP Server."""

import os
import logging
from typing import Dict, Any, Optional
from dataclasses import dataclass, field


@dataclass
class ServerConfig:
    """Server configuration settings."""
    
    # Logging configuration
    log_level: str = field(default_factory=lambda: os.getenv('POWERPOINT_MCP_LOG_LEVEL', 'INFO'))
    log_format: str = field(default='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    
    # Server configuration
    server_name: str = field(default='powerpoint-mcp-server')
    server_version: str = field(default='0.1.0')
    
    # File processing limits
    max_file_size_mb: int = field(default_factory=lambda: int(os.getenv('POWERPOINT_MCP_MAX_FILE_SIZE', '100')))
    processing_timeout_seconds: int = field(default_factory=lambda: int(os.getenv('POWERPOINT_MCP_TIMEOUT', '300')))
    
    # Cache configuration
    cache_enabled: bool = field(default_factory=lambda: os.getenv('POWERPOINT_MCP_CACHE_ENABLED', 'true').lower() == 'true')
    cache_ttl_seconds: int = field(default_factory=lambda: int(os.getenv('POWERPOINT_MCP_CACHE_TTL', '3600')))
    
    # Debug configuration
    debug_mode: bool = field(default_factory=lambda: os.getenv('POWERPOINT_MCP_DEBUG', 'false').lower() == 'true')
    
    def __post_init__(self):
        """Validate configuration after initialization."""
        self._validate_config()
    
    def _validate_config(self):
        """Validate configuration values."""
        # Validate log level
        valid_log_levels = ['DEBUG', 'INFO', 'WARNING', 'ERROR', 'CRITICAL']
        if self.log_level.upper() not in valid_log_levels:
            self.log_level = 'INFO'
        
        # Validate numeric values
        if self.max_file_size_mb <= 0:
            self.max_file_size_mb = 100
        
        if self.processing_timeout_seconds <= 0:
            self.processing_timeout_seconds = 300
        
        if self.cache_ttl_seconds <= 0:
            self.cache_ttl_seconds = 3600
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert configuration to dictionary."""
        return {
            'log_level': self.log_level,
            'log_format': self.log_format,
            'server_name': self.server_name,
            'server_version': self.server_version,
            'max_file_size_mb': self.max_file_size_mb,
            'processing_timeout_seconds': self.processing_timeout_seconds,
            'cache_enabled': self.cache_enabled,
            'cache_ttl_seconds': self.cache_ttl_seconds,
            'debug_mode': self.debug_mode
        }
    
    @classmethod
    def from_env(cls) -> 'ServerConfig':
        """Create configuration from environment variables."""
        return cls()
    
    def get_max_file_size_bytes(self) -> int:
        """Get maximum file size in bytes."""
        return self.max_file_size_mb * 1024 * 1024


class ConfigManager:
    """Configuration manager for the server."""
    
    def __init__(self, config: Optional[ServerConfig] = None):
        """Initialize configuration manager."""
        self.config = config or ServerConfig.from_env()
        self.logger = logging.getLogger(__name__)
    
    def get_config(self) -> ServerConfig:
        """Get current configuration."""
        return self.config
    
    def update_config(self, **kwargs) -> None:
        """Update configuration values."""
        for key, value in kwargs.items():
            if hasattr(self.config, key):
                setattr(self.config, key, value)
                self.logger.debug(f"Configuration updated: {key} = {value}")
            else:
                self.logger.warning(f"Unknown configuration key: {key}")
        
        # Re-validate after updates
        self.config._validate_config()
    
    def log_configuration(self) -> None:
        """Log current configuration (excluding sensitive data)."""
        config_dict = self.config.to_dict()
        self.logger.info("Current server configuration:")
        for key, value in config_dict.items():
            self.logger.info(f"  {key}: {value}")


# Global configuration instance
_config_manager: Optional[ConfigManager] = None


def get_config_manager() -> ConfigManager:
    """Get the global configuration manager instance."""
    global _config_manager
    if _config_manager is None:
        _config_manager = ConfigManager()
    return _config_manager


def get_config() -> ServerConfig:
    """Get the current server configuration."""
    return get_config_manager().get_config()