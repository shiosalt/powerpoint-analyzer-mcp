#!/usr/bin/env python3
"""Main entry point for the PowerPoint MCP Server."""

import asyncio
import logging
import signal
import sys
import os
from typing import Optional
from pathlib import Path

from powerpoint_mcp_server.server import PowerPointMCPServer
from powerpoint_mcp_server.config import get_config, get_config_manager


class ServerApplication:
    """Main server application with configuration and lifecycle management."""
    
    def __init__(self):
        """Initialize the server application."""
        self.server: Optional[PowerPointMCPServer] = None
        self.shutdown_event = asyncio.Event()
        self._setup_logging()
        self._setup_signal_handlers()
    
    def _setup_logging(self):
        """Configure logging for the application."""
        config = get_config()
        
        # Configure root logger
        logging.basicConfig(
            level=getattr(logging, config.log_level.upper(), logging.INFO),
            format=config.log_format,
            handlers=[
                logging.StreamHandler(sys.stderr)
            ]
        )
        
        # Set specific logger levels
        logger = logging.getLogger('powerpoint_mcp_server')
        logger.setLevel(getattr(logging, config.log_level.upper(), logging.INFO))
        
        # Reduce noise from other libraries unless in debug mode
        if not config.debug_mode:
            logging.getLogger('asyncio').setLevel(logging.WARNING)
            logging.getLogger('mcp').setLevel(logging.WARNING)
        
        logger.info(f"Logging configured with level: {config.log_level}")
    
    def _setup_signal_handlers(self):
        """Set up signal handlers for graceful shutdown."""
        def signal_handler(signum, frame):
            """Handle shutdown signals."""
            logger = logging.getLogger(__name__)
            logger.info(f"Received signal {signum}, initiating shutdown...")
            self.shutdown_event.set()
        
        # Register signal handlers
        signal.signal(signal.SIGINT, signal_handler)
        signal.signal(signal.SIGTERM, signal_handler)
        
        # On Windows, SIGBREAK is used instead of SIGTERM
        if hasattr(signal, 'SIGBREAK'):
            signal.signal(signal.SIGBREAK, signal_handler)
    
    async def startup(self):
        """Perform server startup procedures."""
        logger = logging.getLogger(__name__)
        
        try:
            logger.info("Starting PowerPoint MCP Server...")
            
            # Initialize the server
            self.server = PowerPointMCPServer()
            
            # Validate server configuration
            await self._validate_configuration()
            
            logger.info("Server startup completed successfully")
            
        except Exception as e:
            logger.error(f"Server startup failed: {e}")
            raise
    
    async def _validate_configuration(self):
        """Validate server configuration and dependencies."""
        logger = logging.getLogger(__name__)
        
        # Check if required modules are available
        try:
            import xml.etree.ElementTree
            import zipfile
            import json
            logger.debug("All required modules are available")
        except ImportError as e:
            raise RuntimeError(f"Missing required dependency: {e}")
        
        # Validate server instance
        if not self.server:
            raise RuntimeError("Server instance not initialized")
        
        logger.debug("Configuration validation completed")
    
    async def shutdown(self):
        """Perform server shutdown procedures."""
        logger = logging.getLogger(__name__)
        
        try:
            logger.info("Initiating server shutdown...")
            
            if self.server:
                # Perform any server-specific cleanup
                await self._cleanup_server_resources()
            
            logger.info("Server shutdown completed successfully")
            
        except Exception as e:
            logger.error(f"Error during server shutdown: {e}")
            raise
    
    async def _cleanup_server_resources(self):
        """Clean up server resources."""
        logger = logging.getLogger(__name__)
        
        try:
            # Clear any cached data if cache manager exists
            if hasattr(self.server, 'content_extractor') and hasattr(self.server.content_extractor, 'cache_manager'):
                cache_manager = self.server.content_extractor.cache_manager
                if hasattr(cache_manager, 'clear_cache'):
                    cache_manager.clear_cache()
                    logger.debug("Cache cleared")
            
            # Close any open file handles or connections
            logger.debug("Server resources cleaned up")
            
        except Exception as e:
            logger.warning(f"Error cleaning up server resources: {e}")
    
    async def run(self):
        """Run the server application."""
        logger = logging.getLogger(__name__)
        
        try:
            # Perform startup procedures
            await self.startup()
            
            # Create a task to run the server
            server_task = asyncio.create_task(self.server.run())
            
            # Create a task to wait for shutdown signal
            shutdown_task = asyncio.create_task(self.shutdown_event.wait())
            
            # Wait for either the server to complete or shutdown signal
            done, pending = await asyncio.wait(
                [server_task, shutdown_task],
                return_when=asyncio.FIRST_COMPLETED
            )
            
            # Cancel pending tasks
            for task in pending:
                task.cancel()
                try:
                    await task
                except asyncio.CancelledError:
                    pass
            
            # Check if server task completed with an exception
            if server_task in done:
                try:
                    await server_task
                except Exception as e:
                    logger.error(f"Server task failed: {e}")
                    raise
            
        except Exception as e:
            logger.error(f"Application error: {e}")
            raise
        finally:
            # Always perform shutdown procedures
            await self.shutdown()


async def main():
    """Main entry point for the server application."""
    app = ServerApplication()
    await app.run()


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("\nServer stopped by user")
        sys.exit(0)
    except Exception as e:
        print(f"Server error: {e}", file=sys.stderr)
        sys.exit(1)