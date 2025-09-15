"""Core module for Word Document MCP Server.

This module contains the core functionality for the server,
including application context management and server implementation.
"""

from .app_context import AppContext
from .server import mcp_server
from .utils import *

__all__ = [
    "AppContext",
    "mcp_server",
]