"""
Main MCP Server application file, built with the official mcp.server.fastmcp library.

This file serves as the entry point for the MCP server, initializing the server instance,
providing core utilities, and importing all modularized tool implementations.
"""
import os
from typing import Any, Dict, List, Optional, TypeVar, cast

# Import MCP server components
from mcp.server.fastmcp.server import Context, FastMCP

# Import core components
from word_document_server.word_backend import WordBackend
from word_document_server.errors import (
    WordDocumentError,
    format_error_response, validate_input_params
)
from word_document_server.selector import (
    AmbiguousLocatorError, SelectorEngine
)

# --- Import Shared Core Components ---
# Import shared instances initialized in core.py to avoid circular dependencies
from word_document_server.core import mcp_server, selector

# --- Helper Functions ---
# All shared utility functions are now in utils.py

# --- Import Modularized Tools ---
# This will import and register all the tools defined in the modularized files
# Import core_utils module which will make mcp_server available to tools
from word_document_server import core_utils  # This imports and uses the global mcp_server
import word_document_server.tools

# --- End of Main Application File ---




