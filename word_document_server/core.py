"""
Core module for the Word Document MCP Server.

This file initializes shared objects that are used across different modules,
such as the MCP server instance and the selector engine, to avoid circular dependencies.
"""

from mcp.server.fastmcp import FastMCP

from word_document_server.selector import SelectorEngine

# --- MCP Server Initialization ---
# This is the central server instance that tools will be registered against.
mcp_server = FastMCP("Office-Word-MCP-Server")

# --- Selector Engine Initialization ---
# This is the central selector engine instance used by tools to find elements.
selector = SelectorEngine()
