"""
Core module for the Word Document MCP Server.

This file initializes shared objects that are used across different modules,
such as the MCP server instance and the selector engine, to avoid circular dependencies.
"""

from collections.abc import AsyncIterator
from contextlib import asynccontextmanager

from mcp.server.fastmcp import FastMCP

from word_document_server.selector.selector import SelectorEngine
from word_document_server.utils.app_context import AppContext


@asynccontextmanager
async def app_lifespan(server: FastMCP) -> AsyncIterator[AppContext]:
    """Manage application lifecycle with type-safe context."""
    # Initialize AppContext without a Word application instance
    # Word application will be started on-demand when needed
    app_context = AppContext(word_app=None)
    try:
        yield app_context
    finally:
        # Cleanup on shutdown - close any open document but don't quit Word app
        app_context.close_document()


# --- MCP Server Initialization ---
# This is the central server instance that tools will be registered against.
mcp_server = FastMCP("DirectOfficeWordMCP", lifespan=app_lifespan)

# --- Selector Engine Initialization ---
# This is the central selector engine instance used by tools to find objects.
selector = SelectorEngine()