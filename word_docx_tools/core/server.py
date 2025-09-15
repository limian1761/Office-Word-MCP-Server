"""Server module for Word Document MCP Server.

This module initializes shared objects that are used across different modules,
including the MCP server instance, to avoid circular dependencies.
"""

from collections.abc import AsyncIterator
from contextlib import asynccontextmanager

from mcp.server.fastmcp import FastMCP

from .app_context import AppContext


@asynccontextmanager
async def app_lifespan(server: FastMCP) -> AsyncIterator[AppContext]:
    """Manage application lifecycle with type-safe context."""
    # Initialize AppContext
    # Word application will be started on-demand when needed
    app_context = AppContext()
    try:
        yield app_context
    finally:
        # Cleanup on shutdown - close any open document but don't quit Word app
        app_context.close_document()


# --- MCP Server Initialization ---
# This is the central server instance that tools will be registered against.
mcp_server = FastMCP("word-docx-tools", lifespan=app_lifespan)

# --- AppContext Initialization ---
# This is the central AppContext instance used by tools to manage Word application and document operations.
app_context = AppContext()