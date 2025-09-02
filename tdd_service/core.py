"""
Core module for the Test-Driven Development (TDD) Service.

This file initializes shared objects that are used across different modules
in the TDD service.
"""

from collections.abc import AsyncIterator
from contextlib import asynccontextmanager

from mcp.server.fastmcp import FastMCP

# 使用绝对导入路径
from word_docx_tools.utils.app_context import AppContext


@asynccontextmanager
async def tdd_lifespan(server: FastMCP) -> AsyncIterator[AppContext]:
    """Manage TDD service lifecycle with type-safe context."""
    # Initialize AppContext without a Word application instance
    # Word application will be started on-demand when needed
    app_context = AppContext(word_app=None)
    try:
        yield app_context
    finally:
        # Cleanup on shutdown - close any open document but don't quit Word app
        app_context.close_document()


# --- TDD Server Initialization ---
# This is the central server instance for TDD tools.
tdd_server = FastMCP("word-docx-tools-tdd", lifespan=tdd_lifespan)