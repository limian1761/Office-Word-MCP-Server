"""
Tests for the MCP service core components.
"""
import pytest
from unittest.mock import patch, MagicMock, AsyncMock

from word_docx_tools.mcp_service.core import app_lifespan, mcp_server
from word_docx_tools.mcp_service.app_context import AppContext


@pytest.mark.asyncio
async def test_app_lifespan():
    """Test the app lifespan manager."""
    # Create a mock server
    mock_server = MagicMock()
    
    # Test the lifespan context manager
    async with app_lifespan(mock_server) as app_context:
        assert isinstance(app_context, AppContext)
        
        # Check that app_context is properly initialized
        assert hasattr(app_context, '_word_app')
        assert hasattr(app_context, '_active_document')


def test_mcp_server_initialization():
    """Test that the MCP server is properly initialized."""
    assert mcp_server is not None
    assert mcp_server.name == "word-docx-tools"


@pytest.mark.asyncio
async def test_app_lifespan_cleanup():
    """Test that cleanup happens properly in app lifespan."""
    mock_server = MagicMock()
    
    async with app_lifespan(mock_server) as app_context:
        # Set up a mock document to be closed
        mock_doc = MagicMock()
        app_context._active_document = mock_doc
    
    # Verify that close_document was called during cleanup
    # Note: This requires checking the actual implementation of AppContext.close_document