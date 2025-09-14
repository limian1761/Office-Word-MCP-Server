"""
Tests for the AppContext class.
"""
import pytest
from unittest.mock import patch, MagicMock

from word_docx_tools.mcp_service.app_context import AppContext


def test_app_context_singleton():
    """Test that AppContext follows singleton pattern."""
    context1 = AppContext()
    context2 = AppContext()
    assert context1 is context2


def test_app_context_initialization():
    """Test AppContext initialization."""
    context = AppContext()
    assert context._word_app is None
    assert context._active_document is None
    assert context._document_context_tree is None


def test_set_get_word_app():
    """Test setting and getting Word application."""
    context = AppContext()
    mock_app = MagicMock()
    
    context.set_word_app(mock_app)
    assert context._word_app == mock_app
    
    # Test getting the Word app
    assert context.get_word_app() == mock_app


def test_set_get_active_document():
    """Test setting and getting active document."""
    context = AppContext()
    mock_doc = MagicMock()
    
    context.set_active_document(mock_doc)
    assert context.get_active_document() == mock_doc


def test_close_document():
    """Test closing document."""
    context = AppContext()
    mock_doc = MagicMock()
    context.set_active_document(mock_doc)
    
    # Mock the Close method
    mock_doc.Close = MagicMock()
    
    result = context.close_document()
    assert result is True
    mock_doc.Close.assert_called_once_with(SaveChanges=0)
    assert context.get_active_document() is None


@patch('win32com.client.Dispatch')
def test_get_word_app_creates_instance(mock_dispatch):
    """Test that get_word_app creates Word instance when needed."""
    context = AppContext()
    mock_word_app = MagicMock()
    mock_dispatch.return_value = mock_word_app
    
    result = context.get_word_app(create_if_needed=True)
    assert result == mock_word_app
    mock_dispatch.assert_called_once_with("Word.Application")