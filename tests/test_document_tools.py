"""
Tests for document tools.
"""
import pytest
from unittest.mock import patch, MagicMock

from word_docx_tools.tools.document_tools import document_tools


@pytest.fixture
def mock_context():
    """Mock context for testing tools."""
    mock_ctx = MagicMock()
    mock_ctx.request_context.lifespan_context = MagicMock()
    return mock_ctx


@patch('word_docx_tools.tools.document_tools.create_document')
@patch('word_docx_tools.tools.document_tools.save_document')
def test_document_tools_create(mock_save, mock_create, mock_context):
    """Test document creation through tools."""
    mock_doc = MagicMock()
    mock_doc.Name = "test.docx"
    mock_create.return_value = mock_doc
    mock_save.return_value = True
    
    mock_context.request_context.lifespan_context.get_word_app.return_value = MagicMock()
    mock_context.request_context.lifespan_context.get_active_document.return_value = None
    
    # Test create operation
    result = document_tools(
        ctx=mock_context,
        operation_type="create",
        file_path="test.docx"
    )
    
    # Verify the calls were made
    mock_create.assert_called_once()
    mock_save.assert_called_once()


@patch('word_docx_tools.tools.document_tools.open_document')
def test_document_tools_open(mock_open, mock_context):
    """Test document opening through tools."""
    mock_doc = MagicMock()
    mock_doc.Name = "test.docx"
    mock_doc.FullName = "C:\\temp\\test.docx"
    mock_doc.Saved = True
    mock_open.return_value = mock_doc
    
    mock_context.request_context.lifespan_context.get_word_app.return_value = MagicMock()
    mock_context.request_context.lifespan_context.get_active_document.return_value = None
    
    # Test open operation
    result = document_tools(
        ctx=mock_context,
        operation_type="open",
        file_path="test.docx"
    )
    
    # Verify the calls were made
    mock_open.assert_called_once()


def test_document_tools_save(mock_context):
    """Test document saving through tools."""
    mock_doc = MagicMock()
    mock_context.request_context.lifespan_context.get_active_document.return_value = mock_doc
    
    with patch('word_docx_tools.tools.document_tools.save_document') as mock_save:
        mock_save.return_value = True
        
        # Test save operation
        result = document_tools(
            ctx=mock_context,
            operation_type="save"
        )
        
        # Verify the calls were made
        mock_save.assert_called_once_with(mock_doc)


def test_document_tools_close(mock_context):
    """Test document closing through tools."""
    mock_doc = MagicMock()
    mock_context.request_context.lifespan_context.get_active_document.return_value = mock_doc
    
    with patch('word_docx_tools.tools.document_tools.close_document') as mock_close:
        mock_close.return_value = True
        
        # Test close operation
        result = document_tools(
            ctx=mock_context,
            operation_type="close"
        )
        
        # Verify the calls were made
        mock_close.assert_called_once_with(mock_doc)
        mock_context.request_context.lifespan_context.set_active_document.assert_called_once_with(None)