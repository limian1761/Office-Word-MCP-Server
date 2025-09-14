"""
Tests for document operations.
"""
import pytest
from unittest.mock import patch, MagicMock

from word_docx_tools.operations import document_ops


@patch('word_docx_tools.operations.document_ops.create_document')
def test_create_document(mock_create):
    """Test document creation."""
    mock_doc = MagicMock()
    mock_word_app = MagicMock()
    mock_create.return_value = mock_doc
    
    result = document_ops.create_document(mock_word_app)
    assert result == mock_doc
    mock_create.assert_called_once_with(mock_word_app, visible=True, template_path=None)


@patch('word_docx_tools.operations.document_ops.open_document')
def test_open_document(mock_open):
    """Test document opening."""
    mock_doc = MagicMock()
    mock_word_app = MagicMock()
    file_path = "test.docx"
    mock_open.return_value = mock_doc
    
    result = document_ops.open_document(mock_word_app, file_path)
    assert result == mock_doc
    mock_open.assert_called_once_with(mock_word_app, file_path, visible=True, password=None)


def test_save_document():
    """Test document saving."""
    mock_doc = MagicMock()
    mock_doc.Save = MagicMock()
    
    result = document_ops.save_document(mock_doc)
    assert result is True
    mock_doc.Save.assert_called_once()


@patch('word_docx_tools.operations.document_ops.save_document')
def test_save_document_as(mock_save):
    """Test document saving as."""
    mock_doc = MagicMock()
    file_path = "new_test.docx"
    mock_save.return_value = True
    
    result = document_ops.save_document(mock_doc, file_path)
    assert result is True
    mock_save.assert_called_once_with(mock_doc, file_path)


def test_close_document():
    """Test document closing."""
    mock_doc = MagicMock()
    mock_doc.Close = MagicMock()
    
    result = document_ops.close_document(mock_doc)
    assert result is True
    mock_doc.Close.assert_called_once_with(SaveChanges=0)