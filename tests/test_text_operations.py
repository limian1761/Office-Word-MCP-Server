"""
Tests for text operations.
"""
import pytest
from unittest.mock import patch, MagicMock

from word_docx_tools.operations import text_operations


def test_insert_text_before():
    """Test inserting text before a range."""
    mock_range = MagicMock()
    mock_range.Start = 10
    mock_range.End = 20
    
    # Mock the Duplicate property
    mock_duplicate = MagicMock()
    mock_range.Duplicate = mock_duplicate
    mock_duplicate.Collapse = MagicMock()
    mock_duplicate.Text = ""
    
    result = text_operations.insert_text(mock_range, "Hello World", position="before")
    
    # Verify Collapse was called with the correct direction
    mock_duplicate.Collapse.assert_called()
    # Verify text was set
    assert mock_duplicate.Text == "Hello World"


def test_insert_text_after():
    """Test inserting text after a range."""
    mock_range = MagicMock()
    mock_range.Start = 10
    mock_range.End = 20
    
    # Mock the Duplicate property
    mock_duplicate = MagicMock()
    mock_range.Duplicate = mock_duplicate
    mock_duplicate.Collapse = MagicMock()
    mock_duplicate.Text = ""
    
    result = text_operations.insert_text(mock_range, "Hello World", position="after")
    
    # Verify Collapse was called with the correct direction
    mock_duplicate.Collapse.assert_called()
    # Verify text was set
    assert mock_duplicate.Text == "Hello World"


def test_replace_text():
    """Test replacing text in a range."""
    mock_range = MagicMock()
    mock_range.Text = "Old Text"
    
    result = text_operations.replace_text(mock_range, "New Text")
    
    # Verify text was replaced
    assert mock_range.Text == "New Text"
    assert result is True


def test_get_text():
    """Test getting text from a range."""
    mock_range = MagicMock()
    mock_range.Text = "Sample Text Content"
    
    result = text_operations.get_text(mock_range)
    
    # Verify correct text was returned
    assert result == "Sample Text Content"