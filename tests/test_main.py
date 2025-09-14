"""
Tests for the main module.
"""
import pytest
from unittest.mock import patch, MagicMock

from word_docx_tools.main import run_server


def test_run_server():
    """Test that run_server function calls mcp_server.run with correct parameters."""
    with patch('word_docx_tools.main.mcp_server') as mock_server:
        run_server()
        mock_server.run.assert_called_once_with(transport="stdio")