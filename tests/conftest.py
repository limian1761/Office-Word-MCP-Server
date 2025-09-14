"""
Configuration file for pytest.
"""
import sys
from pathlib import Path

# Add the project root directory to Python path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

import pytest
from unittest.mock import MagicMock, patch


@pytest.fixture
def mock_word_app():
    """Mock Word application COM object."""
    mock_app = MagicMock()
    mock_app.Version = "16.0"
    mock_app.Name = "Microsoft Word"
    return mock_app


@pytest.fixture
def mock_document():
    """Mock Word document COM object."""
    mock_doc = MagicMock()
    mock_doc.Name = "test_document.docx"
    mock_doc.FullName = "C:\\temp\\test_document.docx"
    mock_doc.Saved = True
    return mock_doc


@pytest.fixture
def mock_app_context():
    """Mock AppContext for testing."""
    with patch('word_docx_tools.mcp_service.app_context.AppContext') as mock_context:
        instance = mock_context.return_value
        instance.get_word_app.return_value = MagicMock()
        instance.get_active_document.return_value = MagicMock()
        instance.set_active_document.return_value = None
        instance.set_word_app.return_value = None
        yield instance