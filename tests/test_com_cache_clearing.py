"""
Tests for COM cache clearing functionality in AppContext.

This module contains tests that verify the COM cache clearing functionality
in the AppContext class, specifically the _clear_com_cache and get_word_app methods.
"""

import os
import shutil
import sys
import tempfile
import unittest
from unittest.mock import MagicMock, patch

import pythoncom

# Add the project root to the path so we can import our modules
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from word_docx_tools.mcp_service.core_utils import WordDocumentError
from word_docx_tools.mcp_service.app_context import AppContext


class TestComCacheClearing(unittest.TestCase):
    """Tests for COM cache clearing functionality"""

    def setUp(self):
        """Set up test fixtures before each test method."""
        # Initialize COM
        pythoncom.CoInitialize()

        # Create a temporary directory for test files
        self.test_dir = tempfile.mkdtemp()

        # Create an AppContext instance
        self.app_context = AppContext()

        # Reset the _word_app attribute to ensure clean state for each test
        self.app_context._word_app = None

    def tearDown(self):
        """Tear down test fixtures after each test method."""
        # Clean up temporary directory
        try:
            shutil.rmtree(self.test_dir)
        except:
            pass

        # Uninitialize COM
        pythoncom.CoUninitialize()

    @patch("word_docx_tools.utils.app_context.win32com.client.Dispatch")
    def test_get_word_app_success(self, mock_dispatch):
        """Test successful Word application creation"""
        # Mock successful Word application creation
        mock_word_app = MagicMock()
        mock_dispatch.return_value = mock_word_app

        # Reset _word_app attribute to ensure clean state
        self.app_context._word_app = None

        # Test get_word_app with create_if_needed=True
        result = self.app_context.get_word_app(create_if_needed=True)

        # Verify the result
        self.assertIsNotNone(result)
        self.assertEqual(result, mock_word_app)
        mock_dispatch.assert_called_once_with("Word.Application")

        # Verify that subsequent calls return the same instance when already created
        result2 = self.app_context.get_word_app(create_if_needed=True)
        self.assertEqual(result2, mock_word_app)
        # Should not call Dispatch again since app is already created
        mock_dispatch.assert_called_once()

        # Reset _word_app attribute after test
        self.app_context._word_app = None

    @patch("importlib.reload")
    @patch("word_docx_tools.utils.app_context.win32com.client.Dispatch")
    @patch("word_docx_tools.utils.app_context.AppContext._clear_com_cache")
    def test_get_word_app_with_com_error_and_cache_clear_success(
        self, mock_clear_cache, mock_dispatch, mock_reload
    ):
        """Test Word application creation with COM error and successful cache clearing"""
        # Mock successful Word application creation on second call
        mock_word_app = MagicMock()
        mock_dispatch.side_effect = [AttributeError("COM error"), mock_word_app]

        # Mock successful cache clearing
        mock_clear_cache.return_value = True

        # Mock importlib.reload to prevent actual reloading
        mock_reload.return_value = None

        # Reset _word_app attribute to ensure clean state
        self.app_context._word_app = None

        # Test get_word_app with create_if_needed=True
        result = self.app_context.get_word_app(create_if_needed=True)

        # Verify the result
        self.assertEqual(result, mock_word_app)
        self.assertEqual(mock_dispatch.call_count, 2)
        mock_clear_cache.assert_called_once()
        mock_reload.assert_called_once()

        # Reset _word_app attribute after test
        self.app_context._word_app = None

    @patch("word_docx_tools.utils.app_context.win32com.client.Dispatch")
    @patch("word_docx_tools.utils.app_context.AppContext._clear_com_cache")
    def test_get_word_app_with_com_error_and_cache_clear_failure(
        self, mock_clear_cache, mock_dispatch
    ):
        """Test Word application creation with COM error and failed cache clearing"""
        # Mock AttributeError
        mock_dispatch.side_effect = AttributeError("COM error")

        # Mock failed cache clearing
        mock_clear_cache.return_value = False

        # Reset _word_app attribute to ensure clean state
        self.app_context._word_app = None

        # Test get_word_app with create_if_needed=True
        result = self.app_context.get_word_app(create_if_needed=True)

        # Verify the result
        self.assertIsNone(result)
        self.assertEqual(mock_dispatch.call_count, 1)
        mock_clear_cache.assert_called_once()

        # Reset _word_app attribute after test
        self.app_context._word_app = None

    @patch("importlib.reload")
    @patch("word_docx_tools.utils.app_context.win32com.client.Dispatch")
    @patch("word_docx_tools.utils.app_context.AppContext._clear_com_cache")
    def test_get_word_app_with_com_error_and_retry_failure(
        self, mock_clear_cache, mock_dispatch, mock_reload
    ):
        """Test Word application creation with COM error and failed retry after cache clearing"""
        # Mock AttributeError on both calls
        mock_dispatch.side_effect = [
            AttributeError("COM error"),
            AttributeError("COM error after cache clear"),
        ]

        # Mock successful cache clearing
        mock_clear_cache.return_value = True

        # Mock importlib.reload to prevent actual reloading
        mock_reload.return_value = None

        # Reset _word_app attribute to ensure clean state
        self.app_context._word_app = None

        # Test get_word_app with create_if_needed=True
        result = self.app_context.get_word_app(create_if_needed=True)

        # Reset _word_app attribute after test
        self.app_context._word_app = None

        # Verify the result
        self.assertIsNone(result)
        self.assertEqual(mock_dispatch.call_count, 2)
        mock_clear_cache.assert_called_once()
        mock_reload.assert_called_once()

    @patch("word_docx_tools.utils.app_context.os.path.exists")
    @patch("word_docx_tools.utils.app_context.shutil.rmtree")
    @patch("word_docx_tools.utils.app_context.win32com")
    def test_clear_com_cache_success(self, mock_win32com, mock_rmtree, mock_exists):
        """Test successful COM cache clearing"""
        # Mock path existence
        mock_exists.return_value = True

        # Mock win32com.__gen_path__
        mock_win32com.__gen_path__ = "/fake/path"

        # Test _clear_com_cache
        result = self.app_context._clear_com_cache()

        # Verify the result
        self.assertTrue(result)
        mock_exists.assert_called()
        mock_rmtree.assert_called()

    @patch("word_docx_tools.utils.app_context.os.path.exists")
    def test_clear_com_cache_not_found(self, mock_exists):
        """Test COM cache clearing when cache directory doesn't exist"""
        # Mock path not existing
        mock_exists.return_value = False

        # Test _clear_com_cache
        result = self.app_context._clear_com_cache()

        # Verify the result
        self.assertFalse(result)
        mock_exists.assert_called()

    @patch("word_docx_tools.utils.app_context.os.path.exists")
    @patch("word_docx_tools.utils.app_context.shutil.rmtree")
    @patch("word_docx_tools.utils.app_context.win32com")
    def test_clear_com_cache_exception(self, mock_win32com, mock_rmtree, mock_exists):
        """Test COM cache clearing when an exception occurs"""
        # Mock path existence
        mock_exists.return_value = True

        # Mock win32com.__gen_path__
        mock_win32com.__gen_path__ = "/fake/path"

        # Mock shutil.rmtree to raise an exception
        mock_rmtree.side_effect = Exception("Failed to remove directory")

        # Test _clear_com_cache
        result = self.app_context._clear_com_cache()

        # Verify the result
        self.assertFalse(result)
        mock_exists.assert_called()
        mock_rmtree.assert_called()


if __name__ == "__main__":
    unittest.main()
