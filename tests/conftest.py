"""
Pytest configuration for Word Document MCP Server tests.

This file is automatically loaded by pytest and sets up the test environment,
including adding the project root to the Python path so that test files can
import modules from the project.
"""

import os
import sys
import pytest
import pythoncom
import win32com.client

# Add the project root to the Python path so we can import modules
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))


@pytest.fixture(scope='session', autouse=True)
def setup_com():
    """Initialize COM for the test session."""
    pythoncom.CoInitialize()
    yield
    pythoncom.CoUninitialize()


@pytest.fixture(scope='module')
def word_application():
    """Create a Word application instance for tests."""
    app = win32com.client.Dispatch('Word.Application')
    app.Visible = False
    yield app
    # Clean up
    if app.Documents.Count > 0:
        for doc in app.Documents:
            doc.Close(SaveChanges=False)
    app.Quit()


@pytest.fixture(scope='module')
def test_document(word_application):
    """Create a test document for tests."""
    doc = word_application.Documents.Add()
    # Add some test content
    doc.Range(0, 0).Text = "这是测试文档的第一段。\n"
    doc.Range().Collapse(0)  # 0 = wdCollapseEnd
    doc.Range().Text = "这是包含关键词的第二段。\n"
    yield doc
    # Clean up
    doc.Close(SaveChanges=False)