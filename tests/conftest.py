import os
import shutil
import sys
import tempfile
from typing import Optional

import pytest
import win32com.client

# Add project root to the Python path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))


@pytest.fixture(scope="session")
def word_app():
    """
    Fixture to provide a single Word application instance for the entire test session.
    Starts Word once at the beginning of the session and closes it at the end.
    
    Yields:
        A reference to the Word Application COM object.
    """
    # Start Word application
    word = None
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False  # Keep Word invisible for testing
        yield word
    finally:
        # Cleanup: Close Word application
        if word:
            # Ensure all documents are closed before quitting
            for doc in word.Documents:
                doc.Close(SaveChanges=0)  # 0 = wdDoNotSaveChanges
            word.Quit()


@pytest.fixture
def document(word_app, request):
    """
    Fixture to provide a clean, isolated test document for each test function.
    
    This fixture:
    1. Copies the specified test document to a temporary directory
    2. Opens the temporary copy
    3. Yields the Document object to the test function
    4. Closes the document without saving
    5. Cleans up the temporary file
    
    Args:
        word_app: The Word application instance from the word_app fixture
        request: pytest request object containing information about the test function
    
    Yields:
        A reference to the opened Word Document COM object.
    """
    # Get the document name from the test marker or use default
    marker = request.node.get_closest_marker("document_name")
    doc_name = marker.args[0] if marker and marker.args else "test_document.docx"
    
    # Determine the source path
    test_docs_dir = os.path.join(os.path.dirname(__file__), "test_docs")
    source_path = os.path.join(test_docs_dir, doc_name)
    
    # Check if the source document exists
    if not os.path.exists(source_path):
        pytest.skip(f"Test document '{doc_name}' not found at {source_path}")
    
    # Create a temporary directory
    temp_dir = tempfile.mkdtemp()
    
    # Copy the document to the temporary directory
    temp_doc_path = os.path.join(temp_dir, doc_name)
    shutil.copy2(source_path, temp_doc_path)
    
    # Open the document
    doc = None
    try:
        doc = word_app.Documents.Open(temp_doc_path)
        yield doc
    finally:
        # Cleanup: Close the document without saving
        if doc:
            doc.Close(SaveChanges=0)  # 0 = wdDoNotSaveChanges
        
        # Remove the temporary directory and file
        if os.path.exists(temp_dir):
            try:
                shutil.rmtree(temp_dir)
            except OSError:
                # Ignore errors during cleanup (files might be locked by Word)
                pass


# Utility function to get document path for backward compatibility
def get_test_doc_path(doc_name: str = "test_document.docx") -> str:
    """
    Utility function to get the path to a test document.
    
    Args:
        doc_name: The name of the test document
    
    Returns:
        The absolute path to the test document
    """
    test_docs_dir = os.path.join(os.path.dirname(__file__), "test_docs")
    return os.path.join(test_docs_dir, doc_name)


# Add a marker for specifying document name in tests
def pytest_configure(config):
    """Register custom markers."""
    config.addinivalue_line(
        "markers",
        "document_name(name): specify the name of the test document to use"
    )


# Import sys for path manipulation
import sys