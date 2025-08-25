import os
import sys

import pytest
import win32com.client

# Add project root to Python path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from word_document_server.com_backend import WordBackend

# Define test document path
TEST_DOC_PATH = os.path.join(os.path.dirname(__file__), 'test_docs', 'test_document.docx')

def test_context_manager_no_file():
    """Test that the backend can be used as a context manager without a file."""
    try:
        with WordBackend(visible=False) as backend:
            assert backend.word_app is not None
            assert backend.document is not None
            # Check that a new document is created
            assert backend.document.Name.startswith("Document") or backend.document.Name.startswith("文档")
    except Exception as e:
        pytest.fail(f"WordBackend context manager raised an exception: {e}")

@pytest.mark.document_name("test_document.docx")
def test_context_manager_with_file():
    """Test that the backend can open an existing document."""
    try:
        with WordBackend(file_path=TEST_DOC_PATH, visible=False) as backend:
            assert backend.word_app is not None
            assert backend.document is not None
            assert "test_document.docx" in backend.document.Name
    except Exception as e:
        pytest.fail(f"WordBackend context manager raised an exception with file: {e}")

def test_get_all_paragraphs():
    """Test retrieving all paragraphs from a document."""
    if not os.path.exists(TEST_DOC_PATH):
        pytest.skip(f"Test document not found at {TEST_DOC_PATH}")

    with WordBackend(file_path=TEST_DOC_PATH, visible=False) as backend:
        paragraphs = backend.get_all_paragraphs()
        assert isinstance(paragraphs, list)
        # Assuming test_document.docx has at least one paragraph
        assert len(paragraphs) > 0
        assert isinstance(paragraphs[0], win32com.client.CDispatch)

def test_formatting_methods():
    """Test the text formatting methods."""
    with WordBackend(visible=False) as backend:
        doc = backend.document
        # Add a paragraph to format
        p = doc.Paragraphs.Add()
        p.Range.Text = "This is a test paragraph."
        
        test_range = p.Range

        # Test bold
        backend.set_bold_for_range(test_range, True)
        assert test_range.Font.Bold == -1  # -1 means True in COM
        backend.set_bold_for_range(test_range, False)
        assert test_range.Font.Bold == 0

        # Test italic
        backend.set_italic_for_range(test_range, True)
        assert test_range.Font.Italic == -1
        backend.set_italic_for_range(test_range, False)
        assert test_range.Font.Italic == 0

        # Test font size
        backend.set_font_size_for_range(test_range, 14)
        assert test_range.Font.Size == 14

        # Test font name
        backend.set_font_name_for_range(test_range, "Arial")
        assert test_range.Font.Name == "Arial"

def test_insert_paragraph_after():
    """Test inserting a paragraph after a range."""
    with WordBackend(visible=False) as backend:
        doc = backend.document
        initial_p_count = doc.Paragraphs.Count
        
        # Use the first paragraph as the anchor
        if initial_p_count == 0:
            doc.Paragraphs.Add() # Add one if empty
        
        first_p_range = doc.Paragraphs(1).Range
        
        backend.insert_paragraph_after(first_p_range, "Newly inserted paragraph.")
        
        assert doc.Paragraphs.Count == initial_p_count + 1
        
        # This is a bit tricky to verify content due to how ranges work after insertion.
        # A simple check is to see if the new paragraph's text is somewhere.
        # Note: COM paragraph indexing is 1-based.
        new_p_text = doc.Paragraphs(2).Range.Text
        assert "Newly inserted paragraph." in new_p_text


if __name__ == "__main__":
    pytest.main(["-v", __file__])
