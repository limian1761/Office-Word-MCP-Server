import os
import sys

import pytest
import pythoncom

# Add project root to Python path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from create_test_doc import create_test_document
from word_document_server import app
from word_document_server.com_backend import WordBackend

# --- Test Setup ---

# Path to the test document
TEST_DOC_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), 'test_docs', 'test_document.docx'))

class MockContext:
    """A mock context object to simulate the MCP server's context."""
    def __init__(self):
        self._state = {}
    def set_state(self, key, value):
        self._state[key] = value
    def get_state(self, key):
        return self._state.get(key)

# --- Test Cases ---

@pytest.fixture(scope="module")
def test_doc():
    """Fixture to create a fresh test document for the module."""
    create_test_document(TEST_DOC_PATH)
    return TEST_DOC_PATH

def test_set_header_text(test_doc):
    """
    Verify that the set_header_text tool correctly adds text to the document's primary header.
    """
    pythoncom.CoInitialize()
    mock_context = MockContext()
    header_text_to_set = "This is a test header."
    
    try:
        # 1. Open the document using the tool
        result = app.open_document(mock_context, test_doc)
        assert "Active document set to" in result, "Failed to open document."

        # 2. Call the tool to set the header text (this tool doesn't exist yet)
        result = app.set_header_text(mock_context, header_text_to_set)
        assert result == "Header text set successfully.", "Tool did not report success."

        # 3. Manually verify the result using the backend
        backend = mock_context.get_state("word_backend")
        assert backend is not None, "Backend was not initialized in context."
        
        # Access the primary header of the first section
        header_range = backend.document.Sections(1).Headers(1).Range
        # Word adds a final paragraph mark, so we check if the text starts with our string
        assert header_range.Text.strip().startswith(header_text_to_set), "Header text was not set correctly."

    finally:
        # 4. Clean up
        app.shutdown_word(mock_context)
        pythoncom.CoUninitialize()

def test_set_footer_text(test_doc):
    """
    Verify that the set_footer_text tool correctly adds text to the document's primary footer.
    """
    pythoncom.CoInitialize()
    mock_context = MockContext()
    footer_text_to_set = "This is a test footer."
    
    try:
        # 1. Open the document
        result = app.open_document(mock_context, test_doc)
        assert "Active document set to" in result

        # 2. Call the tool to set the footer text
        result = app.set_footer_text(mock_context, footer_text_to_set)
        assert result == "Footer text set successfully."

        # 3. Verify the result
        backend = mock_context.get_state("word_backend")
        assert backend is not None
        footer_range = backend.document.Sections(1).Footers(1).Range
        assert footer_range.Text.strip().startswith(footer_text_to_set)

    finally:
        # 4. Clean up
        app.shutdown_word(mock_context)
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    pytest.main(["-v", __file__])

def test_apply_format_tool(test_doc):
    """
    Verify that the apply_format tool correctly applies formatting to an element.
    """
    pythoncom.CoInitialize()
    mock_context = MockContext()
    
    # Locator for the first paragraph
    locator = {
        "target": {
            "type": "paragraph",
            "filters": [{"index_in_parent": 0}]
        }
    }
    
    # Formatting to apply
    formatting = {
        "bold": True,
        "alignment": "center"  # "left", "center", "right"
    }

    try:
        # 1. Open the document
        app.open_document(mock_context, test_doc)

        # 2. Call the tool to apply formatting (this tool doesn't exist yet)
        result = app.apply_format(mock_context, locator, formatting)
        assert result == "Formatting applied successfully."

        # 3. Manually verify the result using the backend
        backend = mock_context.get_state("word_backend")
        assert backend is not None, "Backend was not initialized in context."
        
        # Re-select the element to check its properties
        p1 = backend.document.Paragraphs(1)
        
        # Verify formatting
        # Note: Word's COM constant for center alignment is 1
        # Note: Word's COM constant for True is -1, so we cast to bool
        assert bool(p1.Range.Font.Bold) is True, "Text was not made bold."
        assert p1.Range.ParagraphFormat.Alignment == 1, "Paragraph was not center-aligned."

    finally:
        # 4. Clean up
        app.shutdown_word(mock_context)
        pythoncom.CoUninitialize()

def test_create_bulleted_list_tool(test_doc):
    """
    Verify that the create_bulleted_list tool correctly inserts a new list.
    """
    pythoncom.CoInitialize()
    mock_context = MockContext()
    
    # Locator for the paragraph we want to insert our list before
    locator = {
        "target": {
            "type": "paragraph",
            "filters": [{"contains_text": "This paragraph is outside the table."}]
        }
    }
    
    items_to_add = ["New item 1", "New item 2"]

    try:
        # 1. Open the document
        app.open_document(mock_context, test_doc)

        # 2. Call the tool to create the list
        result = app.create_bulleted_list(mock_context, locator, items_to_add, position="before")
        assert result == "Bulleted list created successfully."

        # 3. Manually verify the result using the selector
        backend = mock_context.get_state("word_backend")
        selector = app.selector # Get the selector from the app module
        
        # Find all list items in the document
        list_item_locator = {"target": {"type": "paragraph", "filters": [{"is_list_item": True}]}}
        selection = selector.select(backend, list_item_locator)
        
        # We created 2 new items, and there were 2 existing items.
        assert len(selection._elements) == 4, "Did not find the correct total number of list items."
        
        full_text = selection.get_text()
        assert "New item 1" in full_text
        assert "New item 2" in full_text
        assert "List item 1" in full_text # Verify old items are still there

    finally:
        # 4. Clean up
        app.shutdown_word(mock_context)
        pythoncom.CoUninitialize()

def test_accept_all_changes_tool(test_doc):
    """
    Verify that the accept_all_changes tool correctly accepts all revisions.
    """
    pythoncom.CoInitialize()
    mock_context = MockContext()
    
    try:
        # 1. Open the document
        app.open_document(mock_context, test_doc)
        backend = mock_context.get_state("word_backend")
        
        # 2. Verify that there are revisions to accept
        assert len(backend.document.Revisions) > 0, "Test document should have revisions."

        # 3. Call the tool to accept all changes
        result = app.accept_all_changes(mock_context)
        assert result == "All changes accepted successfully."

        # 4. Verify that there are no more revisions
        assert len(backend.document.Revisions) == 0, "Revisions were not accepted."

    finally:
        # 5. Clean up
        app.shutdown_word(mock_context)
        pythoncom.CoUninitialize()

def test_get_document_structure_tool(test_doc):
    """
    Verify that the get_document_structure tool returns the correct heading structure.
    """
    pythoncom.CoInitialize()
    mock_context = MockContext()
    
    try:
        # 1. Open the document
        app.open_document(mock_context, test_doc)

        # 2. Call the tool
        structure = app.get_document_structure(mock_context)

        # 3. Verify the structure
        assert isinstance(structure, list)
        assert len(structure) > 0
        
        heading_found = any(
            item.get("text") == "This is a heading." and item.get("level") == 1
            for item in structure
        )
        assert heading_found, "Did not find the expected 'Heading 1' in the structure."

    finally:
        # 4. Clean up
        app.shutdown_word(mock_context)
        pythoncom.CoUninitialize()
