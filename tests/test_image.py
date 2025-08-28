import pytest
import os
import sys
import json

# Add the project root to the Python path
current_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(current_dir)
sys.path.insert(0, project_root)

import pytest
from word_document_server.tools.image import add_caption, get_image_info, insert_object
from word_document_server.tools.document import open_document, close_document, shutdown_word
from word_document_server.utils.core_utils import MockContext

@pytest.fixture
def image_test_setup():
    # Mock the Context object
    ctx = MockContext()
    
    # Get the absolute path to the test document
    test_doc_path = os.path.join(current_dir, 'test_docs', 'image_test_doc.docx')
    
    # Ensure test document exists
    if not os.path.exists(test_doc_path):
        # Create a simple test document for testing
        # Using WordBackend to create the document would be better
        # but for simplicity, we'll just use a workaround for testing
        temp_ctx = MockContext()
        result = open_document(temp_ctx, test_doc_path)  # Create document with correct path
        try:
            close_document(temp_ctx)
        except:
            pass
        finally:
            shutdown_word(temp_ctx)
    
    # Open the test document
    open_document(ctx, test_doc_path)
    
    yield ctx
    
    # Test cleanup
    try:
        close_document(ctx)
    except:
        pass
    shutdown_word(ctx)


def test_get_image_info(image_test_setup):
    ctx = image_test_setup
    # Test getting image info from a document without images
    result = get_image_info(ctx)
    try:
        result_data = json.loads(result)
        assert isinstance(result_data, list)
    except json.JSONDecodeError:
        # If it's not valid JSON, it's likely an error message
        assert "Error" in result


def test_insert_inline_picture(image_test_setup):
    ctx = image_test_setup
    # Insert a paragraph first to have content
    from word_document_server.tools.text import insert_paragraph
    locator = {"target": {"type": "document_start"}}
    insert_paragraph(ctx, locator=locator, text="Test paragraph for image insertion", position="after")
    
    # Try to insert an image (this will fail without a valid image path, but we can test the error handling)
    locator = {"target": {"type": "document_end"}}
    # Use the correct function name
    result = insert_object(ctx, locator=locator, object_path="nonexistent.jpg")
    # Should return an error message about file not found
    assert "Error" in result


def test_add_picture_caption(image_test_setup):
    ctx = image_test_setup
    # Test adding a caption to a non-existent picture
    result = add_caption(ctx, locator={"target": {"type": "document_end"}}, caption_text="Test caption", position="below")
    # Should return an error message
    assert "Error" in result


def test_image_operation_errors(image_test_setup):
    ctx = image_test_setup
    # Close document to test error handling
    close_document(ctx)
    
    # Test that operations return error messages when no document is open
    result = get_image_info(ctx)
    assert "Error" in result

    
if __name__ == '__main__':
    pytest.main(["-v", __file__])
