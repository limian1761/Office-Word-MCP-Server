import pytest
import os
import sys
import win32com.client

# Add the project root to the Python path
current_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(current_dir)
sys.path.insert(0, project_root)

from word_document_server.tools.image import insert_object
from word_document_server.tools.document import open_document, close_document, shutdown_word
from word_document_server.errors import WordDocumentError
from word_document_server.core_utils import MockContext

@pytest.fixture
def image_test_setup():
    # Mock the Context object
    ctx = MockContext()
    
    # Start Word application for testing
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    
    # Get the absolute path to the test document
    test_doc_path = os.path.join(current_dir, 'test_docs', 'image_test_doc.docx')
    
    # Ensure test document exists
    if not os.path.exists(test_doc_path):
        # Create a simple test document for testing
        doc = word.Documents.Add()
        doc.SaveAs(test_doc_path)
        doc.Close()
    
    # Open the test document
    open_document(ctx, test_doc_path)
    
    yield ctx, word
    
    # Test cleanup
    try:
        close_document(ctx)
    except:
        pass
    shutdown_word(ctx)
    
    # Close Word application
    word.Quit()

@pytest.fixture
def image_test_setup():
    # Mock the Context object
    ctx = MockContext()
    
    # Start Word application for testing
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    
    # Get the absolute path to the test document
    test_doc_path = os.path.join(current_dir, 'test_docs', 'image_test_doc.docx')
    
    # Ensure test document exists
    if not os.path.exists(test_doc_path):
        # Create a simple test document for testing
        doc = word.Documents.Add()
        doc.SaveAs(test_doc_path)
        doc.Close()
    
    # Open the test document
    open_document(ctx, test_doc_path)
    
    yield ctx, word
    
    # Test cleanup
    try:
        close_document(ctx)
    except:
        pass
    shutdown_word(ctx)
    
    # Close Word application
    word.Quit()


def test_get_image_info(image_test_setup):
    ctx = image_test_setup
    # Test getting image info from a document without images
    result = get_image_info(ctx)
    result_data = json.loads(result)
    assert isinstance(result_data, list)


def test_insert_inline_picture(image_test_setup):
    ctx = image_test_setup
    # Insert a paragraph first to have content
    from word_document_server.tools.text import insert_paragraph
    locator = {"target": {"type": "document_start"}}
    insert_paragraph(ctx, locator=locator, text="Test paragraph for image insertion", position="after")
    
    # Try to insert an image (this will fail without a valid image path, but we can test the error handling)
    locator = {"target": {"type": "text", "value": "Test paragraph"}}
    result = insert_inline_picture(ctx, locator=locator, image_path="nonexistent.jpg")
    # Should return an error message about file not found
    assert "Error" in result


def test_set_image_size(image_test_setup):
    ctx = image_test_setup
    # Test setting image size on a document without images
    # This should return an error since no images are found
    locator = {"target": {"type": "document_start"}}
    result = set_image_size(ctx, locator=locator, width=100.0, height=100.0)
    assert "Error" in result


def test_set_image_color_type(image_test_setup):
    ctx = image_test_setup
    # Test setting image color type on a document without images
    # This should return an error since no images are found
    locator = {"target": {"type": "document_start"}}
    result = set_image_color_type(ctx, locator=locator, color_type="Grayscale")
    assert "Error" in result


def test_delete_image(image_test_setup):
    ctx = image_test_setup
    # Test deleting images from a document without images
    # This should return an error since no images are found
    locator = {"target": {"type": "document_start"}}
    result = delete_image(ctx, locator=locator)
    assert "Error" in result


def test_add_picture_caption(image_test_setup):
    ctx = image_test_setup
    # Test adding a caption to a non-existent picture
    result = add_picture_caption(ctx, filename="nonexistent.jpg", caption_text="Test caption")
    # Should return an error message
    assert "Error" in result


def test_image_operation_errors(image_test_setup):
    ctx = image_test_setup
    # Close document to test error handling
    close_document(ctx)
    
    # Test that operations return error messages when no document is open
    result = get_image_info(ctx)
    assert "Error" in result
    
    locator = {"target": {"type": "document_start"}}
    result = set_image_size(ctx, locator=locator, width=100.0)
    assert "Error" in result
    
    result = set_image_color_type(ctx, locator=locator, color_type="Grayscale")
    assert "Error" in result

if __name__ == '__main__':
    pytest.main(["-v", __file__])
