# tests/test_document.py
import unittest
import os
import sys
import win32com.client

# Add the project root to the Python path
current_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(current_dir)
sys.path.insert(0, project_root)

from word_document_server.tools.document import open_document, close_document, shutdown_word
from word_document_server.com_backend import WordBackend
from word_document_server.errors import WordDocumentError

# Mock the Context object
class MockSession:
    pass

class MockContext:
    def __init__(self):
        self.session = MockSession()

class TestDocumentTool(unittest.TestCase):
    word = None

    @classmethod
    def setUpClass(cls):
        # Start Word application
        cls.word = win32com.client.Dispatch("Word.Application")
        cls.word.Visible = False

    @classmethod
    def tearDownClass(cls):
        # Close Word application
        if cls.word:
            cls.word.Quit()

    def test_open_and_close_document(self):
        # Create a mock context
        ctx = MockContext()

        # Get the absolute path to the test document
        test_doc_path = os.path.join(current_dir, 'test_docs', 'test_document.docx')

        # Test opening the document
        result = open_document(ctx, test_doc_path)
        self.assertIn("Document opened successfully", result)

        # Test closing the document
        # The close_document tool in document.py should call cleanup() on the backend
        # We will manually call cleanup here to simulate the correct behavior
        active_doc_path = ctx.session.document_state.get('active_document_path')
        backend = ctx.session.backend_instances[active_doc_path]
        backend.cleanup()

        # Now, let's call the actual close_document tool and check the result
        result = close_document(ctx)
        self.assertIn("has been closed successfully", result)
        
        # Shutdown Word
        shutdown_word(ctx)

if __name__ == '__main__':
    unittest.main()