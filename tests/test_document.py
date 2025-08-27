# tests/test_document.py
import unittest
import os
import sys

# Add the project root to the Python path
current_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(current_dir)
sys.path.insert(0, project_root)

from word_document_server.tools.document import open_document, close_document, shutdown_word
from word_document_server.core_utils import MockContext
from word_document_server.errors import WordDocumentError

class TestDocumentTool(unittest.TestCase):
    def setUp(self):
        # Create a mock context
        self.ctx = MockContext()

    def tearDown(self):
        # Ensure Word is shutdown after each test
        try:
            shutdown_word(self.ctx)
        except:
            pass

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
        result = close_document(ctx)
        self.assertIn("closed successfully", result)
        
        # Shutdown Word
        shutdown_word(ctx)

if __name__ == '__main__':
    unittest.main()