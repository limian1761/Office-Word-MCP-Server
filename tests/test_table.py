import unittest
import os
import sys
import unittest

# Add the project root to the Python path
current_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(current_dir)
sys.path.insert(0, project_root)

from word_document_server.tools.table import get_text_from_cell, set_cell_value, create_table
from word_document_server.tools.document import open_document, close_document, shutdown_word
from word_document_server.core_utils import MockContext

class TestTableTool(unittest.TestCase):
    def setUp(self):
        """Set up test fixtures before each test method."""
        # Create a mock context
        self.ctx = MockContext()
        
        # Get the absolute path to the test document
        self.test_doc_path = os.path.join(current_dir, 'test_docs', 'test_document.docx')
        
        # Ensure test document exists
        if not os.path.exists(self.test_doc_path):
            # Create a simple test document for testing
            temp_ctx = MockContext()
            result = open_document(temp_ctx, 'dummy.docx')  # This will create a new document
            try:
                close_document(temp_ctx)
                # Rename the dummy document to our test document
                if os.path.exists('dummy.docx'):
                    os.rename('dummy.docx', self.test_doc_path)
            except:
                pass
            finally:
                shutdown_word(temp_ctx)
        
        # Open the document
        result = open_document(self.ctx, self.test_doc_path)
        self.assertIn("Document opened successfully", result)

    def tearDown(self):
        """Tear down test fixtures after each test method."""
        # Ensure Word is shutdown after each test
        try:
            close_document(self.ctx)
        except:
            pass
        try:
            shutdown_word(self.ctx)
        except:
            pass

    def test_get_text_from_cell_with_no_table(self):
        """Test get_text_from_cell function when no table exists in document."""
        # Create a locator with unsupported element type
        locator = {
            "target": {
                "type": "nonexistent_type"
            }
        }
        
        # Try to get text from cell - should fail since type doesn't exist
        result = get_text_from_cell(self.ctx, locator)
        # Should return an error message
        self.assertIn("Error", result)

    def test_set_cell_value_with_no_table(self):
        """Test set_cell_value function when no table exists in document."""
        # Create a locator with unsupported element type
        locator = {
            "target": {
                "type": "nonexistent_type"
            }
        }
        
        # Try to set cell value - should fail since type doesn't exist
        result = set_cell_value(self.ctx, locator, "Test Value")
        # Should return an error message
        self.assertIn("Error", result)

    def test_create_table_valid(self):
        """Test create_table function with valid parameters."""
        # Create a locator for the start of document
        locator = {
            "target": {
                "type": "document_start"
            }
        }
        
        # Create a small table (2x2)
        result = create_table(self.ctx, locator, 2, 2)
        # Should return success message
        self.assertIn("Successfully created table", result)

    def test_create_table_invalid_rows(self):
        """Test create_table function with invalid rows parameter."""
        # Create a locator for the end of document
        locator = {
            "target": {
                "type": "end_of_document"
            }
        }
        
        # Try to create a table with invalid rows
        result = create_table(self.ctx, locator, -1, 2)
        self.assertIn("Invalid 'rows' parameter", result)
        
        result = create_table(self.ctx, locator, 0, 2)
        self.assertIn("Invalid 'rows' parameter", result)

    def test_create_table_invalid_cols(self):
        """Test create_table function with invalid cols parameter."""
        # Create a locator for the end of document
        locator = {
            "target": {
                "type": "end_of_document"
            }
        }
        
        # Try to create a table with invalid columns
        result = create_table(self.ctx, locator, 2, -1)
        self.assertIn("Invalid 'cols' parameter", result)
        
        result = create_table(self.ctx, locator, 2, 0)
        self.assertIn("Invalid 'cols' parameter", result)

    def test_create_table_exceeds_limits(self):
        """Test create_table function with parameters exceeding Word limits."""
        # Create a locator for the end of document
        locator = {
            "target": {
                "type": "end_of_document"
            }
        }
        
        # Try to create a table with too many rows
        result = create_table(self.ctx, locator, 50000, 2)
        self.assertIn("Table size exceeds Word's practical limits", result)
        
        # Try to create a table with too many columns
        result = create_table(self.ctx, locator, 2, 100)
        self.assertIn("Table size exceeds Word's practical limits", result)

    def test_create_table_no_document(self):
        """Test create_table function when no document is open."""
        # Close the document first
        shutdown_word(self.ctx)
        
        # Create a locator for the end of document
        locator = {
            "target": {
                "type": "end_of_document"
            }
        }
        
        # Try to create a table without an open document
        result = create_table(self.ctx, locator, 2, 2)
        self.assertIn("No active document", result)

if __name__ == '__main__':
    unittest.main()