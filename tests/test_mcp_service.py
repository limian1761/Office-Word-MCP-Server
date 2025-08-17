import sys
import os
import pytest

# Add project root to Python path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from word_document_server.mcp_service import McpService
from word_document_server.com_backend import WordBackend
from word_document_server.selector import ElementNotFoundError

# Path to the test document
TEST_DOC_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), 'test_docs', 'test_document.docx'))

class TestMcpServiceIntegration:
    """Integration tests for the McpService high-level API."""

    def test_add_table(self):
        """Verify that a table can be added after a specified paragraph."""
        
        service = McpService()
        temp_file_path = os.path.join(os.path.dirname(TEST_DOC_PATH), "temp_test_doc_for_table.docx")

        try:
            # --- Perform all operations within a single Word instance ---
            with WordBackend(visible=False) as backend:
                # 1. Setup the document
                doc = backend.document
                p1 = doc.Paragraphs.Add()
                p1.Range.Text = "First paragraph for anchoring."
                doc.Paragraphs.Add() # Add a second one

                # 2. Define the task to be run (now just a simple function call)
                locator = {"target": {"type": "paragraph", "filters": [{"index_in_parent": 0}]}}
                rows, cols = 2, 3
                data = [
                    ["Header 1", "Header 2", "Header 3"],
                    ["Cell 1", "Cell 2", "Cell 3"]
                ]
                
                # 3. Execute the service method directly
                service.add_table(backend, locator, rows, cols, data)

                # 4. Verification
                tables = backend.get_all_tables()
                assert len(tables) == 1, "Exactly one table should be found."
                
                new_table = tables[0]
                assert new_table.Rows.Count == rows
                assert new_table.Columns.Count == cols
                # COM Text includes a special character at the end of a cell, so we check if it starts with the expected text.
                assert new_table.Cell(1, 1).Range.Text.startswith("Header 1")
                assert new_table.Cell(2, 3).Range.Text.startswith("Cell 3")

                # Save is not done by the service layer, so we don't check file persistence here.
                # The backend context manager will close without saving.

        finally:
            # Clean up the temporary file if it was created (it shouldn't be in this new logic)
            if os.path.exists(temp_file_path):
                os.remove(temp_file_path)

if __name__ == "__main__":
    pytest.main(["-v", __file__])
