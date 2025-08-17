import sys
import os
import pytest

# Add project root to the Python path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from word_document_server.selector import SelectorEngine, ElementNotFoundError
from word_document_server.com_backend import WordBackend

# --- Test Setup ---
TEST_DOC_DIR = os.path.dirname(os.path.abspath(__file__))
TEST_DOC_PATH = os.path.join(TEST_DOC_DIR, 'test_docs', 'test_document.docx')

@pytest.fixture(scope="module")
def test_doc_path():
    """Fixture to provide the path to the test document."""
    if not os.path.exists(TEST_DOC_PATH):
        pytest.fail(f"Test document not found at {TEST_DOC_PATH}. "
                    f"Please run create_test_doc.py first.")
    return TEST_DOC_PATH

# --- Test Class for Selector Engine ---

class TestSelectorEngineIntegration:
    """Integration tests for the SelectorEngine using a real Word document."""

    def test_select_by_positive_index(self, test_doc_path):
        """Verify selecting the first paragraph using index 0."""
        engine = SelectorEngine()
        locator = {"target": {"type": "paragraph", "filters": [{"index_in_parent": 0}]}}
        with WordBackend(file_path=test_doc_path, visible=False) as backend:
            selection = engine.select(backend, locator)
            assert len(selection._elements) == 1
            assert selection.get_text().strip() == "First paragraph."

    def test_select_by_negative_index(self, test_doc_path):
        """Verify selecting the last paragraph using index -1."""
        engine = SelectorEngine()
        locator = {"target": {"type": "paragraph", "filters": [{"index_in_parent": -1}]}}
        with WordBackend(file_path=test_doc_path, visible=False) as backend:
            selection = engine.select(backend, locator)
            assert len(selection._elements) == 1
            assert selection.get_text().strip() == "The last paragraph."

    def test_select_by_contains_text(self, test_doc_path):
        """Verify selecting a paragraph by substring containment."""
        engine = SelectorEngine()
        locator = {"target": {"type": "paragraph", "filters": [{"contains_text": "substring search"}]}}
        with WordBackend(file_path=test_doc_path, visible=False) as backend:
            selection = engine.select(backend, locator)
            assert len(selection._elements) == 1
            assert "substring search" in selection.get_text()

    def test_select_by_is_bold(self, test_doc_path):
        """Verify selecting a paragraph by bold formatting."""
        engine = SelectorEngine()
        locator = {"target": {"type": "paragraph", "filters": [{"is_bold": True}]}}
        with WordBackend(file_path=test_doc_path, visible=False) as backend:
            selection = engine.select(backend, locator)
            assert len(selection._elements) == 1
            assert selection.get_text().strip() == "This is a bold paragraph."

    def test_select_by_style(self, test_doc_path):
        """Verify selecting a paragraph by its style."""
        engine = SelectorEngine()
        locator = {"target": {"type": "paragraph", "filters": [{"style": "标题 1"}]}}
        with WordBackend(file_path=test_doc_path, visible=False) as backend:
            selection = engine.select(backend, locator)
            assert len(selection._elements) == 1
            assert selection.get_text().strip() == "This is a heading."

    def test_select_by_text_matches_regex(self, test_doc_path):
        """Verify selecting a paragraph by a regex pattern."""
        engine = SelectorEngine()
        locator = {"target": {"type": "paragraph", "filters": [{"text_matches_regex": r"unique_word_\d+"}]}}
        with WordBackend(file_path=test_doc_path, visible=False) as backend:
            selection = engine.select(backend, locator)
            assert len(selection._elements) == 1
            assert "unique_word_123" in selection.get_text()

    def test_select_with_multiple_filters(self, test_doc_path):
        """Verify selecting with multiple filters (bold and contains text)."""
        engine = SelectorEngine()
        locator = {
            "target": {
                "type": "paragraph",
                "filters": [
                    {"is_bold": True},
                    {"contains_text": "bold paragraph"}
                ]
            }
        }
        with WordBackend(file_path=test_doc_path, visible=False) as backend:
            selection = engine.select(backend, locator)
            assert len(selection._elements) == 1
            assert selection.get_text().strip() == "This is a bold paragraph."

    def test_element_not_found(self, test_doc_path):
        """Verify that a non-matching locator raises ElementNotFoundError."""
        engine = SelectorEngine()
        locator = {"target": {"type": "paragraph", "filters": [{"contains_text": "non_existent_text"}]}}
        with pytest.raises(ElementNotFoundError):
            with WordBackend(file_path=test_doc_path, visible=False) as backend:
                engine.select(backend, locator)

    def test_relation_first_occurrence_after(self, test_doc_path):
        """Verify selecting the first paragraph after a heading."""
        engine = SelectorEngine()
        locator = {
            "anchor": {
                "type": "paragraph",
                "filters": [{"contains_text": "This is a heading."}]
            },
            "target": {
                "type": "paragraph"
            },
            "relation": {
                "type": "first_occurrence_after"
            }
        }
        with WordBackend(file_path=test_doc_path, visible=False) as backend:
            selection = engine.select(backend, locator)
            assert len(selection._elements) == 1
            assert "unique_word_123" in selection.get_text()

    def test_relation_all_occurrences_within(self, test_doc_path):
        """Verify selecting all paragraphs within the first table."""
        engine = SelectorEngine()
        locator = {
            "anchor": {
                "type": "table",
                "filters": [{"index_in_parent": 0}]
            },
            "target": {
                "type": "paragraph"
            },
            "relation": {
                "type": "all_occurrences_within"
            }
        }
        with WordBackend(file_path=test_doc_path, visible=False) as backend:
            selection = engine.select(backend, locator)
            # The test table has 3 rows and 2 columns, with text in 5 cells.
            # Each cell's text is technically a paragraph.
            assert len(selection._elements) >= 5
            
            full_text = selection.get_text()
            assert "Table Cell 1-1" in full_text
            assert "Last cell." in full_text
            # Ensure text from outside the table is not included
            assert "First paragraph." not in full_text
            
    def test_replace_text(self, test_doc_path):
        """Verify replacing text in a selected element."""
        engine = SelectorEngine()
        locator = {"target": {"type": "paragraph", "filters": [{"contains_text": "First paragraph."}]}}
        with WordBackend(file_path=test_doc_path, visible=False) as backend:
            selection = engine.select(backend, locator)
            assert len(selection._elements) == 1
            
            # Replace the text
            selection.replace_text("Replaced text.")
            
            # Verify the text was replaced
            assert selection.get_text().strip() == "Replaced text."

if __name__ == "__main__":
    pytest.main(["-v", __file__])