import os
from word_document_server.com_backend import WordBackend
from word_document_server.selector import SelectorEngine
from word_document_server.selection import Selection

def debug_selector():
    current_dir = os.path.dirname(os.path.abspath(__file__))
    test_doc_path = os.path.join(current_dir, 'tests', 'test_docs', 'test_document.docx')
    
    engine = SelectorEngine()
    locator = {"target": {"type": "paragraph", "filters": [{"contains_text": "First paragraph."}]}}
    
    with WordBackend(file_path=test_doc_path, visible=False) as backend:
        # Get all paragraphs to see what's in the document
        all_paragraphs = backend.get_all_paragraphs()
        print("All paragraphs in the document:")
        for i, para in enumerate(all_paragraphs):
            print(f"  {i}: '{para.Range.Text.strip()}'")
        
        # Now test the selector
        selection = engine.select(backend, locator)
        print(f"\nSelected elements: {len(selection._elements)}")
        for i, element in enumerate(selection._elements):
            print(f"  {i}: '{element.Range.Text.strip()}'")
        
        # Get the text before replacement
        print(f"\nText before replacement: '{selection.get_text().strip()}'")
        
        # Replace the text
        selection.replace_text("Replaced text.")
        
        # Get the text after replacement
        print(f"\nText after replacement: '{selection.get_text().strip()}'")
        
        # Let's also check all paragraphs again
        all_paragraphs = backend.get_all_paragraphs()
        print("\nAll paragraphs in the document after replacement:")
        for i, para in enumerate(all_paragraphs):
            print(f"  {i}: '{para.Range.Text.strip()}'")

if __name__ == "__main__":
    debug_selector()