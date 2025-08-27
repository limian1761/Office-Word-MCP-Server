# tests/debug_comments.py
import os
import sys
import win32com.client

# Add the project root to the Python path
current_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(current_dir)
sys.path.insert(0, project_root)

from word_document_server.com_backend import WordBackend

print("Starting debug test for comments...")

try:
    # Get the absolute path to the test document
    test_doc_path = os.path.join(current_dir, 'test_docs', 'valid_test_document_v2.docx')
    print(f"Test document path: {test_doc_path}")
    
    # Create a WordBackend instance directly
    print("Creating WordBackend instance...")
    backend = WordBackend(file_path=test_doc_path, visible=False)
    backend.__enter__()
    
    print("WordBackend initialized successfully.")
    
    # Check if document is loaded
    if backend.document is None:
        print("Error: Document not loaded!")
    else:
        print(f"Document loaded successfully: {backend.document.Name}")
        
        # Delete all existing comments to start fresh
        print("Deleting all existing comments...")
        try:
            if hasattr(backend.document, 'Comments'):
                count = backend.document.Comments.Count
                print(f"Found {count} existing comments.")
                # Delete comments in reverse order
                for i in range(count, 0, -1):
                    backend.document.Comments(i).Delete()
                print("All existing comments deleted.")
            else:
                print("Document has no Comments collection.")
        except Exception as e:
            print(f"Error deleting comments: {str(e)}")
        
        # Try to add a comment directly using COM
        print("\nAttempting to add a comment directly using COM...")
        try:
            # Get a range in the document
            doc_range = backend.document.Range(0, 10)  # First 10 characters
            print(f"Got document range: {doc_range.Text}")
            
            # Add a comment
            comment = backend.document.Comments.Add(Range=doc_range, Text="This is a debug comment")
            comment.Author = "Debugger"
            print(f"Comment added successfully: {comment.Range.Text} by {comment.Author}")
            
            # Save the document
            backend.document.Save()
            print("Document saved after adding comment.")
            
            # Check if the comment was really added
            if hasattr(backend.document, 'Comments'):
                count = backend.document.Comments.Count
                print(f"After adding, found {count} comments in the document.")
                if count > 0:
                    added_comment = backend.document.Comments(1)
                    print(f"Added comment text: {added_comment.Range.Text}")
                    print(f"Added comment author: {added_comment.Author}")
            
            # Clean up
            print("Cleaning up - deleting the test comment...")
            comment.Delete()
            print("Test comment deleted.")
        except Exception as e:
            print(f"Error adding comment via COM: {str(e)}")
            
    # Clean up
    print("\nCleaning up resources...")
    backend.cleanup()
    print("Debug test completed.")
    
except Exception as e:
    print(f"Unexpected error in debug test: {str(e)}")
    import traceback
    traceback.print_exc()