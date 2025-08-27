# tests/debug_comment_tools.py
import os
import sys
import json

# Add the project root to the Python path
current_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(current_dir)
sys.path.insert(0, project_root)

from word_document_server.com_backend import WordBackend
from word_document_server.core_utils import MockSession, MockContext
from word_document_server.tools.comment import add_comment, get_comments, delete_comment, delete_all_comments, edit_comment, reply_to_comment, get_comment_thread

print("Starting debug test for comment tools...")

try:
    # Get the absolute path to the test document
    test_doc_path = os.path.join(current_dir, 'test_docs', 'valid_test_document_v2.docx')
    print(f"Test document path: {test_doc_path}")
    
    # Create a WordBackend instance directly
    print("Creating WordBackend instance...")
    backend = WordBackend(file_path=test_doc_path, visible=False)
    backend.__enter__()
    
    # Create a mock context (it will create its own MockSession)
    ctx = MockContext()
    
    # Add the backend to the session using the correct attributes
    ctx.session.backend_instances[test_doc_path] = backend
    ctx.session.document_state['active_document_path'] = test_doc_path
    
    # Delete all existing comments to start fresh
    print("Deleting all existing comments...")
    delete_all_comments(ctx)
    
    # Debug 1: Try add_comment with paragraph locator (correct format)
    print("\nDebug 1: Testing add_comment with correct paragraph locator...")
    try:
        # Use the correct locator format with 'target' and 'filters'
        locator = {
            "target": {
                "type": "paragraph",
                "filters": [{"index_in_parent": 0}]
            }
        }
        comment_text = "Test comment with paragraph locator"
        author = "Debugger"
        
        result = add_comment(ctx, locator, comment_text, author)
        print(f"add_comment result: {result}")
        
        # Check if comment was added
        comments_json = get_comments(ctx)
        print(f"get_comments result: {comments_json}")
        comments = json.loads(comments_json)
        print(f"Number of comments: {len(comments)}")
        if len(comments) > 0:
            print(f"Added comment: {comments[0]}")
    except Exception as e:
        print(f"Error in add_comment test: {str(e)}")
    
    # Delete all comments for next test
    delete_all_comments(ctx)
    
    # Debug 2: Try add_comment with simplified range approach
    print("\nDebug 2: Testing add_comment with simplified approach...")
    try:
        # Let's bypass the locator for now and add a comment directly through the backend
        # to verify that get_comments works correctly
        doc_range = backend.document.Range(0, 10)
        comment = backend.document.Comments.Add(Range=doc_range, Text="Backend added comment")
        comment.Author = "BackendDebugger"
        backend.document.Save()
        print(f"Comment added directly through backend: {comment.Range.Text} by {comment.Author}")
        
        # Check if comment was added using get_comments
        comments_json = get_comments(ctx)
        print(f"get_comments result: {comments_json}")
        comments = json.loads(comments_json)
        print(f"Number of comments: {len(comments)}")
        if len(comments) > 0:
            print(f"Retrieved comment: {comments[0]}")
            
            # Try deleting this comment using delete_comment tool
            delete_result = delete_comment(ctx, 0)
            print(f"delete_comment result: {delete_result}")
            
            # Verify deletion
            comments_json_after = get_comments(ctx)
            print(f"get_comments after deletion: {comments_json_after}")
            comments_after = json.loads(comments_json_after)
            print(f"Number of comments after deletion: {len(comments_after)}")
    except Exception as e:
        print(f"Error in simplified test: {str(e)}")
    
    # Delete all comments for next test
    delete_all_comments(ctx)
    
    # Debug 3: Try adding a comment directly through backend and then using get_comments
    print("\nDebug 3: Adding comment directly through backend and testing get_comments...")
    try:
        # Add a comment directly
        doc_range = backend.document.Range(0, 10)
        comment = backend.document.Comments.Add(Range=doc_range, Text="Direct COM comment")
        comment.Author = "DirectDebugger"
        backend.document.Save()
        print(f"Direct comment added: {comment.Range.Text} by {comment.Author}")
        
        # Use get_comments tool function
        comments_json = get_comments(ctx)
        print(f"get_comments result: {comments_json}")
        comments = json.loads(comments_json)
        print(f"Number of comments: {len(comments)}")
        if len(comments) > 0:
            print(f"Retrieved comment: {comments[0]}")
    except Exception as e:
        print(f"Error in direct comment test: {str(e)}")
    
    # Clean up
    print("\nCleaning up resources...")
    delete_all_comments(ctx)
    backend.cleanup()
    print("Debug test for comment tools completed.")
    
except Exception as e:
    print(f"Unexpected error in debug test: {str(e)}")
    import traceback
    traceback.print_exc()