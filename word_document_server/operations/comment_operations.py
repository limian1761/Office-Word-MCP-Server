"""
Comment operations for Word Document MCP Server.

This module contains functions for comment-related operations.
"""
from typing import Any, Dict, List, Optional

import pythoncom
import win32com.client

from word_document_server.word_backend import WordBackend
from word_document_server.errors import WordDocumentError, ErrorCode

def add_comment(backend: WordBackend, com_range_obj: win32com.client.CDispatch, text: str, author: str = "User") -> win32com.client.CDispatch:
    """
    Adds a comment to the specified range.

    Args:
        backend: The WordBackend instance.
        com_range_obj: The COM Range object where the comment will be inserted.
        text: The text of the comment.
        author: The author of the comment (default: "User").

    Returns:
        The newly created Comment COM object.
    """
    if not backend.document:
        raise RuntimeError("No document open.")
    
    if not com_range_obj:
        raise ValueError("Invalid range object provided.")
    
    try:
        # Add a comment at the specified range
        return backend.document.Comments.Add(Range=com_range_obj, Text=text)
    except Exception as e:
        raise WordDocumentError(f"Failed to add comment: {e}")

def get_comments(backend: WordBackend) -> List[Dict[str, Any]]:
    """
    Retrieves all comments in the document.

    Args:
        backend: The WordBackend instance.

    Returns:
        A list of dictionaries containing comment information, each with "index", "text", "author", "start_pos", "end_pos", and "scope_text" keys.
    """
    if not backend.document:
        raise RuntimeError("No document open.")
        
    comments = []
    try:
        # Check if Comments property exists and is accessible
        if not hasattr(backend.document, 'Comments'):
            return comments
        
        # Get all comments from the document
        comments_count = 0
        try:
            comments_count = backend.document.Comments.Count
        except Exception as e:
            print(f"Warning: Failed to access Comments collection: {e}")
            return comments
        
        for i in range(1, comments_count + 1):
            try:
                comment = backend.document.Comments(i)
                try:
                    comment_info = {
                        "index": i - 1,  # 0-based index
                        "text": comment.Range.Text if hasattr(comment, 'Range') else "",
                        "author": comment.Author if hasattr(comment, 'Author') else "Unknown",
                        "start_pos": comment.Scope.Start if hasattr(comment, 'Scope') and hasattr(comment.Scope, 'Start') else 0,
                        "end_pos": comment.Scope.End if hasattr(comment, 'Scope') and hasattr(comment.Scope, 'End') else 0,
                        "scope_text": comment.Scope.Text.strip() if hasattr(comment, 'Scope') and hasattr(comment.Scope, 'Text') else ""
                    }
                    comments.append(comment_info)
                except Exception as e:
                    print(f"Warning: Failed to retrieve comment information for index {i}: {e}")
                    continue
            except Exception as e:
                print(f"Warning: Failed to access comment at index {i}: {e}")
                continue
    except Exception as e:
        print(f"Error: Failed to retrieve comments: {e}")
        
    return comments

def get_comments_by_range(backend: WordBackend, com_range_obj: win32com.client.CDispatch) -> List[Dict[str, Any]]:
    """
    Retrieves comments within a specific COM Range.

    Args:
        backend: The WordBackend instance.
        com_range_obj: The COM Range object to search within.

    Returns:
        A list of dictionaries containing comment information.
    """
    if not backend.document:
        raise RuntimeError("No document open.")
        
    if not com_range_obj:
        raise ValueError("Invalid range object provided.")
        
    comments = []
    try:
        # Check if Comments property exists and is accessible
        if not hasattr(backend.document, 'Comments'):
            return comments
        
        # Get all comments from the document
        comments_count = 0
        try:
            comments_count = backend.document.Comments.Count
        except Exception as e:
            print(f"Warning: Failed to access Comments collection: {e}")
            return comments
        
        # Check if the range object has Start and End properties
        if not hasattr(com_range_obj, 'Start') or not hasattr(com_range_obj, 'End'):
            print("Warning: Invalid range object - missing Start or End properties")
            return comments
        
        for i in range(1, comments_count + 1):
            try:
                comment = backend.document.Comments(i)
                try:
                    # Check if comment is within the specified range
                    if (hasattr(comment, 'Scope') and 
                        hasattr(comment.Scope, 'Start') and 
                        hasattr(comment.Scope, 'End') and 
                        comment.Scope.Start >= com_range_obj.Start and 
                        comment.Scope.End <= com_range_obj.End):
                        comment_info = {
                            "index": i - 1,  # 0-based index
                            "text": comment.Range.Text if hasattr(comment, 'Range') else "",
                            "author": comment.Author if hasattr(comment, 'Author') else "Unknown",
                            "start_pos": comment.Scope.Start if hasattr(comment, 'Scope') and hasattr(comment.Scope, 'Start') else 0,
                            "end_pos": comment.Scope.End if hasattr(comment, 'Scope') and hasattr(comment.Scope, 'End') else 0,
                            "scope_text": comment.Scope.Text.strip() if hasattr(comment, 'Scope') and hasattr(comment.Scope, 'Text') else ""
                        }
                        comments.append(comment_info)
                except Exception as e:
                    print(f"Warning: Failed to retrieve comment information for index {i}: {e}")
                    continue
            except Exception as e:
                print(f"Warning: Failed to access comment at index {i}: {e}")
                continue
    except Exception as e:
        print(f"Error: Failed to retrieve comments by range: {e}")
        
    return comments
    
def delete_comment(backend: WordBackend, comment_index: int) -> None:
    """
    Deletes a comment by its 0-based index.

    Args:
        backend: The WordBackend instance.
        comment_index: The 0-based index of the comment to delete.
    """
    if not backend.document:
        raise RuntimeError("No document open.")
        
    try:
        # Check if Comments property exists and is accessible
        if not hasattr(backend.document, 'Comments'):
            raise WordDocumentError("Comments collection is not available in this document.")
        
        # Get comments count safely
        comments_count = 0
        try:
            comments_count = backend.document.Comments.Count
        except Exception as e:
            raise WordDocumentError(f"Failed to access Comments collection: {e}")
        
        # Validate comment index
        if comment_index < 0 or comment_index >= comments_count:
            raise ValueError(f"Invalid comment index: {comment_index}. Valid range is 0 to {comments_count - 1}.")
        
        # Comments are 1-based in the COM API
        try:
            backend.document.Comments(comment_index + 1).Delete()
        except Exception as e:
            raise WordDocumentError(f"Failed to delete comment: {e}")
    except WordDocumentError:
        # Re-raise WordDocumentError to maintain consistency
        raise
    except Exception as e:
        raise WordDocumentError(f"Error during comment deletion: {e}")

def delete_all_comments(backend: WordBackend) -> int:
    """
    Deletes all comments in the document.
    
    Args:
        backend: The WordBackend instance.
        
    Returns:
        The number of comments deleted.
    """
    if not backend.document:
        raise RuntimeError("No document open.")
        
    try:
        # Check if Comments property exists and is accessible
        if not hasattr(backend.document, 'Comments'):
            # No comments to delete
            return 0
        
        # Get initial comments count safely
        comments_count = 0
        try:
            comments_count = backend.document.Comments.Count
        except Exception as e:
            raise WordDocumentError(f"Failed to access Comments collection: {e}")
        
        if comments_count == 0:
            # No comments to delete
            return 0
        
        # Store initial count for return value
        deleted_count = comments_count
        
        # Delete comments in reverse order to avoid index shifting issues
        try:
            for i in range(comments_count, 0, -1):
                try:
                    backend.document.Comments(i).Delete()
                except Exception as e:
                    print(f"Warning: Failed to delete comment at index {i}: {e}")
                    # Continue with next comment
                    continue
        except Exception as e:
            raise WordDocumentError(f"Failed to delete all comments: {e}")
        
        return deleted_count
    except WordDocumentError:
        # Re-raise WordDocumentError to maintain consistency
        raise
    except Exception as e:
        raise WordDocumentError(f"Error during deletion of all comments: {e}")

def edit_comment(backend: WordBackend, comment_index: int, new_text: str) -> None:
    """
    Edits an existing comment by its 0-based index.

    Args:
        backend: The WordBackend instance.
        comment_index: The 0-based index of the comment to edit.
        new_text: The new text for the comment.

    Raises:
        IndexError: If the comment index is out of range.
        WordDocumentError: If editing the comment fails.
    """
    if not backend.document:
        raise RuntimeError("No document open.")
    
    try:
        # Check if comment index is valid
        if comment_index < 0 or comment_index >= backend.document.Comments.Count:
            raise IndexError(f"Comment index {comment_index} out of range.")
        
        # Get the comment (COM is 1-based)
        comment = backend.document.Comments(comment_index + 1)
        
        # Update the comment text
        comment.Range.Text = new_text
    except IndexError:
        raise
    except Exception as e:
        raise WordDocumentError(f"Failed to edit comment: {e}")

def reply_to_comment(backend: WordBackend, comment_index: int, reply_text: str, author: str = "User") -> None:
    """
    Replies to an existing comment.

    Args:
        backend: The WordBackend instance.
        comment_index: The 0-based index of the comment to reply to.
        reply_text: The text of the reply.
        author: The author of the reply (default: "User").

    Raises:
        IndexError: If the comment index is out of range.
        WordDocumentError: If replying to the comment fails.
    """
    if not backend.document:
        raise RuntimeError("No document open.")
    
    try:
        # Check if comment index is valid
        if comment_index < 0 or comment_index >= backend.document.Comments.Count:
            raise IndexError(f"Comment index {comment_index} out of range.")
        
        # Get the comment (COM is 1-based)
        comment = backend.document.Comments(comment_index + 1)
        
        # Add a reply to the comment
        # Note: Word COM doesn't have a direct Reply method, so we need to
        # create a new comment at the same range as the original comment
        # and set the author accordingly
        reply = backend.document.Comments.Add(
            Range=comment.Scope, 
            Text=reply_text
        )
        
        # Set the author of the reply
        reply.Author = author
    except IndexError:
        raise
    except Exception as e:
        raise WordDocumentError(f"Failed to reply to comment: {e}")

def get_comment_thread(backend: WordBackend, comment_index: int) -> Dict[str, Any]:
    """
    Retrieves a comment thread including the original comment and all replies.

    Args:
        backend: The WordBackend instance.
        comment_index: The 0-based index of the original comment.

    Returns:
        A dictionary containing the original comment and all replies.

    Raises:
        IndexError: If the comment index is out of range.
        WordDocumentError: If retrieving the comment thread fails.
    """
    if not backend.document:
        raise RuntimeError("No document open.")
    
    try:
        # Check if comment index is valid
        if comment_index < 0 or comment_index >= backend.document.Comments.Count:
            raise IndexError(f"Comment index {comment_index} out of range.")
        
        # Get the original comment (COM is 1-based)
        original_comment = backend.document.Comments(comment_index + 1)
        
        # Get the range of the original comment's scope
        original_scope = original_comment.Scope
        
        # Create the result dictionary with the original comment
        result = {
            "original_comment": {
                "text": original_comment.Range.Text,
                "author": original_comment.Author,
                "date": original_comment.Date
            },
            "replies": []
        }
        
        # Search for replies to this comment
        # We consider a reply as any comment that shares the same scope as the original
        for i in range(1, backend.document.Comments.Count + 1):
            comment = backend.document.Comments(i)
            
            # Skip the original comment
            if comment.Index == original_comment.Index:
                continue
            
            # Check if this comment shares the same scope as the original
            # We compare the start and end positions of the scopes
            if (comment.Scope.Start == original_scope.Start and 
                comment.Scope.End == original_scope.End):
                result["replies"].append({
                    "text": comment.Range.Text,
                    "author": comment.Author,
                    "date": comment.Date
                })
        
        return result
    except IndexError:
        raise
    except Exception as e:
        raise WordDocumentError(f"Failed to get comment thread: {e}")