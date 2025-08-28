"""
Comment operations for Word Document MCP Server.

This module contains functions for comment-related operations.
"""

from typing import Any, Dict, List, Optional

import win32com.client

from word_document_server.errors import ErrorCode, WordDocumentError
from word_document_server.utils.com_utils import handle_com_error, safe_com_call


@handle_com_error(ErrorCode.COMMENT_ERROR, "add comment")
def add_comment(
    document: win32com.client.CDispatch,
    com_range_obj: win32com.client.CDispatch,
    text: str,
    author: Optional[str] = None,
) -> win32com.client.CDispatch:
    """
    Adds a comment to the document at the specified range.

    Args:
        document: The Word document COM object.
        com_range_obj: The range to add the comment to.
        text: The comment text.
        author: Optional author name for the comment.

    Returns:
        The newly created comment COM object.
    """
    # Add the comment
    comment = document.Comments.Add(com_range_obj, text)
    
    # Set the author if provided
    if author:
        comment.Author = author
        
    return comment


def get_comments(document: win32com.client.CDispatch) -> List[Dict[str, Any]]:
    """
    Retrieves all comments from the document.

    Args:
        document: The Word document COM object.

    Returns:
        A list of dictionaries with comment details.
    """
    if not document:
        raise RuntimeError("No document open.")

    comments: List[Dict[str, Any]] = []
    try:
        comments_count = document.Comments.Count
        for i in range(1, comments_count + 1):
            try:
                comment = document.Comments(i)
                comment_info = {
                    "index": i - 1,  # 0-based index
                    "text": comment.Text(),
                    "author": comment.Author,
                    "initials": comment.Initial,
                    "start_pos": comment.Scope.Start,
                    "end_pos": comment.Scope.End,
                    "scope_text": comment.Scope.Text.strip(),
                    "date": str(comment.Date),
                    "is_virtual": comment.IsVirtual,
                    "replies_count": comment.Replies.Count if hasattr(comment, "Replies") else 0,
                }
                comments.append(comment_info)
            except Exception as e:
                print(f"Warning: Failed to retrieve comment at index {i}: {e}")
                continue
    except Exception as e:
        print(f"Error: Failed to retrieve comments: {e}")

    return comments


@handle_com_error(ErrorCode.COMMENT_ERROR, "delete comment")
def delete_comment(document: win32com.client.CDispatch, index: int) -> bool:
    """
    Deletes a comment at the specified index.

    Args:
        document: The Word document COM object.
        index: The 0-based index of the comment to delete.

    Returns:
        True if successful.
    """
    # Get the comment at the specified index
    comment = document.Comments(index + 1)  # COM is 1-based
    # Delete the comment
    comment.Delete()
    return True


@handle_com_error(ErrorCode.COMMENT_ERROR, "delete all comments")
def delete_all_comments(document: win32com.client.CDispatch) -> int:
    """
    Deletes all comments from the document.

    Args:
        document: The Word document COM object.

    Returns:
        Number of comments deleted.
    """
    count = document.Comments.Count
    # Delete all comments
    document.Comments.Delete()
    return count


@handle_com_error(ErrorCode.COMMENT_ERROR, "edit comment")
def edit_comment(document: win32com.client.CDispatch, index: int, new_text: str) -> bool:
    """
    Edits a comment at the specified index.

    Args:
        document: The Word document COM object.
        index: The 0-based index of the comment to edit.
        new_text: The new text for the comment.

    Returns:
        True if successful.
    """
    # Get the comment at the specified index
    comment = document.Comments(index + 1)  # COM is 1-based
    # Edit the comment
    comment.Range.Text = new_text
    return True


@handle_com_error(ErrorCode.COMMENT_ERROR, "reply to comment")
def reply_to_comment(document: win32com.client.CDispatch, index: int, text: str, author: Optional[str] = None) -> bool:
    """
    Adds a reply to a comment at the specified index.

    Args:
        document: The Word document COM object.
        index: The 0-based index of the comment to reply to.
        text: The reply text.
        author: Optional author name for the reply.

    Returns:
        True if successful.
    """
    # Get the comment at the specified index
    comment = document.Comments(index + 1)  # COM is 1-based
    # Add the reply
    reply = comment.Replies.Add(text, author)
    return True


def get_comment_thread(document: win32com.client.CDispatch, index: int) -> List[Dict[str, Any]]:
    """
    Retrieves a comment and all its replies as a thread.

    Args:
        document: The Word document COM object.
        index: The 0-based index of the comment to retrieve.

    Returns:
        A list of dictionaries with comment and reply details.
    """
    if not document:
        raise RuntimeError("No document open.")

    try:
        # Get the comment at the specified index
        comment = document.Comments(index + 1)  # COM is 1-based
        
        # Start with the main comment
        thread = [
            {
                "index": -1,  # Main comment has index -1 to distinguish from replies
                "text": comment.Text(),
                "author": comment.Author,
                "start_pos": comment.Scope.Start,
                "end_pos": comment.Scope.End,
                "scope_text": comment.Scope.Text,
            }
        ]
        
        # Add all replies
        replies_count = comment.Replies.Count
        for i in range(1, replies_count + 1):
            reply = comment.Replies(i)
            thread.append(
                {
                    "index": i,
                    "text": reply.Text,
                    "author": reply.Author,
                    "start_pos": reply.Scope.Start,
                    "end_pos": reply.Scope.End,
                    "scope_text": reply.Scope.Text,
                }
            )

        return thread
    except Exception as e:
        raise WordDocumentError(ErrorCode.COMMENT_ERROR, f"Failed to get comment thread: {e}")
