"""
Comment operations for Word Document MCP Server.

This module contains functions for comment-related operations.
"""

import logging
from typing import Any, Dict, List, Optional

import win32com.client

from word_document_server.utils.core_utils import ErrorCode, WordDocumentError
from word_document_server.com_backend.com_utils import handle_com_error


# === Comment Creation Operations ===

@handle_com_error(ErrorCode.COMMENT_ERROR, "add comment")
def add_comment(
    document: win32com.client.CDispatch,
    com_range_obj: win32com.client.CDispatch,
    text: str,
    author: Optional[str] = None,
) -> Any:
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


# === Comment Retrieval Operations ===

@handle_com_error(ErrorCode.COMMENT_ERROR, "get comments")
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
    comments_count = document.Comments.Count
    for i in range(1, comments_count + 1):
        try:
            comment = document.Comments(i)
            comment_info = {
                "index": i - 1,  # 0-based index
                "text": comment.Text(),
                "author": comment.Author,
                "initials": comment.Initial,
                "date": str(comment.Date),
                "scope_start": comment.Scope.Start,
                "scope_end": comment.Scope.End,
                "scope_text": comment.Scope.Text.strip(),
                "replies_count": comment.Replies.Count if hasattr(comment, "Replies") else 0,
            }
            comments.append(comment_info)
        except Exception as e:
            logging.warning(f"Failed to retrieve comment at index {i}: {e}")
            continue

    return comments


@handle_com_error(ErrorCode.COMMENT_ERROR, "get comment thread")
def get_comment_thread(document: win32com.client.CDispatch, index: int) -> List[Dict[str, Any]]:
    """
    Retrieves a comment thread (a comment and its replies) by index.

    Args:
        document: The Word document COM object.
        index: The 0-based index of the comment.

    Returns:
        A list of dictionaries with comment thread details.
    """
    if not document:
        raise RuntimeError("No document open.")

    thread: List[Dict[str, Any]] = []
    # Get the comment at the specified index
    comment = document.Comments(index + 1)  # COM is 1-based
    # Add the main comment
    thread.append(
        {
            "index": index,
            "text": comment.Text(),
            "author": comment.Author,
            "initials": comment.Initial,
            "date": str(comment.Date),
            "scope_start": comment.Scope.Start,
            "scope_end": comment.Scope.End,
            "scope_text": comment.Scope.Text.strip(),
        }
    )
    # Add any replies
    replies_count = comment.Replies.Count if hasattr(comment, "Replies") else 0
    for i in range(1, replies_count + 1):
        reply = comment.Replies(i)
        thread.append(
            {
                "index": f"{index}-reply-{i-1}",
                "text": reply.Text(),
                "author": reply.Author,
                "initials": reply.Initial,
                "date": str(reply.Date),
            }
        )

    return thread


# === Comment Modification Operations ===

@handle_com_error(ErrorCode.COMMENT_ERROR, "delete comment")
def delete_comment(document: win32com.client.CDispatch, index: int) -> bool:
    """
    Deletes a comment at the specified index.

    Args:
        document: The Word document COM object.
        index: The 0-based index of the comment to delete.

    Returns:
        True if the deletion was successful.
    """
    # Get the comment at the specified index
    comment = document.Comments(index + 1)  # COM is 1-based
    # Delete the comment
    comment.Delete()
    return True


@handle_com_error(ErrorCode.COMMENT_ERROR, "delete all comments")
def delete_all_comments(document: win32com.client.CDispatch) -> Any:
    """
    Deletes all comments from the document.

    Args:
        document: The Word document COM object.

    Returns:
        The number of comments deleted.
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
        True if the edit was successful.
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
        True if the reply was successfully added.
    """
    # Get the comment at the specified index
    comment = document.Comments(index + 1)  # COM is 1-based
    # Add the reply
    reply = comment.Replies.Add(text, author)
    return True