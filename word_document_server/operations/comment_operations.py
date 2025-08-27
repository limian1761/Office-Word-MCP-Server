"""
Comment operations for Word Document MCP Server.

This module contains functions for comment-related operations.
"""

from typing import Any, Dict, List, Optional

import pythoncom
import win32com.client

from word_document_server.errors import ErrorCode, WordDocumentError
from word_document_server.word_backend import WordBackend


def add_comment(
    backend: WordBackend,
    com_range_obj: win32com.client.CDispatch,
    text: str,
    author: str = "User",
) -> win32com.client.CDispatch:
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
        raise WordDocumentError(ErrorCode.COMMENT_ERROR, f"Failed to add comment: {e}")


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

    comments: List[Dict[str, Any]] = []
    try:
        # Check if Comments property exists and is accessible
        if not hasattr(backend.document, "Comments"):
            return comments

        # Get all comments from the document
        comments_count = 0
        try:
            comments_count = backend.document.Comments.Count
        except Exception as e:
            raise WordDocumentError(ErrorCode.COMMENT_ERROR, f"Failed to access Comments collection: {e}")

        for i in range(1, comments_count + 1):
            comment = backend.document.Comments(i)
            comments.append(
                {
                    "index": i,
                    "text": comment.Text,
                    "author": comment.Author,
                    "start_pos": comment.Scope.Start,
                    "end_pos": comment.Scope.End,
                    "scope_text": comment.Scope.Text,
                }
            )
    except Exception as e:
        raise WordDocumentError(ErrorCode.COMMENT_ERROR, f"Failed to retrieve comments: {e}")

    return comments


def delete_all_comments(backend: WordBackend) -> None:
    """
    Deletes all comments in the document.

    Args:
        backend: The WordBackend instance.

    Returns:
        None
    """
    if not backend.document:
        raise RuntimeError("No document open.")

    try:
        backend.document.Comments.DeleteAll()
    except Exception as e:
        raise WordDocumentError(ErrorCode.COMMENT_ERROR, f"Error during deletion of all comments: {e}")


def edit_comment(
    backend: WordBackend,
    comment_index: int,
    new_text: str,
) -> None:
    """
    Edits the text of an existing comment.

    Args:
        backend: The WordBackend instance.
        comment_index: The index of the comment to edit.
        new_text: The new text for the comment.

    Returns:
        None
    """
    if not backend.document:
        raise RuntimeError("No document open.")

    try:
        comment = backend.document.Comments(comment_index)
        comment.Text = new_text
    except Exception as e:
        raise WordDocumentError(ErrorCode.COMMENT_ERROR, f"Failed to edit comment: {e}")


def reply_to_comment(
    backend: WordBackend,
    comment_index: int,
    reply_text: str,
    author: str = "User",
) -> None:
    """
    Replies to an existing comment.

    Args:
        backend: The WordBackend instance.
        comment_index: The index of the comment to reply to.
        reply_text: The text of the reply.
        author: The author of the reply (default: "User").

    Returns:
        None
    """
    if not backend.document:
        raise RuntimeError("No document open.")

    try:
        comment = backend.document.Comments(comment_index)
        comment.Replies.Add(Text=reply_text, Author=author)
    except Exception as e:
        raise WordDocumentError(ErrorCode.COMMENT_ERROR, f"Failed to reply to comment: {e}")


def get_comment_thread(
    backend: WordBackend,
    comment_index: int,
) -> List[Dict[str, Any]]:
    """
    Retrieves the thread of a comment, including the comment and all replies.

    Args:
        backend: The WordBackend instance.
        comment_index: The index of the comment.

    Returns:
        A list of dictionaries containing comment information, each with "index", "text", "author", "start_pos", "end_pos", and "scope_text" keys.
    """
    if not backend.document:
        raise RuntimeError("No document open.")

    try:
        comment = backend.document.Comments(comment_index)
        thread = [
            {
                "index": comment_index,
                "text": comment.Text,
                "author": comment.Author,
                "start_pos": comment.Scope.Start,
                "end_pos": comment.Scope.End,
                "scope_text": comment.Scope.Text,
            }
        ]

        for i in range(1, comment.Replies.Count + 1):
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
