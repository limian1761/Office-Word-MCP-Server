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