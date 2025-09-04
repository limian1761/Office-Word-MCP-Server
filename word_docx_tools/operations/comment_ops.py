"""
Comment operations for Word Document MCP Server.

This module contains functions for comment-related operations.
"""

import logging
from typing import Any, Dict, List, Optional

import win32com.client

from ..com_backend.com_utils import handle_com_error
from ..mcp_service.core_utils import (ErrorCode, WordDocumentError, log_error,
                                       log_info)

logger = logging.getLogger(__name__)

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
            # 创建一个基本的评论信息字典，只包含必要的属性
            comment_info = {
                "index": i - 1,  # 0-based index
                "replies_count": 0
            }
            
            # 尝试获取每个属性，使用try-except包装每个属性访问
            # 尝试多种方法获取Text属性
            try:
                # 方法1: 直接访问Text属性
                comment_info["text"] = str(comment.Text)
            except Exception as e1:
                try:
                    # 方法2: 通过Range属性获取Text
                    if hasattr(comment, "Range"):
                        comment_info["text"] = str(comment.Range.Text)
                    else:
                        raise AttributeError("Range attribute not found")
                except Exception as e2:
                    try:
                        # 方法3: 使用Get_Text()方法（如果存在）
                        if hasattr(comment, "Get_Text") and callable(comment.Get_Text):
                            comment_info["text"] = str(comment.Get_Text())
                        else:
                            raise AttributeError("Get_Text method not found")
                    except Exception as e3:
                        logging.warning(f"Failed to get Text for comment {i} using multiple methods: {e1}, {e2}, {e3}")
                        comment_info["text"] = "[Unable to retrieve text]"
                
            try:
                comment_info["author"] = str(comment.Author)
            except Exception as e:
                logging.warning(f"Failed to get Author for comment {i}: {e}")
                comment_info["author"] = "[Unknown]"
                
            try:
                comment_info["initials"] = str(comment.Initial)
            except Exception as e:
                logging.warning(f"Failed to get Initial for comment {i}: {e}")
                comment_info["author_initial"] = ""
                
            try:
                comment_info["date"] = str(comment.Date)
            except Exception as e:
                logging.warning(f"Failed to get Date for comment {i}: {e}")
                comment_info["date"] = "[Unknown date]"
                
            # 尝试获取Scope属性
            try:
                if hasattr(comment, "Scope") and comment.Scope:
                    scope_info = {
                        "start": comment.Scope.Start,
                        "end": comment.Scope.End,
                        "text": comment.Scope.Text.strip()
                    }
                    comment_info["scope"] = scope_info
            except Exception as e:
                logging.warning(f"Failed to get Scope for comment {i}: {e}")
                
            # 尝试获取Replies.Count
            try:
                if hasattr(comment, "Replies"):
                    comment_info["replies_count"] = comment.Replies.Count
            except Exception as e:
                logging.warning(f"Failed to get Replies for comment {i}: {e}")
                
            # 无论如何都添加评论信息，即使某些属性无法访问
            comments.append(comment_info)
            
        except Exception as e:
            logging.warning(f"Failed to retrieve comment at index {i}: {e}")
            # 仍然添加一个基本的评论信息，以便至少知道有这个评论存在
            comments.append({
                "index": i - 1,
                "text": "[Error retrieving comment]",
                "author": "[Unknown]",
                "replies_count": 0
            })
            continue

    return comments


@handle_com_error(ErrorCode.COMMENT_ERROR, "get comment thread")
def get_comment_thread(
    document: win32com.client.CDispatch, index: int
) -> List[Dict[str, Any]]:
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
            "text": comment.Range.Text,
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
                "text": reply.Range.Text,
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
    # Delete all comments by iterating backwards
    for i in range(count, 0, -1):
        try:
            comment = document.Comments(i)
            comment.Delete()
        except Exception as e:
            logger.warning(f"Failed to delete comment at index {i}: {e}")
            continue
    return count


@handle_com_error(ErrorCode.COMMENT_ERROR, "edit comment")
def edit_comment(
    document: win32com.client.CDispatch, index: int, new_text: str
) -> bool:
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
def reply_to_comment(
    document: win32com.client.CDispatch,
    index: int,
    text: str,
    author: Optional[str] = None,
) -> bool:
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
    reply = comment.Replies.Add(comment.Scope, text, author)
    return True
