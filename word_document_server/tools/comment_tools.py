"""
Comment Integration Tool for Word Document MCP Server.

This module provides a unified MCP tool for comment operations.
"""

import os
from typing import Any, Dict, Optional

# Standard library imports
from dotenv import load_dotenv
# Third-party imports
from mcp.server.fastmcp import Context
from mcp.server.session import ServerSession

from pydantic import Field

# Local imports
from word_document_server.mcp_service.core import mcp_server
from word_document_server.selector.selector import SelectorEngine
from word_document_server.utils.app_context import AppContext
from word_document_server.utils.core_utils import (ErrorCode,
                                                   WordDocumentError,
                                                   get_active_document)

# 在函数内部导入以避免循环导入
def _import_comment_operations():
    """延迟导入comment操作函数以避免循环导入"""
    from word_document_server.operations.comment_ops import (add_comment,
                                                             delete_all_comments,
                                                             delete_comment,
                                                             edit_comment,
                                                             get_comment_thread,
                                                             get_comments,
                                                             reply_to_comment)
    return (add_comment, delete_all_comments, delete_comment, edit_comment, 
            get_comment_thread, get_comments, reply_to_comment)


# Load environment variables from .env file
load_dotenv()


@mcp_server.tool()
async def comment_tools(
    ctx: Context[ServerSession, AppContext] = Field(description="Context object"),
    operation_type: str = Field(
        ..., description="Type of comment operation to perform: add, delete, get_all, reply, get_thread, delete_all, edit"
    ),
    comment_text: Optional[str] = Field(
        default=None, description="Comment text for add and reply operations\n\n    Required for: add, reply, edit\n    "
    ),
    comment_id: Optional[str] = Field(
        default=None, description="Comment ID for delete and reply operations\n\n    Required for: delete, reply, get_thread, edit\n    "
    ),
    locator: Optional[Dict[str, Any]] = Field(
        default=None, description="Object locator for add operation\n\n    Required for: add\n    "
    ),
    author: Optional[str] = Field(
        default=None, description="Comment author for add operation\n\n    Optional for: add\n    "
    ),
) -> Any:
    """Unified comment operation tool.

    This tool provides a single interface for all comment operations:
    - add: Add a comment to an object
      * Required parameters: comment_text
      * Optional parameters: locator, author
    - delete: Delete a comment by ID
      * Required parameters: comment_id
    - get_all: Get all comments in the document
      * No required parameters
    - reply: Reply to an existing comment
      * Required parameters: comment_text, comment_id
    - get_thread: Get a specific comment thread
      * Required parameters: comment_id
    - delete_all: Delete all comments in the document
      * No required parameters
    - edit: Edit an existing comment
      * Required parameters: comment_text, comment_id

    Returns:
        Result of the operation
    """
    # Get the active Word document
    app_context = AppContext.get_instance()
    document = app_context.get_active_document()

    # 延迟导入comment操作函数以避免循环导入
    (add_comment, delete_all_comments, delete_comment, edit_comment, 
     get_comment_thread, get_comments, reply_to_comment) = _import_comment_operations()

    # Handle add comment operation
    if operation_type == "add":
        if comment_text is None:
            raise WordDocumentError(
                ErrorCode.INVALID_INPUT, "Comment text is required for add operation"
            )

        try:
            if locator:
                selector = SelectorEngine()
                selection = selector.select(document, locator)
            if not selection or not selection._com_ranges:
                raise WordDocumentError(
                    ErrorCode.OBJECT_NOT_FOUND,
                    "No object found matching the locator",
                )
            # Selection._com_ranges中只包含Range对象
            else:
                # Use current selection if no locator is provided
                range_obj = document.Application.Selection.Range

            result = add_comment(document, range_obj, comment_text, author)
            return {
                "success": True,
                "comment_id": str(result),
                "message": "Comment added successfully",
            }

        except Exception as e:
            raise WordDocumentError(
                ErrorCode.SERVER_ERROR, f"Failed to add comment: {str(e)}"
            )

    # Handle get all comments operation
    elif operation_type == "get_all":
        try:
            result = get_comments(document)
            return {
                "success": True,
                "comments": result,
                "message": "Comments retrieved successfully",
            }
        except Exception as e:
            raise WordDocumentError(
                ErrorCode.SERVER_ERROR, f"Failed to get comments: {str(e)}"
            )

    # Handle delete all comments operation
    elif operation_type == "delete_all":
        result = delete_all_comments(document)
        return {
            "success": True,
            "deleted_count": result,
            "message": (
                f"Successfully deleted {result} comments"
                if result > 0
                else "No comments to delete"
            ),
        }

    # Handle operations that require a comment_id
    elif operation_type in ["get_thread", "delete", "edit", "reply"]:
        if comment_id is None:
            raise WordDocumentError(
                ErrorCode.INVALID_INPUT,
                f"Comment ID is required for '{operation_type}' operation",
            )

        if operation_type == "get_thread":
            result = get_comment_thread(document, comment_id)
            return {
                "success": True,
                "thread": result,
                "message": "Comment thread retrieved successfully",
            }

        elif operation_type == "delete":
            result = delete_comment(document, comment_id)
            return {
                "success": result,
                "message": (
                    f"Comment {comment_id} deleted successfully"
                    if result
                    else "Failed to delete comment"
                ),
            }

        elif operation_type == "edit":
            if comment_text is None:
                raise WordDocumentError(
                    ErrorCode.INVALID_INPUT,
                    "Comment text is required for edit operation",
                )

            result = edit_comment(document, comment_id, comment_text)
            return {
                "success": result,
                "message": (
                    f"Comment {comment_id} edited successfully"
                    if result
                    else "Failed to edit comment"
                ),
            }

        elif operation_type == "reply":
            if comment_text is None:
                raise WordDocumentError(
                    ErrorCode.INVALID_INPUT,
                    "Comment text is required for reply operation",
                )

            result = reply_to_comment(document, comment_id, comment_text, author)
            return {
                "success": True,
                "reply_id": str(result),
                "message": "Reply added successfully",
            }

    # Handle unknown operation types
    else:
        raise WordDocumentError(
            ErrorCode.INVALID_INPUT, f"Unknown operation type: {operation_type}"
        )
