"""
Comment Integration Tool for Word Document MCP Server.

This module provides a unified MCP tool for comment operations.
"""

import os
from typing import Any, Dict, Optional, Union

# Standard library imports
from dotenv import load_dotenv
# Third-party imports
from mcp.server.fastmcp import Context
from mcp.server.session import ServerSession
from pydantic import Field

# Local imports
from ..mcp_service.core import mcp_server
from ..mcp_service.core_utils import (ErrorCode, WordDocumentError,
                                      get_active_document)
from ..selector.selector import SelectorEngine
from ..mcp_service.app_context import AppContext


# 在函数内部导入以避免循环导入
def _import_comment_operations():
    """延迟导入comment操作函数以避免循环导入"""
    from ..operations.comment_ops import (add_comment, delete_all_comments,
                                          delete_comment, edit_comment,
                                          get_comment_thread, get_comments,
                                          reply_to_comment)

    return (
        add_comment,
        delete_all_comments,
        delete_comment,
        edit_comment,
        get_comment_thread,
        get_comments,
        reply_to_comment,
    )


# Load environment variables from .env file
load_dotenv()


@mcp_server.tool()
async def comment_tools(
    ctx: Context[ServerSession, AppContext] = Field(description="Context object"),
    operation_type: str = Field(
        ...,
        description="Type of comment operation to perform: add, delete, get_all, reply, get_thread, delete_all, edit",
    ),
    comment_text: Optional[str] = Field(
        default=None,
        description="Comment text for add and reply operations\n\n    Required for: add, reply, edit\n    ",
    ),
    comment_id: Optional[Union[str, int]] = Field(
        default=None,
        description="Comment ID for delete and reply operations (can be string or integer)\n\n    Required for: delete, reply, get_thread, edit\n    ",
    ),
    locator: Optional[Dict[str, Any]] = Field(
        default=None,
        description="Object locator for add operation\n\n    Required for: add\n    ",
    ),
    author: Optional[str] = Field(
        default=None,
        description="Comment author for add operation\n\n    Optional for: add\n    ",
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
    document = ctx.request_context.lifespan_context.get_active_document()

    # 延迟导入comment操作函数以避免循环导入
    (
        add_comment,
        delete_all_comments,
        delete_comment,
        edit_comment,
        get_comment_thread,
        get_comments,
        reply_to_comment,
    ) = _import_comment_operations()

    # 恢复正确的缩进结构
    # Handle add comment operation
    if operation_type == "add":
        if comment_text is None:
            raise WordDocumentError(
                ErrorCode.INVALID_INPUT, "Comment text is required for add operation"
            )

        try:
            # 先尝试使用locator定位
            if locator:
                try:
                    selector = SelectorEngine()
                    selection = selector.select(document, locator)
                    if (
                        selection
                        and hasattr(selection, "_com_ranges")
                        and selection._com_ranges
                    ):
                        range_obj = selection._com_ranges[0]  # 使用第一个Range对象
                    else:
                        # 如果locator定位失败，使用文档末尾作为默认位置
                        if document and hasattr(document, "Content"):
                            range_obj = document.Content
                            range_obj.Collapse(0)  # 折叠到文档末尾
                        else:
                            raise WordDocumentError(
                                ErrorCode.SERVER_ERROR,
                                "Document object is invalid or missing Content attribute",
                            )
                except Exception as e:
                    # locator定位失败，使用文档末尾作为默认位置
                    if document and hasattr(document, "Content"):
                        range_obj = document.Content
                        range_obj.Collapse(0)  # 折叠到文档末尾
                    else:
                        raise WordDocumentError(
                            ErrorCode.SERVER_ERROR,
                            "Document object is invalid or missing Content attribute",
                        )
            else:
                # 没有提供locator，使用文档末尾作为默认位置
                if document is None:
                    raise WordDocumentError(
                        ErrorCode.SERVER_ERROR, "No active document found"
                    )
                if hasattr(document, "Content"):
                    range_obj = document.Content
                    range_obj.Collapse(0)  # 折叠到文档末尾
                else:
                    raise WordDocumentError(
                        ErrorCode.SERVER_ERROR,
                        "Document object is invalid or missing Content attribute",
                    )

            # 添加评论
            result = add_comment(document, range_obj, comment_text, author)

            # 获取评论的索引作为ID，而不是直接返回COM对象
            if document is None:
                raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")
            comments_count = document.Comments.Count
            comment_id = comments_count - 1  # 0-based index

            return {
                "success": True,
                "comment_id": comment_id,
                "message": "Comment added successfully",
            }

        except Exception as e:
            raise WordDocumentError(
                ErrorCode.SERVER_ERROR, f"Failed to add comment: {str(e)}"
            )

    # Handle get all comments operation
    elif operation_type == "get_all":
        try:
            # 检查Comments集合是否存在
            if hasattr(document, "Comments"):
                result = get_comments(document)
                return {
                    "success": True,
                    "comments": result,
                    "message": "Comments retrieved successfully",
                }
            else:
                return {
                    "success": True,
                    "comments": [],
                    "message": "No comments available in this document",
                }
        except Exception as e:
            raise WordDocumentError(
                ErrorCode.SERVER_ERROR, f"Failed to get comments: {str(e)}"
            )

    # Handle delete all comments operation
    elif operation_type == "delete_all":
        try:
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
        except Exception as e:
            raise WordDocumentError(
                ErrorCode.SERVER_ERROR, f"Failed to delete all comments: {str(e)}"
            )

    # Handle operations that require a comment_id
    elif operation_type in ["get_thread", "delete", "edit", "reply"]:
        if comment_id is None:
            raise WordDocumentError(
                ErrorCode.INVALID_INPUT,
                f"Comment ID is required for '{operation_type}' operation",
            )

        # 确保comment_id是整数
        try:
            # 允许comment_id是字符串或整数
            if isinstance(comment_id, int):
                comment_index = comment_id
            else:
                comment_index = int(comment_id)
        except (ValueError, TypeError):
            if document is None:
                raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")
            # 如果comment_id是字符串格式的COM对象引用，尝试提取ID
            if isinstance(comment_id, str) and "Add" in comment_id:
                # 使用文档中最后一个评论作为回退
                comment_index = max(0, document.Comments.Count - 1)
            else:
                raise WordDocumentError(
                    ErrorCode.INVALID_INPUT,
                    f"Invalid comment ID format: {comment_id}",
                )

        if operation_type == "get_thread":
            try:
                result = get_comment_thread(document, comment_index)
                return {
                    "success": True,
                    "thread": result,
                    "message": "Comment thread retrieved successfully",
                }
            except Exception as e:
                raise WordDocumentError(
                    ErrorCode.SERVER_ERROR, f"Failed to get comment thread: {str(e)}"
                )

        elif operation_type == "delete":
            try:
                result = delete_comment(document, comment_index)
                return {
                    "success": result,
                    "message": (
                        f"Comment {comment_index} deleted successfully"
                        if result
                        else "Failed to delete comment"
                    ),
                }
            except Exception as e:
                raise WordDocumentError(
                    ErrorCode.SERVER_ERROR, f"Failed to delete comment: {str(e)}"
                )

        elif operation_type == "edit":
            if comment_text is None:
                raise WordDocumentError(
                    ErrorCode.INVALID_INPUT,
                    "Comment text is required for edit operation",
                )

            try:
                result = edit_comment(document, comment_index, comment_text)
                return {
                    "success": result,
                    "message": (
                        f"Comment {comment_index} edited successfully"
                        if result
                        else "Failed to edit comment"
                    ),
                }
            except Exception as e:
                raise WordDocumentError(
                    ErrorCode.SERVER_ERROR, f"Failed to edit comment: {str(e)}"
                )

        elif operation_type == "reply":
            if comment_text is None:
                raise WordDocumentError(
                    ErrorCode.INVALID_INPUT,
                    "Comment text is required for reply operation",
                )

            try:
                result = reply_to_comment(document, comment_index, comment_text, author)
                if document is None:
                    raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")
                # 获取评论的回复数作为回复ID
                comment = document.Comments(comment_index + 1)
                reply_id = comment.Replies.Count - 1  # 0-based index
                return {
                    "success": True,
                    "reply_id": reply_id,
                    "message": "Reply added successfully",
                }
            except Exception as e:
                raise WordDocumentError(
                    ErrorCode.SERVER_ERROR, f"Failed to reply to comment: {str(e)}"
                )

    # Handle unknown operation types
    else:
        raise WordDocumentError(
            ErrorCode.INVALID_INPUT, f"Unknown operation type: {operation_type}"
        )
