import json
from typing import Any, Dict, Optional

from mcp.server.fastmcp.server import Context
from pydantic import Field
from mcp.server.session import ServerSession
from word_document_server.utils.app_context import AppContext

from word_document_server.utils.core_utils import require_active_document_validation
from word_document_server.utils.core_utils import (CommentEmptyError, CommentIndexError,
                                         ElementNotFoundError, ReplyEmptyError,
                                         format_error_response)
from word_document_server.operations.comment_operations import (add_comment as add_comment_op,
                                                                get_comments as get_comments_op,
                                                                delete_comment as delete_comment_op,
                                                                delete_all_comments as delete_all_comments_op,
                                                                edit_comment as edit_comment_op,
                                                                reply_to_comment as reply_to_comment_op,
                                                                get_comment_thread as get_comment_thread_op)


@require_active_document_validation
def add_comment(
    ctx: Context[ServerSession, AppContext] = Field(description="Context object"),
    locator: Dict[str, Any] = Field(
        description="The Locator object to find the target location for the comment"
    ),
    text: str = Field(description="The text of the comment"),
    author: str = Field(description="The author of the comment", default="User"),
) -> str:
    """
    Adds a comment to the document at the location specified by the locator.

    Returns:
        A success or error message.
    """
    try:
        from word_document_server.mcp_service.core import mcp_server
        
        active_doc = ctx.request_context.lifespan_context.get_active_document()

        # Use the shared selector engine from core
        from word_document_server.core import selector

        # Convert locator to Selection object
        selection = selector.select(active_doc, locator, expect_single=True)

        # Call the Selection method to add a comment
        comment_id = selection.add_comment(text, author)

        # Check if document is not None before saving
        if active_doc is None:
            raise ValueError(
                "Failed to save document after adding comment: Document object is None."
            )

        # Save the document
        active_doc.Save()

        return f"Comment added successfully with ID: {comment_id}"
    except ElementNotFoundError as e:
        return format_error_response(e)


@require_active_document_validation
def get_comments(ctx: Context[ServerSession, AppContext] = Field(description="Context object")) -> str:
    """
    Retrieves all comments in the active document.

    Returns:
        A JSON string containing a list of comments with their information.
    """
    try:
        from word_document_server.mcp_service.core import mcp_server
        
        active_doc = ctx.request_context.lifespan_context.get_active_document()

        # Get all comments from the document using Selection method
        from word_document_server.selector.selection import Selection
        selection = Selection([active_doc], active_doc)
        comments = get_comments_op(active_doc)

        # Convert to JSON string
        return json.dumps(comments, ensure_ascii=False)
    except Exception as e:
        return format_error_response(e)


@require_active_document_validation
def delete_comment(
    ctx: Context[ServerSession, AppContext] = Field(description="Context object"),
    comment_index: Optional[int] = Field(
        description="The 0-based index of the comment to delete. If not provided, all comments will be deleted.",
        default=None
    ),
) -> str:
    """
    Deletes a comment by its 0-based index, or all comments if no index is provided.

    Returns:
        A success or error message.
    """
    try:
        from word_document_server.mcp_service.core import mcp_server
        
        active_doc = ctx.request_context.lifespan_context.get_active_document()

        if comment_index is None:
            # Delete all comments
            deleted_count = delete_all_comments_op(active_doc)
            # Save the document
            if active_doc is not None:
                active_doc.Save()
            return f"All {deleted_count} comments deleted successfully."
        else:
            # Delete specific comment
            delete_comment_op(active_doc, comment_index)
            # Save the document
            if active_doc is not None:
                active_doc.Save()
            return f"Comment at index {comment_index} deleted successfully."
    except IndexError:
        return format_error_response(CommentIndexError(comment_index if comment_index is not None else 0))
    except Exception as e:
        return format_error_response(e)


@require_active_document_validation
def edit_comment(
    ctx: Context[ServerSession, AppContext] = Field(description="Context object"),
    comment_index: int = Field(description="The 0-based index of the comment to edit"),
    new_text: str = Field(description="The new text for the comment"),
) -> str:
    """
    Edits an existing comment by its 0-based index.

    Returns:
        A success or error message.
    """
    if not new_text:
        return format_error_response(CommentEmptyError())

    try:
        from word_document_server.mcp_service.core import mcp_server
        
        active_doc = ctx.request_context.lifespan_context.get_active_document()

        # Call the backend method to edit the comment
        edit_comment_op(active_doc, comment_index, new_text)

        # Save the document
        if active_doc is not None:
            active_doc.Save()

        return f"Comment at index {comment_index} edited successfully."
    except IndexError:
        return format_error_response(CommentIndexError(comment_index))
    except Exception as e:
        return format_error_response(e)


@require_active_document_validation
def reply_to_comment(
    ctx: Context[ServerSession, AppContext] = Field(description="Context object"),
    comment_index: int = Field(
        description="The 0-based index of the comment to reply to"
    ),
    reply_text: str = Field(description="The text of the reply"),
    author: str = Field(description="The author of the reply", default="User"),
) -> str:
    """
    Replies to an existing comment.

    Returns:
        A success or error message.
    """

    if not reply_text:
        return format_error_response(ReplyEmptyError())

    try:
        from word_document_server.mcp_service.core import mcp_server
        
        active_doc = ctx.request_context.lifespan_context.get_active_document()

        # Call the backend method to add a reply
        reply_to_comment_op(active_doc, comment_index, reply_text, author)

        # Save the document
        if active_doc is not None:
            active_doc.Save()

        return "Reply added successfully."
    except IndexError:
        return format_error_response(CommentIndexError(comment_index))
    except Exception as e:
        return format_error_response(e)