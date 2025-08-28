import json
from typing import Any, Dict, Optional

from mcp.server.fastmcp.server import Context
from pydantic import Field
from word_document_server.core import ServerSession
from word_document_server.utils.app_context import AppContext

from word_document_server.core_utils import mcp_server
from word_document_server.errors import (CommentEmptyError, CommentIndexError,
                                         ElementNotFoundError, ReplyEmptyError,
                                         format_error_response,
                                         handle_tool_errors)
from word_document_server.operations.comment_operations import (add_comment as add_comment_op,
                                                                get_comments as get_comments_op,
                                                                delete_comment as delete_comment_op,
                                                                delete_all_comments as delete_all_comments_op,
                                                                edit_comment as edit_comment_op,
                                                                reply_to_comment as reply_to_comment_op,
                                                                get_comment_thread as get_comment_thread_op)


@mcp_server.tool()
@require_active_document_validation
@handle_tool_errors
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
        active_doc = ctx.request_context.lifespan_context.get_active_document()

        # Use the shared selector engine from core
        from word_document_server.core import selector

        # Convert locator to Selection object
        selection = selector.select(active_doc, locator, expect_single=True)

        # Get the first element's range
        com_range_obj = selection._elements[0].Range

        # Call the backend method to add a comment
        comment_id = add_comment_op(active_doc, com_range_obj, text, author)

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



@mcp_server.tool()
def get_comments(ctx: Context[ServerSession, AppContext] = Field(description="Context object")) -> str:
    """
    Retrieves all comments in the active document.

    Returns:
        A JSON string containing a list of comments with their information.
    """
    try:
        active_doc = ctx.request_context.lifespan_context.get_active_document()

        # Get all comments from the document
        comments = get_comments_op(backend)

        # Convert to JSON string
        return json.dumps(comments, ensure_ascii=False)
    except Exception as e:
        return format_error_response(e)


@mcp_server.tool()
def delete_comment(
    ctx: Context[ServerSession, AppContext] = Field(description="Context object"),
    comment_index: int = Field(
        description="The 0-based index of the comment to delete"
    ),
) -> str:
    """
    Deletes a comment by its 0-based index.

    Returns:
        A success or error message.
    """
    try:
        backend = get_backend_for_tool(
            ctx, ctx.session.document_state["active_document_path"]
        )

        # Call the backend method to delete the comment
        delete_comment_op(backend, comment_index)

        # Save the document
        backend.document.Save()

        return f"Comment at index {comment_index} deleted successfully."
    except IndexError:
        return format_error_response(CommentIndexError(comment_index))
    except Exception as e:
        return format_error_response(e)


@mcp_server.tool()
def delete_all_comments(ctx: Context[ServerSession, AppContext] = Field(description="Context object")) -> str:
    """
    Deletes all comments in the active document.

    Returns:
        A success or error message.
    """
    try:
        backend = get_backend_for_tool(
            ctx, ctx.session.document_state["active_document_path"]
        )

        # Call the backend method to delete all comments
        deleted_count = delete_all_comments_op(backend)

        # Save the document
        backend.document.Save()

        return f"All {deleted_count} comments deleted successfully."
    except Exception as e:
        return format_error_response(e)


@mcp_server.tool()
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
    # Get active document path from session state
    from word_document_server.core_utils import validate_active_document

    error = validate_active_document(ctx)
    if error:
        return error

    if not new_text:
        return format_error_response(CommentEmptyError())

    try:
        backend = get_backend_for_tool(
            ctx, ctx.session.document_state["active_document_path"]
        )

        # Call the backend method to edit the comment
        edit_comment_op(backend, comment_index, new_text)

        # Save the document
        backend.document.Save()

        return f"Comment at index {comment_index} edited successfully."
    except IndexError:
        return format_error_response(CommentIndexError(comment_index))
    except Exception as e:
        return format_error_response(e)


@mcp_server.tool()
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
    # Get active document path from session state
    from word_document_server.core_utils import validate_active_document

    error = validate_active_document(ctx)
    if error:
        return error

    if not reply_text:
        return format_error_response(ReplyEmptyError())

    try:
        backend = get_backend_for_tool(
            ctx, ctx.session.document_state["active_document_path"]
        )

        # Call the backend method to reply to the comment
        reply_to_comment_op(backend, comment_index, reply_text, author)

        # Save the document
        backend.document.Save()

        return f"Reply added to comment at index {comment_index} successfully."
    except IndexError:
        return format_error_response(CommentIndexError(comment_index))
    except Exception as e:
        return format_error_response(e)


@mcp_server.tool()
def get_comment_thread(
    ctx: Context[ServerSession, AppContext] = Field(description="Context object"),
    comment_index: int = Field(description="The 0-based index of the original comment"),
) -> str:
    """
    Retrieves a comment thread including the original comment and all replies.

    Returns:
        A JSON string containing the original comment and all replies.
    """
    # Get active document path from session state
    from word_document_server.core_utils import validate_active_document

    error = validate_active_document(ctx)
    if error:
        return error

    try:
        backend = get_backend_for_tool(
            ctx, ctx.session.document_state["active_document_path"]
        )

        # Call the backend method to get the comment thread
        thread = get_comment_thread_op(backend, comment_index)

        # Convert to JSON string
        return json.dumps(thread, ensure_ascii=False)
    except IndexError:
        return format_error_response(CommentIndexError(comment_index))
    except Exception as e:
        return format_error_response(e)