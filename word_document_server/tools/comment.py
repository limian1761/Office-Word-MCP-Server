import json
from typing import Dict, Any, Optional
from mcp.server.fastmcp.server import Context
from word_document_server.core_utils import get_backend_for_tool, mcp_server
from word_document_server.errors import ElementNotFoundError, format_error_response


@mcp_server.tool()
def add_comment(ctx: Context, locator: Dict[str, Any], text: str, author: str = "User") -> str:
    """
    Adds a comment to the document at the location specified by the locator.

    Args:
        locator: The Locator object to find the target location for the comment.
        text: The text of the comment.
        author: The author of the comment (default: "User").

    Returns:
        A success or error message.
    """
    # Get active document path from session state
    active_doc_path = None
    if hasattr(ctx.session, 'document_state'):
        active_doc_path = ctx.session.document_state.get('active_document_path')
    if not active_doc_path:
        return "Error [2001]: No active document. Please use 'open_document' first."

    if not text:
        return "Error [1001]: Comment text cannot be empty."

    try:
        backend = get_backend_for_tool(ctx, active_doc_path)
        
        # Get selector engine
        from word_document_server.selector import SelectorEngine
        selector = SelectorEngine()
        
        # Convert locator to Selection object
        selection = selector.select(backend, locator, expect_single=True)
        
        # Get the first element's range
        com_range_obj = selection._elements[0].Range
        
        # Call the backend method to add a comment
        comment_id = backend.add_comment(com_range_obj, text, author)
        
        # Check if document is not None before saving
        if backend.document is None:
            raise ValueError("Failed to save document after adding comment: Document object is None.")
            
        # Save the document
        backend.document.Save()
        
        return f"Comment added successfully with ID: {comment_id}"
    except ElementNotFoundError as e:
        return f"Error [2002]: Error finding target location for comment: {e}" + " Please try simplifying your locator or use get_document_structure to check the actual document structure."
    except Exception as e:
        return format_error_response(e)


@mcp_server.tool()
def get_comments(ctx: Context) -> str:
    """
    Retrieves all comments in the active document.

    Returns:
        A JSON string containing a list of comments with their information.
    """
    # Get active document path from session state
    active_doc_path = None
    if hasattr(ctx.session, 'document_state'):
        active_doc_path = ctx.session.document_state.get('active_document_path')
    if not active_doc_path:
        return "Error [2001]: No active document. Please use 'open_document' first."

    try:
        backend = get_backend_for_tool(ctx, active_doc_path)
        
        # Get all comments from the backend
        comments = backend.get_comments()
        
        # Convert to JSON string
        return json.dumps(comments, ensure_ascii=False)
    except Exception as e:
        return format_error_response(e)


@mcp_server.tool()
def delete_comment(ctx: Context, comment_index: int) -> str:
    """
    Deletes a comment by its 0-based index.

    Args:
        comment_index: The 0-based index of the comment to delete.

    Returns:
        A success or error message.
    """
    # Get active document path from session state
    active_doc_path = None
    if hasattr(ctx.session, 'document_state'):
        active_doc_path = ctx.session.document_state.get('active_document_path')
    if not active_doc_path:
        return "Error [2001]: No active document. Please use 'open_document' first."

    if not isinstance(comment_index, int):
        return "Error [1001]: Comment index must be an integer."

    if comment_index < 0:
        return "Error [1001]: Comment index cannot be negative."

    try:
        backend = get_backend_for_tool(ctx, active_doc_path)
        
        # Delete the comment by index
        backend.delete_comment(comment_index)
        
        # Check if document is not None before saving
        if backend.document is None:
            raise ValueError("Failed to save document after deleting comment: Document object is None.")
            
        # Save the document
        backend.document.Save()
        
        return f"Comment at index {comment_index} has been deleted successfully."
    except IndexError:
        # Get total number of comments to provide better feedback
        total_comments = len(backend.get_comments())
        return f"Error [4004]: Comment index {comment_index} out of range. There are {total_comments} comments in the document."
    except Exception as e:
        return format_error_response(e)


@mcp_server.tool()
def delete_all_comments(ctx: Context) -> str:
    """
    Deletes all comments in the active document.

    Returns:
        A success or error message.
    """
    # Get active document path from session state
    active_doc_path = None
    if hasattr(ctx.session, 'document_state'):
        active_doc_path = ctx.session.document_state.get('active_document_path')
    if not active_doc_path:
        return "Error [2001]: No active document. Please use 'open_document' first."

    try:
        backend = get_backend_for_tool(ctx, active_doc_path)
        
        # Delete all comments
        count = backend.delete_all_comments()
        
        # Check if document is not None before saving
        if backend.document is None:
            raise ValueError("Failed to save document after deleting all comments: Document object is None.")
            
        # Save the document
        backend.document.Save()
        
        return f"Successfully deleted {count} comments."
    except Exception as e:
        return format_error_response(e)


@mcp_server.tool()
def edit_comment(ctx: Context, comment_index: int, new_text: str) -> str:
    """
    Edits an existing comment by its 0-based index.

    Args:
        comment_index: The 0-based index of the comment to edit.
        new_text: The new text for the comment.

    Returns:
        A success or error message.
    """
    # Get active document path from session state
    active_doc_path = None
    if hasattr(ctx.session, 'document_state'):
        active_doc_path = ctx.session.document_state.get('active_document_path')
    if not active_doc_path:
        return "Error [2001]: No active document. Please use 'open_document' first."

    if not isinstance(comment_index, int):
        return "Error [1001]: Comment index must be an integer."

    if comment_index < 0:
        return "Error [1001]: Comment index cannot be negative."

    if not new_text:
        return "Error [1001]: New comment text cannot be empty."

    try:
        backend = get_backend_for_tool(ctx, active_doc_path)
        
        # Edit the comment
        backend.edit_comment(comment_index, new_text)
        
        # Check if document is not None before saving
        if backend.document is None:
            raise ValueError("Failed to save document after editing comment: Document object is None.")
            
        # Save the document
        backend.document.Save()
        
        return f"Comment at index {comment_index} has been updated successfully."
    except IndexError:
        # Get total number of comments to provide better feedback
        total_comments = len(backend.get_comments())
        return f"Error [4004]: Comment index {comment_index} out of range. There are {total_comments} comments in the document."
    except Exception as e:
        return format_error_response(e)


@mcp_server.tool()
def reply_to_comment(ctx: Context, comment_index: int, reply_text: str, author: str = "User") -> str:
    """
    Replies to an existing comment.

    Args:
        comment_index: The 0-based index of the comment to reply to.
        reply_text: The text of the reply.
        author: The author of the reply (default: "User").

    Returns:
        A success or error message.
    """
    # Get active document path from session state
    active_doc_path = None
    if hasattr(ctx.session, 'document_state'):
        active_doc_path = ctx.session.document_state.get('active_document_path')
    if not active_doc_path:
        return "Error [2001]: No active document. Please use 'open_document' first."

    if not isinstance(comment_index, int):
        return "Error [1001]: Comment index must be an integer."

    if comment_index < 0:
        return "Error [1001]: Comment index cannot be negative."

    if not reply_text:
        return "Error [1001]: Reply text cannot be empty."

    try:
        backend = get_backend_for_tool(ctx, active_doc_path)
        
        # Reply to the comment
        backend.reply_to_comment(comment_index, reply_text, author)
        
        # Check if document is not None before saving
        if backend.document is None:
            raise ValueError("Failed to save document after replying to comment: Document object is None.")
            
        # Save the document
        backend.document.Save()
        
        return f"Reply added to comment at index {comment_index} successfully."
    except IndexError:
        # Get total number of comments to provide better feedback
        total_comments = len(backend.get_comments())
        return f"Error [4004]: Comment index {comment_index} out of range. There are {total_comments} comments in the document."
    except Exception as e:
        return format_error_response(e)


@mcp_server.tool()
def get_comment_thread(ctx: Context, comment_index: int) -> str:
    """
    Retrieves a comment thread including the original comment and all replies.

    Args:
        comment_index: The 0-based index of the original comment.

    Returns:
        A JSON string containing the comment thread information.
    """
    # Get active document path from session state
    active_doc_path = None
    if hasattr(ctx.session, 'document_state'):
        active_doc_path = ctx.session.document_state.get('active_document_path')
    if not active_doc_path:
        return "Error [2001]: No active document. Please use 'open_document' first."

    if not isinstance(comment_index, int):
        return "Error [1001]: Comment index must be an integer."

    if comment_index < 0:
        return "Error [1001]: Comment index cannot be negative."

    try:
        backend = get_backend_for_tool(ctx, active_doc_path)
        
        # Get the comment thread
        thread = backend.get_comment_thread(comment_index)
        
        # Convert to JSON string
        return json.dumps(thread, ensure_ascii=False)
    except IndexError:
        # Get total number of comments to provide better feedback
        total_comments = len(backend.get_comments())
        return f"Error [4004]: Comment index {comment_index} out of range. There are {total_comments} comments in the document."
    except Exception as e:
        return format_error_response(e)