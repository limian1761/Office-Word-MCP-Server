import json
import os
from typing import Any, Dict, Optional

from mcp.server.fastmcp.server import Context
from pydantic import Field

from word_document_server import WordBackend, get_selection_info
from word_document_server.core_utils import get_backend_for_tool, mcp_server
from word_document_server.errors import (format_error_response,
                                         handle_tool_errors)
from word_document_server.operations import (enable_track_revisions,
                                             get_all_text,
                                             get_document_styles)


@mcp_server.tool()
@handle_tool_errors
def open_document(
    ctx: Context = Field(description="Context object"),
    file_path: str = Field(description="The absolute path to the .docx file"),
) -> str:
    """
    Opens a Word document and prepares it for editing. This must be the first tool called.

    Returns:
        A confirmation message with document information.
    """
    # Initialize or get session state for document
    if not hasattr(ctx.session, "document_state"):
        ctx.session.document_state = {}
        ctx.session.backend_instances = {}

    # Create WordBackend for this document using get_backend_for_tool utility
    backend = get_backend_for_tool(ctx, file_path)

    # Check if document is not None
    if backend.document is None:
        raise ValueError("Failed to open document: Document object is None.")

    # Enable track changes by default
    enable_track_revisions(backend)

    # Get document info
    document_info = {
        "file_path": file_path,
        "title": backend.document.Name,
        "saved": backend.document.Saved,
    }
    document_info_str = json.dumps(document_info, ensure_ascii=False)

    # Store the document path as the active document
    ctx.session.document_state["active_document_path"] = file_path
    ctx.session.backend_instances[file_path] = backend

    # Read agent guide content
    agent_guide_path = os.path.join(
        os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))),
        "docs",
        "agent_guide.md",
    )
    try:
        with open(agent_guide_path, "r", encoding="utf-8") as f:
            agent_guide_content = f.read()
    except Exception as e:
        agent_guide_content = f"Error reading agent guide: {str(e)}"

    # Return combined response
    return f"Document opened successfully: {document_info}\n\n---\n\n# Office-Word-MCP-Server Agent Guide\n\n{agent_guide_content}"


@mcp_server.tool()
def close_document(ctx: Context = Field(description="Context object")) -> str:
    """
    Closes the active document but keeps the Word application running.

    Returns:
        A success or error message.
    """
    try:
        # Check if we have any open documents
        if not hasattr(ctx.session, "document_state") or not ctx.session.document_state:
            return "No documents are currently open."

        # Get active document path
        active_doc_path = ctx.session.document_state.get("active_document_path")
        if not active_doc_path:
            return "No active document found."

        # Get the backend for the active document
        if active_doc_path in ctx.session.backend_instances:
            backend = ctx.session.backend_instances[active_doc_path]
            try:
                # Close the document without shutting down Word
                backend.document.Close(SaveChanges=True)
                # Remove from backend instances
                del ctx.session.backend_instances[active_doc_path]
                # Clear active document path
                ctx.session.document_state.pop("active_document_path", None)
                return f"Document '{active_doc_path}' closed successfully."
            except Exception as e:
                return f"Error closing document: {e}"

        return "Active document backend not found."
    except Exception as e:
        return format_error_response(e)


@mcp_server.tool()
def shutdown_word(ctx: Context = Field(description="Context object")) -> str:
    """
    Closes the document and shuts down the Word application instance.
    Should be called at the end of a session.

    Returns:
        A success or error message.
    """
    try:
        # Check if we have any open documents
        if not hasattr(ctx.session, "document_state") or not ctx.session.document_state:
            return "No documents are currently open."

        # Close all backend instances
        for file_path, backend in ctx.session.backend_instances.items():
            try:
                backend.shutdown()
            except Exception as e:
                return f"Error closing document '{file_path}': {e}"

        # Clear session state
        ctx.session.document_state = {}
        ctx.session.backend_instances = {}

        return "Word application has been shut down successfully."
    except Exception as e:
        return format_error_response(e)


@mcp_server.tool()
def get_document_styles(ctx: Context = Field(description="Context object")) -> str:
    """
    Retrieves all available styles in the active document.

    Returns:
        A JSON string containing a list of styles with their names and types.
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

        # Get all document styles using the backend method
        styles = get_document_styles(backend)

        # Convert to JSON string
        return json.dumps(styles, ensure_ascii=False)
    except Exception as e:
        return format_error_response(e)




@mcp_server.tool()
def get_all_text(ctx: Context = Field(description="Context object")) -> str:
    """
    Retrieves all text from the active document.

    Returns:
        A string containing all text content from the document.
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

        # Get all text using the backend method
        text = get_all_text(backend)

        return text
    except Exception as e:
        return format_error_response(e)


@mcp_server.tool()
def get_elements(
    ctx: Context = Field(description="Context object"),
    element_type: str = Field(
        description='Type of elements to retrieve. Can be: "paragraphs", "tables", "images", "headings", "styles", "comments"'
    ),
) -> str:
    """
    Retrieves information about elements of a specific type in the document.

    Returns:
        A JSON string containing information about the elements.
    """
    # Get active document path from session state
    from word_document_server.core_utils import validate_active_document

    error = validate_active_document(ctx)
    if error:
        return error

    # Validate element_type parameter
    supported_types = [
        "paragraphs",
        "tables",
        "images",
        "headings",
        "styles",
        "comments",
    ]
    if element_type not in supported_types:
        from word_document_server.errors import ErrorCode, WordDocumentError

        e = WordDocumentError(
            ErrorCode.ELEMENT_TYPE_ERROR,
            f"Unsupported element type '{element_type}'. Supported types: {', '.join(supported_types)}",
        )
        return format_error_response(e)

    try:
        backend = get_backend_for_tool(
            ctx, ctx.session.document_state["active_document_path"]
        )

        # Get elements using the backend method
        elements = get_selection_info(backend, element_type)

        # Convert to JSON string
        return json.dumps(elements, ensure_ascii=False)
    except Exception as e:
        return format_error_response(e)
