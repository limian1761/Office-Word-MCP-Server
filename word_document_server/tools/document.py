import json
import os
from typing import Any, Dict, Optional

from mcp.server.fastmcp.server import Context
from pydantic import Field

from word_document_server.core_utils import get_backend_for_tool, mcp_server
from word_document_server.errors import (format_error_response,
                                         handle_tool_errors)
from word_document_server.operations import (enable_track_revisions,
                                             get_all_text,
                                             get_document_styles)


@mcp_server.tool()
@require_active_document_validation
@require_active_document_validation
@handle_tool_errors
def open_document(
    ctx: Context[ServerSession, AppContext] = Field(description="Context object"),
    file_path: str = Field(description="The absolute path to the .docx file"),
) -> str:
    """
    Opens a Word document and prepares it for editing. This must be the first tool called.

    Returns:
        A confirmation message with document information.
    """
    # Initialize or get session state for document
    try: 
        ctx.request_context.lifespan_context.open_document(file_path)
    except Exception as e:
        return format_error_response(e)
    
    active_doc = ctx.request_context.lifespan_context.get_active_document()

    # Check if document is not None
    if active_doc is None:
        raise ValueError("Failed to open document: Document object is None.")

    # Enable track changes by default
    active_doc.TrackRevisions = True

    # Get document info
    document_info = {
        "file_path": file_path,
        "title": active_doc.Name,
        "saved":active_doc.Saved,
    }
    document_info_str = json.dumps(document_info, ensure_ascii=False)


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
def close_document(ctx: Context[ServerSession, AppContext] = Field(description="Context object")) -> str:
    """
    Closes the active document but keeps the Word application running.

    Returns:
        A success or error message.
    """
    try:
        active_doc = ctx.request_context.lifespan_context.get_active_document()
        try:
            doc_path = active_doc.Path
            active_doc.Close(SaveChanges=True)
            return f"Document '{doc_path}' closed successfully."
        except Exception as e:
            return f"Error closing document: {e}"
        return "Active document backend not found."
    except Exception as e:
        return format_error_response(e)


@mcp_server.tool()
def shutdown_word(ctx: Context[ServerSession, AppContext] = Field(description="Context object"))-> str:
    """
    Closes the document and shuts down the Word application instance.
    Should be called at the end of a session.

    Returns:
        A success or error message.
    """
    try:
        # Check if we have any open documents
        ctx.request_context.lifespan_context.close_document()
        return "Word application has been shut down successfully."
    except Exception as e:
        return format_error_response(e)


@mcp_server.tool()
def get_document_styles(ctx: Context[ServerSession, AppContext] = Field(description="Context object")) -> str:
    """
    Retrieves all available styles in the active document.

    Returns:
        A JSON string containing a list of styles with their names and types.
    """
    # Get active document path from session state
    active_doc = ctx.request_context.lifespan_context.get_active_document()
    styles = get_document_styles(active_doc)

     # Convert to JSON string
    return json.dumps(styles, ensure_ascii=False)




@mcp_server.tool()
def get_all_text(ctx: Context[ServerSession, AppContext] = Field(description="Context object")) -> str:
    """
    Retrieves all text from the active document.

    Returns:
        A string containing all text content from the document.
    """
    try:
        # Check if session exists
        if not hasattr(ctx, 'session'):
            return format_error_response("No session available in context")

        active_doc = ctx.request_context.lifespan_context.get_active_document()
        # Get all text using the document object directly
        text = get_all_text(active_doc)

        return text
    except Exception as e:
        return format_error_response(e)


@mcp_server.tool()
def get_elements(
    ctx: Context[ServerSession, AppContext] = Field(description="Context object"),
    element_type: str = Field(
        description='Type of elements to retrieve. Can be: "paragraphs", "tables", "images", "headings", "styles", "comments"'
    ),
) -> str:
    """
    Retrieves information about elements of a specific type in the document.

    Returns:
        A JSON string containing information about the elements.
    """
    # Validate element_type parameter
    supported_types = [
        "paragraphs",
        "tables",
        "images",
        "headings",
        "styles",
        "comments",
    ]
    try:
        active_doc = ctx.request_context.lifespan_context.get_active_document()
        # Remove backend usage, use document directly

        # Get elements using the backend method
        elements = get_selection_info(backend, element_type)

        # Convert to JSON string
        return json.dumps(elements, ensure_ascii=False)
    except Exception as e:
        return format_error_response(e)
