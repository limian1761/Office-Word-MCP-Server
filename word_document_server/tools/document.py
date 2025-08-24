import json
from typing import Dict, Any, Optional
from mcp.server.fastmcp.server import Context
from word_document_server.core_utils import get_backend_for_tool, mcp_server
from word_document_server.errors import ElementNotFoundError, format_error_response
from word_document_server.com_backend import WordBackend


@mcp_server.tool()
def open_document(ctx: Context, file_path: str) -> str:
    """
    Opens a Word document and prepares it for editing. This must be the first tool called.

    Args:
        file_path: The absolute path to the .docx file.

    Returns:
        A confirmation message with document information.
    """
    try:
        # Initialize or get session state for document
        if not hasattr(ctx.session, 'document_state'):
            ctx.session.document_state = {}
            ctx.session.backend_instances = {}
        
        # Create WordBackend for this document using get_backend_for_tool utility
        backend = get_backend_for_tool(ctx, file_path)
        
        # Check if document is not None
        if backend.document is None:
            raise ValueError("Failed to open document: Document object is None.")
            
        # Get document info
        document_info = {
            'file_path': file_path,
            'title': backend.document.Name,
            'saved': backend.document.Saved
        }
        document_info_str = json.dumps(document_info, ensure_ascii=False)
        
        # Store the document path as the active document
        ctx.session.document_state['active_document_path'] = file_path
        ctx.session.backend_instances[file_path] = backend
        
        return f"Document opened successfully: {document_info}"
    except FileNotFoundError:
        return f"Error [4001]: The file '{file_path}' was not found."
    except PermissionError:
        return f"Error [4002]: Permission denied when trying to open '{file_path}'."
    except ValueError as e:
        return f"Error [1001]: {e}"
    except Exception as e:
        return format_error_response(e)


@mcp_server.tool()
def shutdown_word(ctx: Context) -> str:
    """
    Closes the document and shuts down the Word application instance.
    Should be called at the end of a session.

    Returns:
        A success or error message.
    """
    try:
        # Check if we have any open documents
        if not hasattr(ctx.session, 'document_state') or not ctx.session.document_state:
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
def get_document_styles(ctx: Context) -> str:
    """
    Retrieves all available styles in the active document.

    Returns:
        A JSON string containing a list of styles with their names and types.
    """
    # Get active document path from session state
    active_doc_path = None
    if hasattr(ctx.session, 'document_state'):
        active_doc_path = ctx.session.document_state.get('active_document_path')
    if not active_doc_path:
        return "Error [2001]: No active document. Please use 'open_document' first."

    # Check cache first if available
    if hasattr(ctx.session, 'document_cache') and 'styles' in ctx.session.document_cache:
        return json.dumps(ctx.session.document_cache['styles'], ensure_ascii=False)

    try:
        backend = get_backend_for_tool(ctx, active_doc_path)
        styles = backend.get_document_styles()
        
        # Cache the result
        if not hasattr(ctx.session, 'document_cache'):
            ctx.session.document_cache = {}
        ctx.session.document_cache['styles'] = styles
        
        return json.dumps(styles, ensure_ascii=False)
    except Exception as e:
        return format_error_response(e)


@mcp_server.tool()
def get_document_structure(ctx: Context) -> str:
    """
    Provides a structured overview of the document by listing all headings.

    Returns:
        A JSON string containing a list of dictionaries, each representing a heading with its text and level.
    """
    # Get active document path from session state
    active_doc_path = None
    if hasattr(ctx.session, 'document_state'):
        active_doc_path = ctx.session.document_state.get('active_document_path')
    if not active_doc_path:
        return "Error [2001]: No active document. Please use 'open_document' first."

    # Check cache first if available
    if hasattr(ctx.session, 'document_cache') and 'structure' in ctx.session.document_cache:
        return json.dumps(ctx.session.document_cache['structure'], ensure_ascii=False)

    try:
        backend = get_backend_for_tool(ctx, active_doc_path)
        structure = backend.get_document_structure()
        
        # Cache the result
        if not hasattr(ctx.session, 'document_cache'):
            ctx.session.document_cache = {}
        ctx.session.document_cache['structure'] = structure
        
        return json.dumps(structure, ensure_ascii=False)
    except Exception as e:
        return format_error_response(e)


@mcp_server.tool()
def enable_track_revisions(ctx: Context) -> str:
    """
    Enables track changes (revision mode) in the document.

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
        backend.enable_track_revisions()
        return "Track revisions has been enabled successfully."
    except Exception as e:
        return format_error_response(e)


@mcp_server.tool()
def disable_track_revisions(ctx: Context) -> str:
    """
    Disables track changes (revision mode) in the document.

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
        backend.disable_track_revisions()
        return "Track revisions has been disabled successfully."
    except Exception as e:
        return format_error_response(e)


@mcp_server.tool()
def accept_all_changes(ctx: Context) -> str:
    """
    Accepts all tracked revisions in the document.

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
        backend.accept_all_changes()
        # Add None check for document
        if backend.document is None:
            raise ValueError("Failed to save document: No active document.")
        backend.document.Save()
        return "All tracked revisions have been accepted successfully."
    except Exception as e:
        return format_error_response(e)


@mcp_server.tool()
def set_header_text(ctx: Context, text: str) -> str:
    """
    Sets the text for the primary header in the active document.

    Args:
        text: The text to place in the header.

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
        backend.set_header_text(text)
        # Add None check for document
        if backend.document is None:
            raise ValueError("Failed to save document: No active document.")
        backend.document.Save()
        return "Header text has been set successfully."
    except Exception as e:
        return format_error_response(e)


@mcp_server.tool()
def set_footer_text(ctx: Context, text: str) -> str:
    """
    Sets the text for the primary footer in the active document.

    Args:
        text: The text to place in the footer.

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
        backend.set_footer_text(text)
        # Add None check for document
        if backend.document is None:
            raise ValueError("Failed to save document: No active document.")
        backend.document.Save()
        return "Footer text has been set successfully."
    except Exception as e:
        return format_error_response(e)


@mcp_server.tool()
def save_document(ctx: Context, file_path: Optional[str] = None) -> str:
    """
    Saves the active document, optionally to a new location.

    Args:
        file_path: Optional, the path to save the document to. If not provided, saves to the current location.

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
        
        # Check if document is not None
        if backend.document is None:
            raise ValueError("Failed to save document: Document object is None.")
            
        if file_path:
            # Save to a new location
            backend.document.SaveAs(file_path)
            # Update the active document path in session state
            ctx.session.document_state['active_document_path'] = file_path
            ctx.session.backend_instances[file_path] = backend
            # Remove the old backend reference if it's different
            if file_path != active_doc_path and active_doc_path in ctx.session.backend_instances:
                del ctx.session.backend_instances[active_doc_path]
            
            # Invalidate cache if we have one
            if hasattr(ctx.session, 'document_cache'):
                del ctx.session.document_cache
                
            return f"Document has been saved successfully to '{file_path}'."
        else:
            # Save to the current location
            backend.document.Save()
            return "Document has been saved successfully."
    except FileNotFoundError:
        return f"Error [4001]: The directory for '{file_path}' was not found."
    except PermissionError:
        return f"Error [4002]: Permission denied when trying to save to '{file_path}'."
    except Exception as e:
        return format_error_response(e)


@mcp_server.tool()
def close_document(ctx: Context, file_path: Optional[str] = None) -> str:
    """
    Closes a specific document or the active document if no path is provided.

    Args:
        file_path: Optional, the path of the document to close. If not provided, closes the active document.

    Returns:
        A success or error message.
    """
    # Get active document path from session state
    active_doc_path = None
    if hasattr(ctx.session, 'document_state'):
        active_doc_path = ctx.session.document_state.get('active_document_path')
    
    # Determine which document to close
    doc_to_close = file_path or active_doc_path
    
    if not doc_to_close:
        return "Error [2001]: No document path specified and no active document."
    
    if not hasattr(ctx.session, 'backend_instances') or doc_to_close not in ctx.session.backend_instances:
        return f"Error [2003]: The document '{doc_to_close}' is not open."
    
    try:
        backend = ctx.session.backend_instances[doc_to_close]
        backend.close_document()
        
        # Remove from backend instances
        del ctx.session.backend_instances[doc_to_close]
        
        # If this was the active document, clear the active document path
        if doc_to_close == active_doc_path:
            ctx.session.document_state['active_document_path'] = None
            
            # Invalidate cache if we have one
            if hasattr(ctx.session, 'document_cache'):
                del ctx.session.document_cache
        
        return f"Document '{doc_to_close}' has been closed successfully."
    except Exception as e:
        return format_error_response(e)