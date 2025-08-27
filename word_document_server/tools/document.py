import json
import os
from typing import Dict, Any, Optional
from mcp.server.fastmcp.server import Context
from word_document_server.core_utils import get_backend_for_tool, mcp_server
from word_document_server.errors import format_error_response, handle_tool_errors
from word_document_server import WordBackend
from word_document_server.operations import get_document_styles, get_document_structure, get_all_text, enable_track_revisions


@mcp_server.tool()
@handle_tool_errors
def open_document(ctx: Context, file_path: str) -> str:
    """
    Opens a Word document and prepares it for editing. This must be the first tool called.

    Args:
        file_path: The absolute path to the .docx file.

    Returns:
        A confirmation message with document information.
    """
    # Initialize or get session state for document
    if not hasattr(ctx.session, 'document_state'):
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
        'file_path': file_path,
        'title': backend.document.Name,
        'saved': backend.document.Saved
    }
    document_info_str = json.dumps(document_info, ensure_ascii=False)
    
    # Store the document path as the active document
    ctx.session.document_state['active_document_path'] = file_path
    ctx.session.backend_instances[file_path] = backend
    
    # Read agent guide content
    agent_guide_path = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), 'docs', 'agent_guide.md')
    try:
        with open(agent_guide_path, 'r', encoding='utf-8') as f:
            agent_guide_content = f.read()
    except Exception as e:
        agent_guide_content = f"Error reading agent guide: {str(e)}"

    # Return combined response
    return f"Document opened successfully: {document_info}\n\n---\n\n# Office-Word-MCP-Server Agent Guide\n\n{agent_guide_content}"

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
    from word_document_server.core_utils import validate_active_document
    error = validate_active_document(ctx)
    if error:
        return error

    try:
        backend = get_backend_for_tool(ctx, active_doc_path)
        
        # Get all document styles using the backend method
        styles = get_document_styles(backend)
        
        # Convert to JSON string
        return json.dumps(styles, ensure_ascii=False)
    except Exception as e:
        return format_error_response(e)


@mcp_server.tool()
def get_document_structure(ctx: Context) -> str:
    """
    Provides a structured overview of the document by listing all headings.

    Returns:
        A JSON string containing a list of headings with their text and level.
    """
    # Get active document path from session state
    from word_document_server.core_utils import validate_active_document
    error = validate_active_document(ctx)
    if error:
        return error

    try:
        backend = get_backend_for_tool(ctx, active_doc_path)
        
        # Get document structure using the backend method
        structure = get_document_structure(backend)
        
        # Convert to JSON string
        return json.dumps(structure, ensure_ascii=False)
    except Exception as e:
        return format_error_response(e)


@mcp_server.tool()
def get_all_text(ctx: Context) -> str:
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
        backend = get_backend_for_tool(ctx, active_doc_path)
        
        # Get all text using the backend method
        text = get_all_text(backend)
        
        return text
    except Exception as e:
        return format_error_response(e)


@mcp_server.tool()
def get_elements(ctx: Context, element_type: str) -> str:
    """
    Retrieves information about elements of a specific type in the document.

    Args:
        element_type: Type of elements to retrieve. Can be:
            - "paragraphs": All paragraphs
            - "tables": All tables
            - "images": All inline shapes/images
            - "headings": All headings
            - "styles": All styles
            - "comments": All comments

    Returns:
        A JSON string containing information about the elements.
    """
    # Get active document path from session state
    from word_document_server.core_utils import validate_active_document
    error = validate_active_document(ctx)
    if error:
        return error

    # Validate element_type parameter
    supported_types = ["paragraphs", "tables", "images", "headings", "styles", "comments"]
    if element_type not in supported_types:
        return f"Error [1001]: Unsupported element type '{element_type}'. Supported types: {', '.join(supported_types)}"

    try:
        backend = get_backend_for_tool(ctx, active_doc_path)
        
        # Get elements using the backend method
        elements = get_selection_info(backend, element_type)
        
        # Convert to JSON string
        return json.dumps(elements, ensure_ascii=False)
    except Exception as e:
        return format_error_response(e)