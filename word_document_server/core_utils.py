"""
Utility functions shared across multiple Word document server modules.
"""
from typing import Any, Dict, List, Optional

from mcp.server.fastmcp.server import Context

from word_document_server.com_backend import WordBackend
from word_document_server.errors import WordDocumentError, format_error_response



from word_document_server.core import mcp_server

def get_backend_for_tool(ctx: Context, file_path: str) -> WordBackend:
    """
    Gets or creates a WordBackend instance for the specified file path.
    
    Args:
        ctx: The MCP context object
        file_path: The absolute path to the Word document
        
    Returns:
        A WordBackend instance
        
    Raises:
        WordDocumentError: If creating the backend fails
    """
    # Initialize session state if not exists
    if not hasattr(ctx.session, 'document_state'):
        ctx.session.document_state = {}
        ctx.session.backend_instances = {}
    
    # Check if we already have a backend for this file
    if file_path in ctx.session.backend_instances:
        return ctx.session.backend_instances[file_path]
    
    # Create a new backend instance
    try:
        backend = WordBackend(file_path=file_path, visible=True)
        backend.__enter__()
        
        # Store backend in session state
        ctx.session.backend_instances[file_path] = backend
        
        return backend
    except Exception as e:
        raise WordDocumentError(f"Failed to create backend for file '{file_path}': {e}")


def get_active_document_path(ctx: Context) -> Optional[str]:
    """
    Retrieves the active document path from session state if available.
    
    Args:
        ctx: The MCP context object
        
    Returns:
        The active document path or None if no document is active
    """
    if hasattr(ctx.session, 'document_state'):
        return ctx.session.document_state.get('active_document_path')
    return None


def validate_active_document(ctx: Context) -> Optional[str]:
    """
    Validates that there is an active document and returns its path.
    
    Args:
        ctx: The MCP context object
        
    Returns:
        An error message if no active document, or None if validation passes
    """
    active_doc_path = get_active_document_path(ctx)
    if not active_doc_path:
        return "Error: No active document. Please use 'open_document' first."
    return None


def validate_locator(locator: Dict[str, Any]) -> Optional[str]:
    """
    Validates the structure of a locator dictionary.
    
    Args:
        locator: The locator to validate
        
    Returns:
        An error message if validation fails, or None if validation passes
    """
    if not isinstance(locator, dict):
        return "Error: Locator must be a dictionary"
    
    if 'target' not in locator:
        return "Error: Locator must contain a 'target' field"
    
    target = locator['target']
    if not isinstance(target, dict):
        return "Error: Locator 'target' must be a dictionary"
    
    if 'type' not in target:
        return "Error: Locator target must contain a 'type' field"
    
    return None