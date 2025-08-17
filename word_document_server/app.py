"""
Main MCP Server application file, built with the official mcp.server.fastmcp library.
"""
import os
from typing import Dict, Any

# Correct import for the MCP server components from the local SDK
from mcp.server.fastmcp.server import Context, FastMCP

from word_document_server.com_backend import WordBackend
from word_document_server.selector import SelectorEngine, ElementNotFoundError

# --- MCP Server Initialization ---
mcp_server = FastMCP("Office-Word-MCP-Server")
selector = SelectorEngine()

# --- State Management ---
# A simple dictionary on the context object can manage state per-session.
# We will store the active document path in the context's state.

# --- Tool Definitions ---

def get_backend_for_tool(ctx: Context, file_path: str) -> WordBackend:
    """
    This is a temporary solution for debugging. It creates a new backend
    for every tool call to ensure there is no stale state.
    """
    # Always create a new backend to ensure we get a fresh view of the document.
    # This is inefficient but necessary to isolate the state bug.
    backend = WordBackend(file_path=file_path, visible=True)
    backend.__enter__()
    ctx.set_state("word_backend", backend) # Overwrite the old one
    return backend

@mcp_server.tool()
def open_document(ctx: Context, file_path: str) -> str:
    """
    Opens a Word document and prepares it for editing. This must be the first tool called.

    Args:
        file_path: The absolute path to the .docx file.
    
    Returns:
        A confirmation message.
    """
    if not os.path.isabs(file_path):
        return f"Error: Path '{file_path}' is not an absolute path."
    if not os.path.exists(file_path):
        return f"Error: File '{file_path}' not found."
    if not file_path.lower().endswith('.docx'):
        return f"Error: File '{file_path}' is not a .docx file."
    
    # Get or create the backend instance
    backend = get_backend_for_tool(ctx, file_path)
    
    ctx.set_state("active_document_path", file_path)
    return f"Active document set to: {file_path}"

@mcp_server.tool()
def shutdown_word(ctx: Context) -> str:
    """
    Closes the document and shuts down the Word application instance.
    Should be called at the end of a session.
    """
    backend = ctx.get_state("word_backend")
    if backend and backend.word_app:
        try:
            backend.word_app.Quit()
            ctx.set_state("word_backend", None)
            return "Word application shut down successfully."
        except Exception as e:
            return f"Error shutting down Word: {e}"
    return "No active Word application to shut down."

@mcp_server.tool()
def insert_paragraph(ctx: Context, locator: Dict[str, Any], text: str, position: str = "after") -> str:
    """
    Inserts a new paragraph with the given text relative to the element found by the locator.
    """
    active_doc_path = ctx.get_state("active_document_path")
    if not active_doc_path:
        return "Error: No active document. Please use 'open_document' first."

    try:
        backend = get_backend_for_tool(ctx, active_doc_path)
        selection = selector.select(backend, locator)
        selection.insert_text(text, position)
        backend.document.Save()
        return f"Successfully inserted paragraph."
    except ElementNotFoundError as e:
        return f"Error finding element for insertion: {e}"
    except Exception as e:
        return f"An unexpected error occurred during insertion: {e}"


@mcp_server.tool()
def get_text_from_cell(ctx: Context, locator: Dict[str, Any]) -> str:
    """
    Retrieves the text from a single table cell found by the locator.
    """
    active_doc_path = ctx.get_state("active_document_path")
    if not active_doc_path:
        return "Error: No active document. Please use 'open_document' first."

    try:
        backend = get_backend_for_tool(ctx, active_doc_path)
        selection = selector.select(backend, locator)
        text = selection.get_text()
        return f"Success: '{text.strip().replace(chr(7), '')}'"
            
    except ElementNotFoundError as e:
        return f"Error finding element: {e}"
    except Exception as e:
        return f"An unexpected error occurred: {e}"


@mcp_server.tool()
def delete_element(ctx: Context, locator: Dict[str, Any]) -> str:
    """
    Deletes the element(s) found by the locator.

    Args:
        locator: The Locator object to find the target element(s) to delete.
    
    Returns:
        A success or error message.
    """
    active_doc_path = ctx.get_state("active_document_path")
    if not active_doc_path:
        return "Error: No active document. Please use 'open_document' first."

    try:
        backend = get_backend_for_tool(ctx, active_doc_path)
        selection = selector.select(backend, locator)
        selection.delete()
        backend.document.Save()
        return "Successfully deleted element(s)."
            
    except ElementNotFoundError as e:
        return f"Error finding element to delete: {e}"
    except Exception as e:
        return f"An unexpected error occurred during deletion: {e}"


@mcp_server.tool()
def get_text(ctx: Context, locator: Dict[str, Any]) -> str:
    """
    Retrieves the text from all elements found by the locator.
    """
    active_doc_path = ctx.get_state("active_document_path")
    if not active_doc_path:
        return "Error: No active document. Please use 'open_document' first."

    try:
        backend = get_backend_for_tool(ctx, active_doc_path)
        selection = selector.select(backend, locator)
        text = selection.get_text()
        return f"Success: '{text.strip()}'"
            
    except Exception as e:
        return f"An unexpected error occurred: {e}"


@mcp_server.tool()
def replace_text(ctx: Context, locator: Dict[str, Any], new_text: str) -> str:
    """
    Replaces the text content of the element(s) found by the locator with new text.
    
    Args:
        locator: The Locator object to find the target element(s) to replace.
        new_text: The new text to replace the existing content.
    
    Returns:
        A success or error message.
    """
    active_doc_path = ctx.get_state("active_document_path")
    if not active_doc_path:
        return "Error: No active document. Please use 'open_document' first."

    try:
        backend = get_backend_for_tool(ctx, active_doc_path)
        selection = selector.select(backend, locator)
        selection.replace_text(new_text)
        backend.document.Save()
        return "Successfully replaced text."
            
    except ElementNotFoundError as e:
        return f"Error finding element to replace: {e}"
    except Exception as e:
        return f"An unexpected error occurred during replacement: {e}"


@mcp_server.tool()
def set_cell_value(ctx: Context, locator: Dict[str, Any], text: str) -> str:
    """
    Sets the text value of a single table cell found by the locator.
    
    Args:
        locator: The Locator object to find the target cell.
        text: The text to set in the cell.
    
    Returns:
        A success or error message.
    """
    active_doc_path = ctx.get_state("active_document_path")
    if not active_doc_path:
        return "Error: No active document. Please use 'open_document' first."

    try:
        backend = get_backend_for_tool(ctx, active_doc_path)
        selection = selector.select(backend, locator)
        selection.replace_text(text)
        backend.document.Save()
        return "Successfully set cell value."
            
    except ElementNotFoundError as e:
        return f"Error finding cell: {e}"
    except Exception as e:
        return f"An unexpected error occurred: {e}"


@mcp_server.tool()
def create_table(ctx: Context, locator: Dict[str, Any], rows: int, cols: int) -> str:
    """
    Creates a new table at the location specified by the locator.
    
    Args:
        locator: The Locator object to find the anchor point for the new table.
        rows: Number of rows for the new table.
        cols: Number of columns for the new table.
    
    Returns:
        A success or error message.
    """
    active_doc_path = ctx.get_state("active_document_path")
    if not active_doc_path:
        return "Error: No active document. Please use 'open_document' first."

    try:
        backend = get_backend_for_tool(ctx, active_doc_path)
        # Use the selector to find the anchor point
        selection = selector.select(backend, locator)
        # Get the range of the first element in the selection
        anchor_range = selection._elements[0].Range
        # Add the table after the anchor range
        backend.add_table(anchor_range, rows, cols)
        backend.document.Save()
        return "Successfully created table."
            
    except ElementNotFoundError as e:
        return f"Error finding anchor point: {e}"
    except Exception as e:
        return f"An unexpected error occurred: {e}"




