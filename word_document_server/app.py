"""
Main MCP Server application file, built with the official mcp.server.fastmcp library.
"""
import os
from typing import Dict, Any, List

# Correct import for the MCP server components from the local SDK
from mcp.server.fastmcp.server import Context, FastMCP

from word_document_server.com_backend import WordBackend
from word_document_server.selector import SelectorEngine, ElementNotFoundError

# --- State Management ---
# Using Context.session to store state as recommended by MCP documentation

# --- MCP Server Initialization ---
mcp_server = FastMCP("Office-Word-MCP-Server")
selector = SelectorEngine()

# --- State Management ---
# A simple dictionary on the context object can manage state per-session.
# We will store the active document path in the context's state.

# --- Tool Definitions ---

def get_backend_for_tool(ctx: Context, file_path: str) -> WordBackend:
    """
    Gets or creates a WordBackend instance for the specified file path.
    """
    # Initialize session state if not exists
    if 'document_state' not in ctx.session:
        ctx.session['document_state'] = {}
    
    # Create a new backend instance
    backend = WordBackend(file_path=file_path, visible=True)
    backend.__enter__()
    
    # Store backend in session state
    ctx.session['document_state']['word_backend'] = backend
    
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
    
    # Store active document path in session state
    if 'document_state' not in ctx.session:
        ctx.session['document_state'] = {}
    ctx.session['document_state']['active_document_path'] = file_path
    return f"Active document set to: {file_path}"

@mcp_server.tool()
def shutdown_word(ctx: Context) -> str:
    """
    Closes the document and shuts down the Word application instance.
    Should be called at the end of a session.
    """
    # Get backend from session state
    backend = None
    if 'document_state' in ctx.session:
        backend = ctx.session['document_state'].get('word_backend')
    
    if backend and backend.word_app:
        try:
            backend.word_app.Quit()
            # Clear backend reference from session state
            if 'document_state' in ctx.session:
                ctx.session['document_state']['word_backend'] = None
            return "Word application shut down successfully."
        except Exception as e:
            return f"Error shutting down Word: {e}"
    return "No active Word application to shut down."

@mcp_server.tool()
def insert_paragraph(ctx: Context, locator: Dict[str, Any], text: str, position: str = "after") -> str:
    """
    Inserts a new paragraph with the given text relative to the element found by the locator.
    """
    # Get active document path from session state
    active_doc_path = None
    if 'document_state' in ctx.session:
        active_doc_path = ctx.session['document_state'].get('active_document_path')
    
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
    # Get active document path from session state
    active_doc_path = None
    if 'document_state' in ctx.session:
        active_doc_path = ctx.session['document_state'].get('active_document_path')
    
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
    # Get active document path from session state
    active_doc_path = None
    if 'document_state' in ctx.session:
        active_doc_path = ctx.session['document_state'].get('active_document_path')
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
    # Get active document path from session state
    active_doc_path = None
    if 'document_state' in ctx.session:
        active_doc_path = ctx.session['document_state'].get('active_document_path')
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
    # Get active document path from session state
    active_doc_path = None
    if 'document_state' in ctx.session:
        active_doc_path = ctx.session['document_state'].get('active_document_path')
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
    # Get active document path from session state
    active_doc_path = None
    if 'document_state' in ctx.session:
        active_doc_path = ctx.session['document_state'].get('active_document_path')
    if not active_doc_path:
        return "Error: No active document. Please use 'open_document' first."

    try:
        backend = get_backend_for_tool(ctx, active_doc_path)
        selection = selector.select(backend, locator, expect_single=True)
        selection.replace_text(text)
        backend.document.Save()
        return "Successfully set cell value."
            
    except ElementNotFoundError as e:
        return f"Error finding cell: {e}"
    except AmbiguousLocatorError as e:
        return f"Error: The locator found multiple cells. Please specify a unique cell. Details: {e}"
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
    # Get active document path from session state
    active_doc_path = None
    if 'document_state' in ctx.session:
        active_doc_path = ctx.session['document_state'].get('active_document_path')
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
    if 'document_state' in ctx.session:
        active_doc_path = ctx.session['document_state'].get('active_document_path')
    if not active_doc_path:
        return "Error: No active document. Please use 'open_document' first."

    try:
        backend = get_backend_for_tool(ctx, active_doc_path)
        backend.set_header_text(text)
        backend.document.Save()
        return "Header text set successfully."
    except Exception as e:
        return f"An unexpected error occurred while setting the header: {e}"


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
    if 'document_state' in ctx.session:
        active_doc_path = ctx.session['document_state'].get('active_document_path')
    if not active_doc_path:
        return "Error: No active document. Please use 'open_document' first."

    try:
        backend = get_backend_for_tool(ctx, active_doc_path)
        backend.set_footer_text(text)
        backend.document.Save()
        return "Footer text set successfully."
    except Exception as e:
        return f"An unexpected error occurred while setting the footer: {e}"


@mcp_server.tool()
def create_bulleted_list(ctx: Context, locator: Dict[str, Any], items: List[str], position: str = "after") -> str:
    """
    Creates a new bulleted list relative to the element found by the locator.

    Args:
        locator: The Locator object to find the anchor element.
        items: A list of strings to become the list items.
        position: "before" or "after" the anchor element.
    
    Returns:
        A success or error message.
    """
    # Get active document path from session state
    active_doc_path = None
    if 'document_state' in ctx.session:
        active_doc_path = ctx.session['document_state'].get('active_document_path')
    if not active_doc_path:
        return "Error: No active document. Please use 'open_document' first."

    if not items:
        return "Error: Cannot create an empty list."

    try:
        backend = get_backend_for_tool(ctx, active_doc_path)
        selection = selector.select(backend, locator, expect_single=True)
        
        # Use the range of the single element found as the anchor
        anchor_range = selection._elements[0].Range
        
        backend.create_bulleted_list_relative_to(anchor_range, items, position)
        backend.document.Save()
        return "Bulleted list created successfully."
            
    except (ElementNotFoundError, AmbiguousLocatorError) as e:
        return f"Error finding a unique anchor point: {e}"
    except ValueError as e:
        return f"Error creating list: {e}"
    except Exception as e:
        return f"An unexpected error occurred during list creation: {e}"


@mcp_server.tool()
def get_document_structure(ctx: Context) -> List[Dict[str, Any]]:
    """
    Provides a structured overview of the document by listing all headings.

    Returns:
        A list of dictionaries, each representing a heading with its text and level.
    """
    # Get active document path from session state
    active_doc_path = None
    if 'document_state' in ctx.session:
        active_doc_path = ctx.session['document_state'].get('active_document_path')
    if not active_doc_path:
        return []

    try:
        backend = get_backend_for_tool(ctx, active_doc_path)
        return backend.get_headings()
    except Exception:
        return []


@mcp_server.tool()
def accept_all_changes(ctx: Context) -> str:
    """
    Accepts all tracked revisions in the document.
    """
    # Get active document path from session state
    active_doc_path = None
    if 'document_state' in ctx.session:
        active_doc_path = ctx.session['document_state'].get('active_document_path')
    if not active_doc_path:
        return "Error: No active document. Please use 'open_document' first."

    try:
        backend = get_backend_for_tool(ctx, active_doc_path)
        backend.accept_all_changes()
        backend.document.Save()
        return "All changes accepted successfully."
    except Exception as e:
        return f"An unexpected error occurred: {e}"


@mcp_server.tool()
def apply_format(ctx: Context, locator: Dict[str, Any], formatting: Dict[str, Any]) -> str:
    """
    Applies specified formatting to the element(s) found by the locator.

    Args:
        locator: The Locator object to find the target element(s).
        formatting: A dictionary of formatting options to apply.
                    Example: {"bold": True, "alignment": "center"}
    
    Returns:
        A success or error message.
    """
    # Get active document path from session state
    active_doc_path = None
    if 'document_state' in ctx.session:
        active_doc_path = ctx.session['document_state'].get('active_document_path')
    if not active_doc_path:
        return "Error: No active document. Please use 'open_document' first."

    try:
        backend = get_backend_for_tool(ctx, active_doc_path)
        selection = selector.select(backend, locator)
        selection.apply_format(formatting)
        backend.document.Save()
        return "Formatting applied successfully."
            
    except ElementNotFoundError as e:
        return f"Error finding element to format: {e}"
    except ValueError as e:
        return f"Error applying format: {e}"
    except Exception as e:
        return f"An unexpected error occurred during formatting: {e}"




