from typing import Dict, Any, Optional
from mcp.server.fastmcp.server import Context
from word_document_server.core_utils import get_backend_for_tool, mcp_server
from word_document_server.core import selector
from word_document_server.errors import ElementNotFoundError, format_error_response
from word_document_server.selector import AmbiguousLocatorError


@mcp_server.tool()
def get_text_from_cell(ctx: Context, locator: Dict[str, Any]) -> str:
    """
    Retrieves the text from a single table cell found by the locator.

    Args:
        locator: The Locator object to find the target cell.

    Returns:
        A string containing the retrieved text content or an error message.
    """
    # Get active document path from session state
    active_doc_path = None
    if hasattr(ctx.session, 'document_state'):
        active_doc_path = ctx.session.document_state.get('active_document_path')
    if not active_doc_path:
        return "Error [2001]: No active document. Please use 'open_document' first."

    try:
        backend = get_backend_for_tool(ctx, active_doc_path)
        selection = selector.select(backend, locator, expect_single=True)
        
        # Verify we have exactly one cell
        if len(selection._elements) != 1:
            return "Error [3001]: The locator must match exactly one cell."
        
        # Get the cell text
        cell_text = selection._elements[0].Range.Text.strip()
        return cell_text
    except ElementNotFoundError as e:
        return f"Error [2002]: No cell found matching the locator: {e}" + " Please try simplifying your locator or use get_document_structure to check the actual document structure."
    except AmbiguousLocatorError as e:
        return f"Error [3001]: The locator matched multiple cells: {e}" + " Please refine your locator to match a single cell."
    except AttributeError:
        return "Error [3002]: The selected element is not a table cell."
    except Exception as e:
        # Handle COM errors with more specific messages
        error_message = str(e)
        if "COM" in str(type(e)) or "Dispatch" in str(type(e)):
            return "Error [7001]: Failed to access cell content. This may occur if the table structure is corrupted or Word is in an unstable state."
        return format_error_response(e)


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
    if hasattr(ctx.session, 'document_state'):
        active_doc_path = ctx.session.document_state.get('active_document_path')
    if not active_doc_path:
        return "Error [2001]: No active document. Please use 'open_document' first."

    # Validate that locator is provided
    if not locator:
        return "Error [1001]: Locator parameter is required."

    try:
        backend = get_backend_for_tool(ctx, active_doc_path)
        selection = selector.select(backend, locator, expect_single=True)
        
        # Verify we have exactly one cell
        if len(selection._elements) != 1:
            return "Error [3001]: The locator must match exactly one cell."
        
        # Set the cell text
        cell = selection._elements[0]
        cell.Range.Text = text
        
        # Save the document
        backend.document.Save()
        return "Successfully set cell value."
    except ElementNotFoundError as e:
        return f"Error [2002]: No cell found matching the locator: {e}" + " Please try simplifying your locator or use get_document_structure to check the actual document structure."
    except AmbiguousLocatorError as e:
        return f"Error [3001]: The locator matched multiple cells: {e}" + " Please refine your locator to match a single cell."
    except AttributeError:
        return "Error [3002]: The selected element is not a table cell."
    except Exception as e:
        # Handle COM errors with more specific messages
        error_message = str(e)
        if "COM" in str(type(e)) or "Dispatch" in str(type(e)):
            return "Error [7001]: Failed to update cell content. This may occur if the table structure is corrupted or Word is in an unstable state."
        return format_error_response(e)


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
    if hasattr(ctx.session, 'document_state'):
        active_doc_path = ctx.session.document_state.get('active_document_path')
    if not active_doc_path:
        return "Error [2001]: No active document. Please use 'open_document' first."

    # Validate rows and cols parameters
    if not isinstance(rows, int) or rows <= 0:
        return "Error [1001]: Invalid 'rows' parameter. Must be a positive integer."
    if not isinstance(cols, int) or cols <= 0:
        return "Error [1001]: Invalid 'cols' parameter. Must be a positive integer."

    # Validate row and column limits (Word has practical limits)
    if rows > 32767 or cols > 63:
        return "Error [1001]: Table size exceeds Word's practical limits (max 32767 rows, 63 columns)."

    try:
        backend = get_backend_for_tool(ctx, active_doc_path)
        
        # Find the anchor point using the locator
        try:
            selection = selector.select(backend, locator, expect_single=True)
            anchor_range = selection._elements[0].Range
        except ElementNotFoundError:
            # If no locator or element not found, use the end of the document as default
            anchor_range = backend.document.Content
            anchor_range.Collapse(1)  # Collapse to end
        
        # Create the table
        table = backend.document.Tables.Add(
            Range=anchor_range,
            NumRows=rows,
            NumColumns=cols
        )
        
        # Save the document
        backend.document.Save()
        
        return f"Successfully created table with {rows} rows and {cols} columns."
    except AttributeError:
        return "Error [3002]: Failed to create table at the specified location. The selected element may not support table insertion."
    except Exception as e:
        # Handle COM errors with more specific messages
        error_message = str(e)
        if "COM" in str(type(e)) or "Dispatch" in str(type(e)):
            return "Error [7001]: Failed to create table. This may occur if Word is in an unstable state or if there's insufficient memory."
        return format_error_response(e)