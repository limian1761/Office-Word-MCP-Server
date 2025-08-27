from typing import Any, Dict, Optional

from mcp.server.fastmcp.server import Context
from pydantic import Field

from word_document_server import selector
from word_document_server.core_utils import get_backend_for_tool, mcp_server
from word_document_server.errors import (ElementNotFoundError,
                                         format_error_response,
                                         handle_tool_errors)
from word_document_server.operations import add_table
from word_document_server.selector import AmbiguousLocatorError


@mcp_server.tool()
@handle_tool_errors
def get_text_from_cell(
    ctx: Context = Field(description="Context object"),
    locator: Dict[str, Any] = Field(
        description="The Locator object to find the target cell"
    ),
) -> str:
    """
    Retrieves text from a single table cell found by the locator.

    Returns:
        The text content of the cell or an error message.
    """
    # Validate active document
    from word_document_server.core_utils import validate_active_document

    error = validate_active_document(ctx)
    if error:
        raise Exception(error)

    # Validate that locator is provided
    if not locator:
        raise Exception("Locator parameter is required.")

    try:
        backend = get_backend_for_tool(
            ctx, ctx.session.document_state["active_document_path"]
        )
        selector_engine = selector.SelectorEngine(backend)
        selection = selector_engine.select(locator, expect_single=True)

        # Verify we have exactly one cell
        if len(selection._elements) != 1:
            return "The locator must match exactly one cell."

        # Get the cell text
        cell = selection._elements[0]
        text = cell.Range.Text.strip()

        return text
    except ElementNotFoundError as e:
        return f"No cell found matching the locator: {e}. Please try simplifying your locator or use get_document_outline to check the actual document structure."
    except AmbiguousLocatorError as e:
        return f"The locator matched multiple cells: {e}. Please refine your locator to match a single cell."
    except AttributeError:
        return "The selected element is not a table cell."
    except Exception as e:
        return format_error_response(e)


@mcp_server.tool()
def set_cell_value(
    ctx: Context = Field(description="Context object"),
    locator: Dict[str, Any] = Field(
        description="The Locator object to find the target cell"
    ),
    text: str = Field(description="The text to set in the cell"),
) -> str:
    """
    Sets the text content of a table cell found by the locator.

    Returns:
        A success or error message.
    """
    # Validate active document
    from word_document_server.core_utils import validate_active_document

    error = validate_active_document(ctx)
    if error:
        return error

    # Validate that locator is provided
    if not locator:
        return "Locator parameter is required."

    try:
        backend = get_backend_for_tool(
            ctx, ctx.session.document_state["active_document_path"]
        )
        selector_engine = selector.SelectorEngine(backend)
        selection = selector_engine.select(locator, expect_single=True)

        # Verify we have exactly one cell
        if len(selection._elements) != 1:
            return "The locator must match exactly one cell."

        # Set the cell text
        cell = selection._elements[0]
        cell.Range.Text = text

        # Save the document
        backend.document.Save()
        return "Successfully set cell value."
    except ElementNotFoundError as e:
        return f"No cell found matching the locator: {e}. Please try simplifying your locator or use get_document_structure to check the actual document structure."
    except AmbiguousLocatorError as e:
        return f"The locator matched multiple cells: {e}. Please refine your locator to match a single cell."
    except AttributeError:
        return "The selected element is not a table cell."
    except Exception as e:
        # Handle COM errors with more specific messages
        error_message = str(e)
        if "COM" in str(type(e)) or "Dispatch" in str(type(e)):
            return "Failed to update cell content. This may occur if the table structure is corrupted or Word is in an unstable state."
        return format_error_response(e)


@mcp_server.tool()
def create_table(
    ctx: Context = Field(description="Context object"),
    locator: Dict[str, Any] = Field(
        description="The Locator object to find the anchor point for the new table"
    ),
    rows: int = Field(description="Number of rows for the new table"),
    cols: int = Field(description="Number of columns for the new table"),
) -> str:
    """
    Creates a new table at the location specified by the locator.

    Returns:
        A success or error message.
    """
    # Validate active document
    from word_document_server.core_utils import validate_active_document

    error = validate_active_document(ctx)
    if error:
        return error

    # Validate rows and cols parameters
    if not isinstance(rows, int) or rows <= 0:
        return "Invalid 'rows' parameter. Must be a positive integer."
    if not isinstance(cols, int) or cols <= 0:
        return "Invalid 'cols' parameter. Must be a positive integer."

    # Validate row and column limits (Word has practical limits)
    if rows > 32767 or cols > 63:
        return (
            "Table size exceeds Word's practical limits (max 32767 rows, 63 columns)."
        )

    try:
        backend = get_backend_for_tool(
            ctx, ctx.session.document_state["active_document_path"]
        )

        # Find the anchor point using the locator
        selector_engine = selector.SelectorEngine()
        anchor_selection = selector_engine.select(backend, locator, expect_single=True)

        # Get the COM range object from the selection
        anchor_element = anchor_selection._elements[0]
        # Check if anchor_element is already a Range object
        if hasattr(anchor_element, 'Start') and hasattr(anchor_element, 'End'):
            com_range_obj = anchor_element
        else:
            com_range_obj = anchor_element.Range

        # Add table using the backend function
        add_table(backend, com_range_obj, rows, cols)

        # Save the document
        backend.document.Save()
        return "Successfully created table."
    except ElementNotFoundError as e:
        return f"No elements found matching the locator: {e}. Please try simplifying your locator or use get_document_structure to check the actual document structure."
    except Exception as e:
        return format_error_response(e)
