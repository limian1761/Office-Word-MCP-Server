import os
import json
from typing import Dict, Any, Optional
from mcp.server.fastmcp.server import Context
from word_document_server.core_utils import get_backend_for_tool, mcp_server
from word_document_server.selector import SelectorEngine, AmbiguousLocatorError
from word_document_server.errors import format_error_response, handle_tool_errors
from word_document_server.operations import get_all_inline_shapes, add_picture_caption


@mcp_server.tool()
@handle_tool_errors
def insert_object(ctx: Context, locator: Dict[str, Any], object_path: str, object_type: str = "image", position: str = "after") -> str:
    """
    Inserts an object (image, file, or OLE object) at the location specified by the locator.

    Args:
        locator: The Locator object to find the anchor element.
        object_path: Absolute path to the object file.
        object_type: Type of object to insert ("image", "file", or "ole").
        position: Position relative to the anchor element ("before", "after", or "replace").

    Returns:
        A success or error message.
    """
    # Validate active document
    from word_document_server.core_utils import validate_active_document
    error = validate_active_document(ctx)
    if error:
        raise Exception(error)
    
    # Validate object path
    if not os.path.exists(object_path):
        raise Exception(f"Object path does not exist: {object_path}")

    # Validate object type
    supported_types = ["image", "file", "ole"]
    if object_type not in supported_types:
        return f"Unsupported object type '{object_type}'. Supported types: {', '.join(supported_types)}"
    
    # Validate position
    if position not in ["before", "after", "replace"]:
        return "Invalid position. Must be 'before', 'after', or 'replace'."

    try:
        backend = get_backend_for_tool(ctx, active_doc_path)
        # Create selector engine instance
        selector_engine = SelectorEngine()
        selection = selector_engine.select(backend, locator, expect_single=True)
        
        # Use the Selection's insert_object method which handles position correctly
        selection.insert_object(object_path, object_type, position)
        # Add None check for document
        if backend.document is None:
            raise ValueError("Failed to save document: No active document.")
        backend.document.Save()
        return f"Successfully inserted {object_type} object."
    except ElementNotFoundError as e:
        return f"No elements found matching the locator: {e}. Please try simplifying your locator or use get_document_structure to check the actual document structure."
    except ValueError as e:
        return f"Invalid parameter: {e}"
    except Exception as e:
        return format_error_response(e)


@mcp_server.tool()
def add_caption(ctx: Context, locator: Dict[str, Any], caption_text: str, label: str = "Figure", position: str = "below") -> str:
    """
    Adds a caption to an object (picture, table, etc.) found by the locator.

    Args:
        locator: The Locator object to find the target object.
        caption_text: The caption text to add.
        label: The label for the caption (e.g., "Figure", "Table", "Equation").
        position: Where to place the caption relative to the object. 
                 Supported values: "above", "below".

    Returns:
        A success or error message.
    """
    # Validate active document
    from word_document_server.core_utils import validate_active_document
    error = validate_active_document(ctx)
    if error:
        return error

    # Validate position
    if position not in ["above", "below"]:
        return "Invalid position. Must be 'above' or 'below'."

    try:
        backend = get_backend_for_tool(ctx, active_doc_path)
        
        # Get selector engine
        selector_engine = SelectorEngine()
        
        # Convert locator to Selection object
        selection = selector_engine.select(backend, locator, expect_single=True)
        
        # Add caption to the selected object
        selection.add_caption(caption_text, label, position)
        
        # Save the document
        backend.document.Save()
        
        return f"Successfully added {label} caption."
    except ElementNotFoundError as e:
        return f"No elements found matching the locator: {e}. Please try simplifying your locator or use get_document_structure to check the actual document structure."
    except ValueError as e:
        return f"Invalid parameter: {e}"
    except Exception as e:
        return format_error_response(e)


@mcp_server.tool()
def get_image_info(ctx: Context, locator: Optional[Dict[str, Any]] = None) -> str:
    """
    Retrieves information about images in the document.

    Args:
        locator: Optional, the Locator object to find specific images.
                If not provided, returns information about all images.

    Returns:
        A JSON string containing image information.
    """
    # Validate active document
    from word_document_server.core_utils import validate_active_document
    error = validate_active_document(ctx)
    if error:
        return error

    try:
        backend = get_backend_for_tool(ctx, active_doc_path)
        
        if locator:
            # Find specific images using locator
            selector_engine = SelectorEngine()
            selection = selector_engine.select(backend, locator)
            
            # Filter for only inline shapes (images)
            images = [element for element in selection._elements if hasattr(element, 'Type')]
        else:
            # Get all images
            images = get_all_inline_shapes(backend)
        
        # Collect image information
        image_info = []
        for i, image in enumerate(images):
            try:
                info = {
                    "index": i,
                    "type": image.Type if hasattr(image, 'Type') else "Unknown",
                    "width": image.Width if hasattr(image, 'Width') else "Unknown",
                    "height": image.Height if hasattr(image, 'Height') else "Unknown"
                }
                image_info.append(info)
            except Exception as e:
                # Skip images that cause errors
                continue
        
        # Convert to JSON string
        return json.dumps(image_info, ensure_ascii=False)
    except Exception as e:
        return format_error_response(e)
