import os
import json
from typing import Dict, Any, Optional
from mcp.server.fastmcp.server import Context
from word_document_server.core_utils import get_backend_for_tool, mcp_server
from word_document_server.selector import SelectorEngine
from word_document_server.selector import AmbiguousLocatorError
from word_document_server.errors import ElementNotFoundError, format_error_response


@mcp_server.tool()
def get_image_info(ctx: Context, locator: Optional[Dict[str, Any]] = None) -> str:
    """
    Retrieves information about all inline shapes (including images) in the document or matching the locator.

    Args:
        locator: Optional, the Locator object to find specific images. If not provided, returns all images.

    Returns:
        A JSON string containing a list of image information objects.
    """
    # Get active document path from session state
    active_doc_path = None
    if hasattr(ctx.session, 'document_state'):
        active_doc_path = ctx.session.document_state.get('active_document_path')
    if not active_doc_path:
        return "Error [2001]: No active document. Please use 'open_document' first."

    try:
        backend = get_backend_for_tool(ctx, active_doc_path)
        
        # Create selector engine instance
        selector_engine = SelectorEngine()
        
        if locator:
            # Use locator to select specific images
            try:
                selection = selector_engine.select(backend, locator)
                image_info = selection.get_image_info()
            except ElementNotFoundError as e:
                return f"Error [2002]: No elements found matching the locator: {e}" + " Please try simplifying your locator or use get_document_structure to check the actual document structure."
        else:
            # Get all inline shapes in the document
            image_info = backend.get_all_inline_shapes()
        
        # Convert to JSON string
        return json.dumps(image_info, ensure_ascii=False)
    except ValueError as e:
        return f"Error [1001]: Invalid parameter: {e}"
    except Exception as e:
        return format_error_response(e)


@mcp_server.tool()
def insert_inline_picture(ctx: Context, locator: Dict[str, Any], image_path: str, position: str = "after") -> str:
    """
    Inserts an inline picture at the location specified by the locator.

    Args:
        locator: The Locator object to find the anchor point for image insertion.
        image_path: The absolute path to the image file.
        position: "before", "after", or "replace" to specify where to insert the image relative to the anchor element.
    
    Returns:
        A success or error message.
    """
    # Get active document path from session state
    active_doc_path = None
    if hasattr(ctx.session, 'document_state'):
        active_doc_path = ctx.session.document_state.get('active_document_path')
    if not active_doc_path:
        return "Error [2001]: No active document. Please use 'open_document' first."
    
    # Validate image path
    if not os.path.isabs(image_path):
        return f"Error [1001]: Image path '{image_path}' is not an absolute path."
    if not os.path.exists(image_path):
        return f"Error [1001]: Image file '{image_path}' not found."
    
    # Validate image file extension
    valid_extensions = ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff']
    file_ext = os.path.splitext(image_path)[1].lower()
    if file_ext not in valid_extensions:
        return f"Error [1001]: File '{image_path}' is not a supported image format. Supported formats: {', '.join(valid_extensions)}"
    
    # Validate position
    if position not in ["before", "after", "replace"]:
        return "Error [1001]: Invalid position. Must be 'before', 'after', or 'replace'."

    try:
        backend = get_backend_for_tool(ctx, active_doc_path)
        # Create selector engine instance
        selector_engine = SelectorEngine()
        selection = selector_engine.select(backend, locator, expect_single=True)
        
        # Use the Selection's insert_image method which handles position correctly
        selection.insert_image(image_path, position)
        # Add None check for document
        if backend.document is None:
            raise ValueError("Failed to save document: No active document.")
        backend.document.Save()
        return f"Successfully inserted image from '{image_path}'."
    except ElementNotFoundError as e:
        return f"Error [2002]: Error finding anchor point: {e}" + " Please try simplifying your locator or use get_document_structure to check the actual document structure."
    except AmbiguousLocatorError as e:
        return f"Error [3001]: The locator found multiple elements. Please specify a unique anchor point. Details: {e}"
    except Exception as e:
        return f"An unexpected error occurred during image insertion: {e}"


@mcp_server.tool()
def set_image_size(ctx: Context, locator: Dict[str, Any], width: Optional[float] = None, height: Optional[float] = None, lock_aspect_ratio: bool = True) -> str:
    """
    Sets the size of images matching the locator.

    Args:
        locator: The Locator object to find the images to resize.
        width: The new width in points. If None, width remains unchanged.
        height: The new height in points. If None, height remains unchanged.
        lock_aspect_ratio: Whether to maintain the aspect ratio when changing dimensions.
    
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
        # Create selector engine instance
        selector_engine = SelectorEngine()
        selection = selector_engine.select(backend, locator)
        selection.set_image_size(width, height, lock_aspect_ratio)
        # Add None check for document
        if backend.document is None:
            raise ValueError("Failed to save document: No active document.")
        backend.document.Save()
        return "Successfully resized image(s)."
    except ElementNotFoundError as e:
        return f"Error [2002]: Error finding image(s): {e}" + " Please try simplifying your locator or use get_document_structure to check the actual document structure."
    except Exception as e:
        return f"An unexpected error occurred while resizing image(s): {e}"


@mcp_server.tool()
def set_image_color_type(ctx: Context, locator: Dict[str, Any], color_type: str) -> str:
    """
    Sets the color type of images matching the locator.

    Args:
        locator: The Locator object to find the images to modify.
        color_type: The color type to apply. Can be 'Color', 'Grayscale', 'BlackAndWhite', or 'Watermark'.
    
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
        # Create selector engine instance
        selector_engine = SelectorEngine()
        selection = selector_engine.select(backend, locator)
        selection.set_image_color_type(color_type)
        backend.document.Save()
        return f"Successfully set color type to '{color_type}' for image(s)."
    except ElementNotFoundError as e:
        return f"Error [2002]: Error finding image(s): {e}" + " Please try simplifying your locator or use get_document_structure to check the actual document structure."
    except ValueError as e:
        return f"Error [1001]: {e}"
    except Exception as e:
        return f"An unexpected error occurred while setting image color type: {e}"


@mcp_server.tool()
def delete_image(ctx: Context, locator: Dict[str, Any]) -> str:
    """
    Deletes images matching the locator.

    Args:
        locator: The Locator object to find the images to delete.
    
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
        # Create selector engine instance
        selector_engine = SelectorEngine()
        selection = selector_engine.select(backend, locator)
        selection.delete()
        backend.document.Save()
        return "Successfully deleted image(s)."
    except ElementNotFoundError as e:
        return f"Error [2002]: Error finding image(s): {e}" + " Please try simplifying your locator or use get_document_structure to check the actual document structure."
    except Exception as e:
        return f"An unexpected error occurred while deleting image(s): {e}"


@mcp_server.tool()
def add_picture_caption(ctx: Context, filename: str, caption_text: str, picture_index: Optional[int] = None, paragraph_index: Optional[int] = None) -> str:
    """
    Adds a caption to a picture in the active document.

    Args:
        filename: The name of the picture file (without path).
        caption_text: The text to use as caption.
        picture_index: Optional, the 0-based index of the picture to caption. If not provided, uses the first matching picture.
        paragraph_index: Optional, the 0-based index of the paragraph where the picture is located.

    Returns:
        A success or error message.
    """
    # Get active document path from session state
    active_doc_path = None
    if hasattr(ctx.session, 'document_state'):
        active_doc_path = ctx.session.document_state.get('active_document_path')
    if not active_doc_path:
        return "Error [2001]: No active document. Please use 'open_document' first."

    if not caption_text:
        return "Error [1001]: Caption text cannot be empty."

    try:
        backend = get_backend_for_tool(ctx, active_doc_path)
        
        # Call the backend method to add picture caption
        result = backend.add_picture_caption(filename, caption_text, picture_index, paragraph_index)
        
        # Save the document - add None check
        if backend.document is None:
            raise ValueError("Failed to save document: No active document.")
        backend.document.Save()
        
        return result
    except Exception as e:
        return f"An unexpected error occurred while adding picture caption: {e}"