import json
import os
from typing import Any, Dict, List, Optional

from mcp.server.fastmcp import Context
from mcp.server.session import ServerSession
from pydantic import Field

from word_document_server.mcp_service.core import mcp_server, selector
from word_document_server.utils.app_context import AppContext
from word_document_server.utils.core_utils import (ElementNotFoundError,
                                           format_error_response,
                                          handle_tool_errors)
from word_document_server.utils.core_utils import ErrorCode, WordDocumentError
from word_document_server.utils.core_utils import get_shape_types


def get_color_type(color_code: int) -> str:
    """
    Converts Word picture color type code to human-readable string.
    This function is moved here from core_utils.py to avoid duplication.

    Args:
        color_code: Color type code from Word COM interface.

    Returns:
        Human-readable color type.
    """
    # Word picture color type constants
    color_types = {
        0: "Color",  # msoPictureColorTypeColor
        1: "Grayscale",  # msoPictureColorTypeGrayscale
        2: "BlackAndWhite",  # msoPictureColorTypeBlackAndWhite
        3: "Watermark",  # msoPictureColorTypeWatermark
    }
    return color_types.get(color_code, "Unknown")


def get_all_inline_shapes(document) -> List[Dict[str, Any]]:
    """
    Retrieves all inline shapes (including pictures) in the document.

    Args:
        document: The Word document COM object.

    Returns:
        A list of dictionaries containing shape information, each with "index", "type", and "width" keys.
    """
    if not document:
        raise RuntimeError("No document open.")

    shapes: List[Dict[str, Any]] = []
    try:
        # Check if InlineShapes property exists and is accessible
        if not hasattr(document, "InlineShapes"):
            return shapes

        # Get all inline shapes from the document
        shapes_count = 0
        try:
            shapes_count = document.InlineShapes.Count
        except Exception as e:
            raise WordDocumentError(ErrorCode.IMAGE_NOT_FOUND, f"Failed to access InlineShapes collection: {e}")

        for i in range(1, shapes_count + 1):
            try:
                shape = document.InlineShapes(i)
                try:
                    shape_info = {
                        "index": i - 1,  # 0-based index
                        "type": (
                            get_shape_types().get(shape.Type, "Unknown")
                            if hasattr(shape, "Type")
                            else "Unknown"
                        ),
                        "width": shape.Width if hasattr(shape, "Width") else 0,
                        "height": shape.Height if hasattr(shape, "Height") else 0,
                    }
                    # Add additional properties based on shape type
                    if shape_info["type"] == "Picture":
                        # Try to get picture format information if available
                        if hasattr(shape, "PictureFormat"):
                            if hasattr(shape.PictureFormat, "ColorType"):
                                shape_info["color_type"] = _get_color_type(
                                    shape.PictureFormat.ColorType
                                )
                    shapes.append(shape_info)
                except Exception as e:
                    print(
                        f"Warning: Failed to retrieve shape information for index {i}: {e}"
                    )
                    continue
            except Exception as e:
                print(f"Warning: Failed to access shape at index {i}: {e}")
                continue
    except Exception as e:
        print(f"Error: Failed to retrieve inline shapes: {e}")

    return shapes


@mcp_server.tool()
@require_active_document_validation
def insert_object(
    ctx: Context[ServerSession, AppContext] = Field(description="Context object"),
    locator: Dict[str, Any] = Field(
        description="The Locator object to find the anchor element"
    ),
    object_path: str = Field(description="Absolute path to the object file"),
    object_type: str = Field(
        description='Type of object to insert ("image", "file", or "ole")',
        default="image",
    ),
    position: str = Field(
        description='Position relative to the anchor element ("before", "after", or "replace")',
        default="after",
    ),
) -> str:
    """
    Inserts an object (image, file, or OLE object) at the location specified by the locator.

    Returns:
        A success or error message.
    """

    # Validate object path
    if not os.path.exists(object_path):
        raise Exception(f"Object path does not exist: {object_path}")

    # Validate object type
    supported_types = ["image", "file", "ole"]
    if object_type not in supported_types:
        return f"Unsupported object type '{object_type}'. Supported types: {', '.join(supported_types)}"

    from word_document_server.utils.core_utils import validate_insert_position
    validation_error = validate_insert_position(position)
    if validation_error:
        return validation_error

    try:
        active_doc = ctx.request_context.lifespan_context.get_active_document()
        if active_doc is None:
            raise ValueError("No active document.")

        # Convert locator to Selection object
        selection = selector.select(active_doc, locator, expect_single=True)

        # Use the Selection's insert_object method which handles position correctly
        selection.insert_object(object_path, object_type, position)
        # Add None check for document
        if active_doc is None:
            raise ValueError("Failed to save document: No active document.")
        active_doc.Save()
        return f"Successfully inserted {object_type} object."
    except Exception as e:
        return f"Error inserting {object_type} object: {str(e)}"



@mcp_server.tool()
@require_active_document_validation
def add_caption(
    ctx: Context[ServerSession, AppContext] = Field(description="Context object"),
    locator: Dict[str, Any] = Field(
        description="The Locator object to find the target object"
    ),
    caption_text: str = Field(description="The caption text to add"),
    label: str = Field(
        description='The label for the caption (e.g., "Figure", "Table", "Equation")',
        default="Figure",
    ),
    position: str = Field(
        description='Where to place the caption relative to the object ("above" or "below")',
        default="below",
    ),
) -> str:
    """
    Adds a caption to an object (picture, table, etc.) found by the locator.

    Returns:
        A success or error message.
    """
    # Validate position
    from word_document_server.utils.core_utils import validate_position
    pos_error = validate_position(position)
    if pos_error:
        return pos_error

    try:
        active_doc = ctx.request_context.lifespan_context.get_active_document()
        selector_engine = selector.SelectorEngine()
        selection = selector_engine.select(active_doc, locator, expect_single=True)

        # Add caption using the Selection method
        selection.add_caption(caption_text, label, position)

        # Save the document
        active_doc.Save()
        return "Successfully added caption."
    except Exception as e:
        return f"Error adding caption: {str(e)}"


@mcp_server.tool()
@require_active_document_validation
def get_images_info(ctx: Context[ServerSession, AppContext] = Field(description="Context object")) -> str:
    """
    Retrieves information about all images (inline shapes) in the active document.

    Returns:
        A JSON string containing a list of images with their information.
    """
    try:
        active_doc = ctx.request_context.lifespan_context.get_active_document()

        # Get all inline shapes using Selection method
        from word_document_server.selector.selection import Selection
        selection = Selection([active_doc], active_doc)
        images_info = selection.get_image_info()

        # Convert to JSON string
        return json.dumps(images_info, ensure_ascii=False)
    except Exception as e:
        return format_error_response(e)


@mcp_server.tool()
@require_active_document_validation
def set_image_color_type(
    ctx: Context[ServerSession, AppContext] = Field(description="Context object"),
    locator: Dict[str, Any] = Field(
        description="The Locator object to find the target image(s)"
    ),
    color_type: int = Field(
        description="The color type to set (0=Color, 1=Grayscale, 2=BlackAndWhite, 3=Watermark)"
    ),
) -> str:
    """
    Sets the color type for images found by the locator.

    Returns:
        A success or error message.
    """
    try:
        active_doc = ctx.request_context.lifespan_context.get_active_document()
        
        # Use the shared selector engine from core
        from word_document_server.mcp_service.core import selector

        # Convert locator to Selection object
        selection = selector.select(active_doc, locator)

        # Validate that we have elements to modify
        if not selection._elements:
            return "No images found matching the locator. Please try simplifying your locator."

        # Use the new Selection method to set color type
        selection.set_picture_color_type(color_type)

        # Save the document
        active_doc.Save()

        return "Image color type updated successfully."
    except Exception as e:
        return format_error_response(e)

from word_document_server.tools.tool_imports import *
from word_document_server.tools.base_tool import BaseWordTool
