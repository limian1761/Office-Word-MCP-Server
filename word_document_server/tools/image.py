import json
import os
from typing import Any, Dict, List, Optional

from mcp.server.fastmcp.server import Context
from pydantic import Field
from word_document_server.core import ServerSession
from word_document_server.utils.app_context import AppContext

from word_document_server.core import selector as selector_engine
from word_document_server.core_utils import mproxy_server
from word_document_server.errors import (ElementNotFoundError,
                                         format_error_response,
                                         handle_tool_errors)
from word_document_server.errors import ErrorCode, WordDocumentError
from word_document_server.utils.core_utils import get_shape_types


def _get_color_type(color_code: int) -> str:
    """
    Converts Word picture color type code to human-readable string.

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
            raise WordDocumentError(ErrorCode.IMAGE_ERROR, f"Failed to access InlineShapes collection: {e}")

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
@standardize_tool_errors
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

        selection = selector_engine.select(active_doc, locator, expect_single=True)

        # Use the Selection's insert_object method which handles position correctly
        selection.insert_object(object_path, object_type, position)
        # Add None check for document
        if active_doc is None:
            raise ValueError("Failed to save document: No active document.")
        active_doc.Save()
        return f"Successfully inserted {object_type} object."



@mcp_server.tool()
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
        description='Where to place the caption relative to the object. Supported values: "above", "below"',
        default="below",
    ),
) -> str:
    """
    Adds a caption to an object (picture, table, etc.) found by the locator.

    Returns:
        A success or error message.
    """

    # Validate position
    if position not in ["above", "below"]:
        return "Invalid position. Must be 'above' or 'below'."

    try:
        active_doc = ctx.request_context.lifespan_context.get_active_document()

        # Convert locator to Selection object
        selection = selector_engine.select(active_doc, locator, expect_single=True)

        # Add caption to the selected object
        selection.add_caption(caption_text, label, position)

        # Save the document
        active_doc.Save()

        return f"Successfully added {label} caption."
    except ElementNotFoundError as e:
        return f"No elements found matching the locator: {e}. Please try simplifying your locator or use get_document_outline to check the actual document structure."
    except ValueError as e:
        return f"Invalid parameter: {e}"
    except Exception as e:
        return format_error_response(e)


@mcp_server.tool()
def get_image_info(
    ctx: Context[ServerSession, AppContext] = Field(description="Context object"),
    locator: Optional[Dict[str, Any]] = Field(
        description="Optional, the Locator object to find specific images. If not provided, returns information about all images",
        default=None,
    ),
) -> str:
    """
    Retrieves information about images in the document.

    Returns:
        A JSON string containing image information.
    """


    try:
        active_doc = ctx.request_context.lifespan_context.get_active_document()

        if locator:
            # Find specific images using locator
            from word_document_server.core import selector as selector_engine

            selection = selector_engine.select(active_doc, locator)

            # Filter for only inline shapes (images)
            images = [
                element for element in selection._elements if hasattr(element, "Type")
            ]
        else:
            # Get all images
            images = get_all_inline_shapes(active_doc)

        # Collect image information
        image_info = []
        # Check if images are already dictionaries (from get_all_inline_shapes)
        if images and isinstance(images[0], dict):
            # Already processed by get_all_inline_shapes
            image_info = images
        else:
            # Process raw COM objects
            for i, image in enumerate(images):
                try:
                    info = {
                        "index": i,
                        "type": image.Type if hasattr(image, "Type") else "Unknown",
                        "width": image.Width if hasattr(image, "Width") else "Unknown",
                        "height": image.Height if hasattr(image, "Height") else "Unknown",
                    }
                    image_info.append(info)
                except Exception as e:
                    # Skip images that cause errors
                    continue

        # Convert to JSON string
        return json.dumps(image_info, ensure_ascii=False)
    except Exception as e:
        return format_error_response(e)
