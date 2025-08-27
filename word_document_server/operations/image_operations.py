"""
Image operations for Word Document MCP Server.

This module contains functions for image-related operations.
"""

from typing import Any, Dict, List, Optional

import pythoncom
import win32com.client

from word_document_server.errors import ErrorCode, WordDocumentError
from word_document_server.word_backend import WordBackend


def get_all_inline_shapes(backend: WordBackend) -> List[Dict[str, Any]]:
    """
    Retrieves all inline shapes (including pictures) in the document.

    Args:
        backend: The WordBackend instance.

    Returns:
        A list of dictionaries containing shape information, each with "index", "type", and "width" keys.
    """
    if not backend.document:
        raise RuntimeError("No document open.")

    shapes: List[Dict[str, Any]] = []
    try:
        # Check if InlineShapes property exists and is accessible
        if not hasattr(backend.document, "InlineShapes"):
            return shapes

        # Get all inline shapes from the document safely
        shapes_count = 0
        try:
            shapes_count = backend.document.InlineShapes.Count
        except Exception as e:
            print(f"Warning: Failed to access InlineShapes collection: {e}")
            return shapes

        for i in range(1, shapes_count + 1):
            try:
                shape = backend.document.InlineShapes(i)
                try:
                    shape_info = {
                        "index": i - 1,  # 0-based index
                        "type": (
                            _get_shape_type(shape.Type)
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


def insert_inline_picture(
    backend: WordBackend,
    com_range_obj: win32com.client.CDispatch,
    image_path: str,
    position: str = "after",
) -> win32com.client.CDispatch:
    """
    Inserts an inline picture at the specified range.

    Args:
        backend: The WordBackend instance.
        com_range_obj: The COM Range object where the picture will be inserted.
        image_path: The absolute path to the image file.
        position: "before", "after", or "replace" to specify where to insert the picture relative to the range.

    Returns:
        The newly inserted InlineShape COM object.
    """
    if not backend.document:
        raise RuntimeError("No document open.")

    if not image_path or not isinstance(image_path, str):
        raise ValueError("Invalid image path provided.")

    if not com_range_obj:
        raise ValueError("Invalid range object provided.")

    if position not in ["before", "after", "replace"]:
        raise ValueError("Invalid position. Must be 'before', 'after', or 'replace'.")

    try:
        # Create a duplicate of the range to avoid modifying the original
        insert_range = com_range_obj.Duplicate

        if position == "replace":
            # Delete the content of the range
            insert_range.Text = ""
            # Insert the picture
            return backend.document.InlineShapes.AddPicture(
                FileName=image_path,
                LinkToFile=False,
                SaveWithDocument=True,
                Range=insert_range,
            )
        elif position == "before":
            # Collapse the range to its start point
            insert_range.Collapse(1)  # wdCollapseStart
            # Insert the picture
            return backend.document.InlineShapes.AddPicture(
                FileName=image_path,
                LinkToFile=False,
                SaveWithDocument=True,
                Range=insert_range,
            )
        else:  # position == "after"
            # Collapse the range to its end point
            insert_range.Collapse(0)  # wdCollapseEnd
            # Insert the picture
            return backend.document.InlineShapes.AddPicture(
                FileName=image_path,
                LinkToFile=False,
                SaveWithDocument=True,
                Range=insert_range,
            )
    except Exception as e:
        raise WordDocumentError(f"Failed to insert picture '{image_path}': {e}")


def _get_shape_type(type_code: int) -> str:
    """
    Converts Word shape type code to human-readable string.

    Args:
        type_code: Shape type code from Word COM interface.

    Returns:
        Human-readable shape type.
    """
    # Word inline shape type constants
    shape_types = {
        1: "Picture",  # wdInlineShapePicture
        2: "LinkedPicture",  # wdInlineShapeLinkedPicture
        3: "Chart",  # wdInlineShapeChart
        4: "Diagram",  # wdInlineShapeDiagram
        5: "OLEControlObject",  # wdInlineShapeOLEControlObject
        6: "OLEObject",  # wdInlineShapeOLEObject
        7: "ActiveXControl",  # wdInlineShapeActiveXControl
        8: "SmartArt",  # wdInlineShapeSmartArt
        9: "3DModel",  # wdInlineShape3DModel
    }
    return shape_types.get(type_code, "Unknown")


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


def add_picture_caption(
    backend: WordBackend,
    filename: str,
    caption_text: str,
    picture_index: Optional[int] = None,
    paragraph_index: Optional[int] = None,
) -> None:
    """
    Adds a caption to a picture in the document.

    Args:
        backend: The WordBackend instance.
        filename: The filename of the document.
        caption_text: The caption text to add.
        picture_index: Optional index of the picture (0-based). If not specified, adds to first picture.
        paragraph_index: Optional index of the paragraph to add caption after.
    """
    if not backend.document:
        raise RuntimeError("No document open.")

    try:
        # Get all inline shapes (pictures)
        inline_shapes = backend.document.InlineShapes
        shape_count = inline_shapes.Count

        if shape_count == 0:
            raise WordDocumentError("No pictures found in the document")

        # Determine which picture to add caption to
        target_index = picture_index if picture_index is not None else 0
        if target_index < 0 or target_index >= shape_count:
            raise ValueError(
                f"Invalid picture index: {target_index}. Valid range is 0 to {shape_count - 1}."
            )

        # Get the target picture
        picture = inline_shapes(target_index + 1)  # COM is 1-based

        # Create a range after the picture for the caption
        caption_range = picture.Range
        caption_range.Collapse(0)  # wdCollapseEnd
        caption_range.InsertAfter("\n" + caption_text)

        # Apply caption style if available
        try:
            caption_range.Style = backend.document.Styles("Caption")
        except:
            # If Caption style doesn't exist, continue without it
            pass

    except Exception as e:
        raise WordDocumentError(f"Failed to add picture caption: {e}")
