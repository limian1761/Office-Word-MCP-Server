"""
Element operations for Word Document MCP Server.
This module contains operations that work on single document elements.
"""

from typing import Any, Dict, List, Optional
import logging

import win32com.client

from word_document_server.utils.errors import ErrorCode, WordDocumentError
from word_document_server.com_backend.com_utils import handle_com_error


# === Text Element Operations ===

def get_element_text(element: win32com.client.CDispatch) -> str:
    """
    Get the text content of a single element.

    Args:
        element: The COM object representing the element.

    Returns:
        str: The text content of the element.
    """
    element_text = ""
    if hasattr(element, "Text"):
        element_text = element.Text()
    elif hasattr(element, "Range") and hasattr(element.Range, "Text"):
        element_text = element.Range.Text
    return element_text


@handle_com_error(ErrorCode.SERVER_ERROR, "insert text before element")
def insert_text_before_element(document: win32com.client.CDispatch, element: win32com.client.CDispatch, text: str, style: Optional[str] = None) -> bool:
    """Insert text before a single element

    Args:
        document: Document object.
        element: Single document element object.
        text: Text to insert.
        style: Optional style name.

    Returns:
        bool: Operation success status.
    """
    anchor_range = element.Range
    new_range = anchor_range.Duplicate
    new_range.Collapse(1)  # wdCollapseStart = 1
    new_range.InsertAfter(text + "\r")
    
    # Apply style
    if style:
        paragraph = new_range.Paragraphs(1)
        paragraph.Style = style
    
    return True


@handle_com_error(ErrorCode.SERVER_ERROR, "insert text after element")
def insert_text_after_element(document: win32com.client.CDispatch, element: win32com.client.CDispatch, text: str, style: Optional[str] = None) -> bool:
    """Insert text after a single element

    Args:
        document: Document object.
        element: Single document element object.
        text: Text to insert.
        style: Optional style name.

    Returns:
        bool: Operation success status.
    """
    anchor_range = element.Range
    new_range = anchor_range.Duplicate
    new_range.Collapse(0)  # wdCollapseEnd = 0
    new_range.InsertAfter("\r" + text)
    
    # Apply style
    if style:
        paragraph = new_range.Paragraphs(1)
        paragraph.Style = style
    
    return True


@handle_com_error(ErrorCode.SERVER_ERROR, "replace element text")
def replace_element_text(document: win32com.client.CDispatch, element: win32com.client.CDispatch, new_text: str, style: Optional[str] = None) -> bool:
    """Replace the text content of a single element

    Args:
        document: Document object.
        element: Single document element object.
        new_text: Replacement text.
        style: Optional style name.

    Returns:
        bool: Operation success status.
    """
    if hasattr(element, "Range"):
        element.Range.Text = new_text
        
        # Apply style
        if style and hasattr(element, "Style"):
            element.Style = style
                
        return True
    return False

@handle_com_error(ErrorCode.SERVER_ERROR, "delete element")
def delete_element(element: win32com.client.CDispatch) -> bool:
    """Delete a single element

    Args:
        element: The COM object representing the element to delete.

    Returns:
        bool: Operation success status.
    """
    if hasattr(element, "Range"):
        element.Range.Delete()
        return True
    return False

@handle_com_error(ErrorCode.IMAGE_FORMAT_ERROR, "get element image info")
def get_element_image_info(element: win32com.client.CDispatch) -> Dict[str, Any]:
    """Get information about an image element

    Args:
        element: The COM object representing the image element.

    Returns:
        Dict: Dictionary containing image information.
    """
    image_info = {}
    if hasattr(element, "Width"):
        image_info["width"] = element.Width
    if hasattr(element, "Height"):
        image_info["height"] = element.Height
    if hasattr(element, "PictureFormat"):
        image_info["has_picture_format"] = True
    return image_info

@handle_com_error(ErrorCode.SERVER_ERROR, "insert object relative to element")
def insert_object_relative_to_element(
    document: win32com.client.CDispatch,
    target_element: win32com.client.CDispatch,
    object_path: str,
    position: str = "after"
) -> bool:
    """Insert an object relative to a target element

    Args:
        document: Document object.
        target_element: Target element to insert relative to.
        object_path: Path to the object file.
        position: Position relative to target element ("before" or "after").

    Returns:
        bool: Operation success status.
    """
    try:
        range_obj = target_element.Range.Duplicate
        if position == "after":
            range_obj.Collapse(0)  # wdCollapseEnd
        else:
            range_obj.Collapse(1)  # wdCollapseStart
        range_obj.InsertFile(FileName=object_path)
        return True
    except Exception:
        return False


# === Image Element Operations ===

@handle_com_error(ErrorCode.SERVER_ERROR, "set picture element color type")
def set_picture_element_color_type(document: win32com.client.CDispatch, element: win32com.client.CDispatch, color_code: int) -> bool:
    """Set the color type of a single image element

    Args:
        document: Document object.
        element: Single image element object.
        color_code: Color type code (0-3).

    Returns:
        bool: Operation success status.
    """
    if hasattr(element, "Type") and (element.Type == 1 or element.Type == 2):
        if hasattr(element, "PictureFormat") and hasattr(element.PictureFormat, "ColorType"):
            element.PictureFormat.ColorType = color_code
            return True
    return False


# === Caption Operations ===

@handle_com_error(ErrorCode.SERVER_ERROR, "add element caption")
def add_element_caption(document: win32com.client.CDispatch, element: win32com.client.CDispatch, caption_text: str, label: str = "Figure", position: str = "below") -> bool:
    """Add a caption to an element

    Args:
        document: Document object.
        element: Element object to add caption to.
        caption_text: Caption text.
        label: Label (e.g. "Figure", "Table").
        position: Position ("above" or "below").

    Returns:
        bool: Operation success status.
    """
    anchor_range = element.Range
    new_range = anchor_range.Duplicate
    
    if position.lower() == "above":
        new_range.Collapse(1)  # wdCollapseStart = 1
    else:  # below
        new_range.Collapse(0)  # wdCollapseEnd = 0
        
    # Add caption paragraph
    caption_paragraph = new_range.Paragraphs.Add()
    caption_range = caption_paragraph.Range
    
    # Insert caption text
    if position.lower() == "above":
        caption_range.InsertBefore(f"{label} {caption_text}")
    else:  # below
        caption_range.InsertAfter(f"{label} {caption_text}")
        
    return True