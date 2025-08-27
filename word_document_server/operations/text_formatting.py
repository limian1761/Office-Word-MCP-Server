"""
Text formatting operations for Word Document MCP Server.

This module contains functions for text formatting operations.
"""
from typing import Any, Dict, List, Optional

import pythoncom
import win32com.client

from word_document_server.word_backend import WordBackend
from word_document_server.errors import WordDocumentError, ErrorCode

def set_bold_for_range(backend: WordBackend, com_range_obj: win32com.client.CDispatch, is_bold: bool):
    """
    Set bold formatting for a range.

    Args:
        backend: The WordBackend instance.
        com_range_obj: COM Range object to format.
        is_bold: Whether to set bold formatting.
    """
    com_range_obj.Font.Bold = is_bold

def set_italic_for_range(backend: WordBackend, com_range_obj: win32com.client.CDispatch, is_italic: bool):
    """
    Set italic formatting for a range.

    Args:
        backend: The WordBackend instance.
        com_range_obj: COM Range object to format.
        is_italic: Whether to set italic formatting.
    """
    com_range_obj.Font.Italic = is_italic

def set_font_size_for_range(backend: WordBackend, com_range_obj: win32com.client.CDispatch, size: int):
    """
    Set font size for a range.

    Args:
        backend: The WordBackend instance.
        com_range_obj: COM Range object to format.
        size: The font size in points.
    """
    com_range_obj.Font.Size = size

def set_font_color_for_range(backend: WordBackend, com_range_obj: win32com.client.CDispatch, color: str):
    """
    Set font color for a range.

    Args:
        backend: The WordBackend instance.
        com_range_obj: COM Range object to format.
        color: Named color (e.g., 'blue') or hex code (e.g., '#0000FF').
    """
    # Convert color name to Word's RGB color value or use hex code
    color_map = {
        'black': 0,
        'white': 16777215,
        'red': 255,
        'green': 65280,
        'blue': 16711680,
        'yellow': 65535
    }
    if color.lower() in color_map:
        com_range_obj.Font.Color = color_map[color.lower()]
    else:
        # Try to parse hex color (e.g., '#RRGGBB' or 'RRGGBB')
        color = color.lstrip('#')
        if len(color) == 6:
            try:
                rgb = int(color, 16)
                com_range_obj.Font.Color = rgb
            except ValueError:
                raise ValueError(f"Invalid hex color format: {color}")
        else:
            raise ValueError(f"Unsupported color: {color}. Use named color or 6-digit hex code.")

def set_font_name_for_range(backend: WordBackend, com_range_obj: win32com.client.CDispatch, name: str):
    """
    Set font name for a range.

    Args:
        backend: The WordBackend instance.
        com_range_obj: COM Range object to format.
        name: The name of the font.
    """
    com_range_obj.Font.Name = name

def insert_paragraph_after(backend: WordBackend, com_range_obj: win32com.client.CDispatch, text: str, style: str = None):
    """
    Insert a paragraph after a given range using the document's Paragraphs collection.

    Args:
        backend: The WordBackend instance.
        com_range_obj: COM Range object after which to insert.
        text: Text to insert.
        style: Optional, paragraph style name to apply.

    Returns:
        The newly created paragraph COM object.
    """
    # Collapse the range to its end point before inserting
    com_range_obj.Collapse(0)  # 0 corresponds to wdCollapseEnd
    insert_range = backend.document.Range(com_range_obj.Start, com_range_obj.End)
    
    # Add a new paragraph at this range.
    new_para = backend.document.Paragraphs.Add(insert_range)
    
    # Set the text for the new paragraph.
    new_para.Range.Text = text
    
    # Apply style if specified
    if style:
        try:
            new_para.Style = style
        except Exception as e:
            print(f"Warning: Failed to apply paragraph style '{style}': {e}")
            
    # Return the newly created paragraph object
    return new_para

def create_bulleted_list_relative_to(backend: WordBackend, com_range_obj: win32com.client.CDispatch, items: List[str], position: str):
    """
    Creates a new bulleted list relative to a given range.

    Args:
        backend: The WordBackend instance.
        com_range_obj: The range to insert the list before or after.
        items: A list of strings, where each string is a list item.
        position: "before" or "after".
    """
    if position == "before":
        insertion_point = com_range_obj.Start
    elif position == "after":
        insertion_point = com_range_obj.End
    else:
        raise ValueError("Position must be 'before' or 'after'.")

    # Collapse the range to the desired insertion point
    target_range = backend.document.Range(insertion_point, insertion_point)
    
    # Join items and insert the text block
    full_text = "\n".join(items) + "\n"
    target_range.InsertAfter(full_text)

    # Select the newly inserted text
    new_text_range = backend.document.Range(insertion_point, insertion_point + len(full_text))
    
    # Apply list format to each paragraph in the new range
    for para in new_text_range.Paragraphs:
        para.Range.ListFormat.ApplyBulletDefault()

def set_alignment_for_range(backend: WordBackend, com_range_obj: win32com.client.CDispatch, alignment: str):
    """
    Set paragraph alignment for a range.

    Args:
        backend: The WordBackend instance.
        com_range_obj: COM Range object to format.
        alignment: "left", "center", or "right".
    """
    alignment_map = {
        "left": 0,    # wdAlignParagraphLeft
        "center": 1,  # wdAlignParagraphCenter
        "right": 2    # wdAlignParagraphRight
    }
    if alignment.lower() in alignment_map:
        com_range_obj.ParagraphFormat.Alignment = alignment_map[alignment.lower()]
    else:
        raise ValueError(f"Invalid alignment value: {alignment}. Must be 'left', 'center', or 'right'.")