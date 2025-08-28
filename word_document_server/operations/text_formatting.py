"""
Text formatting operations for Word Document MCP Server.
"""

from typing import Any, Dict, List, Optional, Union

import win32com.client

from word_document_server.errors import ErrorCode, WordDocumentError
from word_document_server.utils.com_utils import handle_com_error, safe_com_call
import logging


def replace_element_text(element, new_text: str, style: Optional[str] = None) -> bool:
    """元操作：替换单个元素的文本内容

    Args:
        element: 单个文档元素对象
        new_text: 替换后的文本
        style: 可选样式名称

    Returns:
        bool: 操作成功状态
    """
    try:
        if hasattr(element, "Range"):
            element.Range.Text = new_text
            
        # 应用样式
        if style and hasattr(element, "Style"):
            try:
                element.Style = style
            except Exception as e:
                logging.warning(f"应用样式失败: {str(e)}")
                
        return True
    except Exception as e:
        logging.error(f"文本替换失败: {str(e)}")
        return False


def set_paragraph_style(element: win32com.client.CDispatch, style_name: str) -> None:
    """
    Applies a paragraph style to the element.

    Args:
        element: The COM object representing the element (e.g., a Paragraph).
        style_name: The name of the style to apply.
    """
    try:
        element.Style = style_name
    except Exception as e:
        # Log the error but continue processing
        import logging
        logging.error(
            f"Failed to apply paragraph style '{style_name}': {str(e)}"
        )


@handle_com_error(ErrorCode.TEXT_FORMATTING_ERROR, "set bold formatting")
def set_bold_for_range(
    com_range_obj: win32com.client.CDispatch, is_bold: bool
):
    """
    Set bold formatting for a range.

    Args:
        com_range_obj: COM Range object to format.
        is_bold: Whether to set bold formatting.
    """
    com_range_obj.Bold = 1 if is_bold else 0


@handle_com_error(ErrorCode.TEXT_FORMATTING_ERROR, "set italic formatting")
def set_italic_for_range(
    com_range_obj: win32com.client.CDispatch, is_italic: bool
):
    """
    Set italic formatting for a range.

    Args:
        com_range_obj: COM Range object to format.
        is_italic: Whether to set italic formatting.
    """
    com_range_obj.Font.Italic = is_italic


@handle_com_error(ErrorCode.TEXT_FORMATTING_ERROR, "set font size")
def set_font_size_for_range(
    com_range_obj: win32com.client.CDispatch, size: int
):
    """
    Set font size for a range.

    Args:
        com_range_obj: COM Range object to format.
        size: The font size in points.
    """
    com_range_obj.Font.Size = size


def set_font_color_for_range(
    document: win32com.client.CDispatch, com_range_obj: win32com.client.CDispatch, color: str
):
    """
    Set font color for a range.

    Args:
        document: The Word document COM object.
        com_range_obj: COM Range object to format.
        color: Named color (e.g., 'blue') or hex code (e.g., '#0000FF').
    """
    color_map = {
        "black": 0,
        "white": 16777215,
        "red": 255,
        "green": 65280,
        "blue": 16711680,
        "yellow": 65535,
    }
    if color.lower() in color_map:
        with safe_com_call(ErrorCode.TEXT_FORMATTING_ERROR, "set font color"):
            com_range_obj.Font.Color = color_map[color.lower()]
    else:
        # Try to parse hex color (e.g., '#RRGGBB' or 'RRGGBB')
        color = color.lstrip("#")
        if len(color) == 6:
            try:
                rgb = int(color, 16)
                with safe_com_call(ErrorCode.TEXT_FORMATTING_ERROR, "set font color"):
                    com_range_obj.Font.Color = rgb
            except ValueError:
                raise WordDocumentError(
                    ErrorCode.TEXT_FORMATTING_ERROR,
                    f"Invalid hex color format: {color}"
                )
        else:
            raise WordDocumentError(
                ErrorCode.TEXT_FORMATTING_ERROR,
                f"Unsupported color: {color}. Use named color or 6-digit hex code."
            )


@handle_com_error(ErrorCode.TEXT_FORMATTING_ERROR, "set font name")
def set_font_name_for_range(
    com_range_obj: win32com.client.CDispatch, font_name: str
):
    """
    Set font name for a range.

    Args:
        com_range_obj: COM Range object to format.
        font_name: The font name to set.
    """
    com_range_obj.Font.Name = font_name


@handle_com_error(ErrorCode.TEXT_FORMATTING_ERROR, "insert paragraph")
def insert_paragraph_after(
    com_range_obj: win32com.client.CDispatch
):
    """
    Insert a new paragraph after the given range.

    Args:
        com_range_obj: The range after which to insert the paragraph.
    """
    # Collapse the range to its end
    com_range_obj.Collapse(Direction=0)  # wdCollapseEnd
    # Insert paragraph marks
    com_range_obj.InsertParagraphAfter()


@handle_com_error(ErrorCode.TEXT_FORMATTING_ERROR, "insert bulleted list")
def insert_bulleted_list(
    document: win32com.client.CDispatch,
    com_range_obj: win32com.client.CDispatch,
    items: list,
    style: str = "BulletDefault",
):
    """
    Insert a bulleted list at the given range.

    Args:
        document: The Word document COM object.
        com_range_obj: The range where to insert the list.
        items: List of strings to add as bullet items.
        style: The bullet style to use.
    """
    # Insert paragraphs for each item
    for i, item in enumerate(items):
        if i > 0:
            com_range_obj.InsertParagraphAfter()
            com_range_obj.Collapse(Direction=0)  # wdCollapseEnd
        com_range_obj.Text = item
        # Apply bullet style
        com_range_obj.ParagraphFormat.ListTemplate = document.ListTemplates(1)
        com_range_obj.ParagraphFormat.ListTemplate.Apply(com_range_obj)
        com_range_obj.Collapse(Direction=0)  # wdCollapseEnd


@handle_com_error(ErrorCode.TEXT_FORMATTING_ERROR, "set alignment")
def set_alignment_for_range(
    document: win32com.client.CDispatch, com_range_obj: win32com.client.CDispatch, alignment: str
):
    """
    Set paragraph alignment for a range.

    Args:
        document: The Word document COM object.
        com_range_obj: The range to align.
        alignment: One of "left", "center", "right", "justify".
    """
    alignment_map = {
        "left": 0,  # wdAlignParagraphLeft
        "center": 1,  # wdAlignParagraphCenter
        "right": 2,  # wdAlignParagraphRight
        "justify": 3,  # wdAlignParagraphJustify
    }
    if alignment.lower() in alignment_map:
        com_range_obj.ParagraphFormat.Alignment = alignment_map[alignment.lower()]
    else:
        raise WordDocumentError(
            ErrorCode.TEXT_FORMATTING_ERROR,
            f"Invalid alignment: {alignment}. Use 'left', 'center', 'right', or 'justify'."
        )


@handle_com_error(ErrorCode.TEXT_FORMATTING_ERROR, "set underline")
def set_underline_for_range(
    com_range_obj: win32com.client.CDispatch, is_underline: bool
):
    """
    Set underline formatting for a range.

    Args:
        com_range_obj: COM Range object to format.
        is_underline: Whether to set underline formatting.
    """
    com_range_obj.Font.Underline = 1 if is_underline else 0  # wdUnderlineSingle or wdUnderlineNone


@handle_com_error(ErrorCode.TEXT_FORMATTING_ERROR, "add heading")
def add_heading(
    document: win32com.client.CDispatch,
    com_range_obj: win32com.client.CDispatch,
    text: str,
    level: int = 1,
):
    """
    Add a heading at the given range.

    Args:
        document: The Word document COM object.
        com_range_obj: The range where to insert the heading.
        text: The heading text.
        level: The heading level (1-9).
    """
    # Insert the text
    com_range_obj.Text = text
    # Apply heading style
    style_name = f"Heading {level}" if 1 <= level <= 9 else "Heading 1"
    com_range_obj.Style = document.Styles(style_name)
    # Insert paragraph after for spacing
    com_range_obj.InsertParagraphAfter()


def get_runs_in_range(
    range_obj: win32com.client.CDispatch
):
    """
    Get runs in a given range.

    Args:
        range_obj: COM Range object to get runs from.
    """
    return range_obj.Runs

