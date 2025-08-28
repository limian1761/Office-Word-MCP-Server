"""
Element operations for Word Document MCP Server.
This module contains operations that work on single document elements.
"""

from typing import Any, Dict, List, Optional
import logging

import win32com.client

from word_document_server.errors import ErrorCode, WordDocumentError


def insert_text_before_element(element, text: str, style: Optional[str], document) -> bool:
    """元操作：在单个元素前插入文本

    Args:
        element: 单个文档元素对象
        text: 插入文本内容
        style: 可选样式名称
        document: 文档对象

    Returns:
        bool: 操作成功状态
    """
    try:
        anchor_range = element.Range
        new_range = anchor_range.Duplicate
        new_range.Collapse(1)  # wdCollapseStart = 1
        new_range.InsertAfter(text + "\r")
        
        # 应用样式
        if style:
            paragraph = new_range.Paragraphs(1)
            paragraph.Style = style
        
        return True
    except Exception as e:
        logging.error(f"在元素前插入文本失败: {str(e)}")
        return False


def insert_text_after_element(element, text: str, style: Optional[str], document) -> bool:
    """元操作：在单个元素后插入文本

    Args:
        element: 单个文档元素对象
        text: 插入文本内容
        style: 可选样式名称
        document: 文档对象

    Returns:
        bool: 操作成功状态
    """
    try:
        anchor_range = element.Range
        new_range = anchor_range.Duplicate
        new_range.Collapse(0)  # wdCollapseEnd = 0
        new_range.InsertAfter("\r" + text)
        
        # 应用样式
        if style:
            paragraph = new_range.Paragraphs(1)
            paragraph.Style = style
        
        return True
    except Exception as e:
        logging.error(f"在元素后插入文本失败: {str(e)}")
        return False


def set_picture_element_color_type(element, color_code: int) -> bool:
    """元操作：设置单个图片元素的颜色类型

    Args:
        element: 单个图片元素对象
        color_code: 颜色类型代码（0-3）

    Returns:
        bool: 操作成功状态
    """
    try:
        if hasattr(element, "Type") and (element.Type == 1 or element.Type == 2):
            if hasattr(element, "PictureFormat") and hasattr(element.PictureFormat, "ColorType"):
                element.PictureFormat.ColorType = color_code
                return True
        return False
    except Exception as e:
        logging.error(f"设置图片颜色类型失败: {str(e)}")
        return False


def add_table(
    document: win32com.client.CDispatch, com_range_obj: win32com.client.CDispatch, rows: int, cols: int
):
    """
    Adds a table after a given range.

    Args:
        document: The Word document COM object.
        com_range_obj: The range to insert the table after.
        rows: Number of rows for the table.
        cols: Number of columns for the table.
    """
    try:
        com_range_obj.Tables.Add(com_range_obj, rows, cols)
    except Exception as e:
        raise WordDocumentError(ErrorCode.TABLE_ERROR, f"Failed to add table: {e}")
sm

def get_element_text(element: win32com.client.CDispatch) -> str:
    """
    Gets the text content of a single element.

    Args:
        element: The COM object representing the element.

    Returns:
        The text content of the element.
    """
    element_text = ""
    if hasattr(element, "Text"):
        element_text = element.Text()
    elif hasattr(element, "Range") and hasattr(element.Range, "Text"):
        element_text = element.Range.Text
    return element_text
