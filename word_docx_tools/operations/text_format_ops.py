"""
Text formatting operations for Word Document MCP Server.
This module contains functions for text formatting operations.
"""

import logging
from typing import Any, Dict, List, Optional, Union

import win32com.client

from ..com_backend.com_utils import handle_com_error
from ..selector.selector import SelectorEngine
from ..mcp_service.core_utils import (ErrorCode,
                                                          ObjectNotFoundError,
                                                          WordDocumentError,
                                                          log_error, log_info)

logger = logging.getLogger(__name__)


@handle_com_error(ErrorCode.FORMATTING_ERROR, "set bold")
def set_bold_for_range(range_obj: Any, is_bold: bool) -> bool:
    """设置文本范围的粗体格式

    Args:
        range_obj: Word文本范围对象
        is_bold: 是否设置为粗体

    Returns:
        操作是否成功
    """
    try:
        if hasattr(range_obj, "Font"):
            range_obj.Font.Bold = is_bold
            return True
        return False
    except Exception as e:
        log_error(f"Failed to set bold for range: {e}")
        return False


@handle_com_error(ErrorCode.FORMATTING_ERROR, "set italic")
def set_italic_for_range(range_obj: Any, is_italic: bool) -> bool:
    """设置文本范围的斜体格式

    Args:
        range_obj: Word文本范围对象
        is_italic: 是否设置为斜体

    Returns:
        操作是否成功
    """
    try:
        if hasattr(range_obj, "Font"):
            range_obj.Font.Italic = is_italic
            return True
        return False
    except Exception as e:
        log_error(f"Failed to set italic for range: {e}")
        return False


@handle_com_error(ErrorCode.FORMATTING_ERROR, "set font size")
def set_font_size_for_range(range_obj: Any, font_size: float) -> bool:
    """设置文本范围的字体大小

    Args:
        range_obj: Word文本范围对象
        font_size: 字体大小

    Returns:
        操作是否成功
    """
    try:
        if hasattr(range_obj, "Font"):
            range_obj.Font.Size = font_size
            return True
        return False
    except Exception as e:
        log_error(f"Failed to set font size for range: {e}")
        return False


@handle_com_error(ErrorCode.FORMATTING_ERROR, "set font name")
def set_font_name_for_range(range_obj: Any, font_name: str) -> bool:
    """设置文本范围的字体名称

    Args:
        range_obj: Word文本范围对象
        font_name: 字体名称

    Returns:
        操作是否成功
    """
    try:
        if hasattr(range_obj, "Font"):
            range_obj.Font.Name = font_name
            return True
        return False
    except Exception as e:
        log_error(f"Failed to set font name for range: {e}")
        return False


@handle_com_error(ErrorCode.FORMATTING_ERROR, "set font color")
def set_font_color_for_range(document: Any, range_obj: Any, color: str) -> bool:
    """设置文本范围的字体颜色

    Args:
        document: Word文档COM对象
        range_obj: Word文本范围对象
        color: 颜色值

    Returns:
        操作是否成功
    """
    try:
        if hasattr(range_obj, "Font"):
            # 尝试将颜色字符串转换为RGB值
            # 这里简化处理，实际项目中可能需要更复杂的颜色解析
            if color.lower() == "red":
                range_obj.Font.ColorIndex = 6  # wdRed = 6
            elif color.lower() == "blue":
                range_obj.Font.ColorIndex = 5  # wdBlue = 5
            elif color.lower() == "green":
                range_obj.Font.ColorIndex = 4  # wdGreen = 4
            else:
                # 默认使用黑色
                range_obj.Font.ColorIndex = 1  # wdBlack = 1
            return True
        return False
    except Exception as e:
        log_error(f"Failed to set font color for range: {e}")
        return False


@handle_com_error(ErrorCode.FORMATTING_ERROR, "set alignment")
def set_alignment_for_range(document: Any, range_obj: Any, alignment: str) -> bool:
    """设置文本范围的对齐方式

    Args:
        document: Word文档COM对象
        range_obj: Word文本范围对象
        alignment: 对齐方式 (left, center, right, justify)

    Returns:
        操作是否成功
    """
    try:
        if hasattr(range_obj, "ParagraphFormat"):
            # Word的对齐常量
            wdAlignParagraphLeft = 0
            wdAlignParagraphCenter = 1
            wdAlignParagraphRight = 2
            wdAlignParagraphJustify = 3

            if alignment.lower() == "center":
                range_obj.ParagraphFormat.Alignment = wdAlignParagraphCenter
            elif alignment.lower() == "right":
                range_obj.ParagraphFormat.Alignment = wdAlignParagraphRight
            elif alignment.lower() == "justify":
                range_obj.ParagraphFormat.Alignment = wdAlignParagraphJustify
            else:
                # 默认左对齐
                range_obj.ParagraphFormat.Alignment = wdAlignParagraphLeft
            return True
        return False
    except Exception as e:
        log_error(f"Failed to set alignment for range: {e}")
        return False


@handle_com_error(ErrorCode.FORMATTING_ERROR, "set paragraph style")
def set_paragraph_style(object: Any, style_name: str) -> bool:
    """设置段落样式

    Args:
        object: 段落元素
        style_name: 样式名称

    Returns:
        操作是否成功
    """
    try:
        if hasattr(object, "Style"):
            # 尝试直接设置样式
            try:
                object.Style = style_name
                return True
            except Exception:
                # 如果失败，尝试在文档中查找样式
                if hasattr(object, "Document") and hasattr(object.Document, "Styles"):
                    styles = object.Document.Styles
                    for i in range(1, styles.Count + 1):
                        try:
                            if styles(i).NameLocal.lower() == style_name.lower():
                                object.Style = styles(i)
                                return True
                        except Exception:
                            continue
        return False
    except Exception as e:
        log_error(f"Failed to set paragraph style: {e}")
        return False


@handle_com_error(ErrorCode.FORMATTING_ERROR, "create bulleted list")
def create_bulleted_list_relative_to(
    document: win32com.client.CDispatch,
    anchor_range: Any,
    items: List[str],
    position: str = "after",
) -> bool:
    """在指定范围附近创建项目符号列表

    Args:
        document: Word文档COM对象
        anchor_range: 锚点范围对象
        items: 列表项内容
        position: 插入位置 (before, after)

    Returns:
        操作是否成功
    """
    try:
        if not document:
            raise RuntimeError("No document open.")

        if not items:
            raise ValueError("Items list cannot be empty.")

        # 创建一个新的范围用于插入列表
        insertion_range = anchor_range.Duplicate

        if position == "before":
            # 在锚点前插入
            insertion_range.Collapse(1)  # wdCollapseStart = 1
        else:
            # 在锚点后插入
            insertion_range.Collapse(0)  # wdCollapseEnd = 0

        # 如果不是在文档开头插入，先添加一个段落标记
        if position == "after" and insertion_range.Start > 0:
            insertion_range.InsertAfter("\r")
            insertion_range.Collapse(0)

        # 插入列表项
        for i, item in enumerate(items):
            # 插入列表项文本
            insertion_range.InsertAfter(item)

            # 如果不是最后一项，添加段落标记
            if i < len(items) - 1:
                insertion_range.Collapse(False)  # wdCollapseEnd
                insertion_range.InsertAfter("\r")
                insertion_range.Collapse(False)  # wdCollapseEnd

        # 为新插入的文本应用项目符号列表格式
        # 获取刚刚插入的文本范围
        list_start = insertion_range.Start - sum(
            len(item) + 2 for item in items
        )  # 2 for \r
        list_end = insertion_range.Start
        list_range = document.Range(list_start, list_end)

        # 应用项目符号列表
        list_range.ParagraphFormat.Bullet.Enabled = True

        return True
    except Exception as e:
        log_error(f"Failed to create bulleted list: {e}")
        raise WordDocumentError(
            ErrorCode.FORMATTING_ERROR, f"Failed to create bulleted list: {str(e)}"
        )


@handle_com_error(ErrorCode.FORMATTING_ERROR, "create bulleted list")
def create_bulleted_list(
    document: win32com.client.CDispatch,
    locator: Dict[str, Any],
    items: List[str],
    position: str = "after",
) -> bool:
    """在指定元素附近创建项目符号列表

    Args:
        document: Word文档COM对象
        locator: 定位器对象
        items: 列表项内容
        position: 插入位置 (before, after, replace)

    Returns:
        操作是否成功
    """
    from .selector.selector import SelectorEngine

    try:
        if not document:
            raise RuntimeError("No document open.")

        if not items:
            raise ValueError("Items list cannot be empty.")

        # 使用选择器引擎选择元素
        selector = SelectorEngine()
        selection = selector.select(document, locator, expect_single=True)

        # Selection._com_ranges中只包含Range对象
        for object in selection._com_ranges:
            if position == "replace":
                # 删除元素首先
                object.Range.Delete()
                # 使用元素的范围作为插入点
                insertion_range = object.Range
            elif position == "before":
                # 折叠范围到开始
                insertion_range = object.Range.Duplicate
                insertion_range.Collapse(True)  # wdCollapseStart
            else:  # position == "after"
                # 折叠范围到结束
                insertion_range = object.Range.Duplicate
                insertion_range.Collapse(False)  # wdCollapseEnd

            # 在插入点创建项目符号列表
            create_bulleted_list_relative_to(document, insertion_range, items, "after")

        return True
    except Exception as e:
        log_error(f"Failed to create bulleted list: {e}")
        raise WordDocumentError(
            ErrorCode.FORMATTING_ERROR, f"Failed to create bulleted list: {str(e)}"
        )
