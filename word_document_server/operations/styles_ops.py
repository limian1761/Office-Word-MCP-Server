"""
Styles operations for Word Document MCP Server.
This module contains functions for style-related operations.
"""

import json
import logging
from typing import Any, Dict, List, Optional, Union

import win32com.client

from ..com_backend.com_utils import handle_com_error
from ..selector.selector import SelectorEngine
from ..utils.core_utils import ErrorCode, WordDocumentError, log_error, log_info
from . import text_format_ops

logger = logging.getLogger(__name__)


def set_paragraph_alignment(
    document: win32com.client.CDispatch,
    alignment: str,
    locator: Optional[Dict[str, Any]] = None,
) -> str:
    """设置段落对齐方式

    Args:
        document: Word文档COM对象
        alignment: 对齐方式 (left, center, right, justify)
        locator: 定位器对象，用于指定要设置对齐方式的元素

    Returns:
        设置对齐方式成功的消息
    """
    if not document:
        raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")

    selector = SelectorEngine()

    # 获取要设置对齐方式的范围
    aligned_count = 0

    if locator:
        # 使用定位器找到要设置对齐方式的元素
        selection = selector.select(document, locator)

        if (
            not selection
            or not hasattr(selection, "_com_ranges")
            or not selection._com_ranges
        ):
            raise WordDocumentError(
                ErrorCode.OBJECT_NOT_FOUND, "No object found matching the locator"
            )

        # 对每个元素设置对齐方式
        # Selection._com_ranges中只包含Range对象
        for range_obj in selection._com_ranges:
            try:
                text_format_ops.set_alignment_for_range(document, range_obj, alignment)
                aligned_count += 1
            except Exception as e:
                log_error(f"Failed to apply alignment to object: {str(e)}")
    else:
        # 如果没有定位器，使用当前选区
        try:
            range_obj = document.Application.Selection.Range
            text_format_ops.set_alignment_for_range(document, range_obj, alignment)
            aligned_count = 1
        except Exception as e:
            raise WordDocumentError(
                ErrorCode.FORMATTING_ERROR,
                f"Failed to apply alignment to selection: {str(e)}",
            )

    log_info(
        f"Successfully applied alignment '{alignment}' to {aligned_count} paragraph(s)"
    )
    return json.dumps(
        {
            "success": True,
            "message": f"Successfully applied alignment '{alignment}'",
            "alignment": alignment,
            "paragraph_count": aligned_count,
        },
        ensure_ascii=False,
    )


@handle_com_error(ErrorCode.FORMATTING_ERROR, "apply formatting")
def apply_formatting(
    document: win32com.client.CDispatch,
    formatting: Dict[str, Any],
    locator: Optional[Dict[str, Any]] = None,
) -> str:
    """应用文本格式化

    Args:
        document: Word文档COM对象
        formatting: 格式化参数字典
        locator: 定位器对象，用于指定要格式化的元素

    Returns:
        格式化成功的消息

    Raises:
        ValueError: 当参数无效时抛出
        WordDocumentError: 当应用格式化失败时抛出
    """
    if not document:
        raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")

    selector = SelectorEngine()

    # 验证格式化参数
    if not formatting or not isinstance(formatting, dict):
        raise ValueError("Formatting parameter must be a non-empty dictionary")

    # 获取要格式化的范围
    range_obj = None

    if locator:
        # 使用定位器获取范围
        try:
            selection = selector.select(document, locator)
            if hasattr(selection, "_com_ranges") and selection._com_ranges:
                # Selection._com_ranges中只包含Range对象
                range_obj = selection._com_ranges[0]
            else:
                raise WordDocumentError(
                    ErrorCode.OBJECT_NOT_FOUND, "No object found matching the locator"
                )
        except Exception as e:
            raise WordDocumentError(
                ErrorCode.FORMATTING_ERROR,
                f"Failed to locate object for formatting: {str(e)}",
            )
    else:
        # 如果没有提供定位器，格式化整个文档
        range_obj = document.Range()

    try:
        # 应用格式化选项
        if "bold" in formatting:
            text_format_ops.set_bold_for_range(range_obj, formatting["bold"])

        if "italic" in formatting:
            text_format_ops.set_italic_for_range(range_obj, formatting["italic"])

        if "font_size" in formatting:
            text_format_ops.set_font_size_for_range(range_obj, formatting["font_size"])

        if "font_name" in formatting:
            text_format_ops.set_font_name_for_range(range_obj, formatting["font_name"])

        if "font_color" in formatting:
            text_format_ops.set_font_color_for_range(
                document, range_obj, formatting["font_color"]
            )

        if "alignment" in formatting:
            text_format_ops.set_alignment_for_range(
                document, range_obj, formatting["alignment"]
            )

        if "paragraph_style" in formatting:
            # 对于段落样式，我们需要对整个段落应用样式
            try:
                range_obj.Paragraphs(1).Style = formatting["paragraph_style"]
            except Exception:
                # 如果直接设置失败，尝试在文档样式中查找
                style_found = False
                for i in range(1, document.Styles.Count + 1):
                    if (
                        document.Styles(i).NameLocal.lower()
                        == formatting["paragraph_style"].lower()
                    ):
                        range_obj.Paragraphs(1).Style = document.Styles(i)
                        style_found = True
                        break

                if not style_found:
                    raise WordDocumentError(
                        ErrorCode.FORMATTING_ERROR,
                        f"Style '{formatting['paragraph_style']}' not found in document",
                    )

        # 添加成功日志
        log_info("Successfully applied formatting to document")

        return json.dumps(
            {"success": True, "message": "Formatting applied successfully"},
            ensure_ascii=False,
        )

    except Exception as e:
        log_error(f"Failed to apply formatting: {str(e)}", exc_info=True)
        raise WordDocumentError(
            ErrorCode.FORMATTING_ERROR, f"Failed to apply formatting: {str(e)}"
        )


@handle_com_error(ErrorCode.FORMATTING_ERROR, "set font")
def set_font(
    document: win32com.client.CDispatch,
    font_name: str,
    font_size: Optional[float] = None,
    bold: Optional[bool] = None,
    italic: Optional[bool] = None,
    underline: Optional[str] = None,
    color: Optional[str] = None,
    locator: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    """设置文本字体属性

    Args:
        document: Word文档COM对象
        font_name: 字体名称
        font_size: 字体大小
        bold: 是否粗体
        italic: 是否斜体
        underline: 下划线类型
        color: 字体颜色
        locator: 定位器对象，用于指定要设置字体的元素

    Returns:
        包含操作结果的字典

    Raises:
        ValueError: 当参数无效时抛出
        WordDocumentError: 当设置字体失败时抛出
    """
    if not document:
        raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")

    selector = SelectorEngine()

    if not font_name:
        raise ValueError("Font name parameter must be provided")

    # 验证字体是否存在
    font_exists = False
    available_fonts = list(document.Application.FontNames)
    for font in available_fonts:
        if font == font_name:
            font_exists = True
            break

    if not font_exists:
        # 准备可用字体列表
        if len(available_fonts) <= 10:
            fonts_list = ", ".join(available_fonts)
        else:
            fonts_list = ", ".join(available_fonts[:10]) + f", and {len(available_fonts)-10} more fonts"
        raise WordDocumentError(
            ErrorCode.FORMATTING_ERROR,
            f"Font '{font_name}' not found. Available fonts: {fonts_list}"
        )

    range_obj = None
    object_count = 0

    if locator:
        selection = selector.select(document, locator)
        if not selection._com_ranges:
            raise WordDocumentError(
                ErrorCode.OBJECT_NOT_FOUND, "No object found matching the locator"
            )

        # Selection._com_ranges中只包含Range对象
        for range_obj in selection._com_ranges:
            text_format_ops.set_font_name_for_range(range_obj, font_name)
            if font_size is not None:
                text_format_ops.set_font_size_for_range(range_obj, font_size)
            if bold is not None:
                text_format_ops.set_bold_for_range(range_obj, bold)
            if italic is not None:
                text_format_ops.set_italic_for_range(range_obj, italic)
            if color is not None:
                text_format_ops.set_font_color_for_range(document, range_obj, color)
            # Underline is not yet in text_format_ops, so we handle it here for now.
            if underline is not None:
                font = range_obj.Font
                underline_map = {
                    "none": 0,
                    "single": 1,
                    "double": 2,
                    "dotted": 4,
                    "dashed": 5,
                    "wave": 16,
                }
                font.Underline = underline_map.get(underline, 0)
        object_count = len(selection._com_ranges)
    else:
        try:
            range_obj = document.Application.Selection.Range
        except Exception:
            range_obj = document.Content
            range_obj.Collapse(0)

        text_format_ops.set_font_name_for_range(range_obj, font_name)
        if font_size is not None:
            text_format_ops.set_font_size_for_range(range_obj, font_size)
        if bold is not None:
            text_format_ops.set_bold_for_range(range_obj, bold)
        if italic is not None:
            text_format_ops.set_italic_for_range(range_obj, italic)
        if color is not None:
            text_format_ops.set_font_color_for_range(document, range_obj, color)
        # Underline is not yet in text_format_ops, so we handle it here for now.
        if underline is not None:
            font = range_obj.Font
            underline_map = {
                "none": 0,
                "single": 1,
                "double": 2,
                "dotted": 4,
                "dashed": 5,
                "wave": 16,
            }
            font.Underline = underline_map.get(underline, 0)
        object_count = 1

    log_info(f"Successfully set font properties for {object_count} object(s)")
    return {
        "success": True,
        "message": f"Successfully set font properties for {object_count} object(s)",
        "font_name": font_name,
        "object_count": object_count,
    }


@handle_com_error(ErrorCode.SERVER_ERROR, "set paragraph style")
def set_paragraph_style(
    document: win32com.client.CDispatch,
    style_name: str,
    locator: Optional[Dict[str, Any]] = None,
) -> str:
    """设置段落样式

    Args:
        document: Word文档COM对象
        style_name: 段落样式名称
        locator: 定位器对象，用于指定要设置样式的元素

    Returns:
        设置样式成功的消息

    Raises:
        ValueError: 当参数无效时抛出
        WordDocumentError: 当设置样式失败时抛出
    """
    if not document:
        raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")

    selector = SelectorEngine()

    # 验证样式名称参数
    if not style_name:
        raise ValueError("Style name parameter must be provided")

    # 检查样式是否存在（使用Name属性而非本地化名称以确保兼容性）
    style_exists = False
    paragraph_styles = []
    for style in document.Styles:
        if style.Type == 1:  # wdStyleTypeParagraph = 1
            paragraph_styles.append(style.Name)
            if style.Name == style_name:
                style_exists = True
                break

    if not style_exists:
        # 准备可用段落样式列表
        if len(paragraph_styles) <= 10:
            styles_list = ", ".join(paragraph_styles)
        else:
            styles_list = ", ".join(paragraph_styles[:10]) + f", and {len(paragraph_styles)-10} more styles"
        
        raise WordDocumentError(
            ErrorCode.SERVER_ERROR,
            f"Style '{style_name}' not found. Available paragraph styles: {styles_list}",
        )

    # 获取要设置样式的范围
    styled_count = 0

    if locator:
        # 使用定位器找到要设置样式的元素
        selection = selector.select(document, locator)

        if (
            not selection
            or not hasattr(selection, "_com_ranges")
            or not selection._com_ranges
        ):
            raise WordDocumentError(
                ErrorCode.OBJECT_NOT_FOUND, "No object found matching the locator"
            )

        # 对每个元素设置样式
        # Selection._com_ranges中只包含Range对象
        for range_obj in selection._com_ranges:
            try:
                range_obj.Paragraphs(1).Style = style_name
                styled_count += 1
            except Exception as e:
                log_error(f"Failed to apply style to object: {str(e)}")
    else:
        # 如果没有定位器，使用当前选区
        try:
            range_obj = document.Application.Selection.Range
            range_obj.Paragraphs(1).Style = style_name
            styled_count = 1
        except Exception as e:
            raise WordDocumentError(
                ErrorCode.FORMATTING_ERROR,
                f"Failed to apply style to selection: {str(e)}",
            )

    log_info(
        f"Successfully applied style '{style_name}' to {styled_count} paragraph(s)"
    )
    return json.dumps(
        {
            "success": True,
            "message": f"Successfully applied style '{style_name}'",
            "style_name": style_name,
            "paragraph_count": styled_count,
        },
        ensure_ascii=False,
    )