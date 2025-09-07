"""
Styles operations for Word Document MCP Server.
This module contains functions for style-related operations.
"""

import json
import logging
from typing import Any, Dict, List, Optional, Union

import win32com.client

from ..com_backend.com_utils import handle_com_error
from ..mcp_service.core_utils import (ErrorCode, WordDocumentError, log_error,
                                      log_info)
from ..selector.selector import SelectorEngine
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

    # Check if document has Application property
    if not hasattr(document, "Application") or document.Application is None:
        raise WordDocumentError(
            ErrorCode.DOCUMENT_ERROR, "Document Application object not available"
        )

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
    ranges_to_format = []

    if locator:
        # 使用定位器获取范围
        try:
            selection = selector.select(document, locator)
            if hasattr(selection, "_com_ranges") and selection._com_ranges:
                # 获取所有匹配的Range对象
                ranges_to_format = selection._com_ranges
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
        ranges_to_format = [document.Range()]

    try:
        # 应用格式化选项到所有匹配的范围
        formatted_count = 0
        for range_obj in ranges_to_format:
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
            formatted_count += 1

        # 添加成功日志
        log_info(f"Successfully applied formatting to {formatted_count} object(s)")

        return json.dumps(
            {"success": True, "message": "Formatting applied successfully", "formatted_count": formatted_count},
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
            fonts_list = (
                ", ".join(available_fonts[:10])
                + f", and {len(available_fonts)-10} more fonts"
            )
        raise WordDocumentError(
            ErrorCode.FORMATTING_ERROR,
            f"Font '{font_name}' not found. Available fonts: {fonts_list}",
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
                # Check if range_obj has Font property
                if not hasattr(range_obj, "Font"):
                    log_error("Range object does not have Font property")
                    continue
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
            range_obj.Collapse(False)  # wdCollapseEnd

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

    # Check if document has Styles property
    if not hasattr(document, "Styles") or document.Styles is None:
        raise WordDocumentError(
            ErrorCode.DOCUMENT_ERROR, "Document Styles collection not available"
        )

    selector = SelectorEngine()

    # 验证样式名称参数
    if not style_name:
        raise ValueError("Style name parameter must be provided")

    # 检查样式是否存在（使用NameLocal属性并添加异常处理以提高兼容性）
    style_exists = False
    paragraph_styles = []
    target_style = None
    for style in document.Styles:
        try:
            if style.Type == 1:  # wdStyleTypeParagraph = 1
                # 优先使用NameLocal属性，这在不同语言环境下更可靠
                style_name_local = style.NameLocal
                paragraph_styles.append(style_name_local)
                # 同时检查Name和NameLocal，以增加兼容性
                if style_name_local == style_name or (
                    hasattr(style, "Name") and style.Name == style_name
                ):
                    style_exists = True
                    target_style = style
                    break
        except Exception as e:
            log_error(f"Error accessing style property: {str(e)}")
            continue

    if not style_exists:
        # 准备可用段落样式列表
        if len(paragraph_styles) <= 10:
            styles_list = ", ".join(paragraph_styles)
        else:
            styles_list = (
                ", ".join(paragraph_styles[:10])
                + f", and {len(paragraph_styles)-10} more styles"
            )

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
                # 首先尝试使用样式对象
                if target_style:
                    range_obj.Paragraphs(1).Style = target_style
                    styled_count += 1
                else:
                    # 如果没有找到样式对象，尝试直接使用样式名称
                    range_obj.Paragraphs(1).Style = style_name
                    styled_count += 1
            except Exception as e:
                log_error(f"Failed to apply style to object: {str(e)}")
    else:
        # 如果没有定位器，使用当前选区
        try:
            range_obj = document.Application.Selection.Range
            # 首先尝试使用样式对象
            if target_style:
                range_obj.Paragraphs(1).Style = target_style
                styled_count = 1
            else:
                # 如果没有找到样式对象，尝试直接使用样式名称
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


@handle_com_error(ErrorCode.FORMATTING_ERROR, "set paragraph formatting")
def set_paragraph_formatting(
    document: win32com.client.CDispatch,
    alignment: Optional[str] = None,
    line_spacing: Optional[float] = None,
    line_spacing_type: Optional[str] = None,  # 'multiple' 或 'exact'
    space_before: Optional[float] = None,
    space_after: Optional[float] = None,
    first_line_indent: Optional[float] = None,
    left_indent: Optional[float] = None,
    right_indent: Optional[float] = None,
    locator: Optional[Dict[str, Any]] = None,
) -> str:
    """设置段落格式

    Args:
        document: Word文档COM对象
        alignment: 对齐方式 (left, center, right, justify)
        line_spacing: 行距值
        line_spacing_type: 行距类型 ('multiple' 表示倍数，'exact' 表示精确磅值，默认为'multiple')
        space_before: 段前间距
        space_after: 段后间距
        first_line_indent: 首行缩进
        left_indent: 左缩进
        right_indent: 右缩进
        locator: 定位器对象，用于指定要设置格式的元素

    Returns:
        设置格式成功的消息

    Raises:
        ValueError: 当参数无效时抛出
        WordDocumentError: 当设置格式失败时抛出
    """
    if not document:
        raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")

    # Check if document has Application property
    if not hasattr(document, "Application") or document.Application is None:
        raise WordDocumentError(
            ErrorCode.DOCUMENT_ERROR, "Document Application object not available"
        )

    selector = SelectorEngine()
    formatting_count = 0
    successfully_applied = {}

    # 获取要设置格式的范围
    if locator:
        # 使用定位器找到要设置格式的元素
        selection = selector.select(document, locator)

        if (
            not selection
            or not hasattr(selection, "_com_ranges")
            or not selection._com_ranges
        ):
            raise WordDocumentError(
                ErrorCode.OBJECT_NOT_FOUND, "No object found matching the locator"
            )

        # 对每个元素设置格式
        # Selection._com_ranges中只包含Range对象
        for range_obj in selection._com_ranges:
            try:
                # 获取段落对象
                paragraphs = range_obj.Paragraphs
                if not paragraphs:
                    continue

                for para in paragraphs:
                    # 为每个段落创建一个记录字典
                    para_applied = {}

                    # 设置对齐方式
                    if alignment:
                        try:
                            text_format_ops.set_alignment_for_range(
                                document, range_obj, alignment
                            )
                            para_applied["alignment"] = alignment
                        except Exception as e:
                            log_error(f"Failed to set alignment: {str(e)}")

                    # 设置行距
                    if line_spacing is not None:
                        try:
                            if hasattr(para, "LineSpacingRule") and hasattr(
                                para, "LineSpacing"
                            ):
                                # 根据line_spacing_type决定如何设置行距
                                if line_spacing_type == "exact":
                                    # 设置为精确磅值
                                    para.LineSpacingRule = 4  # wdLineSpaceExactly = 6
                                else:
                                    # 默认设置为倍数
                                    para.LineSpacingRule = 5  # wdLineSpaceMultiple = 4

                                para.LineSpacing = line_spacing
                                para_applied["line_spacing"] = line_spacing
                                if line_spacing_type:
                                    para_applied["line_spacing_type"] = (
                                        line_spacing_type
                                    )
                        except Exception as e:
                            log_error(f"Failed to set line spacing: {str(e)}")

                    # 设置段前间距
                    if space_before is not None:
                        try:
                            if hasattr(para, "SpaceBefore"):
                                para.SpaceBefore = space_before
                                para_applied["space_before"] = space_before
                        except Exception as e:
                            log_error(f"Failed to set space before: {str(e)}")

                    # 设置段后间距
                    if space_after is not None:
                        try:
                            if hasattr(para, "SpaceAfter"):
                                para.SpaceAfter = space_after
                                para_applied["space_after"] = space_after
                        except Exception as e:
                            log_error(f"Failed to set space after: {str(e)}")

                    # 设置首行缩进
                    if first_line_indent is not None:
                        try:
                            if hasattr(para, "FirstLineIndent"):
                                para.FirstLineIndent = first_line_indent
                                para_applied["first_line_indent"] = first_line_indent
                        except Exception as e:
                            log_error(f"Failed to set first line indent: {str(e)}")

                    # 设置左缩进
                    if left_indent is not None:
                        try:
                            if hasattr(para, "LeftIndent"):
                                para.LeftIndent = left_indent
                                para_applied["left_indent"] = left_indent
                        except Exception as e:
                            log_error(f"Failed to set left indent: {str(e)}")

                    # 设置右缩进
                    if right_indent is not None:
                        try:
                            if hasattr(para, "RightIndent"):
                                para.RightIndent = right_indent
                                para_applied["right_indent"] = right_indent
                        except Exception as e:
                            log_error(f"Failed to set right indent: {str(e)}")

                    # 如果这个段落有成功应用的设置，增加计数
                    if para_applied:
                        formatting_count += 1
                        # 合并成功应用的设置
                        for key, value in para_applied.items():
                            if (
                                key not in successfully_applied
                                or successfully_applied[key] < value
                            ):
                                successfully_applied[key] = value
            except Exception as e:
                log_error(f"Failed to apply formatting to object: {str(e)}")
    else:
        # 如果没有定位器，使用当前选区
        try:
            range_obj = document.Application.Selection.Range
            paragraphs = range_obj.Paragraphs
            if not paragraphs:
                raise WordDocumentError(
                    ErrorCode.FORMATTING_ERROR,
                    "No paragraphs found in selection",
                )

            for para in paragraphs:
                # 为每个段落创建一个记录字典
                para_applied = {}

                # 设置对齐方式
                if alignment:
                    try:
                        text_format_ops.set_alignment_for_range(
                            document, range_obj, alignment
                        )
                        para_applied["alignment"] = alignment
                    except Exception as e:
                        log_error(f"Failed to set alignment: {str(e)}")

                # 设置行距
                if line_spacing is not None:
                    try:
                        if hasattr(para, "LineSpacingRule") and hasattr(
                            para, "LineSpacing"
                        ):
                            # 根据line_spacing_type决定如何设置行距
                            if line_spacing_type == "exact":
                                # 设置为精确磅值
                                para.LineSpacingRule = 4  # wdLineSpaceExactly = 6
                            else:
                                # 默认设置为倍数
                                para.LineSpacingRule = 5  # wdLineSpaceMultiple = 4

                            para.LineSpacing = line_spacing
                            para_applied["line_spacing"] = line_spacing
                            if line_spacing_type:
                                para_applied["line_spacing_type"] = line_spacing_type
                    except Exception as e:
                        log_error(f"Failed to set line spacing: {str(e)}")

                # 设置段前间距
                if space_before is not None:
                    try:
                        if hasattr(para, "SpaceBefore"):
                            para.SpaceBefore = space_before
                            para_applied["space_before"] = space_before
                    except Exception as e:
                        log_error(f"Failed to set space before: {str(e)}")

                # 设置段后间距
                if space_after is not None:
                    try:
                        if hasattr(para, "SpaceAfter"):
                            para.SpaceAfter = space_after
                            para_applied["space_after"] = space_after
                    except Exception as e:
                        log_error(f"Failed to set space after: {str(e)}")

                # 设置首行缩进
                if first_line_indent is not None:
                    try:
                        if hasattr(para, "FirstLineIndent"):
                            para.FirstLineIndent = first_line_indent
                            para_applied["first_line_indent"] = first_line_indent
                    except Exception as e:
                        log_error(f"Failed to set first line indent: {str(e)}")

                # 设置左缩进
                if left_indent is not None:
                    try:
                        if hasattr(para, "LeftIndent"):
                            para.LeftIndent = left_indent
                            para_applied["left_indent"] = left_indent
                    except Exception as e:
                        log_error(f"Failed to set left indent: {str(e)}")

                # 设置右缩进
                if right_indent is not None:
                    try:
                        if hasattr(para, "RightIndent"):
                            para.RightIndent = right_indent
                            para_applied["right_indent"] = right_indent
                    except Exception as e:
                        log_error(f"Failed to set right indent: {str(e)}")

                # 如果这个段落有成功应用的设置，增加计数
                if para_applied:
                    formatting_count += 1
                    # 合并成功应用的设置
                    for key, value in para_applied.items():
                        if (
                            key not in successfully_applied
                            or successfully_applied[key] < value
                        ):
                            successfully_applied[key] = value
        except Exception as e:
            raise WordDocumentError(
                ErrorCode.FORMATTING_ERROR,
                f"Failed to apply formatting to selection: {str(e)}",
            )

    # 构建应用设置的字符串表示
    applied_settings = []
    for key, value in successfully_applied.items():
        if isinstance(value, str):
            applied_settings.append(f"{key}='{value}'")
        else:
            applied_settings.append(f"{key}={value}")

    settings_str = ", ".join(applied_settings)
    log_info(
        f"Successfully applied paragraph formatting ({settings_str}) to {formatting_count} object(s)"
    )

    return json.dumps(
        {
            "success": len(successfully_applied) > 0,
            "message": (
                "Successfully applied paragraph formatting"
                if formatting_count > 0
                else "No formatting applied"
            ),
            "applied_settings": applied_settings,
            "object_count": formatting_count,
            "successfully_applied": successfully_applied,
        },
        ensure_ascii=False,
    )
