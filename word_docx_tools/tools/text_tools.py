"""
Text Integration Tool for Word Document MCP Server.

This module provides a unified MCP tool for text operations.
"""

import json
import os
from typing import Any, Dict, List, Optional, Union

# Standard library imports
from dotenv import load_dotenv
# Third-party imports
from mcp.server.fastmcp import Context
from mcp.server.session import ServerSession
from pydantic import Field

# Local imports
from ..mcp_service.core import mcp_server
from ..mcp_service.core_utils import (ErrorCode, WordDocumentError,
                                      format_error_response,
                                      get_active_document, handle_tool_errors,
                                      log_error, log_info,
                                      require_active_document_validation)
from ..operations.text_ops import (apply_formatting_to_object,
                                   get_character_count, get_object_text,
                                   insert_text_after_range,
                                   insert_text_before_range,
                                   replace_object_text,
                                   set_alignment_for_range, set_bold_for_range,
                                   set_font_color_for_range,
                                   set_font_name_for_range,
                                   set_font_size_for_range,
                                   set_italic_for_range, set_paragraph_style)
from ..selector.selector import SelectorEngine
from ..mcp_service.app_context import AppContext


@mcp_server.tool()
@require_active_document_validation
@handle_tool_errors
def text_tools(
    ctx: Context[ServerSession, AppContext] = Field(description="Context object"),
    operation_type: Optional[str] = Field(
        default=None,
        description="Type of text operation: get_text, insert_text, replace_text, get_char_count, apply_formatting, get_paragraphs, insert_paragraph, get_paragraphs_info, get_all_paragraphs, get_paragraphs_in_range, format_text",
    ),
    locator: Optional[Dict[str, Any]] = Field(
        default=None,
        description="Locator object for object selection. Returns all text when empty.\n\n    Required for: insert_text, replace_text, apply_formatting, format_text, insert_paragraph\n",
    ),
    text: Optional[str] = Field(
        default=None,
        description="Text content for insert or replace operations\n\n    Required for: insert_text, replace_text, insert_paragraph\n",
    ),
    position: str = Field(
        default="after",
        description="Position for insert operations: before, after, replace\n\n    Used by: insert_text\n",
    ),
    style: Optional[str] = Field(
        default=None,
        description="Paragraph style name\n\n    Optional for: insert_paragraph\n",
    ),
    is_independent_paragraph: bool = Field(
        default=True,
        description="Whether to insert the paragraph as an independent paragraph\n\n    Optional for: insert_paragraph\n",
    ),
    formatting: Optional[Dict[str, Any]] = Field(
        default=None,
        description="Formatting options: bold, italic, font_size, font_name, font_color, alignment\n\n    Required for: apply_formatting\n",
    ),
    format_type: Optional[str] = Field(
        default=None,
        description="Text format type: bold, italic, font_size, font_name, font_color, alignment, paragraph_style\n\n    Required for: format_text\n",
    ),
    format_value: Optional[Union[str, int, float, bool]] = Field(
        default=None,
        description="Value for the text format operation\n\n    Required for: format_text\n",
    ),
) -> Any:
    """Unified text operation tool.

    This tool provides a single interface for all text operations:
    - get_text: Get text from document or specific object
    - insert_text: Insert text at specific object
    - replace_text: Replace text in specific object
    - get_char_count: Get character count of document or specific object
    - apply_formatting: Apply multiple formatting options to an object
    - format_text: Apply a single formatting option to an object
    - get_paragraphs: Get paragraphs in a specific range
    - get_paragraphs_info: Get paragraph statistics
    - get_all_paragraphs: Get all paragraphs in the document
    - insert_paragraph: Insert a new paragraph at specific object
    - get_paragraphs_in_range: Get paragraphs within a specific range

    Returns:
        Operation result based on the operation type
    """
    try:
        log_info(f"Starting text operation: {operation_type}")

        # 获取活动文档
        active_doc = ctx.request_context.lifespan_context.get_active_document()
        if not active_doc:
            raise WordDocumentError(
                ErrorCode.DOCUMENT_NOT_OPEN, "No active document found"
            )

        # 根据操作类型调用相应的处理函数
        if operation_type == "get_text":
            return _handle_get_text_operation(active_doc, locator)

        elif operation_type == "insert_text":
            return _handle_insert_text_operation(active_doc, text, locator, position)

        elif operation_type == "replace_text":
            return _handle_replace_text_operation(active_doc, text, locator)

        elif operation_type == "get_char_count":
            # 获取字符数
            if locator:
                range_obj = _get_selection_range(active_doc, locator, "character count")
                text_content = range_obj.Text
            else:
                text_content = active_doc.Content.Text
            
            char_count = len(text_content)
            return json.dumps(
                {"success": True, "character_count": char_count}, ensure_ascii=False
            )

        elif operation_type == "apply_formatting":
            return _handle_apply_formatting_operation(active_doc, formatting, locator)

        elif operation_type == "format_text":
            return _handle_format_text_operation(active_doc, format_type, format_value, locator)

        elif operation_type == "get_paragraphs":
            # 获取特定范围内的段落
            if locator:
                range_obj = _get_selection_range(active_doc, locator, "paragraphs")
                paragraphs = []
                for i in range(1, range_obj.Paragraphs.Count + 1):
                    paragraph = range_obj.Paragraphs(i)
                    paragraphs.append({
                        "text": paragraph.Range.Text,
                        "index": i
                    })
                return json.dumps(
                    {"success": True, "paragraphs": paragraphs}, ensure_ascii=False
                )
            else:
                return _handle_get_all_paragraphs_operation(active_doc)

        elif operation_type == "get_paragraphs_info":
            return _handle_get_paragraphs_info_operation(active_doc)

        elif operation_type == "get_all_paragraphs":
            return _handle_get_all_paragraphs_operation(active_doc)

        elif operation_type == "insert_paragraph":
            return _handle_insert_paragraph_operation(
                active_doc, text, locator, style, is_independent_paragraph
            )

        elif operation_type == "get_paragraphs_in_range":
            # 获取特定范围内的段落
            if not locator:
                raise ValueError("locator parameter must be provided for get_paragraphs_in_range operation")
            
            range_obj = _get_selection_range(active_doc, locator, "paragraphs in range")
            paragraphs = []
            
            for i in range(1, range_obj.Paragraphs.Count + 1):
                paragraph = range_obj.Paragraphs(i)
                paragraphs.append({
                    "text": paragraph.Range.Text,
                    "index": i,
                    "start": paragraph.Range.Start,
                    "end": paragraph.Range.End
                })
            
            return json.dumps(
                {"success": True, "paragraphs": paragraphs}, ensure_ascii=False
            )

        else:
            raise ValueError(f"Unsupported operation type: {operation_type}")

    except Exception as e:
        log_error(f"Error in text_tools: {e}", exc_info=True)
        return format_error_response(e)


# 辅助函数
def _get_selection_range(active_doc, locator, operation_name):
    """获取选择范围，处理错误"""
    if locator is None:
        raise ValueError(
            f"locator parameter must be provided for {operation_name} operation"
        )

    selector_engine = SelectorEngine()
    selection = selector_engine.select(active_doc, locator)

    if not selection or not selection.get_object_types():
        raise WordDocumentError(
            ErrorCode.OBJECT_TYPE_ERROR,
            f"Failed to locate object for {operation_name}"
        )

    if not hasattr(selection, "_com_ranges") or not selection._com_ranges:
        raise WordDocumentError(
            ErrorCode.OBJECT_TYPE_ERROR,
            f"Failed to get objects from selection for {operation_name}"
        )

    return selection._com_ranges[0]


def _validate_required_params(params, operation_name):
    """验证必需参数"""
    for param_name, param_value in params.items():
        if param_value is None:
            raise ValueError(
                f"{param_name} parameter must be provided for {operation_name} operation"
            )


def _handle_get_text_operation(active_doc, locator):
    """处理获取文本操作"""
    log_info("Getting text from document")

    if locator:
        # 如果提供了定位器，获取特定元素的文本
        selector_engine = SelectorEngine()
        try:
            selection = selector_engine.select(active_doc, locator)

            if not selection or not selection.get_object_types():
                # 如果找不到元素，返回空文本而不是抛出异常
                return json.dumps(
                    {"success": True, "text": ""}, ensure_ascii=False
                )

            # 获取选择区域的文本
            # Selection._com_ranges中只包含Range对象
            range_obj = selection._com_ranges[0]
            result = range_obj.Text
        except Exception as e:
            # 如果选择过程出错，也返回空文本
            log_error(f"Error selecting object: {e}")
            return json.dumps({"success": True, "text": ""}, ensure_ascii=False)
    else:
        # 如果没有提供定位器，获取整个文档的文本
        result = active_doc.Content.Text

    return json.dumps({"success": True, "text": result}, ensure_ascii=False)


def _handle_insert_text_operation(active_doc, text, locator, position):
    """处理插入文本操作"""
    _validate_required_params({"text": text, "locator": locator}, "insert_text")
    log_info(f"Inserting text: {text}")

    range_obj = _get_selection_range(active_doc, locator, "text insertion")

    # 插入文本
    if position.lower() == "before":
        result = insert_text_before_range(com_range=range_obj, text=text)
    else:
        result = insert_text_after_range(com_range=range_obj, text=text)

    # 检查返回结果是否为字符串（JSON格式），如果是则直接返回
    if isinstance(result, str):
        try:
            result_dict = json.loads(result)
            if not result_dict.get("success", False):
                return json.dumps(result_dict, ensure_ascii=False)
        except json.JSONDecodeError:
            pass
        return result

    # 如果函数返回的是布尔值，则构造结果消息
    return json.dumps(
        {"success": True, "message": "Text inserted successfully"},
        ensure_ascii=False,
    )


def _handle_replace_text_operation(active_doc, text, locator):
    """处理替换文本操作"""
    _validate_required_params({"text": text, "locator": locator}, "replace_text")
    log_info(f"Replacing text with: {text}")

    range_obj = _get_selection_range(active_doc, locator, "text replacement")

    # 替换文本
    result = replace_object_text(range_obj=range_obj, new_text=text)

    return json.dumps(
        {"success": True, "message": "Text replaced successfully"},
        ensure_ascii=False,
    )


def _handle_apply_formatting_operation(active_doc, formatting, locator):
    """处理应用格式操作"""
    _validate_required_params({"formatting": formatting, "locator": locator}, "apply_formatting")
    log_info("Applying formatting")

    range_obj = _get_selection_range(active_doc, locator, "formatting")

    # 应用格式
    result = apply_formatting_to_object(range_obj=range_obj, formatting=formatting)

    return json.dumps(
        {"success": True, "message": "Formatting applied successfully"},
        ensure_ascii=False,
    )


def _handle_format_text_operation(active_doc, format_type, format_value, locator):
    """处理格式化文本操作"""
    _validate_required_params({
        "format_type": format_type, 
        "format_value": format_value, 
        "locator": locator
    }, "format_text")
    log_info(f"Applying text format: {format_type}")

    range_obj = _get_selection_range(active_doc, locator, "text formatting")

    # 应用文本格式
    if format_type.lower() == "bold":
        result = set_bold_for_range(range_obj, format_value)
    elif format_type.lower() == "italic":
        result = set_italic_for_range(range_obj, format_value)
    elif format_type.lower() == "font_size":
        result = set_font_size_for_range(range_obj, format_value)
    elif format_type.lower() == "font_name":
        result = set_font_name_for_range(range_obj, format_value)
    elif format_type.lower() == "font_color":
        result = set_font_color_for_range(active_doc, range_obj, format_value)
    elif format_type.lower() == "alignment":
        result = set_alignment_for_range(active_doc, range_obj, format_value)
    elif format_type.lower() == "paragraph_style":
        result = set_paragraph_style(range_obj, format_value)
    else:
        raise ValueError(f"Unsupported format type: {format_type}")

    return json.dumps(
        {"success": True, "message": "Text formatted successfully"},
        ensure_ascii=False,
    )


def _handle_insert_paragraph_operation(active_doc, text, locator, style, is_independent_paragraph):
    """处理插入段落操作"""
    _validate_required_params({"text": text, "locator": locator}, "insert_paragraph")
    log_info(f"Inserting paragraph: {text}")

    range_obj = _get_selection_range(active_doc, locator, "paragraph insertion")

    # 如果需要作为独立段落插入
    if is_independent_paragraph:
        try:
            # 检查当前范围是否已经在段落末尾
            if (
                hasattr(range_obj, "Paragraphs")
                and range_obj.Paragraphs.Count > 0
            ):
                current_paragraph = range_obj.Paragraphs(1)
                # 如果范围不在段落末尾，创建新段落
                if range_obj.Start != current_paragraph.Range.End - 1:
                    # 在当前范围前插入段落标记创建新段落
                    range_obj.InsertBefore("\n")
                    # 更新范围到新段落
                    range_obj.Start = range_obj.Start
                    range_obj.End = range_obj.Start
        except Exception as e:
            log_error(f"Failed to prepare independent paragraph: {str(e)}")

    # 插入段落
    result = insert_text_after_range(com_range=range_obj, text=f"\n{text}")

    if style:
        # 应用段落样式
        selector_engine = SelectorEngine()
        selection = selector_engine.select(
            active_doc, {"type": "paragraph", "index": -1}  # 最后一个段落
        )
        if (
            selection
            and hasattr(selection, "_com_ranges")
            and selection._com_ranges
        ):
            set_paragraph_style(selection._com_ranges[0], style)

    return result


def _handle_get_all_paragraphs_operation(active_doc):
    """处理获取所有段落操作"""
    log_info("Getting all paragraphs from document")

    try:
        paragraphs = []
        # 获取文档中的所有段落
        for i in range(1, active_doc.Paragraphs.Count + 1):
            paragraph = active_doc.Paragraphs(i)
            paragraph_range = paragraph.Range
            paragraphs.append({"text": paragraph_range.Text, "index": i})

        return json.dumps(
            {"success": True, "paragraphs": paragraphs}, ensure_ascii=False
        )
    except Exception as e:
        log_error(f"Error getting all paragraphs: {e}")
        raise WordDocumentError(
            ErrorCode.SERVER_ERROR, f"Failed to get all paragraphs: {str(e)}"
        )


def _handle_get_paragraphs_info_operation(active_doc):
    """处理获取段落信息操作"""
    log_info("Getting paragraph statistics")

    try:
        total_paragraphs = active_doc.Paragraphs.Count
        empty_paragraphs = 0
        total_characters = 0
        total_words = 0
        paragraphs_info = []

        # 获取所有段落的详细信息
        for i in range(1, total_paragraphs + 1):
            paragraph = active_doc.Paragraphs(i)
            paragraph_range = paragraph.Range
            text = paragraph_range.Text.strip()
            
            # 计算字符数和字数
            paragraph_chars = len(text)
            paragraph_words = len(text.split()) if text else 0
            
            total_characters += paragraph_chars
            total_words += paragraph_words
            
            # 统计空段落
            if not text:
                empty_paragraphs += 1
            
            # 获取段落样式信息
            style_name = ""
            try:
                if hasattr(paragraph, "Style") and paragraph.Style:
                    style_name = paragraph.Style.NameLocal
            except:
                pass
            
            # 获取段落开头和结尾的句子
            opening_sentence = ""
            closing_sentence = ""
            if text:
                sentences = text.split('.')
                if sentences:
                    opening_sentence = sentences[0].strip() + ('.' if len(sentences) > 1 else '')
                    closing_sentence = sentences[-1].strip() + '.' if sentences and len(sentences) > 1 else ""
            
            # 构建段落信息
            paragraph_info = {
                "index": i,
                "text_length": paragraph_chars,
                "word_count": paragraph_words,
                "style_name": style_name,
                "opening_sentence": opening_sentence,
                "closing_sentence": closing_sentence,
                "is_empty": not bool(text)
            }
            
            # 如果段落包含文字，添加文字内容
            if text:
                paragraph_info["text_preview"] = text[:100] + "..." if len(text) > 100 else text
            else:
                # 对于空段落，添加其他属性信息
                paragraph_info["has_formatting"] = hasattr(paragraph, "Format") and paragraph.Format is not None
                try:
                    paragraph_info["outline_level"] = paragraph.OutlineLevel if hasattr(paragraph, "OutlineLevel") else 0
                except:
                    paragraph_info["outline_level"] = 0
                
            paragraphs_info.append(paragraph_info)
        
        # 计算平均值
        non_empty_paragraphs = total_paragraphs - empty_paragraphs
        avg_chars_per_paragraph = total_characters / non_empty_paragraphs if non_empty_paragraphs > 0 else 0
        avg_words_per_paragraph = total_words / non_empty_paragraphs if non_empty_paragraphs > 0 else 0
        
        # 构建结果
        stats = {
            "total_paragraphs": total_paragraphs,
            "empty_paragraphs": empty_paragraphs,
            "non_empty_paragraphs": non_empty_paragraphs,
            "total_characters": total_characters,
            "total_words": total_words,
            "avg_characters_per_paragraph": round(avg_chars_per_paragraph, 2),
            "avg_words_per_paragraph": round(avg_words_per_paragraph, 2)
        }
