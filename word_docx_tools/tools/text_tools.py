"""
Text Integration Tool for Word Document MCP Server.

This module provides a unified MCP tool for text operations.
"""

import json
import os
from typing import Any, Dict, List, Optional

# Standard library imports
from dotenv import load_dotenv
# Third-party imports
from mcp.server.fastmcp import Context
from mcp.server.session import ServerSession
from pydantic import Field

# Local imports
from ..mcp_service.core import mcp_server
from ..operations.text_ops import (
    apply_formatting_to_object, get_character_count, get_object_text,
    insert_text_after_range, insert_text_before_range, replace_object_text,
    set_alignment_for_range, set_bold_for_range, set_font_color_for_range,
    set_font_name_for_range, set_font_size_for_range, set_italic_for_range,
    set_paragraph_style)
from ..selector.selector import SelectorEngine
from ..utils.app_context import AppContext
from ..mcp_service.core_utils import (
    ErrorCode, WordDocumentError, format_error_response, get_active_document,
    handle_tool_errors, log_error, log_info,
    require_active_document_validation)


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
        default=False,
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
    format_value: Optional[Any] = Field(
        default=None,
        description="Value for the text format operation\n\n    Required for: format_text\n",
    ),
) -> Any:
    """Unified text operation tool.

    This tool provides a single interface for all text operations:
    - get_text: Get text from document or specific object
      * Optional parameters: locator
    - insert_text: Insert text at specific object
      * Required parameters: text, locator
      * Optional parameters: position
    - replace_text: Replace text in specific object
      * Required parameters: text, locator
    - get_char_count: Get character count of document or specific object
      * Optional parameters: locator
    - apply_formatting: Apply multiple formatting options to an object
      * Required parameters: formatting, locator
    - format_text: Apply a single formatting option to an object
      * Required parameters: format_type, format_value, locator
    - get_paragraphs: Get paragraphs in a specific range
      * No required parameters
    - get_paragraphs_info: Get paragraph statistics
      * No required parameters
    - get_all_paragraphs: Get all paragraphs in the document
      * No required parameters
    - insert_paragraph: Insert a new paragraph at specific object
      * Required parameters: text, locator
      * Optional parameters: style, is_independent_paragraph
    - get_paragraphs_in_range: Get paragraphs within a specific range
      * No required parameters

    Returns:
        Operation result based on the operation type
    """
    try:
        # 获取活动文档
        active_doc = ctx.request_context.lifespan_context.get_active_document()
        # 根据操作类型执行相应的操作
        if operation_type and operation_type.lower() == "get_text":
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

        elif operation_type and operation_type.lower() == "insert_text":
            # 验证插入文本所需的参数
            if text is None:
                raise ValueError(
                    "text parameter must be provided for insert_text operation"
                )
            if locator is None:
                raise ValueError(
                    "locator parameter must be provided for insert_text operation"
                )

            log_info(f"Inserting text: {text}")

            # 使用选择器引擎定位元素
            selector_engine = SelectorEngine()
            selection = selector_engine.select(active_doc, locator)

            if not selection or not selection.get_object_types():
                raise WordDocumentError(
                    ErrorCode.OBJECT_TYPE_ERROR,
                    "Failed to locate object for text insertion",
                )

            # 获取选择区域
            if not selection._com_ranges:
                raise WordDocumentError(
                    ErrorCode.OBJECT_TYPE_ERROR,
                    "Failed to get objects from selection",
                )

            # 获取第一个元素
            # Selection._com_ranges中只包含Range对象
            object = selection._com_ranges[0]
            range_obj = None
            if hasattr(object, "Range"):
                range_obj = object.Range
            else:
                range_obj = active_doc.Range(0, 0)
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

        elif operation_type and operation_type.lower() == "replace_text":
            # 验证替换文本所需的参数
            if text is None:
                raise ValueError(
                    "text parameter must be provided for replace_text operation"
                )
            if locator is None:
                raise ValueError(
                    "locator parameter must be provided for replace_text operation"
                )

            log_info(f"Replacing text with: {text}")

            # 使用选择器引擎定位元素
            selector_engine = SelectorEngine()
            selection = selector_engine.select(active_doc, locator)

            if not selection or not selection.get_object_types():
                raise WordDocumentError(
                    ErrorCode.OBJECT_TYPE_ERROR,
                    "Failed to locate object for text replacement",
                )

            # 获取选择区域
            if not selection._com_ranges:
                raise WordDocumentError(
                    ErrorCode.OBJECT_TYPE_ERROR,
                    "Failed to get objects from selection",
                )
            # Selection._com_ranges中只包含Range对象
            object = selection._com_ranges[0]

            # 替换文本
            result = replace_object_text(range_obj=object, new_text=text)

            return json.dumps(
                {"success": True, "message": "Text replaced successfully"},
                ensure_ascii=False,
            )

        elif operation_type and operation_type.lower() == "get_char_count":
            log_info("Getting character count")

            # 获取字符数
            result = get_character_count(document=active_doc, locator=locator)

            return json.dumps({"success": True, "count": result}, ensure_ascii=False)

        elif operation_type and operation_type.lower() == "apply_formatting":
            # 验证应用格式所需的参数
            if formatting is None:
                raise ValueError(
                    "formatting parameter must be provided for apply_formatting operation"
                )
            if locator is None:
                raise ValueError(
                    "locator parameter must be provided for apply_formatting operation"
                )

            log_info("Applying formatting")

            # 使用选择器引擎定位元素
            selector_engine = SelectorEngine()
            selection = selector_engine.select(active_doc, locator)

            if not selection or not selection.get_object_types():
                raise WordDocumentError(
                    ErrorCode.OBJECT_TYPE_ERROR,
                    "Failed to locate object for formatting",
                )

            # 获取选择区域
            if not hasattr(selection, '_com_ranges') or not selection._com_ranges:
                raise WordDocumentError(
                    ErrorCode.OBJECT_TYPE_ERROR,
                    "Failed to get objects from selection",
                )
            # Selection._com_ranges中只包含Range对象
            range_obj = selection._com_ranges[0]

            # 应用格式
            result = apply_formatting_to_object(range_obj=range_obj, formatting=formatting)

            return json.dumps(
                {"success": True, "message": "Formatting applied successfully"},
                ensure_ascii=False,
            )

        elif operation_type and operation_type.lower() == "replace_text":
            # 验证替换文本所需的参数
            if text is None:
                raise ValueError(
                    "text parameter must be provided for replace_text operation"
                )
            if locator is None:
                raise ValueError(
                    "locator parameter must be provided for replace_text operation"
                )

            log_info(f"Replacing text with: {text}")

            # 使用选择器引擎定位元素
            selector_engine = SelectorEngine()
            selection = selector_engine.select(active_doc, locator)

            if not selection or not selection.get_object_types():
                raise WordDocumentError(
                    ErrorCode.OBJECT_TYPE_ERROR,
                    "Failed to locate object for text replacement",
                )

            # 获取选择区域
            if not hasattr(selection, '_com_ranges') or not selection._com_ranges:
                raise WordDocumentError(
                    ErrorCode.OBJECT_TYPE_ERROR,
                    "Failed to get objects from selection",
                )
            # Selection._com_ranges中只包含Range对象
            range_obj = selection._com_ranges[0]

            # 替换文本
            result = replace_object_text(range_obj=range_obj, new_text=text)

            return json.dumps(
                {"success": True, "message": "Text replaced successfully"},
                ensure_ascii=False,
            )

        elif operation_type and operation_type.lower() == "format_text":
            # 验证格式化文本所需的参数
            if format_type is None:
                raise ValueError(
                    "format_type parameter must be provided for format_text operation"
                )
            if format_value is None:
                raise ValueError(
                    "format_value parameter must be provided for format_text operation"
                )
            if locator is None:
                raise ValueError(
                    "locator parameter must be provided for format_text operation"
                )

            log_info(f"Applying text format: {format_type}")

            # 使用选择器引擎定位元素
            selector_engine = SelectorEngine()
            selection = selector_engine.select(active_doc, locator)

            if not selection or not selection.get_object_types():
                raise WordDocumentError(
                    ErrorCode.OBJECT_TYPE_ERROR,
                    "Failed to locate object for text formatting",
                )

            # 获取选择区域
            if not hasattr(selection, '_com_ranges') or not selection._com_ranges:
                raise WordDocumentError(
                    ErrorCode.OBJECT_TYPE_ERROR,
                    "Failed to get objects from selection",
                )
            # Selection._com_ranges中只包含Range对象
            range_obj = selection._com_ranges[0]
            
            # 检查Range属性
            if not hasattr(range_obj, "Range"):
                raise WordDocumentError(
                    ErrorCode.OBJECT_TYPE_ERROR,
                    "Selected object does not have a Range property",
                )
            com_range = range_obj.Range

            # 应用文本格式
            if format_type.lower() == "bold":
                result = set_bold_for_range(com_range, format_value)
            elif format_type.lower() == "italic":
                result = set_italic_for_range(com_range, format_value)
            elif format_type.lower() == "font_size":
                result = set_font_size_for_range(com_range, format_value)
            elif format_type.lower() == "font_name":
                result = set_font_name_for_range(com_range, format_value)
            elif format_type.lower() == "font_color":
                result = set_font_color_for_range(
                    active_doc, com_range, format_value
                )
            elif format_type.lower() == "alignment":
                result = set_alignment_for_range(
                    active_doc, com_range, format_value
                )
            elif format_type.lower() == "paragraph_style":
                result = set_paragraph_style(range_obj, format_value)
            else:
                raise ValueError(f"Unsupported format type: {format_type}")

            return json.dumps(
                {"success": True, "message": "Text formatted successfully"},
                ensure_ascii=False,
            )

        elif operation_type and operation_type.lower() == "insert_paragraph":
            # 验证插入段落所需的参数
            if text is None:
                raise ValueError(
                    "text parameter must be provided for insert_paragraph operation"
                )
            if locator is None:
                raise ValueError(
                    "locator parameter must be provided for insert_paragraph operation"
                )

            log_info(f"Inserting paragraph: {text}")

            # 使用选择器引擎定位元素
            selector_engine = SelectorEngine()
            selection = selector_engine.select(active_doc, locator)

            if not selection or not selection.get_object_types():
                raise WordDocumentError(
                    ErrorCode.OBJECT_TYPE_ERROR,
                    "Failed to locate object for paragraph insertion",
                )

            # 获取选择区域
            if not hasattr(selection, '_com_ranges') or not selection._com_ranges:
                raise WordDocumentError(
                    ErrorCode.OBJECT_TYPE_ERROR,
                    "Failed to get objects from selection",
                )
            # Selection._com_ranges中只包含Range对象
            range_obj = selection._com_ranges[0]

            # 如果需要作为独立段落插入
            if is_independent_paragraph:
                try:
                    # 检查当前范围是否已经在段落末尾
                    if hasattr(range_obj, 'Paragraphs') and range_obj.Paragraphs.Count > 0:
                        current_paragraph = range_obj.Paragraphs(1)
                        # 如果范围不在段落末尾，创建新段落
                        if range_obj.Start != current_paragraph.Range.End - 1:
                            # 在当前范围前插入段落标记创建新段落
                            range_obj.InsertBefore('\n')
                            # 更新范围到新段落
                            range_obj.Start = range_obj.Start
                            range_obj.End = range_obj.Start
                except Exception as e:
                    log_error(f"Failed to prepare independent paragraph: {str(e)}")

            # 插入段落
            # 改进Range对象检测逻辑，确保document_end返回的Range对象能被正确处理
            try:
                # 先尝试直接使用object作为Range对象
                result = insert_text_after_range(com_range=range_obj, text=f"\n{text}")
            except Exception:
                # 如果失败，再尝试使用object.Range
                if hasattr(range_obj, "Range"):
                    result = insert_text_after_range(
                        com_range=range_obj.Range, text=f"\n{text}"
                    )
                else:
                    raise WordDocumentError(
                        ErrorCode.OBJECT_TYPE_ERROR,
                        "Cannot insert paragraph: Invalid object type",
                    )

            if style:
                # 应用段落样式
                selector_engine = SelectorEngine()
                selection = selector_engine.select(
                    active_doc, {"type": "paragraph", "index": -1}  # 最后一个段落
                )
                if selection and hasattr(selection, '_com_ranges') and selection._com_ranges:
                    # Selection._com_ranges中只包含Range对象
                    set_paragraph_style(selection._com_ranges[0], style)

            return json.dumps(
                {"success": True, "message": "Paragraph inserted successfully"},
                ensure_ascii=False,
            )

        elif operation_type and operation_type.lower() == "get_all_paragraphs":
            log_info("Getting all paragraphs from document")
            
            try:
                paragraphs = []
                # 获取文档中的所有段落
                for i in range(1, active_doc.Paragraphs.Count + 1):
                    paragraph = active_doc.Paragraphs(i)
                    paragraph_range = paragraph.Range
                    paragraphs.append({
                        "text": paragraph_range.Text,
                        "index": i
                    })
                
                return json.dumps({
                    "success": True,
                    "paragraphs": paragraphs
                }, ensure_ascii=False)
            except Exception as e:
                log_error(f"Error getting all paragraphs: {e}")
                raise WordDocumentError(
                    ErrorCode.SERVER_ERROR,
                    f"Failed to get all paragraphs: {str(e)}"
                )

        else:
            raise ValueError(f"Unsupported operation type: {operation_type}")

    except Exception as e:
        log_error(f"Error in text_tools: {e}", exc_info=True)
        return format_error_response(e)
