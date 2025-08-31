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
from word_document_server.mcp_service.core import mcp_server
from word_document_server.operations.text_ops import (
    get_character_count,
    get_element_text,
    insert_text_before_range,
    insert_text_after_range,
    apply_formatting_to_element,
    replace_element_text,
    set_bold_for_range,
    set_italic_for_range,
    set_font_size_for_range,
    set_font_name_for_range,
    set_font_color_for_range,
    set_alignment_for_range,
    set_paragraph_style
)
from word_document_server.selector.selector import SelectorEngine
from word_document_server.utils.app_context import AppContext
from word_document_server.utils.core_utils import (
    ErrorCode, WordDocumentError, format_error_response, get_active_document,
    handle_tool_errors, log_error, log_info, require_active_document_validation)


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
        default=None, description="Locator object for element selection. Returns all text when empty."
    ),
    text: Optional[str] = Field(
        default=None, description="Text content for insert or replace operations"
    ),
    position: str = Field(
        default="after",
        description="Position for insert operations: before, after, replace",
    ),
    style: Optional[str] = Field(default=None, description="Paragraph style name"),
    formatting: Optional[Dict[str, Any]] = Field(
        default=None,
        description="Formatting options: bold, italic, font_size, font_name, font_color, alignment",
    ),
    format_type: Optional[str] = Field(
        default=None,
        description="Text format type: bold, italic, font_size, font_name, font_color, alignment, paragraph_style",
    ),
    format_value: Optional[Any] = Field(
        default=None, description="Value for the text format operation"
    ),
) -> Any:
    """
    Unified text operation tool.

    This tool provides a single interface for all text operations:
    - get_text: Get text from document or specific element
    - insert_text: Insert text at specific element
    - replace_text: Replace text in specific element
    - get_char_count: Get character count of document or specific element
    - apply_formatting: Apply multiple formatting options to an element
    - format_text: Apply a single formatting option to an element
    - get_paragraphs: Get paragraphs in a specific range
    - get_paragraphs_info: Get paragraph statistics
    - get_all_paragraphs: Get all paragraphs in the document
    - insert_paragraph: Insert a new paragraph at specific element
    - get_paragraphs_in_range: Get paragraphs within a specific range

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
                    
                    if not selection or not selection.get_element_types():
                        # 如果找不到元素，返回空文本而不是抛出异常
                        return json.dumps({"success": True, "text": ""}, ensure_ascii=False)

                    # 获取选择区域的文本
                    element = selection._elements[0]
                    com_range_obj = element.Range
                    result = com_range_obj.Text
                except Exception as e:
                    # 如果选择过程出错，也返回空文本
                    log_error(f"Error selecting element: {e}")
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

            if not selection or not selection.get_element_types():
                raise WordDocumentError(
                    ErrorCode.ELEMENT_TYPE_ERROR,
                    "Failed to locate element for text insertion",
                )

            # 获取选择区域
            if not selection._elements:
                raise WordDocumentError(
                    ErrorCode.ELEMENT_TYPE_ERROR,
                    "Failed to get elements from selection",
                )
            
            # 获取第一个元素
            element = selection._elements[0]
            range_obj = None
            if hasattr(element, "Range"):
                range_obj = element.Range
            else:
                range_obj = active_doc.Range(0, 0)

            
            # 检查元素是Range对象还是普通元素
            if hasattr(element, 'Start') and hasattr(element, 'End'):
                # 这是一个Range对象
                if position.lower() == "before":
                    result = insert_text_before_range(
                        com_range=range_obj, text=text
                    )
                else:
                    result = insert_text_after_range(
                        com_range=range_obj, text=text
                    )
            else:
                # 这是一个普通元素，检查是否有Range属性
                if not hasattr(element, 'Range'):
                    range_obj = document.Range(0, 0)
                    
                # 插入文本
                if position.lower() == "before":
                    result = insert_text_before_range(
                        element=element, text=text
                    )
                else:
                    result = insert_text_before_range(
                        element=element, text=text
                    )

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

            if not selection or not selection.get_element_types():
                raise WordDocumentError(
                    ErrorCode.ELEMENT_TYPE_ERROR,
                    "Failed to locate element for text replacement",
                )

            # 获取选择区域
            if not selection._elements:
                raise WordDocumentError(
                    ErrorCode.ELEMENT_TYPE_ERROR,
                    "Failed to get elements from selection",
                )
            element = selection._elements[0]

            # 替换文本
            result = replace_element_text(element=element, new_text=text)

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

            if not selection or not selection.get_element_types():
                raise WordDocumentError(
                    ErrorCode.ELEMENT_TYPE_ERROR,
                    "Failed to locate element for formatting",
                )

            # 获取选择区域
            if not selection._elements:
                raise WordDocumentError(
                    ErrorCode.ELEMENT_TYPE_ERROR,
                    "Failed to get elements from selection",
                )
            element = selection._elements[0]

            # 应用格式
            result = apply_formatting_to_element(
                element=element, formatting=formatting
            )

            return json.dumps(
                {"success": True, "message": "Formatting applied successfully"},
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

            if not selection or not selection.get_element_types():
                raise WordDocumentError(
                    ErrorCode.ELEMENT_TYPE_ERROR,
                    "Failed to locate element for text formatting",
                )

            # 获取选择区域
            element = selection._elements[0]
            com_range_obj = element.Range

            # 应用文本格式
            if format_type.lower() == "bold":
                result = set_bold_for_range(com_range_obj, format_value)
            elif format_type.lower() == "italic":
                result = set_italic_for_range(com_range_obj, format_value)
            elif format_type.lower() == "font_size":
                result = set_font_size_for_range(com_range_obj, format_value)
            elif format_type.lower() == "font_name":
                result = set_font_name_for_range(com_range_obj, format_value)
            elif format_type.lower() == "font_color":
                result = set_font_color_for_range(
                    active_doc, com_range_obj, format_value
                )
            elif format_type.lower() == "alignment":
                result = set_alignment_for_range(
                    active_doc, com_range_obj, format_value
                )
            elif format_type.lower() == "paragraph_style":
                result = set_paragraph_style(element, format_value)
            else:
                raise ValueError(f"Unsupported format type: {format_type}")

            return json.dumps(
                {"success": True, "message": "Text formatted successfully"},
                ensure_ascii=False,
            )

        elif operation_type and operation_type.lower() == "get_paragraphs":
            log_info("Getting paragraphs")

            # 获取段落
            selector_engine = SelectorEngine()
            selection = selector_engine.select(active_doc, {"type": "paragraph"})
            result = [get_element_text(element) for element in selection._elements]

            return json.dumps(
                {"success": True, "paragraphs": result}, ensure_ascii=False
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

            if not selection or not selection.get_element_types():
                raise WordDocumentError(
                    ErrorCode.ELEMENT_TYPE_ERROR,
                    "Failed to locate element for paragraph insertion",
                )

            # 获取选择区域
            element = selection._elements[0]

            # 插入段落
            result = insert_text_after_element(
                element=element, text=f"\n{text}"
            )

            if style:
                # 应用段落样式
                selector_engine = SelectorEngine()
                selection = selector_engine.select(
                    active_doc, {"type": "paragraph", "index": -1}
                )  # 最后一个段落
                if selection and selection._elements:
                    set_paragraph_style(selection._elements[0], style)

            return json.dumps(
                {"success": True, "message": "Paragraph inserted successfully"},
                ensure_ascii=False,
            )

        elif operation_type and operation_type.lower() == "get_paragraphs_info":
            log_info("Getting paragraphs info")

            # 使用选择器引擎获取段落
            selector_engine = SelectorEngine()
            selection = selector_engine.select(active_doc, {"type": "paragraph"})

            # 获取段落信息
            paragraphs_info = []
            for i, element in enumerate(selection._elements):
                paragraphs_info.append(
                    {
                        "index": i,
                        "text": (
                            get_element_text(element)[:100] + "..."
                            if len(get_element_text(element)) > 100
                            else get_element_text(element)
                        ),
                        "characters": len(get_element_text(element)),
                    }
                )

            result = {
                "total_count": len(selection._elements),
                "paragraphs": paragraphs_info,
            }

            return json.dumps({"success": True, "info": result}, ensure_ascii=False)

        elif operation_type and operation_type.lower() == "get_all_paragraphs":
            log_info("Getting all paragraphs")

            # 使用选择器引擎获取所有段落
            selector_engine = SelectorEngine()
            selection = selector_engine.select(active_doc, {"type": "paragraph"})

            # 获取所有段落文本
            result = [get_element_text(element) for element in selection._elements]

            return json.dumps(
                {"success": True, "paragraphs": result}, ensure_ascii=False
            )

        elif operation_type and operation_type.lower() == "get_paragraphs_in_range":
            log_info("Getting paragraphs in range")

            # 使用选择器引擎获取所有段落
            selector_engine = SelectorEngine()
            selection = selector_engine.select(active_doc, {"type": "paragraph"})

            # 获取范围内的段落（这里简化为前10个段落）
            result = [get_element_text(element) for element in selection._elements[:10]]

            return json.dumps(
                {"success": True, "paragraphs": result}, ensure_ascii=False
            )

        else:
            raise ValueError(f"Unsupported operation type: {operation_type}")

    except Exception as e:
        log_error(f"Error in text_tools: {e}", exc_info=True)
        return format_error_response(str(e))
