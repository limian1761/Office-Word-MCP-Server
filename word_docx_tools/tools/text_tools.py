"""
Text Integration Tool for Word Document MCP Server.

This module provides a unified MCP tool interface for text operations,
which delegates the actual implementation to the operations layer.
"""

import json
import os
from typing import Any, Dict, List, Optional, Union

# Third-party imports
from mcp.server.fastmcp import Context
from mcp.server.session import ServerSession
from pydantic import Field

# Local imports
from ..mcp_service.core import mcp_server
from ..mcp_service.core_utils import (
    format_error_response,
    handle_tool_errors,
    log_error,
    log_info,
    require_active_document_validation
)
from ..operations.text_operations import (
    get_text_from_document,
    insert_text_into_document,
    replace_text_in_document,
    get_character_count_from_document,
    apply_formatting_to_document_text,
    format_document_text,
    validate_required_params
)
from ..selector.locator_parser import LocatorParser
from ..selector.exceptions import LocatorSyntaxError


@mcp_server.tool()
def get_locator_guid(
    ctx: Context[ServerSession, Any] = Field(description="Context object"),
) -> Dict[str, str]:
    """获取定位器指南内容
    
    此工具提供关于定位器（Locator）格式的完整指南，包括语法、元素类型、过滤器和使用示例。
    
    Returns:
        Dict[str, str]: 包含定位器指南内容的字典
    """
    try:
        # 读取LOCATOR_GUIDE.md文件内容
        guide_path = os.path.join(
            os.path.dirname(__file__), 
            "..", "selector", "LOCATOR_GUIDE.md"
        )
        
        try:
            with open(guide_path, 'r', encoding='utf-8') as f:
                guide_content = f.read()
            return {"guide_content": guide_content}
        except Exception as e:
            log_error(f"无法读取定位器指南文件: {e}")
            return {"error": f"无法读取定位器指南文件: {str(e)}"}
    except Exception as e:
        log_error(f"获取定位器指南时发生错误: {e}")
        return {"error": str(e)}


@mcp_server.tool()
@require_active_document_validation
@handle_tool_errors
def text_tools(
    ctx: Context[ServerSession, Any] = Field(description="Context object"),
    operation_type: Optional[str] = Field(
        default=None,
        description="Type of text operation: get_text, insert_text, replace_text, get_char_count, apply_formatting",
    ),
    locator: Optional[Dict[str, Any]] = Field(
        default=None,
        description="Locator object for object selection. This is a specially defined format with specific syntax requirements. Optional for get_text, Required for:  insert_text, replace_text, apply_formatting\n",
    ),
    text: Optional[str] = Field(
        default=None,
        description="Text content for insert or replace operations\n\n    Required for: insert_text, replace_text\n",
    ),
    position: str = Field(
        default="after",
        description="Position for insert operations: before, after, replace\n\n    Used by: insert_text\n",
    ),
    formatting: Optional[Dict[str, Any]] = Field(
        default=None,
        description="Formatting options: bold, italic, font_size, font_name, font_color, alignment\n\n    Used for: apply_formatting (multiple formats)\n",
    ),
    format_type: Optional[str] = Field(
        default=None,
        description="Single text format type: bold, italic, font_size, font_name, font_color, alignment, paragraph_style\n\n    Used for: apply_formatting (single format)\n",
    ),
    format_value: Optional[Union[str, int, float, bool]] = Field(
        default=None,
        description="Value for the single text format operation\n\n    Used for: apply_formatting (single format)\n",
    ),
) -> Any:
    """Unified text operation tool interface.

    This tool provides a single interface for all text operations, delegating
    the actual implementation to the operations layer:
    - get_text: Get text from document or specific object
    - insert_text: Insert text at specific object
    - replace_text: Replace text in specific object
    - get_char_count: Get character count of document or specific object
    - apply_formatting: Apply formatting options to an object

    Returns:
        Operation result based on the operation type
    """
    try:
        log_info(f"Starting text operation: {operation_type}")

        # 获取活动文档
        active_doc = ctx.request_context.lifespan_context.get_active_document()
        
        # 导入并使用通用的locator参数检查函数
        from .utils import check_locator_param
        
        # 根据操作类型调用相应的处理函数
        if operation_type == "get_text":
            check_locator_param(locator)
            return get_text_from_document(active_doc, locator)

        elif operation_type == "insert_text":
            check_locator_param(locator)
            validate_required_params({"text": text, "locator": locator}, "insert_text")
            return insert_text_into_document(active_doc, text, locator, position)

        elif operation_type == "replace_text":
            check_locator_param(locator)
            validate_required_params({"text": text, "locator": locator}, "replace_text")
            return replace_text_in_document(active_doc, text, locator)

        elif operation_type == "get_char_count":
            check_locator_param(locator)
            return get_character_count_from_document(active_doc, locator)

        elif operation_type == "apply_formatting":
            check_locator_param(locator)
            # 验证必需参数
            validate_required_params({"locator": locator}, "apply_formatting")
            
            # 检查是使用多个格式还是单个格式
            if formatting is not None:
                # 使用多个格式
                return apply_formatting_to_document_text(active_doc, formatting, locator)
            elif format_type is not None and format_value is not None:
                # 使用单个格式
                return format_document_text(active_doc, format_type, format_value, locator)
            else:
                raise ValueError("Either formatting (for multiple formats) or format_type and format_value (for single format) must be provided for apply_formatting operation")

        else:
            raise ValueError(f"Unsupported operation type: {operation_type}")

    except Exception as e:
        log_error(f"Error in text_tools: {e}", exc_info=True)
        return format_error_response(e)
