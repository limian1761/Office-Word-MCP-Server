"""
Range Selection Tool for Word Document MCP Server.

This module provides a simplified MCP tool for handling user selection ranges.
"""

import json
from typing import Any, Dict, Optional

# Third-party imports
from mcp.server.fastmcp import Context
from mcp.server.session import ServerSession
from pydantic import Field

# Local imports
from ..mcp_service.core import mcp_server
from ..mcp_service.app_context import AppContext
from ..mcp_service.core_utils import (
    ErrorCode,
    WordDocumentError,
    get_active_document,
    log_error,
    log_info
)
from ..operations.text_ops import apply_formatting_to_object


@mcp_server.tool()
async def range_tools(
    ctx: Context[ServerSession, AppContext],
    operation_type: str = Field(
        ...,
        description="Type of range operation: get_current_selection, modify_selection_text, modify_selection_style",
    ),
    text: Optional[str] = Field(
        default=None,
        description="Text content for modify_selection_text operation",
    ),
    formatting: Optional[Dict[str, Any]] = Field(
        default=None,
        description="Formatting options for modify_selection_style operation",
    ),
) -> str:
    """
    Simplified range operation tool focused on user selection.

    This tool provides interfaces for:
    - get_current_selection: Get information about the current user selection
      * Required parameters: None
    - modify_selection_text: Modify the text content of the current selection
      * Required parameters: text
    - modify_selection_style: Apply formatting to the current selection
      * Required parameters: formatting

    Returns:
        Result of the operation in JSON format
    """
    # Get the active Word document
    active_doc = ctx.request_context.lifespan_context.get_active_document()

    # Check if there is an active document
    if active_doc is None:
        raise WordDocumentError(
            ErrorCode.NO_ACTIVE_DOCUMENT, "没有找到活动文档"
        )

    try:
        # 获取当前选择
        if operation_type == "get_current_selection":
            return _get_current_selection(active_doc)

        # 修改当前选择的文字
        elif operation_type == "modify_selection_text":
            if text is None:
                raise WordDocumentError(
                    ErrorCode.INVALID_INPUT, "修改选择文字操作需要提供text参数"
                )
            return _modify_selection_text(active_doc, text)

        # 修改当前选择的样式
        elif operation_type == "modify_selection_style":
            if formatting is None:
                raise WordDocumentError(
                    ErrorCode.INVALID_INPUT, "修改选择样式操作需要提供formatting参数"
                )
            return _modify_selection_style(active_doc, formatting)

        else:
            error_msg = f"不支持的操作类型: {operation_type}"
            log_error(error_msg)
            raise WordDocumentError(ErrorCode.INVALID_INPUT, error_msg)

    except Exception as e:
        log_error(f"range_tools错误: {str(e)}", exc_info=True)
        raise


def _get_current_selection(document) -> str:
    """
    获取当前用户选择的信息
    """
    try:
        # 获取Word应用程序的Selection对象
        app = document.Application
        selection = app.Selection
        
        # 检查是否有选择内容
        if selection.Start == selection.End:
            return json.dumps({
                "success": False,
                "message": "没有选择任何内容"
            }, ensure_ascii=False)
        
        # 构建选择信息
        selection_info = {
            "success": True,
            "text": selection.Text,
            "length": len(selection.Text),
            "start_position": selection.Start,
            "end_position": selection.End
        }
        
        # 添加样式信息（如果可用）
        try:
            if hasattr(selection, "Style") and hasattr(selection.Style, "NameLocal"):
                selection_info["style"] = selection.Style.NameLocal
        except Exception as e:
            log_error(f"获取样式信息失败: {e}")
        
        log_info(f"成功获取当前选择信息，长度: {selection_info['length']}")
        return json.dumps(selection_info, ensure_ascii=False)
        
    except Exception as e:
        log_error(f"获取当前选择失败: {e}")
        raise WordDocumentError(ErrorCode.OBJECT_NOT_FOUND, f"获取当前选择失败: {str(e)}")


def _modify_selection_text(document, text: str) -> str:
    """
    修改当前选择的文字内容
    """
    try:
        # 获取Word应用程序的Selection对象
        app = document.Application
        selection = app.Selection
        
        # 检查是否有选择内容
        if selection.Start == selection.End:
            return json.dumps({
                "success": False,
                "message": "没有选择任何内容，无法修改文字"
            }, ensure_ascii=False)
        
        # 保存原始长度用于日志
        original_length = len(selection.Text)
        
        # 修改选择的文字
        selection.Text = text
        
        log_info(f"成功修改选择文字，原长度: {original_length}，新长度: {len(text)}")
        return json.dumps({
            "success": True,
            "message": "选择文字修改成功",
            "new_text": text,
            "new_length": len(text)
        }, ensure_ascii=False)
        
    except Exception as e:
        log_error(f"修改选择文字失败: {e}")
        raise WordDocumentError(ErrorCode.OBJECT_MODIFICATION_ERROR, f"修改选择文字失败: {str(e)}")


def _modify_selection_style(document, formatting: Dict[str, Any]) -> str:
    """
    修改当前选择的样式
    """
    try:
        # 获取Word应用程序的Selection对象
        app = document.Application
        selection = app.Selection
        
        # 检查是否有选择内容
        if selection.Start == selection.End:
            return json.dumps({
                "success": False,
                "message": "没有选择任何内容，无法修改样式"
            }, ensure_ascii=False)
        
        # 应用格式到选择范围
        result = apply_formatting_to_object(selection.Range, formatting)
        
        # 解析结果
        try:
            result_dict = json.loads(result)
            if result_dict.get("success", False):
                log_info(f"成功应用样式到选择范围")
                return json.dumps({
                    "success": True,
                    "message": "选择样式修改成功",
                    "applied_formatting": formatting
                }, ensure_ascii=False)
            else:
                log_error(f"应用样式失败: {result_dict.get('message', '未知错误')}")
                return result
        except json.JSONDecodeError:
            # 如果结果不是有效的JSON，假设成功
            log_info(f"成功应用样式到选择范围")
            return json.dumps({
                "success": True,
                "message": "选择样式修改成功",
                "applied_formatting": formatting
            }, ensure_ascii=False)
            
    except Exception as e:
        log_error(f"修改选择样式失败: {e}")
        raise WordDocumentError(ErrorCode.FORMATTING_ERROR, f"修改选择样式失败: {str(e)}")
