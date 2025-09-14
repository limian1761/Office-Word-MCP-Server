"""
导航工具模块，用于Word文档的上下文和活动对象管理。

此模块提供了设置上下文和活动对象的功能，是上下文控制功能的精简版本。
"""
import os
import logging
from typing import Dict, Any, Optional

# 标准库导入
from dotenv import load_dotenv
# 第三方导入
from mcp.server.fastmcp import Context
from mcp.server.session import ServerSession
from pydantic import Field

# 本地导入
from ..mcp_service.core import mcp_server
from ..mcp_service.app_context import AppContext
from ..mcp_service.core_utils import (
    ErrorCode,
    WordDocumentError,
    format_error_response,
    get_active_document,
    handle_tool_errors,
    log_error,
    log_info,
    require_active_document_validation
)
from ..operations.navigate_tools import set_active_context, set_active_object

# 加载.env文件中的环境变量
load_dotenv()

logger = logging.getLogger(__name__)


@mcp_server.tool()
async def navigate_tools(
    ctx: Context[ServerSession, AppContext] = Field(description="上下文对象"),
    operation_type: str = Field(
        ...,
        description="导航工具操作类型: set_active_context, set_active_object",
    ),
    context_type: Optional[str] = Field(
        default=None,
        description="上下文类型 (section, paragraph, table, image, comment, bookmark)，set_active_context操作必需",
    ),
    context_id: Optional[str] = Field(
        default=None,
        description="上下文ID，set_active_context操作必需",
    ),
    object_type: Optional[str] = Field(
        default=None,
        description="对象类型 (paragraph, table, image, comment, bookmark)，set_active_object操作必需",
    ),
    object_id: Optional[str] = Field(
        default=None,
        description="对象ID，set_active_object操作必需",
    ),
    params: Optional[Dict[str, Any]] = Field(
        default=None,
        description="用于测试兼容性的参数"
    )
) -> Dict[str, Any]:
    """导航工具

    支持的操作类型：
    - set_active_context: 设置活动上下文
      * 必需参数：context_type, context_id
    - set_active_object: 设置活动对象
      * 必需参数：object_type, object_id
    """
    try:
        # 处理params参数，兼容测试用例
        if params:
            context_type = params.get('context_type', context_type)
            context_id = params.get('context_id', context_id)
            object_type = params.get('object_type', object_type)
            object_id = params.get('object_id', object_id)
            operation_type = params.get('operation_type', operation_type)
        
        # 获取活动文档
        document = get_active_document(ctx)
        
        if not document:
            raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "没有活动文档")
        
        # 根据操作类型执行相应的操作
        if operation_type == 'set_active_context':
            # 验证必需参数
            if not context_type or not context_id:
                raise WordDocumentError(
                    ErrorCode.INVALID_INPUT,
                    "set_active_context操作需要context_type和context_id参数"
                )
            
            # 调用操作层函数
            result = set_active_context(document, context_type, context_id)
            log_info(f"成功设置活动上下文: {context_type} {context_id}")
            return result
            
        elif operation_type == 'set_active_object':
            # 验证必需参数
            if not object_type or not object_id:
                raise WordDocumentError(
                    ErrorCode.INVALID_INPUT,
                    "set_active_object操作需要object_type和object_id参数"
                )
            
            # 调用操作层函数
            result = set_active_object(document, object_type, object_id)
            log_info(f"成功设置活动对象: {object_type} {object_id}")
            return result
            
        else:
            raise WordDocumentError(
                ErrorCode.INVALID_INPUT,
                f"不支持的操作类型: {operation_type}，支持的类型为: set_active_context, set_active_object"
            )
            
    except WordDocumentError as e:
        log_error(f"导航工具错误: {str(e)}")
        return format_error_response(e.code, str(e))
    except Exception as e:
        log_error(f"导航工具未预期错误: {str(e)}")
        return format_error_response(ErrorCode.SERVER_ERROR, f"服务器错误: {str(e)}")