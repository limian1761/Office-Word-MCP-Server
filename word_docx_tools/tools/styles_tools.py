"""
Style Integration Tool for Word Document MCP Server.

This module provides a simplified tool for style operations, including querying available styles, list numbering styles and fonts.
"""

import json
from typing import List, Optional

# Standard library imports
from dotenv import load_dotenv
# Third-party imports
from mcp.server.fastmcp import Context
from mcp.server.session import ServerSession
from pydantic import Field

# Local imports
from ..mcp_service.core import mcp_server
from ..mcp_service.core_utils import (
    ErrorCode, WordDocumentError,
    format_error_response,
    handle_tool_errors,
    log_error, log_info,
    require_active_document_validation)
from ..mcp_service.app_context import AppContext


@mcp_server.tool()
@handle_tool_errors
@require_active_document_validation
def styles_tools(
    ctx: Context[ServerSession, AppContext] = Field(description="Context object"),
    operation_type: Optional[str] = Field(
        default=None,
        description="Type of style operation: get_available_styles, get_numbering_styles, get_font_names",
    ),
    style_type: Optional[str] = Field(
        default=None,
        description="Filter styles by type: 'paragraph', 'character', 'numbering', 'table', 'list' (only used with get_available_styles)",
    ),
    limit: Optional[int] = Field(
        default=None,
        description="Limit the number of results returned (only used with get_font_names)",
    ),
) -> str:
    """样式操作工具，支持查询文档中的可用样式、编号样式和字体名称。

    支持的操作类型：
    - get_available_styles: 获取文档中所有可用的样式
      * 必需参数：无
      * 可选参数：style_type - 过滤样式类型
    - get_numbering_styles: 获取文档中所有可用的编号样式
      * 必需参数：无
      * 可选参数：无
    - get_font_names: 获取系统中所有可用的字体名称
      * 必需参数：无
      * 可选参数：limit - 限制返回的字体数量

    返回：
        操作结果的JSON字符串
    """
    try:
        # 验证输入参数
        if not operation_type:
            raise ValueError("operation_type参数必须提供")

        # 获取活动文档
        active_doc = ctx.request_context.lifespan_context.get_active_document()
        if not active_doc:
            raise WordDocumentError(
                ErrorCode.DOCUMENT_ERROR, "未找到活动文档"
            )

        # 处理get_available_styles操作
        if operation_type and operation_type.lower() == "get_available_styles":
            log_info("获取可用样式")
            result = []
            try:
                # Word中样式类型常量
                STYLE_TYPES = {
                    'paragraph': 1,    # wdStyleTypeParagraph
                    'character': 2,    # wdStyleTypeCharacter
                    'table': 3,        # wdStyleTypeTable
                    'list': 4          # wdStyleTypeList
                }
                
                for style in active_doc.Styles:
                    # 如果指定了style_type，则过滤样式类型
                    if style_type and style_type.lower() in STYLE_TYPES:
                        if style.Type == STYLE_TYPES[style_type.lower()]:
                            result.append({
                                "name": style.NameLocal,
                                "type": style.Type,
                                "built_in": style.BuiltIn,
                                "type_name": style_type.lower()
                            })
                    else:
                        # 获取样式类型的名称
                        type_name = "unknown"
                        for key, value in STYLE_TYPES.items():
                            if value == style.Type:
                                type_name = key
                                break
                        
                        result.append({
                            "name": style.NameLocal,
                            "type": style.Type,
                            "built_in": style.BuiltIn,
                            "type_name": type_name
                        })
            except Exception as e:
                raise WordDocumentError(
                    ErrorCode.SERVER_ERROR, f"获取样式失败: {str(e)}"
                )

            return json.dumps(result, ensure_ascii=False)
        
        # 处理get_numbering_styles操作
        elif operation_type and operation_type.lower() == "get_numbering_styles":
            log_info("获取编号样式")
            result = []
            try:
                # Word中列表样式类型常量为4 (wdStyleTypeList)
                for style in active_doc.Styles:
                    if style.Type == 4:  # wdStyleTypeList
                        result.append({
                            "name": style.NameLocal,
                            "built_in": style.BuiltIn,
                            "is_numbering_style": True
                        })
                
                # 如果没有找到列表样式，尝试另一种方式识别编号样式
                if not result:
                    numbering_styles = []
                    for style in active_doc.Styles:
                        # 检查样式名称是否包含编号相关关键词
                        name_lower = style.NameLocal.lower()
                        if any(keyword in name_lower for keyword in 
                              ['number', '编号', 'list', '列表', 'bullet', '项目符号']):
                            numbering_styles.append({
                                "name": style.NameLocal,
                                "built_in": style.BuiltIn,
                                "is_numbering_style": True,
                                "original_type": style.Type
                            })
                    result = numbering_styles
            except Exception as e:
                raise WordDocumentError(
                    ErrorCode.SERVER_ERROR, f"获取编号样式失败: {str(e)}"
                )

            return json.dumps(result, ensure_ascii=False)
        
        # 处理get_font_names操作
        elif operation_type and operation_type.lower() == "get_font_names":
            log_info("获取字体名称")
            result = []
            try:
                if not hasattr(active_doc, "Application") or active_doc.Application is None:
                    raise WordDocumentError(
                        ErrorCode.SERVER_ERROR, "无法访问Application对象获取字体列表"
                    )
                
                # 获取所有可用字体名称
                font_names = list(active_doc.Application.FontNames)
                
                # 如果设置了限制，只返回指定数量的字体
                if limit and isinstance(limit, int) and limit > 0:
                    font_names = font_names[:limit]
                
                # 构建返回结果
                for font_name in font_names:
                    result.append({
                        "font_name": font_name
                    })
            except Exception as e:
                raise WordDocumentError(
                    ErrorCode.SERVER_ERROR, f"获取字体名称失败: {str(e)}"
                )

            return json.dumps(result, ensure_ascii=False)
        else:
            raise ValueError(f"不支持的操作类型: {operation_type}")

    except Exception as e:
        log_error(f"styles_tools中的错误: {e}", exc_info=True)
        # 确保错误响应可以正确序列化
        error_response = format_error_response(e)
        if isinstance(error_response, dict) and "error_code" in error_response:
            # 确保error_code是整数类型
            error_response["error_code"] = int(error_response["error_code"])
        # 如果已经是JSON字符串，则直接返回，否则转换为JSON字符串
        if isinstance(error_response, str):
            return error_response
        else:
            return json.dumps(error_response, ensure_ascii=False)
