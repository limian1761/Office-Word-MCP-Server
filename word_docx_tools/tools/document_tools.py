"""
Document Integration Tool for Word Document MCP Server.

This module provides a unified MCP tool for document operations.
"""

# Standard library imports
import json
import os
from typing import Any, Dict, List, Optional, Union

import win32com.client
from dotenv import load_dotenv
from mcp.server.fastmcp import Context
from mcp.server.session import ServerSession
from pydantic import Field

# Local imports
from ..mcp_service.core import mcp_server
from ..mcp_service.core_utils import (
    ErrorCode, WordDocumentError,
    format_error_response,
    get_active_document, handle_tool_errors,
    log_error, log_info, log_warning,
    require_active_document_validation
)
from ..operations.document_ops import (
    close_document, create_document,
    get_document_outline, open_document,
    save_document
)
from ..mcp_service.app_context import AppContext

# 加载环境变量
try:
    load_dotenv()
except Exception as e:
    log_info("python-dotenv not installed, skipping .env file loading")


@mcp_server.tool()
@handle_tool_errors
def document_tools(
    ctx: Context[ServerSession, AppContext] = Field(description="Context object"),
    operation_type: Optional[str] = Field(
        default="open",
        description="Type of document operation: create, open, save, save_as, close, get_outline, set_property, get_property",
    ),
    file_path: Optional[str] = Field(
        default=None,
        description="File path for document operations. Required for: open,save_as, create. Optional for: None",
    ),
    template_path: Optional[str] = Field(
        default=None,
        description="Template path for create operation. Required for: None. Optional for: create",
    ),
    document_properties: Optional[Dict[str, Any]] = Field(
        default=None,
        description="Document properties for set_property operation. Required for: set_property. Optional for: None",
    ),
    property_name: Optional[str] = Field(
        default=None,
        description="Property name for get/set operations. Required for: get_property, set_property. Optional for: None",
    ),
    property_value: Optional[Union[str, int, float, bool]] = Field(
        default=None,
        description="Property value for set operation. Required for: set_property. Optional for: None",
    ),
    password: Optional[str] = Field(
        default=None,
        description="Password for opening protected documents. Required for: open (when document is password protected). Optional for: None",
    ),
) -> Any:
    """Unified document operation tool.

    This tool provides a single interface for all document operations:
    - create: Create a new document
      * Required parameters: file_path
      * Optional parameters: template_path,
    - open: Open an existing document
      * Required parameters: file_path
      * Optional parameters: password
    - save: Save the current document
      * Required parameters: None
      * Optional parameters: None
    - save_as: Save the current document to a new path
      * Required parameters: file_path
      * Optional parameters: None
    - close: Close the current document
      * Required parameters: None
      * Optional parameters: None
    - get_outline: Get document outline
      * Required parameters: None
      * Optional parameters: None
    - set_property: Set document property
      * Required parameters: property_name, property_value
      * Optional parameters: document_properties
    - get_property: Get document property
      * Required parameters: property_name
      * Optional parameters: Nonecreate

    Returns:
        Operation result based on the operation type
    """
    try:
        # 获取活动文档
        active_doc = ctx.request_context.lifespan_context.get_active_document()

        # 根据操作类型执行相应的操作
        if operation_type:
            # 确保operation_type是字符串类型
            # 处理operation_type可能是字典的情况
            if isinstance(operation_type, dict):
                # 从字典中提取operation_type值
                if 'operation_type' in operation_type:
                    operation_type_str = str(operation_type['operation_type']).lower()
                    # 如果params中包含其他参数，也提取出来
                    if 'params' in operation_type:
                        params = operation_type['params']
                        if 'template_path' in params:
                            template_path = params['template_path']
                        if 'file_path' in params:
                            file_path = params['file_path']
                else:
                    operation_type_str = str(operation_type).lower()
            else:
                operation_type_str = str(operation_type).lower()
                
            # 处理特殊的操作类型映射
            if operation_type_str == "create_document":
                operation_type_str = "create"
            elif operation_type_str == "save_document":
                operation_type_str = "save_as"
            elif operation_type_str == "open_document":
                operation_type_str = "open"
                
            if operation_type_str == "create":
                log_info("Creating new document")
                # 创建新文档的逻辑
                word_app = ctx.request_context.lifespan_context.get_word_app(
                    create_if_needed=True
                )
                if word_app is None:
                    log_error("Failed to get or create Word application instance")
                    raise RuntimeError("Failed to get or create Word application instance")
                doc = create_document(word_app, visible=True, template_path=template_path)

                # 更新上下文中的活动文档
                ctx.request_context.lifespan_context.set_active_document(doc)

                # 检查文件是否已存在
                if file_path and isinstance(file_path, (str, bytes, os.PathLike)) and os.path.exists(file_path):
                    # 文件已存在，返回友好的错误信息
                    return json.dumps(
                        {
                            "success": False,
                            "message": f"文件已存在: {file_path}",
                            "error_code": "FILE_ALREADY_EXISTS",
                        },
                        ensure_ascii=False,
                    )

                # 保存新文档
                # 确保file_path是字符串类型或None
                save_path = str(file_path) if file_path and not isinstance(file_path, (str, bytes, os.PathLike)) else file_path
                save_document(doc, save_path)

                # 返回agent_guide.md文件内容
                agent_guide_path = os.path.join(
                    os.path.dirname(os.path.dirname(os.path.dirname(__file__))),
                    "docs",
                    "agent_guide.md",
                )
                agent_guide_content = ""
                try:
                    if os.path.exists(agent_guide_path):
                        with open(agent_guide_path, "r", encoding="utf-8") as f:
                            # 由于文件可能很大，只读取前10000个字符
                            agent_guide_content = f.read(10000)
                except Exception as e:
                    log_error(f"Failed to read agent_guide.md: {e}")
                    agent_guide_content = ""

                return json.dumps(
                    {
                        "success": True,
                        "message": "New document created successfully",
                        "document_name": doc.Name,
                        "document_created": True,
                        "agent_guide_content": agent_guide_content,
                    },
                    ensure_ascii=False,
                )

            elif operation_type_str == "open":
                if file_path is None:
                    raise ValueError(
                        "file_path parameter must be provided for open operation"
                    )

                log_info(f"Opening document: {file_path}")
                # 获取Word应用实例
                word_app = ctx.request_context.lifespan_context.get_word_app(
                    create_if_needed=True
                )
                if word_app is None:
                    raise RuntimeError("Failed to get or create Word application instance")

                # 使用document_ops中的open_document函数
                doc = open_document(word_app, file_path, visible=True, password=password)

                # 更新上下文中的活动文档
                ctx.request_context.lifespan_context.set_active_document(doc)

                # 返回文档对象的基本信息
                return json.dumps(
                    {
                        "success": True,
                        "message": f"Document opened successfully: {file_path}",
                        "document_opened": True,
                        "document": {
                            "name": doc.Name,
                            "path": file_path,
                            "full_name": doc.FullName,
                            "saved": doc.Saved,
                        },
                    },
                    ensure_ascii=False,
                )

            elif operation_type_str == "save":
                if not active_doc:
                    raise WordDocumentError(
                        ErrorCode.DOCUMENT_ERROR, "No active document found"
                    )

                log_info("Saving document")
                result = save_document(active_doc)

                return json.dumps(
                    {"success": result, "message": "Document saved successfully", "document_saved": True},
                    ensure_ascii=False,
                )

            elif operation_type_str == "save_as":
                if not active_doc:
                    raise WordDocumentError(
                        ErrorCode.DOCUMENT_ERROR, "No active document found"
                    )
                if file_path is None:
                    raise ValueError(
                        "file_path parameter must be provided for save_as operation"
                    )

                log_info(f"Saving document as: {file_path}")
                result = save_document(active_doc, file_path)

                return json.dumps(
                    {"success": result, "message": f"Document saved as: {file_path}", "document_saved": True},
                    ensure_ascii=False,
                )

            elif operation_type_str == "close":
                if not active_doc:
                    raise WordDocumentError(
                        ErrorCode.DOCUMENT_ERROR, "No active document found"
                    )

                log_info("Closing document")
                result = close_document(active_doc)

                # 清除上下文中的活动文档
                ctx.request_context.lifespan_context.set_active_document(None)

                return json.dumps(
                    {"success": result, "message": "Document closed successfully"},
                    ensure_ascii=False,
                )
            elif operation_type_str == "get_outline":
                if not active_doc:
                    raise WordDocumentError(
                        ErrorCode.DOCUMENT_ERROR, "No active document found"
                    )

                log_info("Getting document outline")
                outline = get_document_outline(active_doc)

                return outline

            elif operation_type_str == "set_property":
                if not active_doc:
                    raise WordDocumentError(
                        ErrorCode.DOCUMENT_ERROR, "No active document found"
                    )

                if property_name is None or property_value is None:
                    raise ValueError(
                        "property_name and property_value parameters must be provided for set_property operation"
                    )

                log_info(f"Setting document property: {property_name}")
                
                # 属性名称映射字典，用于支持不同语言的属性名
                property_name_map = {
                    # 英文到英文的映射
                    "Title": "Title",
                    "Subject": "Subject",
                    "Author": "Author",
                    "Keywords": "Keywords",
                    "Comments": "Comments",
                    "Template": "Template",
                    "Last Author": "Last Author",
                    "Revision Number": "Revision Number",
                    "Application Name": "Application Name",
                    "Last Print Date": "Last Print Date",
                    "Creation Date": "Creation Date",
                    "Last Save Time": "Last Save Time",
                    "Total Editing Time": "Total Editing Time",
                    "Number of Pages": "Number of Pages",
                    "Number of Words": "Number of Words",
                    "Number of Characters": "Number of Characters",
                    "Security": "Security",
                }

                # 获取标准化的属性名称
                standard_property_name = property_name_map.get(
                    property_name, property_name
                )

                # 尝试设置文档内置属性
                try:
                    if active_doc is None:
                        raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")
                    # Word文档属性需要通过名称访问，而不是作为对象的属性
                    property_obj = active_doc.BuiltInDocumentProperties(
                        standard_property_name
                    )
                    property_obj.Value = property_value
                    return json.dumps(
                        {
                            "success": True,
                            "property_name": property_name,
                            "standard_property_name": standard_property_name,
                            "property_value": property_value,
                            "is_built_in": True,
                        },
                        ensure_ascii=False,
                    )
                except Exception as e:
                    # 如果内置属性访问失败，尝试检查自定义属性
                    try:
                        custom_properties = active_doc.CustomDocumentProperties
                        # 检查属性是否已存在
                        prop_exists = False
                        for i in range(1, custom_properties.Count + 1):
                            if custom_properties(i).Name == property_name:
                                custom_properties(i).Value = property_value
                                prop_exists = True
                                break

                        # 如果不存在，则添加新的自定义属性
                        if not prop_exists:
                            # 对于自定义属性，需要先检查属性类型
                            property_type = 4  # 默认设置为文本类型
                            if isinstance(property_value, bool):
                                property_type = 1  # 布尔类型
                            elif isinstance(property_value, int):
                                property_type = 2  # 整数类型
                            elif isinstance(property_value, float):
                                property_type = 3  # 浮点数类型

                            # 添加新的自定义属性
                            custom_properties.Add(
                                Name=property_name,
                                Type=property_type,
                                Value=property_value,
                            )

                        return json.dumps(
                            {
                                "success": True,
                                "property_name": property_name,
                                "property_value": property_value,
                                "is_custom_property": True,
                            },
                            ensure_ascii=False,
                        )
                    except Exception as inner_e:
                        raise WordDocumentError(
                            ErrorCode.SERVER_ERROR,
                            f"Failed to set property: {str(inner_e)}",
                        )
                except Exception as e:
                    # 更友好的错误处理
                    error_message = str(e)
                    if "Property not found" in error_message:
                        supported_properties = ", ".join(list(property_name_map.keys()))
                        raise WordDocumentError(
                            ErrorCode.NOT_FOUND,
                            f"Property not found: {property_name}. Supported built-in properties: {supported_properties}",
                        )
                    else:
                        raise WordDocumentError(
                            ErrorCode.SERVER_ERROR, f"Failed to set property: {str(e)}"
                        )

            elif operation_type_str == "get_property":
                if property_name is None:
                    raise ValueError(
                        "property_name parameter must be provided for get_property operation"
                    )

                log_info(f"Getting document property: {property_name}")
                try:
                    if active_doc is None:
                        raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")
                    # 尝试获取文档内置属性
                    try:
                        # Word文档属性需要通过名称访问，而不是作为对象的属性
                        property_obj = active_doc.BuiltInDocumentProperties(property_name)
                        value = property_obj.Value
                        return json.dumps(
                            {
                                "success": True,
                                "property_name": property_name,
                                "value": value,
                                "is_built_in": True,
                            },
                            ensure_ascii=False,
                        )
                    except Exception as e:
                        # 如果内置属性访问失败，尝试检查自定义属性
                        try:
                            custom_properties = active_doc.CustomDocumentProperties
                            # 遍历自定义属性查找指定名称的属性
                            value = None
                            for i in range(1, custom_properties.Count + 1):
                                if custom_properties(i).Name == property_name:
                                    value = custom_properties(i).Value
                                    return json.dumps(
                                        {
                                            "success": True,
                                            "property_name": property_name,
                                            "value": value,
                                            "is_custom_property": True,
                                        },
                                        ensure_ascii=False,
                                    )

                            # 如果未找到属性，返回None值
                            return json.dumps(
                                {
                                    "success": True,
                                    "property_name": property_name,
                                    "value": None,
                                    "message": "Property not found",
                                },
                                ensure_ascii=False,
                            )
                        except Exception as inner_e:
                            raise WordDocumentError(
                                ErrorCode.SERVER_ERROR,
                                f"Failed to get property: {str(inner_e)}",
                            )
                except Exception as e:
                    raise WordDocumentError(
                        ErrorCode.SERVER_ERROR, f"Failed to get property: {str(e)}"
                    )

            else:
                raise ValueError(f"Unsupported operation type: {operation_type}")
        else:
            raise ValueError("operation_type parameter is required")

    except Exception as e:
        log_error(f"Error in document_tools: {e}", exc_info=True)
        return format_error_response(e)
