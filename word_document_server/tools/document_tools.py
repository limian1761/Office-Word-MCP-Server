"""
Document Integration Tool for Word Document MCP Server.

This module provides a unified MCP tool for document operations.
"""

# Standard library imports
import json
import os
from typing import Any, Dict, List, Optional

import win32com.client
from dotenv import load_dotenv
from mcp.server.fastmcp import Context
from mcp.server.session import ServerSession
from pydantic import Field

# Local imports
from word_document_server.mcp_service.core import mcp_server
from word_document_server.operations.document_ops import (
    close_document, create_document, get_document_structure, open_document,
    save_document)
from word_document_server.operations.others_ops import (protect_document,
                                                        unprotect_document)
from word_document_server.utils.app_context import AppContext
from word_document_server.mcp_service.core_utils import (ErrorCode,
                                                   WordDocumentError,
                                                   format_error_response,
                                                   log_error, log_info,
                                                   log_warning)

# 加载环境变量
try:
    load_dotenv()
except Exception as e:
    log_info("python-dotenv not installed, skipping .env file loading")


@mcp_server.tool()
def document_tools(
    ctx: Context[ServerSession, AppContext] = Field(description="Context object"),
    operation_type: Optional[str] = Field(
        default="open",
        description="Type of document operation: create, open, save, save_as, close, get_info, set_property, get_property, print, protect, unprotect",
    ),
    file_path: Optional[str] = Field(
        default=None,
        description="File path for document operations. Required for: open,save_as. Optional for: None.  Optional for: create, ",
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
    property_value: Optional[Any] = Field(
        default=None,
        description="Property value for set operation. Required for: set_property. Optional for: None",
    ),
    print_settings: Optional[Dict[str, Any]] = Field(
        default=None,
        description="Print settings for print operation. Required for: print. Optional for: None",
    ),
    protection_type: Optional[str] = Field(
        default=None,
        description="Protection type for protect operation. Required for: protect. Optional for: None",
    ),
    protection_password: Optional[str] = Field(
        default=None,
        description="Password for protect/unprotect operations. Required for: protect, unprotect. Optional for: None",
    ),
    password: Optional[str] = Field(
        default=None,
        description="Password for opening protected documents. Required for: open (when document is password protected). Optional for: None",
    ),
) -> Any:
    """Unified document operation tool.

    This tool provides a single interface for all document operations:
    - create: Create a new document
      * Required parameters: None
      * Optional parameters: template_path, file_path
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
    - get_info: Get document information
      * Required parameters: None
      * Optional parameters: None
    - set_property: Set document property
      * Required parameters: property_name, property_value
      * Optional parameters: document_properties
    - get_property: Get document property
      * Required parameters: property_name
      * Optional parameters: None
    - print: Print the document
      * Required parameters: None
      * Optional parameters: print_settings
    - protect: Protect the document
      * Required parameters: protection_type
      * Optional parameters: protection_password
    - unprotect: Unprotect the document
      * Required parameters: None
      * Optional parameters: protection_password

    Returns:
        Operation result based on the operation type
    """
    try:
        # 获取活动文档
        active_doc = ctx.request_context.lifespan_context.get_active_document()

        # 根据操作类型执行相应的操作
        if operation_type and operation_type.lower() == "create":
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
            if file_path and os.path.exists(file_path):
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
            save_document(doc, file_path)

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
                    "agent_guide_content": agent_guide_content,
                },
                ensure_ascii=False,
            )

        elif operation_type and operation_type.lower() == "open":
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

            # 尝试打开文档，添加错误处理
            max_retries = 3
            retry_count = 0
            doc = None

            while retry_count < max_retries and doc is None:
                try:
                    # 尝试使用word_app.Documents.Open打开文档
                    if password:
                        doc = word_app.Documents.Open(
                            FileName=file_path, PasswordDocument=password
                        )
                    else:
                        doc = word_app.Documents.Open(FileName=file_path)
                except AttributeError as e:
                    retry_count += 1
                    if retry_count >= max_retries:
                        log_error(
                            f"Failed to open document after {max_retries} retries: {str(e)}"
                        )
                        raise RuntimeError(
                            f"Failed to access Word Documents collection: {str(e)}"
                        )

                    # 尝试重新创建Word应用实例
                    log_warning(
                        f"Retrying document opening (attempt {retry_count}/{max_retries}) after AttributeError"
                    )
                    try:
                        # 释放当前实例并创建新实例
                        if word_app is not None:
                            word_app.Quit()
                        word_app = win32com.client.Dispatch("Word.Application")
                        ctx.request_context.lifespan_context._word_app = (
                            word_app  # 更新上下文
                        )
                    except Exception as inner_e:
                        log_error(
                            f"Failed to recreate Word application: {str(inner_e)}"
                        )
                except Exception as e:
                    # 处理其他异常
                    log_error(f"Error opening document: {str(e)}")
                    raise RuntimeError(f"Failed to open document: {str(e)}")

            # 更新上下文中的活动文档
            ctx.request_context.lifespan_context.set_active_document(doc)

            # 读取agent_guide.md文件内容
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
                        if len(agent_guide_content) == 10000:
                            agent_guide_content += (
                                "\n\n...文档内容过长，已从10000个字符处截断... "
                            )
            except Exception as e:
                log_error(f"Failed to read agent_guide.md: {e}")
                agent_guide_content = "无法读取agent_guide.md文件"

            # 返回文档对象的基本信息和agent_guide.md内容
            return json.dumps(
                {
                    "success": True,
                    "message": f"Document opened successfully: {file_path}",
                    "document": {
                        "name": doc.Name,
                        "path": file_path,
                        "full_name": doc.FullName,
                        "saved": doc.Saved,
                    },
                    "agent_guide": agent_guide_content,
                },
                ensure_ascii=False,
            )

        elif operation_type and operation_type.lower() == "save":
            if not active_doc:
                raise WordDocumentError(
                    ErrorCode.DOCUMENT_ERROR, "No active document found"
                )

            log_info("Saving document")
            result = save_document(active_doc)

            return json.dumps(
                {"success": result, "message": "Document saved successfully"},
                ensure_ascii=False,
            )

        elif operation_type and operation_type.lower() == "save_as":
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
                {"success": result, "message": f"Document saved as: {file_path}"},
                ensure_ascii=False,
            )

        elif operation_type and operation_type.lower() == "close":
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

        elif operation_type and operation_type.lower() == "get_info":
            if not active_doc:
                raise WordDocumentError(
                    ErrorCode.DOCUMENT_ERROR, "No active document found"
                )

            log_info("Getting document info")
            structure = get_document_structure(active_doc)

            return structure

        elif operation_type and operation_type.lower() == "set_property":
            raise NotImplementedError("set_property operation not implemented")

        elif operation_type and operation_type.lower() == "get_property":
            raise NotImplementedError("get_property operation not implemented")

        elif operation_type and operation_type.lower() == "print":
            raise NotImplementedError("print operation not implemented")

        elif operation_type and operation_type.lower() == "protect":
            if not active_doc:
                raise WordDocumentError(
                    ErrorCode.DOCUMENT_ERROR, "No active document found"
                )

            if protection_type is None:
                raise WordDocumentError(
                    ErrorCode.INVALID_INPUT,
                    "Protection type is required for 'protect' operation",
                )

            # Map protection types to match Word's constants
            protection_type_map = {
                "readonly": "readonly",
                "read_only": "readonly",
                "comments": "comments",
                "tracked_changes": "tracked_changes",
                "tracking": "tracked_changes",
                "forms": "forms",
            }

            mapped_protection_type = protection_type_map.get(protection_type.lower())
            if mapped_protection_type is None:
                raise WordDocumentError(
                    ErrorCode.INVALID_INPUT,
                    f"Invalid protection type: {protection_type}. Valid types are: readonly, comments, tracked_changes, forms",
                )

            result = protect_document(
                active_doc, protection_password or "", mapped_protection_type
            )

            return json.dumps(
                {"success": True, "message": "Document protected successfully"},
                ensure_ascii=False,
            )

        elif operation_type and operation_type.lower() == "unprotect":
            if not active_doc:
                raise WordDocumentError(
                    ErrorCode.DOCUMENT_ERROR, "No active document found"
                )

            result = unprotect_document(active_doc, protection_password or "")

            return json.dumps(
                {"success": True, "message": "Document unprotected successfully"},
                ensure_ascii=False,
            )

        else:
            raise ValueError(f"Unsupported operation type: {operation_type}")

    except Exception as e:
        log_error(f"Error in document_tools: {e}", exc_info=True)
        return format_error_response(e)
