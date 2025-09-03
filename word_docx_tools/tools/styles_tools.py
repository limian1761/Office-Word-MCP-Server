"""
Style Integration Tool for Word Document MCP Server.

This module provides a unified MCP tool for style operations.
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
from ..operations.styles_ops import (
    apply_formatting, set_font, set_paragraph_alignment, set_paragraph_style, set_paragraph_formatting)
from ..selector.selector import SelectorEngine
from ..utils.app_context import AppContext
from ..mcp_service.core_utils import (
    ErrorCode, WordDocumentError, format_error_response, get_active_document,
    handle_tool_errors, log_error, log_info,
    require_active_document_validation)


@mcp_server.tool()
@handle_tool_errors
@require_active_document_validation
def styles_tools(
    ctx: Context[ServerSession, AppContext] = Field(description="Context object"),
    operation_type: Optional[str] = Field(
        default=None,
        description="Type of style operation: apply_formatting, set_font, set_paragraph_style, set_alignment, set_paragraph_formatting",
    ),
    formatting: Optional[Dict[str, Any]] = Field(
        default=None,
        description="Formatting parameters dictionary containing various formatting options. Required for: apply_formatting",
    ),
    font_name: Optional[str] = Field(
        default=None, description="Font name. Optional for: set_font"
    ),
    font_size: Optional[float] = Field(
        default=None, description="Font size. Optional for: set_font"
    ),
    bold: Optional[bool] = Field(
        default=None, description="Whether bold. Optional for: set_font"
    ),
    italic: Optional[bool] = Field(
        default=None, description="Whether italic. Optional for: set_font"
    ),
    underline: Optional[str] = Field(
        default=None,
        description="Underline type, options: 'none', 'single', 'double', 'dotted', 'dashed', 'wave. Optional for: set_font",
    ),
    color: Optional[str] = Field(
        default=None, description="Font color. Optional for: set_font"
    ),
    style_name: Optional[str] = Field(
        default=None,
        description="Paragraph style name. Required for: set_paragraph_style, create_style",
    ),
    alignment: Optional[str] = Field(
        default=None,
        description="Alignment, options: 'left', 'center', 'right', 'justify'. Required for: set_alignment. Optional for: set_paragraph_formatting",
    ),
    line_spacing: Optional[float] = Field(
        default=None, description="Line spacing. Optional for: set_paragraph_formatting"
    ),
    space_before: Optional[float] = Field(
        default=None,
        description="Space before paragraph. Optional for: set_paragraph_formatting",
    ),
    space_after: Optional[float] = Field(
        default=None,
        description="Space after paragraph. Optional for: set_paragraph_formatting",
    ),
    first_line_indent: Optional[float] = Field(
        default=None,
        description="First line indent. Optional for: set_paragraph_formatting",
    ),
    left_indent: Optional[float] = Field(
        default=None, description="Left indent. Optional for: set_paragraph_formatting"
    ),
    right_indent: Optional[float] = Field(
        default=None, description="Right indent. Optional for: set_paragraph_formatting"
    ),
    locator: Optional[Dict[str, Any]] = Field(
        default=None,
        description="Object locator for specifying objects to apply styles to. Required for: apply_formatting, set_font, set_paragraph_style, set_alignment, set_paragraph_formatting",
    ),
) -> str:
    """Unified style operation tool.

    This tool provides a single interface for all style operations:
    - apply_formatting: Apply text formatting
      * Required parameters: formatting, locator
      * Optional parameters: None
    - set_font: Set text font properties
      * Required parameters: locator
      * Optional parameters: font_name, font_size, bold, italic, underline, color
    - set_paragraph_style: Set paragraph style
      * Required parameters: style_name, locator
      * Optional parameters: None
    - set_alignment: Set paragraph alignment
      * Required parameters: alignment, locator
      * Optional parameters: None
    - set_paragraph_formatting: Set paragraph formatting
      * Required parameters: locator
      * Optional parameters: alignment, line_spacing, space_before, space_after, first_line_indent, left_indent, right_indent
    - get_available_styles: Get available styles
      * Required parameters: None
      * Optional parameters: None
    - create_style: Create a new style
      * Required parameters: style_name
      * Optional parameters: None

    Returns:
        Operation result based on the operation type
    """
    try:
        # 获取活动文档
        active_doc = ctx.request_context.lifespan_context.get_active_document()
        if not active_doc:
            raise WordDocumentError(
                ErrorCode.DOCUMENT_ERROR, "No active document found"
            )

        # 根据操作类型执行相应的操作
        if operation_type and operation_type.lower() == "apply_formatting":
            if formatting is None or locator is None:
                raise ValueError(
                    "formatting and locator parameters must be provided for apply_formatting operation"
                )

            log_info("Applying formatting")
            result = apply_formatting(
                document=active_doc, formatting=formatting, locator=locator
            )
            return str(result)

        elif operation_type and operation_type.lower() == "set_font":
            if locator is None:
                raise ValueError(
                    "locator parameter must be provided for set_font operation"
                )

            log_info("Setting font")
            set_font(
                document=active_doc,
                font_name=font_name,
                font_size=font_size,
                bold=bold,
                italic=italic,
                underline=underline,
                color=color,
                locator=locator,
            )
            return json.dumps(
                {"success": True, "message": "Font settings applied successfully"},
                ensure_ascii=False,
            )

        elif operation_type and operation_type.lower() == "set_paragraph_style":
            if style_name is None or locator is None:
                raise ValueError(
                    "style_name and locator parameters must be provided for set_paragraph_style operation"
                )

            log_info(f"Setting paragraph style: {style_name}")
            try:
                # 修复段落样式设置逻辑 - 正确处理Selection对象
                engine = SelectorEngine()
                selection = engine.select(active_doc, locator)
                if not selection:
                    raise WordDocumentError(ErrorCode.SELECTOR_ERROR, "No paragraphs found for the given locator")
                
                # 检查selection是否有_com_ranges属性，如果有则使用它
                if hasattr(selection, "_com_ranges") and selection._com_ranges:
                    paragraphs = selection._com_ranges
                else:
                    # 如果是单个段落对象，包装成列表
                    try:
                        # 尝试迭代selection，如果成功则直接使用
                        iter(selection)
                        paragraphs = selection
                    except TypeError:
                        # 如果不能迭代，则包装成列表
                        paragraphs = [selection]
                
                for para in paragraphs:
                    # 确保样式名称正确并处理可能的异常
                    try:
                        para.Style = active_doc.Styles(style_name)
                    except Exception as style_err:
                        # 尝试使用NameLocal属性
                        found = False
                        for s in active_doc.Styles:
                            try:
                                if s.NameLocal == style_name:
                                    para.Style = s
                                    found = True
                                    break
                            except Exception:
                                continue
                        if not found:
                            raise WordDocumentError(ErrorCode.STYLE_NOT_FOUND, f"Style '{style_name}' not found")
                
                return json.dumps({"success": True, "message": f"Successfully applied style '{style_name}'"}, ensure_ascii=False)
            except WordDocumentError as e:
                if e.error_code == ErrorCode.STYLE_NOT_FOUND:
                    # 获取所有可用样式
                    available_styles = []
                    for style in active_doc.Styles:
                        try:
                            if style.InUse:
                                available_styles.append(style.NameLocal)
                        except Exception as ex:
                            log_error(f"Failed to get style name: {ex}")
                    return json.dumps({
                        "success": False,
                        "error_code": e.error_code,
                        "error_message": str(e),
                        "available_styles": available_styles
                    }, ensure_ascii=False)
                raise

        elif operation_type and operation_type.lower() == "set_paragraph_formatting":
            if locator is None:
                raise ValueError(
                    "locator parameter must be provided for set_paragraph_formatting operation"
                )

            log_info("Setting paragraph formatting")
            result = set_paragraph_formatting(
                document=active_doc,
                alignment=alignment,
                line_spacing=line_spacing,
                space_before=space_before,
                space_after=space_after,
                first_line_indent=first_line_indent,
                left_indent=left_indent,
                right_indent=right_indent,
                locator=locator
            )
            return str(result)

        elif operation_type and operation_type.lower() == "set_alignment":
            if alignment is None or locator is None:
                raise ValueError(
                    "alignment and locator parameters must be provided for set_alignment operation"
                )

            log_info("Setting paragraph alignment")
            result = set_paragraph_alignment(
                document=active_doc, alignment=alignment, locator=locator
            )
            return str(result)

        elif operation_type and operation_type.lower() == "get_available_styles":
            log_info("Getting available styles")
            result = []
            try:
                for style in active_doc.Styles:
                    result.append(
                        {
                            "name": style.NameLocal,
                            "type": style.Type,
                            "built_in": style.BuiltIn,
                        }
                    )
            except Exception as e:
                raise WordDocumentError(
                    ErrorCode.SERVER_ERROR, f"Failed to get styles: {str(e)}"
                )

            return json.dumps(result, ensure_ascii=False)

        elif operation_type and operation_type.lower() == "create_style":
            if style_name is None:
                raise ValueError(
                    "style_name parameter must be provided for create_style operation"
                )

            log_info(f"Creating style: {style_name}")
            # 直接在文档中创建样式
            try:
                # 检查样式是否已存在
                style_exists = False
                for style in active_doc.Styles:
                    if style.NameLocal == style_name:
                        style_exists = True
                        break

                if not style_exists:
                    # 创建新样式
                    new_style = active_doc.Styles.Add(
                        Name=style_name, Type=1
                    )  # 1 = Paragraph style
                    result = {
                        "success": True,
                        "message": f"Style '{style_name}' created successfully",
                        "style_name": style_name,
                    }
                else:
                    result = {
                        "success": False,
                        "message": f"Style '{style_name}' already exists",
                        "style_name": style_name,
                    }
            except Exception as e:
                raise WordDocumentError(
                    ErrorCode.SERVER_ERROR, f"Failed to create style: {str(e)}"
                )

            return json.dumps(result, ensure_ascii=False)

        else:
            raise ValueError(f"Unsupported operation type: {operation_type}")

    except Exception as e:
        log_error(f"Error in styles_tools: {e}", exc_info=True)
        return str(format_error_response(e))
