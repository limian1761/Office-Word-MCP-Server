"""
Image Integration Tool for Word Document MCP Server.

This module provides a unified MCP tool for image operations.
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
from ..mcp_service.core_utils import (ErrorCode, WordDocumentError,
                                      format_error_response,
                                      get_active_document, handle_tool_errors,
                                      log_error, log_info,
                                      require_active_document_validation)

from ..mcp_service.app_context import AppContext

# Custom exception class to replace the one from selector.exceptions
class LocatorSyntaxError(Exception):
    """Exception raised for invalid locator syntax."""
    pass


# 延迟导入以避免循环导入
def _import_image_operations():
    """延迟导入image操作函数以避免循环导入"""
    from ..operations.image_ops import (add_caption, get_image_info,
                                        insert_image, resize_image,
                                        set_image_color_type)

    return (
        add_caption,
        get_image_info,
        insert_image,
        resize_image,
        set_image_color_type,
    )


# 加载环境变量
try:
    load_dotenv()
except Exception as e:
    log_info("python-dotenv not installed, skipping .env file loading")


@mcp_server.tool()
async def image_tools(
    ctx: Context[ServerSession, AppContext],
    operation_type: str = Field(
        ...,
        description="Type of image operation: get_info, insert, add_caption, resize, set_color_type",
    ),
    image_path: Optional[str] = Field(
        default=None,
        description="Path to the image file for insert operation. Required for: insert",
    ),
    width: Optional[float] = Field(
        default=None,
        description="Image width in points for resize operation. Required for: resize (if height not provided)",
    ),
    height: Optional[float] = Field(
        default=None,
        description="Image height in points for resize operation. Required for: resize (if width not provided)",
    ),
    color_type: Optional[str] = Field(
        default=None,
        description="Image color type for set_color_type operation. Required for: set_color_type",
    ),
    caption_text: Optional[str] = Field(
        default=None,
        description="Caption text for add_caption operation. Required for: add_caption",
    ),
    label: Optional[str] = Field(
        default=None, description="Caption label. Optional for: add_caption"
    ),
    locator: Optional[Dict[str, Any]] = Field(
        default=None,
        description="Object locator for specifying insertion position. Optional for: insert, add_caption, resize, set_color_type",
    ),
    position: Optional[str] = Field(
        default=None,
        description="Insertion position, options: 'before', 'after' for insert; 'above', 'below' for add_caption. Optional for: insert, add_caption",
    ),
    is_independent_paragraph: Optional[bool] = Field(
        default=True,
        description="Whether the inserted image or caption should be an independent paragraph. Optional for: insert, add_caption",
    ),
    exclude_label: Optional[bool] = Field(
        default=False,
        description="Whether to exclude the caption label when adding a caption. Optional for: add_caption",
    ),
) -> str:
    """图像操作工具

    支持的操作类型：
    - get_info: 获取文档中所有图像的信息
      * 必需参数：无
      * 可选参数：无
    - insert: 插入图像
      * 必需参数：image_path
      * 可选参数：locator, position
    - add_caption: 为图像添加说明文字
      * 必需参数：caption_text
      * 可选参数：locator, label
    - resize: 调整图像大小
      * 必需参数：width或height（至少提供一个）
      * 可选参数：locator
    - set_color_type: 设置图像颜色类型
      * 必需参数：color_type
      * 可选参数：locator

    返回：
        操作结果的JSON字符串
    """
    # 检查locator参数类型和规范
    def check_locator_param(locator_value):
        if locator_value is not None:
            # 检查是否为字典类型
            if not isinstance(locator_value, dict):
                raise TypeError("locator parameter must be a dictionary")
            
            # 简单验证locator结构
            if 'type' not in locator_value:
                raise LocatorSyntaxError("Locator must contain 'type' field")
            
            # 验证定位器类型是否有效
            valid_types = ['paragraph', 'table', 'image', 'selection']
            if locator_value['type'] not in valid_types:
                raise LocatorSyntaxError(f"Invalid locator type. Must be one of: {', '.join(valid_types)}")
            
            # 验证position参数（如果提供）
            if 'position' in locator_value and locator_value['position'] not in ['before', 'after', 'inside']:
                raise LocatorSyntaxError("Invalid position. Must be 'before', 'after', or 'inside'")
            
            # 验证index参数（如果提供）
            if 'index' in locator_value:
                if not isinstance(locator_value['index'], int) or locator_value['index'] < 1:
                    raise LocatorSyntaxError("Index must be a positive integer")
    
    # Get the active Word document from the context
    document = ctx.request_context.lifespan_context.get_active_document()

    # 延迟导入image操作函数以避免循环导入
    (add_caption, get_image_info, insert_image, resize_image, set_image_color_type) = (
        _import_image_operations()
    )

    try:
        if operation_type == "get_info":
            log_info("Getting image information")
            result = get_image_info(document)
            log_info(f"Retrieved information for {len(result) if result else 0} images")
            return json.dumps(
                {
                    "success": True,
                    "images": result,
                    "message": "Image information retrieved successfully",
                },
                ensure_ascii=False,
            )

        elif operation_type == "insert":
            if image_path is None:
                raise WordDocumentError(
                    ErrorCode.INVALID_INPUT,
                    "Image path is required for insert operation",
                )

            # 检查locator参数
            check_locator_param(locator)
            
            if locator is None:
                raise WordDocumentError(
                    ErrorCode.INVALID_INPUT,
                    "Locator is required for insert operation",
                )

            if not os.path.exists(image_path):
                raise WordDocumentError(
                    ErrorCode.NOT_FOUND, f"Image file not found: {image_path}"
                )

            log_info(f"Inserting image from path: {image_path}")

            # 调用operations模块中的insert_image函数
            result = insert_image(document, image_path, locator, position, is_independent_paragraph)

            # 解析结果并返回
            result_dict = json.loads(result)
            log_info("Image inserted successfully")
            return json.dumps(
                {
                    "success": True,
                    "result": {
                        "success": True,
                        "shape_id": result_dict.get("image_index", 0),
                    },
                    "message": "Image inserted successfully",
                },
                ensure_ascii=False,
            )

        elif operation_type == "add_caption":
            if caption_text is None:
                raise WordDocumentError(
                    ErrorCode.INVALID_INPUT,
                    "Caption text is required for add_caption operation",
                )
            # 检查locator参数
            check_locator_param(locator)
            if locator is None:
                raise WordDocumentError(
                    ErrorCode.INVALID_INPUT,
                    "Locator is required for add_caption operation",
                )

            log_info(f"Adding caption: {caption_text}")
            # 确定题注位置，如果未指定则默认为'below'
            caption_position = position if position in ['above', 'below'] else 'below'
            result = add_caption(document, caption_text, locator, label, caption_position, exclude_label, is_independent_paragraph)
            log_info("Caption added successfully")
            return json.dumps(
                {
                    "success": True,
                    "result": result,
                    "message": "Caption added successfully",
                },
                ensure_ascii=False,
            )

        elif operation_type == "resize":
            if width is None and height is None:
                raise WordDocumentError(
                    ErrorCode.INVALID_INPUT,
                    "Width or height is required for resize operation",
                )

            # 检查locator参数
            check_locator_param(locator)
            
            # 获取图片索引（如果有定位器，则需要先获取图片索引）
            image_index = 1  # 默认调整第一个图片
            if locator:
                # 使用AppContext获取Word应用程序对象
                word_app = ctx.request_context.lifespan_context.get_word_app()
                # 尝试根据定位器获取目标图片
                try:
                    # 简单的定位器实现，实际应用中可能需要更复杂的逻辑
                    if locator.get('type') == 'image' and 'index' in locator:
                        image_index = locator['index']
                    elif locator.get('type') == 'selection':
                        # 尝试获取选中的图片
                        selection = word_app.Selection
                        if selection.InlineShapes.Count > 0:
                            # 获取选中的内联形状（通常是图片）
                            for i, shape in enumerate(document.InlineShapes, 1):
                                if shape.Range.Start == selection.Range.Start:
                                    image_index = i
                                    break
                except Exception as e:
                    log_error(f"Failed to get image index from locator: {str(e)}")
                    # 如果获取失败，继续使用默认值

            log_info(f"Resizing image with width: {width}, height: {height}")
            # 调用resize_image函数，注意参数顺序
            result = resize_image(document, image_index, width, height)
            log_info("Image resized successfully")
            return json.dumps(
                {
                    "success": True,
                    "result": result,
                    "message": "Image resized successfully",
                },
                ensure_ascii=False,
            )

        elif operation_type == "set_color_type":
            if color_type is None:
                raise WordDocumentError(
                    ErrorCode.INVALID_INPUT,
                    "Color type is required for set_color_type operation",
                )

            # 检查locator参数
            check_locator_param(locator)
            
            # 获取图片索引（如果有定位器，则需要先获取图片索引）
            image_index = 1  # 默认调整第一个图片
            if locator:
                # 使用AppContext获取Word应用程序对象
                word_app = ctx.request_context.lifespan_context.get_word_app()
                # 尝试根据定位器获取目标图片
                try:
                    # 简单的定位器实现，实际应用中可能需要更复杂的逻辑
                    if locator.get('type') == 'image' and 'index' in locator:
                        image_index = locator['index']
                    elif locator.get('type') == 'selection':
                        # 尝试获取选中的图片
                        selection = word_app.Selection
                        if selection.InlineShapes.Count > 0:
                            # 获取选中的内联形状（通常是图片）
                            for i, shape in enumerate(document.InlineShapes, 1):
                                if shape.Range.Start == selection.Range.Start:
                                    image_index = i
                                    break
                except Exception as e:
                    log_error(f"Failed to get image index from locator: {str(e)}")
                    # 如果获取失败，继续使用默认值

            log_info(f"Setting image color type to: {color_type}")
            # 调用set_image_color_type函数，注意参数顺序
            result = set_image_color_type(document, image_index, color_type)
            log_info("Image color type set successfully")
            return json.dumps(
                {
                    "success": True,
                    "result": result,
                    "message": "Image color type set successfully",
                },
                ensure_ascii=False,
            )

        else:
            error_msg = f"Unknown operation type: {operation_type}"
            log_error(error_msg)
            raise WordDocumentError(ErrorCode.INVALID_INPUT, error_msg)

    except Exception as e:
        log_error(f"Image operation failed: {str(e)}")
        raise WordDocumentError(
            ErrorCode.SERVER_ERROR, f"Image operation failed: {str(e)}"
        )
