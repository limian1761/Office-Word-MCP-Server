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
from ..selector.selector import SelectorEngine
from ..mcp_service.app_context import AppContext


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
    """
    Unified image operation tool.

    This tool provides a single interface for all image operations:
    - get_info: Get information about all images in the document
      * No required parameters
    - insert: Insert an image
      * Required parameters: image_path
      * Optional parameters: locator, position
    - add_caption: Add a caption to an image
      * Required parameters: caption_text
      * Optional parameters: locator, label
    - resize: Resize an image
      * Required parameters: width or height (at least one)
      * Optional parameters: locator
    - set_color_type: Set image color type
      * Required parameters: color_type
      * Optional parameters: locator

    Returns:
        Operation result based on the operation type
    """
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

            # 获取图片索引（如果有定位器，则需要先获取图片索引）
            image_index = 1  # 默认调整第一个图片
            if locator:
                # 在实际实现中，这里应该使用定位器来查找图片并获取其索引
                # 为了简化，我们暂时使用默认值
                pass

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

            # 获取图片索引（如果有定位器，则需要先获取图片索引）
            image_index = 1  # 默认调整第一个图片
            if locator:
                # 在实际实现中，这里应该使用定位器来查找图片并获取其索引
                # 为了简化，我们暂时使用默认值
                pass

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
