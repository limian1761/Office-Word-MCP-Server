"""
Image operations for Word Document MCP Server.
This module contains functions for image-related operations.
"""

import json
import logging
import os
from typing import Any, Dict, List, Optional

import win32com.client

from ..com_backend.com_utils import handle_com_error
from ..selector.selector import SelectorEngine
from ..mcp_service.core_utils import (ErrorCode, WordDocumentError, log_error,
                                log_info)

logger = logging.getLogger(__name__)


def get_image_info(document: win32com.client.CDispatch) -> List[Dict[str, Any]]:
    """获取文档中所有图片的信息

    Args:
        document: Word文档COM对象

    Returns:
        包含所有图片信息的列表

    Raises:
        WordDocumentError: 当获取图片信息失败时抛出
    """
    try:
        if not document:
            raise WordDocumentError(
                ErrorCode.DOCUMENT_ERROR, "No active document found"
            )

        # 检查文档是否包含所需属性
        if not hasattr(document, 'InlineShapes'):
             raise WordDocumentError(
                ErrorCode.DOCUMENT_ERROR,
                "Document object missing required 'InlineShapes' property"
            )
        if not hasattr(document, 'Shapes'):
            raise WordDocumentError(
                ErrorCode.DOCUMENT_ERROR,
                "Document object missing required 'Shapes' property"
            )

        image_info_list = []

        # 获取所有内嵌图片
        inline_shapes = document.InlineShapes
        for i in range(1, inline_shapes.Count + 1):
            try:
                shape = inline_shapes(i)
                image_info = {
                    "index": i,
                    "type": "InlineShape",
                    "width": shape.Width,
                    "height": shape.Height,
                    "name": getattr(shape, "Name", f"Image_{i}"),
                    "range_start": shape.Range.Start,
                    "range_end": shape.Range.End,
                    "has_picture": shape.Type == 1,  # wdInlineShapePicture
                    "position": "inline",
                }

                # 获取更多属性（如果可用）
                if shape.Type == 1 and hasattr(shape, "PictureFormat"):
                    image_info["format"] = "Picture"
                    if hasattr(shape.PictureFormat, "FileSize"):
                        image_info["file_size"] = shape.PictureFormat.FileSize

                image_info_list.append(image_info)
            except Exception as e:
                log_error(f"Error processing inline shape {i}: {e}")
                continue

        # 获取所有浮动图片
        shapes = document.Shapes
        for i in range(1, shapes.Count + 1):
            try:
                shape = shapes(i)
                if (
                    shape.Type == 1 or shape.Type == 13
                ):  # wdShapePicture or wdShapeLinkedPicture
                    image_info = {
                        "index": len(image_info_list) + 1,
                        "type": "Shape",
                        "width": shape.Width,
                        "height": shape.Height,
                        "name": getattr(shape, "Name", f"FloatingImage_{i}"),
                        "left": shape.Left,
                        "top": shape.Top,
                        "position": "floating",
                    }

                    if hasattr(shape, "LinkFormat") and shape.LinkFormat.SourceFullName:
                        image_info["is_linked"] = True
                        image_info["source_path"] = shape.LinkFormat.SourceFullName
                    else:
                        image_info["is_linked"] = False

                    image_info_list.append(image_info)
            except Exception as e:
                log_error(f"Error processing shape {i}: {e}")
                continue

        return image_info_list
    except Exception as e:
        log_error(f"Failed to get image info: {e}")
        raise WordDocumentError(
            ErrorCode.DOCUMENT_ERROR, f"Failed to get image info: {str(e)}"
        )


def set_picture_object_color_type(object: Any, color_type: int) -> bool:
    """设置图片元素的颜色类型

    Args:
        object: 图片元素
        color_type: 颜色类型

    Returns:
        操作是否成功
    """
    try:
        if hasattr(object, "PictureFormat"):
            object.PictureFormat.ColorType = color_type
            return True
        return False
    except Exception:
        return False


def get_object_image_info(object: Any, index: int = 0) -> Optional[Dict[str, Any]]:
    """获取元素的图片信息

    Args:
        object: 元素对象
        index: 元素索引

    Returns:
        包含图片信息的字典，如果元素不是图片则返回None
    """
    try:
        # 检查元素是否为图片
        if not (
            hasattr(object, "PictureFormat")
            or (
                hasattr(object, "InlineShape")
                and hasattr(object.InlineShape, "PictureFormat")
            )
        ):
            return None

        info = {
            "index": index,
            "type": type(object).__name__,
        }

        # 添加尺寸信息
        if hasattr(object, "Width"):
            info["width"] = object.Width
        if hasattr(object, "Height"):
            info["height"] = object.Height

        # 添加图片格式信息
        if hasattr(object, "PictureFormat"):
            info["picture_format"] = {
                "color_type": object.PictureFormat.ColorType,
            }

        # 添加范围信息
        if hasattr(object, "Range"):
            info["range_start"] = object.Range.Start
            info["range_end"] = object.Range.End

        return info
    except Exception:
        return None


@handle_com_error(ErrorCode.IMAGE_NOT_FOUND, "insert image")
def insert_image(
    document: win32com.client.CDispatch,
    image_path: str,
    locator: Optional[Dict[str, Any]] = None,
    position: str = "after",
) -> str:
    """在文档中插入图片

    Args:
        document: Word文档COM对象
        image_path: 图片文件路径
        locator: 定位器对象，用于指定插入位置
        position: 插入位置，可选值：'before', 'after', 'replace'

    Returns:
        插入图片成功的消息

    Raises:
        WordDocumentError: 当插入图片失败时抛出
    """
    if not document:
        raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")

    if not os.path.exists(image_path):
        raise WordDocumentError(
            ErrorCode.NOT_FOUND, f"Image file not found: {image_path}"
        )

    selector = SelectorEngine()
    range_obj = None

    if locator:
        # 使用定位器获取范围
        try:
            selection = selector.select(document, locator)
            if hasattr(selection, "_com_ranges") and selection._com_ranges:
                # 获取第一个对象
                selected_object = selection._com_ranges[0]
                
                # 确保我们有一个有效的Range对象
                if hasattr(selected_object, "Range"):
                    range_obj = selected_object.Range
                else:
                    # 如果对象本身就是Range对象，直接使用
                    range_obj = selected_object
                
                # 根据位置参数调整范围
                if position == "before":
                    range_obj.Collapse(Direction=1)  # wdCollapseStart
                elif position == "after":
                    range_obj.Collapse(Direction=0)  # wdCollapseEnd
                # 如果是"replace"，则不折叠范围，直接替换
            else:
                raise WordDocumentError(
                    ErrorCode.OBJECT_NOT_FOUND, "No object found matching the locator"
                )
        except Exception as e:
            raise WordDocumentError(
                ErrorCode.SERVER_ERROR, f"Failed to locate position for image: {str(e)}"
            )
    else:
        # 如果没有提供定位器，在文档末尾插入图片
         if not hasattr(document, 'Range'):
             raise WordDocumentError(
                 ErrorCode.DOCUMENT_ERROR,
                 "Document object missing required 'Range' method"
             )
         range_obj = document.Range()
         if not range_obj:
             raise WordDocumentError(
                 ErrorCode.OBJECT_TYPE_ERROR,
                 "Failed to create valid Range object"
             )
         range_obj.Collapse(Direction=0)  # wdCollapseEnd

    try:
        # 增强Range对象验证
        if range_obj is None or not hasattr(range_obj, 'Start') or not hasattr(range_obj, 'End'):
            raise WordDocumentError(
                ErrorCode.OBJECT_TYPE_ERROR,
                "Invalid Range object for image insertion"
            )
        
        # 插入图片
        # 如果直接使用Range参数失败，尝试先选择范围再插入
        try:
            picture = document.InlineShapes.AddPicture(
                FileName=image_path,
                LinkToFile=False,
                SaveWithDocument=True,
                Range=range_obj,
            )
        except Exception as first_exception:
            # 备用方法：先选择Range，然后插入图片
            try:
                range_obj.Select()
                picture = document.InlineShapes.AddPicture(
                    FileName=image_path,
                    LinkToFile=False,
                    SaveWithDocument=True,
                )
            except Exception as second_exception:
                log_error(f"Both methods failed to insert image {image_path}: {str(first_exception)}, {str(second_exception)}", exc_info=True)
                raise WordDocumentError(
                    ErrorCode.SERVER_ERROR, f"Failed to insert image: {str(second_exception)}"
                )

        # 添加成功日志
        log_info(f"Successfully inserted image: {image_path}")

        return json.dumps(
            {
                "success": True,
                "message": f"Image inserted successfully: {image_path}",
                "image_index": picture.Index if hasattr(picture, 'Index') else 1,
            },
            ensure_ascii=False,
        )

    except Exception as e:
        log_error(f"Failed to insert image {image_path}: {str(e)}", exc_info=True)
        raise WordDocumentError(
            ErrorCode.SERVER_ERROR, f"Failed to insert image: {str(e)}"
        )


@handle_com_error(ErrorCode.IMAGE_FORMAT_ERROR, "add caption")
def add_caption(
    document: win32com.client.CDispatch,
    caption_text: str,
    locator: Optional[Dict[str, Any]] = None,
    label: Optional[str] = None,
) -> str:
    """为文档元素添加题注

    Args:
        document: Word文档COM对象
        caption_text: 题注文本
        locator: 定位器对象，用于指定添加题注的位置
        label: 题注标签，可选

    Returns:
        添加题注成功的消息

    Raises:
        ValueError: 当参数无效时抛出
        WordDocumentError: 当添加题注失败时抛出
    """
    if not document:
        raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")

    # 验证参数
    if not caption_text:
        raise ValueError("Caption text cannot be empty")

    try:
        if locator:
            # 使用定位器获取范围
            selector = SelectorEngine()
            selection = selector.select(document, locator)
            if hasattr(selection, "_com_ranges") and selection._com_ranges:
                range_obj = selection._com_ranges[0]  # _com_ranges 已包含 Range 对象
                range_obj.Collapse(Direction=0)  # wdCollapseEnd
            else:
                raise WordDocumentError(
                    ErrorCode.OBJECT_NOT_FOUND, "No object found matching the locator"
                )
        else:
            # 如果没有提供定位器，在文档末尾添加题注
            if not hasattr(document, 'Range'):
                raise WordDocumentError(
                    ErrorCode.DOCUMENT_ERROR,
                    "Document object missing required 'Range' method"
                )
            range_obj = document.Range()
            if not range_obj:
                raise WordDocumentError(
                    ErrorCode.OBJECT_TYPE_ERROR,
                    "Failed to create valid Range object"
                )
            range_obj.Collapse(Direction=0)  # wdCollapseEnd

        # 添加题注
        try:
            # 构建完整的题注文本，包含可选的标签
            full_caption_text = caption_text
            if label:
                full_caption_text = f"{label}: {caption_text}"
            
            # 先尝试使用AutoTextEntries方法
            try:
                if hasattr(document.Application, 'ActiveDocument') and hasattr(document.Application.ActiveDocument, 'AttachedTemplate'):
                    document.Application.ActiveDocument.AttachedTemplate.AutoTextEntries(
                        "Caption Figure"
                    ).Insert(Where=range_obj)
                else:
                    raise Exception("AutoTextEntries not available")
                
                # 设置题注文本
                if hasattr(document.Application, 'Selection'):
                    caption_range = document.Application.Selection.Range
                    caption_range.Collapse(0)  # wdCollapseEnd
                    caption_range.Text = f" {full_caption_text}"
                else:
                    # 备用方法：直接在range_obj后插入文本
                    range_obj.InsertAfter(f" {full_caption_text}")
            except Exception:
                # 如果AutoTextEntries方法失败，使用直接文本插入方式
                range_obj.InsertAfter(full_caption_text)
        except Exception as e:
            log_error(f"Failed to insert caption: {str(e)}")
            raise WordDocumentError(
                ErrorCode.SERVER_ERROR, f"Failed to add caption: {str(e)}"
            )
    except Exception as e:
        raise WordDocumentError(
            ErrorCode.SERVER_ERROR, f"Failed to add caption: {str(e)}"
        )

    log_info(f"Successfully added caption to object")
    return json.dumps(
        {
            "success": True,
            "message": "Successfully added caption",
            "caption_text": caption_text,
            "label": label,
        },
        ensure_ascii=False,
    )


@handle_com_error(ErrorCode.IMAGE_LOAD_ERROR, "resize image")
def resize_image(
    document: win32com.client.CDispatch,
    image_index: int,
    width: Optional[int] = None,
    height: Optional[int] = None,
    maintain_aspect_ratio: bool = True,
) -> str:
    """调整图片大小

    Args:
        document: Word文档COM对象
        image_index: 图片索引（从1开始）
        width: 可选的图片宽度（像素）
        height: 可选的图片高度（像素）
        maintain_aspect_ratio: 是否保持宽高比，默认为True

    Returns:
        调整图片大小成功的消息

    Raises:
        ValueError: 当参数无效时抛出
        WordDocumentError: 当调整图片大小失败时抛出
    """
    if not document:
        raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")

    # 验证参数
    if image_index <= 0:
        raise ValueError("Image index must be a positive integer")

    # 至少需要指定宽度或高度
    if width is None and height is None:
        raise ValueError("At least one of width or height must be specified")

    if width is not None and width <= 0:
        raise ValueError("Width must be a positive integer")
    if height is not None and height <= 0:
        raise ValueError("Height must be a positive integer")

    # 获取图片
    inline_shapes = document.InlineShapes
    shapes = document.Shapes

    total_images = inline_shapes.Count + shapes.Count
    if image_index > total_images:
        raise WordDocumentError(
            ErrorCode.IMAGE_LOAD_ERROR,
            f"Image index {image_index} out of range. There are {total_images} images in the document",
        )

    try:
        # 获取图片对象
        if image_index <= inline_shapes.Count:
            image = inline_shapes(image_index)
        else:
            shape_index = image_index - inline_shapes.Count
            image = shapes(shape_index)
            if image.Type != 13:  # wdShapePicture
                raise WordDocumentError(
                    ErrorCode.IMAGE_NOT_FOUND,
                    f"The specified shape at index {shape_index} is not an image",
                )

        # 调整图片大小
        original_width = image.Width
        original_height = image.Height

        if maintain_aspect_ratio:
            if width is not None and height is not None:
                # 同时指定宽度和高度，但保持比例
                # 计算按照哪个尺寸来缩放
                width_ratio = width / original_width
                height_ratio = height / original_height

                if width_ratio < height_ratio:
                    # 按照宽度缩放
                    image.Width = width
                    image.Height = original_height * width_ratio
                else:
                    # 按照高度缩放
                    image.Height = height
                    image.Width = original_width * height_ratio
            elif width is not None:
                # 只指定宽度，保持比例
                ratio = original_height / original_width
                image.Width = width
                image.Height = width * ratio
            else:
                # 只指定高度，保持比例
                ratio = original_width / original_height
                image.Height = height
                image.Width = height * ratio
        else:
            # 不保持比例，直接设置指定的尺寸
            if width is not None:
                image.Width = width
            if height is not None:
                image.Height = height
    except Exception as e:
        raise WordDocumentError(
            ErrorCode.DOCUMENT_ERROR, f"Failed to resize image: {str(e)}"
        )

    log_info(f"Successfully resized image {image_index}")
    return json.dumps(
        {
            "success": True,
            "message": "Successfully resized image",
            "image_index": image_index,
            "new_width": image.Width,
            "new_height": image.Height,
            "maintain_aspect_ratio": maintain_aspect_ratio,
        },
        ensure_ascii=False,
    )


@handle_com_error(ErrorCode.OBJECT_TYPE_ERROR, "set image color type")
def set_image_color_type(
    document: win32com.client.CDispatch, image_index: int, color_type: str
) -> str:
    """
    设置图片的颜色类型

    Args:
        document: Word文档COM对象
        image_index: 图片索引（从1开始）
        color_type: 颜色类型，可选值：'color', 'grayscale', 'black_and_white', 'recolor'

    Returns:
        设置颜色类型成功的消息

    Raises:
        ValueError: 当参数无效时抛出
        WordDocumentError: 当设置颜色类型失败时抛出
    """
    if not document:
        raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")

    # 验证参数
    if image_index <= 0:
        raise ValueError("Image index must be a positive integer")

    # 验证颜色类型
    valid_color_types = ["color", "grayscale", "black_and_white", "recolor"]
    if color_type not in valid_color_types:
        raise ValueError(
            f"Invalid color type: {color_type}. Valid types: {', '.join(valid_color_types)}"
        )

    # 获取图片
    inline_shapes = document.InlineShapes
    shapes = document.Shapes

    total_images = inline_shapes.Count + shapes.Count
    if image_index > total_images:
        raise WordDocumentError(
            ErrorCode.IMAGE_NOT_FOUND,
            f"Image index {image_index} out of range. There are {total_images} images in the document",
        )

    # 映射颜色类型到Word常量
    color_type_map = {
        "color": 0,  # msoPictureColorTypeColor
        "grayscale": 1,  # msoPictureColorTypeGrayscale
        "black_and_white": 2,  # msoPictureColorTypeBlackAndWhite
        "recolor": 3,  # msoPictureColorTypeMixed (使用recolor表示自定义颜色)
    }

    try:
        # 获取图片对象
        if image_index <= inline_shapes.Count:
            image = inline_shapes(image_index)
        else:
            shape_index = image_index - inline_shapes.Count
            image = shapes(shape_index)
            if image.Type != 13:  # wdShapePicture
                raise WordDocumentError(
                    ErrorCode.IMAGE_NOT_FOUND,
                    f"The specified shape at index {shape_index} is not an image",
                )

        # 设置颜色类型
        image.PictureFormat.ColorType = color_type_map[color_type]
    except Exception as e:
        raise WordDocumentError(
            ErrorCode.DOCUMENT_ERROR, f"Failed to set image color type: {str(e)}"
        )

    log_info(f"Successfully set color type for image {image_index} to {color_type}")
    return json.dumps(
        {
            "success": True,
            "message": f"Successfully set image color type to {color_type}",
            "image_index": image_index,
            "color_type": color_type,
        },
        ensure_ascii=False,
    )


def _get_inline_image_details(inline_shape: Any) -> Dict[str, Any]:
    """获取内嵌图片的详细信息

    Args:
        inline_shape: Word内嵌形状对象

    Returns:
        包含图片详细信息的字典
    """
    details = {
        "type": "inline",
        "width": inline_shape.Width,
        "height": inline_shape.Height,
        "name": inline_shape.Name,
        "index": inline_shape.Index,
    }

    # 尝试获取更多信息
    try:
        if hasattr(inline_shape, "PictureFormat"):
            details["has_picture_format"] = True
            # 图片格式相关信息
            if hasattr(inline_shape.PictureFormat, "ColorType"):
                color_type_map = {
                    0: "color",
                    1: "grayscale",
                    2: "black_and_white",
                    3: "mixed",
                }
                color_type = inline_shape.PictureFormat.ColorType
                details["color_type"] = color_type_map.get(color_type, "unknown")
    except Exception as e:
        log_error(f"Failed to get additional image details: {str(e)}")

    return details


def _get_shape_image_details(shape: Any) -> Dict[str, Any]:
    """获取浮动图片的详细信息

    Args:
        shape: Word形状对象

    Returns:
        包含图片详细信息的字典
    """
    details = {
        "type": "floating",
        "width": shape.Width,
        "height": shape.Height,
        "name": shape.Name,
        "left": shape.Left,
        "top": shape.Top,
        "index": shape.Index,
    }

    # 尝试获取更多信息
    try:
        if hasattr(shape, "PictureFormat"):
            details["has_picture_format"] = True
            # 图片格式相关信息
            if hasattr(shape.PictureFormat, "ColorType"):
                color_type_map = {
                    0: "color",
                    1: "grayscale",
                    2: "black_and_white",
                    3: "mixed",
                }
                color_type = shape.PictureFormat.ColorType
                details["color_type"] = color_type_map.get(color_type, "unknown")
    except Exception as e:
        log_error(f"Failed to get additional shape details: {str(e)}")

    return details


@handle_com_error(ErrorCode.IMAGE_FORMAT_ERROR, "add image caption")
def add_image_caption(
    document: win32com.client.CDispatch,
    caption_text: str,
    locator: Optional[Dict[str, Any]] = None,
) -> str:
    """为文档元素添加题注

    Args:
        document: Word文档COM对象
        caption_text: 题注文本
        locator: 定位器对象，用于指定添加题注的位置

    Returns:
        添加题注成功的消息

    Raises:
        ValueError: 当参数无效时抛出
        WordDocumentError: 当添加题注失败时抛出
    """
    if not document:
        raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")

    # 验证参数
    if not caption_text:
        raise ValueError("Caption text cannot be empty")

    try:
        if locator:
            # 使用定位器获取范围
            selector = SelectorEngine()
            selection = selector.select(document, locator)
            if hasattr(selection, "_com_ranges") and selection._com_ranges:
                range_obj = selection._com_ranges[0]  # _com_ranges 已包含 Range 对象
                range_obj.Collapse(Direction=0)  # wdCollapseEnd
            else:
                raise WordDocumentError(
                    ErrorCode.OBJECT_NOT_FOUND, "No object found matching the locator"
                )
        else:
            # 如果没有提供定位器，在文档末尾添加题注
            if not hasattr(document, 'Range'):
                raise WordDocumentError(
                    ErrorCode.DOCUMENT_ERROR,
                    "Document object missing required 'Range' method"
                )
            range_obj = document.Range()
            if not range_obj:
                raise WordDocumentError(
                    ErrorCode.OBJECT_TYPE_ERROR,
                    "Failed to create valid Range object"
                )
            range_obj.Collapse(Direction=0)  # wdCollapseEnd

        # 添加题注
        try:
            document.Application.ActiveDocument.AttachedTemplate.AutoTextEntries(
                "Caption Figure"
            ).Insert(Where=range_obj)
        except Exception as e:
            log_error(f"Failed to insert caption: {str(e)}")
            raise WordDocumentError(
                ErrorCode.SERVER_ERROR, f"Failed to add caption: {str(e)}"
            )

        # 设置题注文本
        caption_range = document.Application.Selection.Range
        caption_range.Collapse(0)  # wdCollapseEnd
        caption_range.Text = f" {caption_text}"
    except Exception as e:
        raise WordDocumentError(
            ErrorCode.SERVER_ERROR, f"Failed to add caption: {str(e)}"
        )

    log_info(f"Successfully added caption to object")
    return json.dumps(
        {
            "success": True,
            "message": "Successfully added caption",
            "caption_text": caption_text,
        },
        ensure_ascii=False,
    )
