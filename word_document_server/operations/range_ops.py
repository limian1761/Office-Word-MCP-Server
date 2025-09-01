"""
Object selection and manipulation operations for Word Document MCP Server.
This module contains operations for selecting and working with document objects.
"""

import json
import logging
from typing import Any, Dict, List, Optional

import win32com.client

from ..com_backend.com_utils import handle_com_error
from ..selector.selector import SelectorEngine
from ..utils.core_utils import ErrorCode, WordDocumentError, log_error, log_info
from .text_format_ops import (set_alignment_for_range, set_bold_for_range,
                              set_font_color_for_range,
                              set_font_name_for_range, set_font_size_for_range,
                              set_italic_for_range, set_paragraph_style)
from .text_ops import (get_object_text, insert_text_after_range,
                       insert_text_before_range)

logger = logging.getLogger(__name__)


# === Object Selection Operations ===


def select_objects(
    document: win32com.client.CDispatch, locator: Dict[str, Any]
) -> str:
    """根据定位器选择文档中的元素

    Args:
        document: Word文档COM对象
        locator: 定位器对象，包含元素类型和选择条件

    Returns:
        包含选定元素信息的JSON字符串
    """
    try:
        if not document:
            raise RuntimeError("No document open.")

        # 使用选择器引擎选择元素
        selector = SelectorEngine()
        selection = selector.select(document, locator)

        # 获取元素信息
        objects_info = []
        for i, range_obj in enumerate(selection._com_ranges):
            try:
                info = {
                    "index": i,
                    "type": "Range",
                }

                # 添加文本内容
                try:
                    # 所有对象都是Range对象，可以直接访问Text属性
                    info["text"] = (
                        range_obj.Text[:100] + "..."
                        if len(range_obj.Text) > 100
                        else range_obj.Text
                    )
                except Exception as text_e:
                    logger.warning(f"Failed to get text for object: {text_e}")

                # 添加样式属性
                if hasattr(range_obj, "Style") and hasattr(range_obj.Style, "NameLocal"):
                    info["style"] = range_obj.Style.NameLocal

                objects_info.append(info)
            except Exception as e:
                logger.warning(f"Failed to get info for object at index {i}: {e}")
                continue

        return json.dumps(objects_info, ensure_ascii=False, indent=2)

    except Exception as e:
        logger.error(f"Error in select_objects: {e}")
        raise WordDocumentError(
            ErrorCode.OBJECT_NOT_FOUND, f"Failed to select objects: {str(e)}")


def get_object_by_id(document: win32com.client.CDispatch, object_id: str) -> str:
    """根据ID获取元素

    Args:
        document: Word文档COM对象
        object_id: 元素ID

    Returns:
        包含元素信息的JSON字符串
    """
    try:
        if not document:
            raise RuntimeError("No document open.")

        # 尝试解析ID为整数索引
        try:
            index = int(object_id)
        except ValueError:
            raise WordDocumentError(
                ErrorCode.INVALID_INPUT,
                f"Invalid object ID: {object_id}. Must be an integer.",
            )

        # 获取所有段落
        paragraphs = list(document.Paragraphs)

        if index < 0 or index >= len(paragraphs):
            raise WordDocumentError(
                    ErrorCode.OBJECT_NOT_FOUND,
                    f"Object with index {index} not found. Document has {len(paragraphs)} paragraphs.",
                )

        # 获取指定索引的段落并转换为Range对象
        range_obj = paragraphs[index].Range

        # 构建元素信息
        object_info = {
            "index": index,
            "type": "Range",
            "text": (
                range_obj.Text[:200] + "..."
                if len(range_obj.Text) > 200
                else range_obj.Text
            ),
            "style": (
                range_obj.Style.NameLocal
                if hasattr(range_obj.Style, "NameLocal")
                else "Unknown"
            ),
        }

        return json.dumps(object_info, ensure_ascii=False, indent=2)

    except Exception as e:
        if isinstance(e, WordDocumentError):
            raise
        logger.error(f"Error in get_object_by_id: {e}")
        raise WordDocumentError(
                ErrorCode.OBJECT_NOT_FOUND, f"Failed to get object by ID: {str(e)}")


# === Object Manipulation Operations ===


def delete_object_by_locator(
    document: win32com.client.CDispatch, locator: Dict[str, Any]
) -> bool:
    """根据定位器删除元素

    Args:
        document: Word文档COM对象
        locator: 定位器对象

    Returns:
        操作是否成功
    """
    try:
        if not document:
            raise RuntimeError("No document open.")

        # 使用选择器引擎选择元素
        selector = SelectorEngine()
        selection = selector.select(document, locator, expect_single=True)

        # 删除元素 - 所有对象都是Range对象，可以直接调用Delete方法
        for range_obj in selection._com_ranges:
            range_obj.Delete()

        return True

    except Exception as e:
        logger.error(f"Error in delete_object_by_locator: {e}")
        raise WordDocumentError(
                ErrorCode.OBJECT_NOT_FOUND, f"Failed to delete object: {str(e)}")


# === Batch Operations ===


def batch_select_objects(
    document: win32com.client.CDispatch, locators: List[Dict[str, Any]]
) -> str:
    """批量选择多个元素

    Args:
        document: Word文档COM对象
        locators: 定位器对象列表

    Returns:
        包含所有选定元素信息的JSON字符串
    """
    try:
        if not document:
            raise RuntimeError("No document open.")

        all_objects_info = []

        # 依次处理每个定位器
        for i, locator in enumerate(locators):
            try:
                # 使用选择器引擎选择元素
                selector = SelectorEngine()
                selection = selector.select(document, locator)

                # 获取元素信息
                objects_info = []
                # Selection._com_ranges中只包含Range对象
                for j, range_obj in enumerate(selection._com_ranges):
                    try:
                        info = {
                            "batch_index": i,
                            "object_index": j,
                            "type": type(object).__name__,
                        }

                        # 添加文本内容（如果可用）
                        try:
                            # 所有对象都是Range对象，可以直接访问Text属性
                            try:
                                info["text"] = (
                                    range_obj.Text[:100] + "..."
                                    if len(range_obj.Text) > 100
                                    else range_obj.Text
                                )
                            except Exception as text_e:
                                logger.warning(f"Failed to get text for object: {text_e}")
                        except Exception as text_e:
                            logger.warning(f"Failed to get text for object: {text_e}")

                        # 添加其他属性（如果可用）
                        if hasattr(range_obj, "Style") and hasattr(
                            range_obj.Style, "NameLocal"
                        ):
                            info["style"] = range_obj.Style.NameLocal

                        objects_info.append(info)
                    except Exception as e:
                        logger.warning(
                            f"Failed to get info for object at batch {i}, index {j}: {e}"
                        )
                        continue

                all_objects_info.extend(objects_info)

            except Exception as e:
                logger.warning(f"Failed to select objects for locator {i}: {e}")
                continue

        return json.dumps(all_objects_info, ensure_ascii=False, indent=2)

    except Exception as e:
        logger.error(f"Error in batch_select_objects: {e}")
        raise WordDocumentError(
            ErrorCode.OBJECT_NOT_FOUND, f"Failed to batch select objects: {str(e)}"
        )


def batch_apply_formatting(
    document: win32com.client.CDispatch, operations: List[Dict[str, Any]]
) -> str:
    """批量应用格式化

    Args:
        document: Word文档COM对象
        operations: 操作列表，每个操作包含locator和formatting

    Returns:
        操作结果的JSON字符串
    """
    try:
        if not document:
            raise RuntimeError("No document open.")

        from .text_ops import apply_formatting_to_object

        results = []

        # 执行每个格式化操作
        for i, operation in enumerate(operations):
            try:
                if "locator" not in operation or "formatting" not in operation:
                    raise ValueError(
                        f"Operation {i} must contain 'locator' and 'formatting' keys"
                    )

                locator = operation["locator"]
                formatting = operation["formatting"]

                # 应用格式
                selector = SelectorEngine()
                selection = selector.select(document, locator)
                if selection and hasattr(selection, '_com_ranges') and selection._com_ranges:
                    # selection._com_ranges中的所有对象都是Range对象
                    range_obj = selection._com_ranges[0]
                    result = apply_formatting_to_object(range_obj, formatting)
                    # 解析结果以检查是否成功
                    import json
                    result_dict = json.loads(result)
                    if not result_dict.get("success", False):
                        raise Exception(result_dict.get("message", "Formatting failed"))

                results.append({"operation_index": i, "status": "success"})

            except Exception as e:
                logger.warning(f"Failed to apply formatting in operation {i}: {e}")
                results.append(
                    {"operation_index": i, "status": "failed", "error": str(e)}
                )
                continue

        return json.dumps(results, ensure_ascii=False, indent=2)

    except Exception as e:
        logger.error(f"Error in batch_apply_formatting: {e}")
        raise WordDocumentError(
            ErrorCode.FORMATTING_ERROR, f"Failed to batch apply formatting: {str(e)}"
        )
