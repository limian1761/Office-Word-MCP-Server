"""
Range operations for Word Document MCP Server.
This module contains functions for manipulating document ranges and object selection.
"""

import json
import logging
from typing import Any, Dict, List, Optional

import win32com.client

from ..com_backend.com_utils import handle_com_error, iter_com_collection
from ..mcp_service.core_utils import (ErrorCode, WordDocumentError, log_error,
                                      log_info)
from ..selector.selector import SelectorEngine
from .text_operations import apply_formatting_to_object

logger = logging.getLogger(__name__)


@handle_com_error(ErrorCode.OBJECT_NOT_FOUND, "select objects")
def select_objects(document: win32com.client.CDispatch, locator: Dict[str, Any]) -> str:
    """根据定位器选择元素

    Args:
        document: Word文档COM对象
        locator: 定位器对象

    Returns:
        包含所选元素信息的JSON字符串
    """
    try:
        if not document:
            raise RuntimeError("No document open.")

        # 使用选择器引擎选择元素
        selector = SelectorEngine()
        selection = selector.select(document, locator)

        # 构建元素信息
        objects_info = []
        # Selection._com_ranges中只包含Range对象
        for i, range_obj in enumerate(selection._com_ranges):
            try:
                info = {
                    "index": i,
                    "type": "Range",
                }

                # 添加文本内容（如果可用）
                try:
                    # 所有对象都是Range对象，可以直接访问Text属性
                    info["text"] = (
                        range_obj.Text[:200] + "..."
                        if len(range_obj.Text) > 200
                        else range_obj.Text
                    )
                except Exception as text_e:
                    logger.warning(f"Failed to get text for object: {text_e}")

                # 添加样式信息（如果可用）
                if hasattr(range_obj, "Style") and hasattr(
                    range_obj.Style, "NameLocal"
                ):
                    info["style"] = range_obj.Style.NameLocal

                # 添加位置信息（如果可用）
                if hasattr(range_obj, "Start") and hasattr(range_obj, "End"):
                    info["start_position"] = range_obj.Start
                    info["end_position"] = range_obj.End

                objects_info.append(info)
            except Exception as e:
                logger.warning(f"Failed to get info for object {i}: {e}")
                continue

        return json.dumps(objects_info, ensure_ascii=False, indent=2)

    except Exception as e:
        logger.error(f"Error in select_objects: {e}")
        raise WordDocumentError(
            ErrorCode.OBJECT_NOT_FOUND, f"Failed to select objects: {str(e)}"
        )


@handle_com_error(ErrorCode.OBJECT_NOT_FOUND, "get object by id")
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

        # 尝试将object_id转换为整数索引
        try:
            index = int(object_id)
        except ValueError:
            raise WordDocumentError(
                ErrorCode.INVALID_INPUT,
                f"Invalid object ID: {object_id}. Must be an integer.",
            )

        # 获取段落总数
        paragraph_count = document.Paragraphs.Count

        if index < 0 or index >= paragraph_count:
            raise WordDocumentError(
                ErrorCode.OBJECT_NOT_FOUND,
                f"Object with index {index} not found. Document has {paragraph_count} paragraphs.",
            )

        # 获取指定索引的段落并转换为Range对象
        range_obj = document.Paragraphs(index + 1).Range

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
            ErrorCode.OBJECT_NOT_FOUND, f"Failed to get object by ID: {str(e)}"
        )


@handle_com_error(ErrorCode.OBJECT_NOT_FOUND, "batch select objects")
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
                            "type": type(range_obj).__name__,
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
                                logger.warning(
                                    f"Failed to get text for object: {text_e}"
                                )
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


@handle_com_error(ErrorCode.FORMATTING_ERROR, "batch apply formatting")
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
                if (
                    selection
                    and hasattr(selection, "_com_ranges")
                    and selection._com_ranges
                ):
                    # 遍历所有Range对象并应用格式
                    all_success = True
                    for range_obj in selection._com_ranges:
                        try:
                            result = apply_formatting_to_object(range_obj, formatting)
                            # 解析结果以检查是否成功
                            try:
                                result_dict = json.loads(result)
                                if not result_dict.get("success", False):
                                    all_success = False
                                    logger.warning(
                                        f"Formatting failed for range object: {result_dict.get('message', 'Unknown error')}"
                                    )
                            except json.JSONDecodeError:
                                # 如果结果不是有效的JSON，尝试检查字符串内容
                                if (
                                    "error" in result.lower()
                                    or "failed" in result.lower()
                                ):
                                    all_success = False
                                    logger.warning(
                                        f"Formatting may have failed (invalid JSON response): {result}"
                                    )
                        except Exception as inner_e:
                            all_success = False
                            logger.warning(
                                f"Error applying formatting to range object: {inner_e}"
                            )

                    if not all_success:
                        raise Exception("Some formatting operations failed")

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


@handle_com_error(ErrorCode.OBJECT_NOT_FOUND, "delete object by locator")
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
        # 对于段落类型的定位器，无论使用什么参数名（value、index、id等），都强制要求单个对象
        is_paragraph = locator.get('type') == 'paragraph'
        has_position_param = any(param in locator for param in ['value', 'index', 'id'])
        expect_single = is_paragraph and has_position_param
        
        selector = SelectorEngine()
        selection = selector.select(document, locator, expect_single=expect_single)

        # 删除元素 - 所有对象都是Range对象，可以直接调用Delete方法
        for range_obj in selection._com_ranges:
            range_obj.Delete()

        return True

    except Exception as e:
        logger.error(f"Error in delete_object_by_locator: {e}")
        raise WordDocumentError(
            ErrorCode.OBJECT_NOT_FOUND, f"Failed to delete object: {str(e)}")
