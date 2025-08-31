"""
Element selection and manipulation operations for Word Document MCP Server.
This module contains operations for selecting and working with document elements.
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
from .text_ops import (get_element_text, insert_text_after_range,
                       insert_text_before_range)

logger = logging.getLogger(__name__)


# === Element Selection Operations ===


def select_elements(
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
        elements_info = []
        for i, element in enumerate(selection._elements):
            try:
                info = {
                    "index": i,
                    "type": type(element).__name__,
                }

                # 添加文本内容（如果可用）
                if hasattr(element, "Range") and hasattr(element.Range, "Text"):
                    info["text"] = (
                        element.Range.Text[:100] + "..."
                        if len(element.Range.Text) > 100
                        else element.Range.Text
                    )

                # 添加其他属性（如果可用）
                if hasattr(element, "Style") and hasattr(element.Style, "NameLocal"):
                    info["style"] = element.Style.NameLocal

                elements_info.append(info)
            except Exception as e:
                logger.warning(f"Failed to get info for element at index {i}: {e}")
                continue

        return json.dumps(elements_info, ensure_ascii=False, indent=2)

    except Exception as e:
        logger.error(f"Error in select_elements: {e}")
        raise WordDocumentError(
            ErrorCode.ELEMENT_NOT_FOUND, f"Failed to select elements: {str(e)}"
        )


def get_element_by_id(document: win32com.client.CDispatch, element_id: str) -> str:
    """根据ID获取元素

    Args:
        document: Word文档COM对象
        element_id: 元素ID

    Returns:
        包含元素信息的JSON字符串
    """
    try:
        if not document:
            raise RuntimeError("No document open.")

        # 尝试解析ID为整数索引
        try:
            index = int(element_id)
        except ValueError:
            raise WordDocumentError(
                ErrorCode.INVALID_INPUT,
                f"Invalid element ID: {element_id}. Must be an integer.",
            )

        # 获取所有段落
        paragraphs = list(document.Paragraphs)

        if index < 0 or index >= len(paragraphs):
            raise WordDocumentError(
                ErrorCode.ELEMENT_NOT_FOUND,
                f"Element with index {index} not found. Document has {len(paragraphs)} paragraphs.",
            )

        # 获取指定索引的段落
        element = paragraphs[index]

        # 构建元素信息
        element_info = {
            "index": index,
            "type": "Paragraph",
            "text": (
                element.Range.Text[:200] + "..."
                if len(element.Range.Text) > 200
                else element.Range.Text
            ),
            "style": (
                element.Style.NameLocal
                if hasattr(element.Style, "NameLocal")
                else "Unknown"
            ),
        }

        return json.dumps(element_info, ensure_ascii=False, indent=2)

    except Exception as e:
        if isinstance(e, WordDocumentError):
            raise
        logger.error(f"Error in get_element_by_id: {e}")
        raise WordDocumentError(
            ErrorCode.ELEMENT_NOT_FOUND, f"Failed to get element by ID: {str(e)}"
        )


# === Element Manipulation Operations ===


def delete_element_by_locator(
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

        # 删除元素
        from .text_ops import delete_element

        for element in selection._elements:
            delete_element(element)

        return True

    except Exception as e:
        logger.error(f"Error in delete_element_by_locator: {e}")
        raise WordDocumentError(
            ErrorCode.ELEMENT_NOT_FOUND, f"Failed to delete element: {str(e)}"
        )


# === Batch Operations ===


def batch_select_elements(
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

        all_elements_info = []

        # 依次处理每个定位器
        for i, locator in enumerate(locators):
            try:
                # 使用选择器引擎选择元素
                selector = SelectorEngine()
                selection = selector.select(document, locator)

                # 获取元素信息
                elements_info = []
                for j, element in enumerate(selection._elements):
                    try:
                        info = {
                            "batch_index": i,
                            "element_index": j,
                            "type": type(element).__name__,
                        }

                        # 添加文本内容（如果可用）
                        if hasattr(element, "Range") and hasattr(element.Range, "Text"):
                            info["text"] = (
                                element.Range.Text[:100] + "..."
                                if len(element.Range.Text) > 100
                                else element.Range.Text
                            )

                        # 添加其他属性（如果可用）
                        if hasattr(element, "Style") and hasattr(
                            element.Style, "NameLocal"
                        ):
                            info["style"] = element.Style.NameLocal

                        elements_info.append(info)
                    except Exception as e:
                        logger.warning(
                            f"Failed to get info for element at batch {i}, index {j}: {e}"
                        )
                        continue

                all_elements_info.extend(elements_info)

            except Exception as e:
                logger.warning(f"Failed to select elements for locator {i}: {e}")
                continue

        return json.dumps(all_elements_info, ensure_ascii=False, indent=2)

    except Exception as e:
        logger.error(f"Error in batch_select_elements: {e}")
        raise WordDocumentError(
            ErrorCode.ELEMENT_NOT_FOUND, f"Failed to batch select elements: {str(e)}"
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

        from .text_ops import apply_formatting_to_element

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
                if selection and hasattr(selection, '_elements') and selection._elements:
                    element = selection._elements[0]
                    result = apply_formatting_to_element(element, formatting)
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
