"""
Selection Abstraction Layer for Word Document MCP Server.
"""

import json
import os
from typing import Any, Dict, List, Optional

import win32com.client

from word_document_server.utils.core_utils import ErrorCode, WordDocumentError


class Selection:
    """Represents a selection of document elements.

    For guidance on proper locator syntax, please refer to:
    word_document_server/selector/LOCATOR_GUIDE.md
    """

    def __init__(
        self,
        raw_com_elements: List[win32com.client.CDispatch],
        document: win32com.client.CDispatch,
    ):
        """Initialize a Selection with COM elements and document reference.

        Args:
            raw_com_elements: List of raw COM objects representing selected elements.
            document: Word document COM object for executing operations.

        Raises:
            ValueError: If raw_com_elements is empty.

        For guidance on proper locator syntax, please refer to:
        word_document_server/selector/LOCATOR_GUIDE.md
        """
        if not raw_com_elements:
            raise ValueError("Selection cannot be empty.")
        self._elements = raw_com_elements
        self._document = document

    def get_element_types(self) -> List[Dict[str, Any]]:
        """
        获取选择集中所有元素的详细类型信息。

        Returns:
            包含每个元素详细类型信息的字典列表。

        For guidance on proper locator syntax, please refer to:
        word_document_server/selector/LOCATOR_GUIDE.md
        """
        element_types = []

        for i, element in enumerate(self._elements):
            element_info = {
                "index": i,
                "com_type": str(type(element)),
                "element_type": "unknown",
                "properties": {},
            }

            # 检测元素类型并收集详细属性
            try:
                # 检查是否为段落
                if hasattr(element, "Style") and hasattr(element, "Range"):
                    element_info["element_type"] = "paragraph"
                    element_info["properties"]["is_paragraph"] = True
                    try:
                        element_info["properties"][
                            "style_name"
                        ] = element.Style.NameLocal
                    except:
                        pass
                    try:
                        element_info["properties"]["text_preview"] = element.Range.Text[
                            :50
                        ] + ("..." if len(element.Range.Text) > 50 else "")
                    except:
                        pass

                # 检查是否为表格
                elif hasattr(element, "Rows") and hasattr(element, "Columns"):
                    element_info["element_type"] = "table"
                    element_info["properties"]["is_table"] = True
                    try:
                        element_info["properties"]["rows_count"] = element.Rows.Count
                        element_info["properties"][
                            "columns_count"
                        ] = element.Columns.Count
                    except:
                        pass

                # 检查是否为图片
                elif hasattr(element, "Type") and element.Type in (
                    1,
                    3,
                ):  # 1=InlineShape, 3=Shape
                    element_info["element_type"] = "image"
                    element_info["properties"]["is_image"] = True
                    try:
                        if hasattr(element, "Width") and hasattr(element, "Height"):
                            element_info["properties"]["width"] = element.Width
                            element_info["properties"]["height"] = element.Height
                    except:
                        pass
                    try:
                        if hasattr(element, "Name"):
                            element_info["properties"]["name"] = element.Name
                    except:
                        pass

                # 检查是否为文本范围
                elif (
                    hasattr(element, "Text")
                    and hasattr(element, "Start")
                    and hasattr(element, "End")
                ):
                    element_info["element_type"] = "text_range"
                    element_info["properties"]["is_range"] = True
                    try:
                        element_info["properties"]["text_length"] = len(element.Text)
                        element_info["properties"]["text_preview"] = element.Text[
                            :50
                        ] + ("..." if len(element.Text) > 50 else "")
                    except:
                        pass

                # 检查是否为书签
                elif hasattr(element, "Name") and hasattr(element, "Range"):
                    element_info["element_type"] = "bookmark"
                    try:
                        element_info["properties"]["name"] = element.Name
                    except:
                        pass

                # 检查是否为评论
                elif hasattr(element, "Initial") and hasattr(element, "Range"):
                    element_info["element_type"] = "comment"
                    try:
                        element_info["properties"]["author"] = element.Author
                        element_info["properties"]["text_preview"] = element.Range.Text[
                            :50
                        ] + ("..." if len(element.Range.Text) > 50 else "")
                    except:
                        pass

                # 检查是否为超链接
                elif hasattr(element, "Address") and hasattr(element, "Range"):
                    element_info["element_type"] = "hyperlink"
                    try:
                        element_info["properties"]["address"] = element.Address
                        element_info["properties"]["text_preview"] = element.Range.Text[
                            :50
                        ] + ("..." if len(element.Range.Text) > 50 else "")
                    except:
                        pass

                else:
                    # 默认类型
                    element_info["element_type"] = "default"

                # 获取元素ID（如果可用）
                if hasattr(element, "ID"):
                    try:
                        element_info["properties"]["id"] = element.ID
                    except:
                        pass

                # 获取元素的起始和结束位置（如果适用）
                if (
                    hasattr(element, "Range")
                    and hasattr(element.Range, "Start")
                    and hasattr(element.Range, "End")
                ):
                    try:
                        element_info["properties"]["range_start"] = element.Range.Start
                        element_info["properties"]["range_end"] = element.Range.End
                    except:
                        pass

            except Exception as e:
                # 忽略获取属性时的错误
                element_info["properties"]["error"] = str(e)

            element_types.append(element_info)

        return element_types
