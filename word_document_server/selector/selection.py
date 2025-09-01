"""
Selection Abstraction Layer for Word Document MCP Server.
"""

import json
import os
from typing import Any, Dict, List, Optional

import win32com.client

from word_document_server.utils.core_utils import ErrorCode, WordDocumentError


class Selection:
    """Represents a selection of document objects.

    For guidance on proper locator syntax, please refer to:
    word_document_server/selector/LOCATOR_GUIDE.md
    """

    def __init__(
        self,
        com_ranges: List[win32com.client.CDispatch],
        document: win32com.client.CDispatch,
    ):
        """Initialize a Selection with COM objects and document reference.

        Args:
            com_ranges: List of COM range objects representing selected objects.
            document: Word document COM object for executing operations.

        Raises:
            ValueError: If com_ranges is empty.

        For guidance on proper locator syntax, please refer to:
        word_document_server/selector/LOCATOR_GUIDE.md
        """
        if not com_ranges:
            raise ValueError("Selection cannot be empty.")
        self._com_ranges = com_ranges
        self._document = document

    def get_object_types(self) -> List[Dict[str, Any]]:
        """
        获取选择集中所有元素的详细类型信息。

        Returns:
            包含每个元素详细类型信息的字典列表。

        For guidance on proper locator syntax, please refer to:
        word_document_server/selector/LOCATOR_GUIDE.md
        """
        object_types = []

        for i, object in enumerate(self._com_ranges):
            object_info = {
                "index": i,
                "com_type": str(type(object)),
                "object_type": "unknown",
                "properties": {},
            }

            # 检测元素类型并收集详细属性
            try:
                # 检查是否为段落
                if hasattr(object, "Style") and (hasattr(object, "Range") or hasattr(object, "Text")):
                    object_info["object_type"] = "paragraph"
                    object_info["properties"]["is_paragraph"] = True
                    try:
                        object_info["properties"][
                            "style_name"
                        ] = object.Style.NameLocal
                    except:
                        pass
                    try:
                        # 先尝试直接使用object作为Range对象（通过检查Text属性）
                        if hasattr(object, 'Text'):
                            text_preview = object.Text[:50] + ("..." if len(object.Text) > 50 else "")
                        # 否则尝试访问Range属性
                        elif hasattr(object, 'Range'):
                            text_preview = object.Range.Text[:50] + ("..." if len(object.Range.Text) > 50 else "")
                        else:
                            text_preview = ""
                        object_info["properties"]["text_preview"] = text_preview
                    except:
                        pass

                # 检查是否为表格
                elif hasattr(object, "Rows") and hasattr(object, "Columns"):
                    object_info["object_type"] = "table"
                    object_info["properties"]["is_table"] = True
                    try:
                        object_info["properties"]["rows_count"] = object.Rows.Count
                        object_info["properties"][
                            "columns_count"
                        ] = object.Columns.Count
                    except:
                        pass

                # 检查是否为图片
                elif hasattr(object, "Type") and object.Type in (
                    1,
                    3,
                ):  # 1=InlineShape, 3=Shape
                    object_info["object_type"] = "image"
                    object_info["properties"]["is_image"] = True
                    try:
                        if hasattr(object, "Width") and hasattr(object, "Height"):
                            object_info["properties"]["width"] = object.Width
                            object_info["properties"]["height"] = object.Height
                    except:
                        pass
                    try:
                        if hasattr(object, "Name"):
                            object_info["properties"]["name"] = object.Name
                    except:
                        pass

                # 检查是否为文本范围
                elif (
                    hasattr(object, "Text")
                    and hasattr(object, "Start")
                    and hasattr(object, "End")
                ):
                    object_info["object_type"] = "text_range"
                    object_info["properties"]["is_range"] = True
                    try:
                        object_info["properties"]["text_length"] = len(object.Text)
                        object_info["properties"]["text_preview"] = object.Text[
                            :50
                        ] + ("..." if len(object.Text) > 50 else "")
                    except:
                        pass

                # 检查是否为书签
                elif hasattr(object, "Name") and (hasattr(object, "Range") or hasattr(object, "Text")):
                    object_info["object_type"] = "bookmark"
                    try:
                        object_info["properties"]["name"] = object.Name
                    except:
                        pass

                # 检查是否为评论
                elif hasattr(object, "Initial") and (hasattr(object, "Range") or hasattr(object, "Text")):
                    object_info["object_type"] = "comment"
                    try:
                        object_info["properties"]["author"] = object.Author
                        # 先尝试直接使用object作为Range对象（通过检查Text属性）
                        if hasattr(object, 'Text'):
                            text_preview = object.Text[:50] + ("..." if len(object.Text) > 50 else "")
                        # 否则尝试访问Range属性
                        elif hasattr(object, 'Range'):
                            text_preview = object.Range.Text[:50] + ("..." if len(object.Range.Text) > 50 else "")
                        else:
                            text_preview = ""
                        object_info["properties"]["text_preview"] = text_preview
                    except:
                        pass

                # 检查是否为超链接
                elif hasattr(object, "Address") and (hasattr(object, "Range") or hasattr(object, "Text")):
                    object_info["object_type"] = "hyperlink"
                    try:
                        object_info["properties"]["address"] = object.Address
                        # 先尝试直接使用object作为Range对象（通过检查Text属性）
                        if hasattr(object, 'Text'):
                            text_preview = object.Text[:50] + ("..." if len(object.Text) > 50 else "")
                        # 否则尝试访问Range属性
                        elif hasattr(object, 'Range'):
                            text_preview = object.Range.Text[:50] + ("..." if len(object.Range.Text) > 50 else "")
                        else:
                            text_preview = ""
                        object_info["properties"]["text_preview"] = text_preview
                    except:
                        pass

                else:
                    # 默认类型
                    object_info["object_type"] = "default"

                # 获取元素ID（如果可用）
                if hasattr(object, "ID"):
                    try:
                        object_info["properties"]["id"] = object.ID
                    except:
                        pass

                # 获取元素的起始和结束位置（如果适用）
                try:
                    # 先尝试直接使用object作为Range对象（通过检查Start和End属性）
                    if hasattr(object, 'Start') and hasattr(object, 'End'):
                        object_info["properties"]["range_start"] = object.Start
                        object_info["properties"]["range_end"] = object.End
                    # 否则尝试访问Range属性
                    elif hasattr(object, 'Range') and hasattr(object.Range, 'Start') and hasattr(object.Range, 'End'):
                        object_info["properties"]["range_start"] = object.Range.Start
                        object_info["properties"]["range_end"] = object.Range.End
                except:
                    pass

            except Exception as e:
                # 忽略获取属性时的错误
                object_info["properties"]["error"] = str(e)

            object_types.append(object_info)

        return object_types
