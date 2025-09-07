"""
Selection Abstraction Layer for Word Document MCP Server.
"""

import json
import os
from typing import Any, Dict, List, Optional

import win32com.client

from ..mcp_service.core_utils import ErrorCode, WordDocumentError


class Selection:
    """Represents a selection of document objects.

    For guidance on proper locator syntax, please refer to:
    word_docx_tools/selector/LOCATOR_GUIDE.md
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
            ValueError: If com_ranges is empty or contains non-Range objects.

        For guidance on proper locator syntax, please refer to:
        word_docx_tools/selector/LOCATOR_GUIDE.md
        """
        if not com_ranges:
            raise ValueError("Selection cannot be empty.")

        # 验证所有对象都是有效的Range对象
        validated_ranges = []
        for obj in com_ranges:
            if self._is_valid_range(obj):
                validated_ranges.append(obj)
            else:
                # 如果对象无效，尝试转换为Range对象
                try:
                    # 检查是否有Range属性
                    if hasattr(obj, "Range"):
                        range_obj = obj.Range
                        if self._is_valid_range(range_obj):
                            validated_ranges.append(range_obj)
                        else:
                            raise ValueError(
                                f"Object's Range property is not a valid Range"
                            )
                    else:
                        raise ValueError(f"Object has no valid Range representation")
                except Exception as e:
                    raise ValueError(f"Invalid object in selection: {e}")

        if not validated_ranges:
            raise ValueError("No valid Range objects in selection.")

        self._com_ranges = validated_ranges
        self._document = document

    def _is_valid_range(self, obj: Any) -> bool:
        """Check if an object is a valid Range object.

        Args:
            obj: The object to check.

        Returns:
            True if the object is a valid Range object, False otherwise.
        """
        try:
            # 检查Range对象的核心属性
            return (
                hasattr(obj, "Text")
                and hasattr(obj, "Start")
                and hasattr(obj, "End")
                and
                # 简单验证属性的可访问性
                isinstance(obj.Text, str)
                and isinstance(obj.Start, (int, float))
                and isinstance(obj.End, (int, float))
            )
        except:
            return False

    def get_object_types(self) -> List[Dict[str, Any]]:
        """
        获取选择集中所有元素的详细类型信息。

        Returns:
            包含每个元素详细类型信息的字典列表。

        For guidance on proper locator syntax, please refer to:
        word_docx_tools/selector/LOCATOR_GUIDE.md
        """
        object_types = []

        for i, range_obj in enumerate(self._com_ranges):
            object_info: dict[str, Any] = {"type": "unknown", "properties": {}}
            object_info["properties"]["is_range"] = True

            try:
                # 获取Range基本信息
                object_info["properties"]["text_length"] = len(range_obj.Text)
                object_info["properties"]["text_preview"] = range_obj.Text[:50] + (
                    "..." if len(range_obj.Text) > 50 else ""
                )
                object_info["properties"]["range_start"] = range_obj.Start
                object_info["properties"]["range_end"] = range_obj.End
                object_info["properties"]["range_length"] = (
                    range_obj.End - range_obj.Start
                )

                # 检测元素类型并收集详细属性
                # 检查是否为段落
                try:
                    # 尝试通过Range获取段落信息
                    if (
                        hasattr(range_obj, "Paragraphs")
                        and range_obj.Paragraphs.Count > 0
                    ):
                        paragraph = range_obj.Paragraphs(1)
                        object_info["object_type"] = "paragraph"
                        object_info["properties"]["is_paragraph"] = True
                        if hasattr(paragraph, "Style"):
                            object_info["properties"][
                                "style_name"
                            ] = paragraph.Style.NameLocal
                except:
                    pass

                # 检查是否为表格
                try:
                    # 尝试通过Range获取表格信息
                    if hasattr(range_obj, "Tables") and range_obj.Tables.Count > 0:
                        table = range_obj.Tables(1)
                        object_info["object_type"] = "table"
                        object_info["properties"]["is_table"] = True
                        if hasattr(table, "Rows") and hasattr(table, "Columns"):
                            object_info["properties"]["rows_count"] = table.Rows.Count
                            object_info["properties"][
                                "columns_count"
                            ] = table.Columns.Count
                except:
                    pass

                # 检查是否为图片
                try:
                    # 尝试通过Range获取内嵌图片
                    if (
                        hasattr(range_obj, "InlineShapes")
                        and range_obj.InlineShapes.Count > 0
                    ):
                        shape = range_obj.InlineShapes(1)
                        object_info["object_type"] = "image"
                        object_info["properties"]["is_image"] = True
                        if hasattr(shape, "Width") and hasattr(shape, "Height"):
                            object_info["properties"]["width"] = shape.Width
                            object_info["properties"]["height"] = shape.Height
                        if hasattr(shape, "Name"):
                            object_info["properties"]["name"] = shape.Name
                except:
                    pass

                # 检查是否为书签
                try:
                    # 尝试通过Range获取书签
                    if hasattr(self._document, "Bookmarks"):
                        for bookmark in self._document.Bookmarks:
                            if (
                                hasattr(bookmark, "Range")
                                and bookmark.Range.Start == range_obj.Start
                                and bookmark.Range.End == range_obj.End
                            ):
                                object_info["object_type"] = "bookmark"
                                if hasattr(bookmark, "Name"):
                                    object_info["properties"]["name"] = bookmark.Name
                                break
                except:
                    pass

                # 检查是否为评论
                try:
                    # 尝试通过Range获取评论
                    if hasattr(self._document, "Comments"):
                        for comment in self._document.Comments:
                            if (
                                hasattr(comment, "Range")
                                and comment.Range.Start == range_obj.Start
                                and comment.Range.End == range_obj.End
                            ):
                                object_info["object_type"] = "comment"
                                if hasattr(comment, "Author"):
                                    object_info["properties"]["author"] = comment.Author
                                break
                except:
                    pass

                # 检查是否为超链接
                try:
                    # 尝试通过Range获取超链接
                    if (
                        hasattr(range_obj, "Hyperlinks")
                        and range_obj.Hyperlinks.Count > 0
                    ):
                        hyperlink = range_obj.Hyperlinks(1)
                        object_info["object_type"] = "hyperlink"
                        if hasattr(hyperlink, "Address"):
                            object_info["properties"]["address"] = hyperlink.Address
                except:
                    pass

                # 如果没有识别出具体类型，默认类型为text_range
                if object_info["object_type"] == "unknown":
                    object_info["object_type"] = "text_range"

                # 检查是否为表格
                elif hasattr(range_obj, "Rows") and hasattr(range_obj, "Columns"):
                    object_info["object_type"] = "table"
                    object_info["properties"]["is_table"] = True
                    try:
                        object_info["properties"]["rows_count"] = range_obj.Rows.Count
                        object_info["properties"][
                            "columns_count"
                        ] = range_obj.Columns.Count
                    except:
                        pass

                # 检查是否为图片
                elif hasattr(range_obj, "Type") and range_obj.Type in (
                    1,
                    3,
                ):  # 1=InlineShape, 3=Shape
                    object_info["object_type"] = "image"
                    object_info["properties"]["is_image"] = True
                    try:
                        if hasattr(range_obj, "Width") and hasattr(range_obj, "Height"):
                            object_info["properties"]["width"] = range_obj.Width
                            object_info["properties"]["height"] = range_obj.Height
                    except:
                        pass
                    try:
                        if hasattr(range_obj, "Name"):
                            object_info["properties"]["name"] = range_obj.Name
                    except:
                        pass

                # 检查是否为文本范围
                elif (
                    hasattr(range_obj, "Text")
                    and hasattr(range_obj, "Start")
                    and hasattr(range_obj, "End")
                ):
                    object_info["object_type"] = "text_range"
                    object_info["properties"]["is_range"] = True
                    try:
                        object_info["properties"]["text_length"] = len(range_obj.Text)
                        object_info["properties"]["text_preview"] = range_obj.Text[
                            :50
                        ] + ("..." if len(range_obj.Text) > 50 else "")
                    except:
                        pass

                # 检查是否为书签
                elif hasattr(range_obj, "Name") and (
                    hasattr(range_obj, "Range") or hasattr(range_obj, "Text")
                ):
                    object_info["object_type"] = "bookmark"
                    try:
                        object_info["properties"]["name"] = range_obj.Name
                    except:
                        pass

                # 检查是否为评论
                elif hasattr(range_obj, "Initial") and (
                    hasattr(range_obj, "Range") or hasattr(range_obj, "Text")
                ):
                    object_info["object_type"] = "comment"
                    try:
                        object_info["properties"]["author"] = range_obj.Author
                        # 先尝试直接使用range_obj作为Range对象（通过检查Text属性）
                        if hasattr(range_obj, "Text"):
                            text_preview = range_obj.Text[:50] + (
                                "..." if len(range_obj.Text) > 50 else ""
                            )
                        # 否则尝试访问Range属性
                        elif hasattr(range_obj, "Range"):
                            text_preview = range_obj.Range.Text[:50] + (
                                "..." if len(range_obj.Range.Text) > 50 else ""
                            )
                        else:
                            text_preview = ""
                        object_info["properties"]["text_preview"] = text_preview
                    except:
                        pass

                # 检查是否为超链接
                elif hasattr(range_obj, "Address") and (
                    hasattr(range_obj, "Range") or hasattr(range_obj, "Text")
                ):
                    object_info["object_type"] = "hyperlink"
                    try:
                        object_info["properties"]["address"] = range_obj.Address
                        # 先尝试直接使用range_obj作为Range对象（通过检查Text属性）
                        if hasattr(range_obj, "Text"):
                            text_preview = range_obj.Text[:50] + (
                                "..." if len(range_obj.Text) > 50 else ""
                            )
                        # 否则尝试访问Range属性
                        elif hasattr(range_obj, "Range"):
                            text_preview = range_obj.Range.Text[:50] + (
                                "..." if len(range_obj.Range.Text) > 50 else ""
                            )
                        else:
                            text_preview = ""
                        object_info["properties"]["text_preview"] = text_preview
                    except:
                        pass

                else:
                    # 默认类型
                    object_info["object_type"] = "default"

                # 获取元素ID（如果可用）
                if hasattr(range_obj, "ID"):
                    try:
                        object_info["properties"]["id"] = range_obj.ID
                    except:
                        pass

                # 获取元素的起始和结束位置（如果适用）
                try:
                    # 先尝试直接使用range_obj作为Range对象（通过检查Start和End属性）
                    if hasattr(range_obj, "Start") and hasattr(range_obj, "End"):
                        properties = object_info["properties"]
                        properties["range_start"] = range_obj.Start
                        properties["range_end"] = range_obj.End
                    # 否则尝试访问Range属性
                    elif (
                        hasattr(range_obj, "Range")
                        and hasattr(range_obj.Range, "Start")
                        and hasattr(range_obj.Range, "End")
                    ):
                        properties = object_info["properties"]
                        properties["range_start"] = range_obj.Range.Start
                        properties["range_end"] = range_obj.Range.End
                except:
                    pass

            except Exception as e:
                # 忽略获取属性时的错误
                object_info["properties"]["error"] = str(e)

            object_types.append(object_info)

        return object_types
