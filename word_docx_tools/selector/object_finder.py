"""Object finder for the selector engine.

This module contains the functionality for finding and selecting
objects within a Word document.
"""

from typing import Any, Dict, List, Optional, Set, TypeVar, Union, cast

import win32com.client
from win32com.client import CDispatch

from .exceptions import AmbiguousLocatorError
from .filter_handlers import FilterHandlers
from ..com_backend.com_utils import iter_com_collection

# Type variables for better type hinting
ObjectT = TypeVar("ObjectT")


class ObjectFinder(FilterHandlers):
    """Component responsible for finding and selecting objects in a Word document."""

    def __init__(self, document: CDispatch):
        """Initialize the ObjectFinder with a document reference.

        Args:
            document: The Word document COM object.
        """
        self.document = document

    def select_core(self, locator: Dict[str, Any]) -> List[CDispatch]:
        """核心选择方法，用于SelectorEngine调用"""
        object_type = locator.get("type", "paragraph")
        value = locator.get("value")
        filters = locator.get("filters", [])

        # 根据元素类型选择不同的获取方法
        if object_type == "paragraph":
            objects = self.get_all_paragraphs()
        elif object_type == "table":
            objects = self.get_all_tables()
        elif object_type == "comment":
            objects = self.get_all_comments()
        elif object_type == "image" or object_type == "inline_shape":
            # 处理image和inline_shape类型，都是获取所有内嵌图片
            objects = self.get_all_images()
        else:
            # 默认返回所有段落
            objects = self.get_all_paragraphs()

        # 应用过滤器
        if filters:
            objects = self.apply_filters(objects, filters)

        # 调试日志
        print(f"[DEBUG] 初始对象数量: {len(objects)}")
        print(f"[DEBUG] 定位器: {locator}")

        # 处理可能的不同参数名（value、index、id等），支持多种类型
        # 定义可能的参数名列表，按优先级排序
        possible_param_names = ["value", "index", "id"]
        
        # 存储找到的对象
        found_objects = None
        
        for param_name in possible_param_names:
            param_value = locator.get(param_name)
            if param_value is not None and param_value != "":
                try:
                    # 尝试作为索引处理（1-based）
                    print(f"[DEBUG] 尝试将{param_name}={param_value}作为索引处理")
                    index = int(param_value)
                    if 0 < index <= len(objects):
                        print(f"[DEBUG] 索引{index}有效，返回单个对象")
                        return [objects[index - 1]]
                    else:
                        # 索引超出范围，设置found_objects为空列表
                        print(f"[DEBUG] 索引{index}超出范围（对象数量: {len(objects)}），设置found_objects为空列表")
                        found_objects = []
                except (ValueError, TypeError):
                    # 如果不能作为索引，将其作为文本内容处理
                    # 添加contains_text过滤器并重新应用所有过滤器
                    # 确保text_filters是一个新的列表
                    text_filters = list(filters) if filters else []
                    text_filters.append({"contains_text": param_value})
                    
                    # 重新从所有对象开始应用过滤，而不是在已过滤的对象上继续过滤
                    # 根据object_type获取所有对象
                    if object_type == "paragraph":
                        all_objects = self.get_all_paragraphs()
                    elif object_type == "table":
                        all_objects = self.get_all_tables()
                    elif object_type == "comment":
                        all_objects = self.get_all_comments()
                    elif object_type == "image" or object_type == "inline_shape":
                        all_objects = self.get_all_images()
                    else:
                        all_objects = self.get_all_paragraphs()
                    
                    text_matched_objects = self.apply_filters(all_objects, text_filters)
                    print(f"[DEBUG] 文本过滤后对象数量: {len(text_matched_objects)}")
                    
                    # 如果有匹配的对象，返回它们
                    if text_matched_objects:
                        print(f"[DEBUG] 文本过滤找到匹配对象，返回")
                        return text_matched_objects
                    else:
                        # 没有匹配的对象，设置found_objects为空列表
                        print(f"[DEBUG] 文本过滤没有找到匹配对象，设置found_objects为空列表")
                        found_objects = []
                    
                    # 继续尝试下一个参数名
                    continue

        # 如果尝试了所有参数名都没有匹配，但设置了found_objects为空列表
        # 则返回空列表，而不是所有对象
        print(f"[DEBUG] 循环结束后，found_objects: {found_objects}")
        if found_objects is not None:
            print(f"[DEBUG] 返回found_objects: {len(found_objects)}个对象")
            return found_objects
            
        # 如果没有提供任何参数名，返回应用过滤器后的对象
        print(f"[DEBUG] 返回应用过滤器后的对象: {len(objects)}个对象")
        return objects

    def find_anchor(self, anchor_id: str) -> Optional[CDispatch]:
        """Find an anchor object in the document based on its identifier.

        Args:
            anchor_id: The identifier of the anchor to find.

        Returns:
            The anchor object if found, None otherwise.

        Raises:
            AmbiguousLocatorError: If multiple objects match the anchor identifier.
        """
        # Handle special anchor types
        if anchor_id == "document_start":
            start_range = self.document.Content
            start_range.Collapse(True)  # wdCollapseStart
            return cast(CDispatch, self.document.Content)
        elif anchor_id == "document_end":
            end_range = self.document.Content
            end_range.Collapse(False)  # wdCollapseEnd
            return cast(CDispatch, end_range)
        elif anchor_id == "current_selection":
            return cast(CDispatch, self.document.Application.Selection)
        elif anchor_id == "active_range":
            return cast(CDispatch, self.document.ActiveWindow.Selection.Range)

        # Handle bookmark anchors
        if anchor_id.startswith("bookmark:"):
            bookmark_name = anchor_id.replace("bookmark:", "", 1)
            try:
                return cast(CDispatch, self.document.Bookmarks(bookmark_name).Range)
            except Exception:
                return None

        # Handle heading anchors
        if anchor_id.startswith("heading:"):
            heading_text = anchor_id.replace("heading:", "", 1)
            for para in self.document.Paragraphs:
                if para.Style.NameLocal.startswith(
                    "Heading"
                ) or para.Style.NameLocal.startswith("标题"):
                    if heading_text.lower() in para.Range.Text.lower():
                        return cast(CDispatch, para.Range)
            return None

        # Try to find by ID or other attributes
        try:
            # Attempt to find by index
            index = int(anchor_id)
            for object_type in ["Paragraphs", "Tables", "Comments"]:
                if hasattr(self.document, object_type):
                    objects = getattr(self.document, object_type)
                    if 0 <= index < objects.Count:
                        return cast(CDispatch, objects(index + 1))
            return None
        except ValueError:
            # Not an index, try other methods
            pass

        # Default: try to find by text content
        for para in self.document.Paragraphs:
            if anchor_id.lower() in para.Range.Text.lower():
                return cast(CDispatch, para.Range)

        return None

    def get_initial_candidates(self, locator_type: str, **kwargs) -> List[Any]:
        """Get the initial set of candidate objects based on the locator type.

        Args:
            locator_type: The type of objects to retrieve.
            **kwargs: Additional parameters for filtering.

        Returns:
            A list of candidate objects.
        """
        if "within_range" in kwargs:
            return self._get_range_specific_candidates(
                locator_type, kwargs["within_range"]
            )
        else:
            return self._get_global_candidates(locator_type)

    def _get_global_candidates(self, object_type: str) -> List[Any]:
        """Retrieve objects of a specific type from the entire document.

        Args:
            object_type: The type of objects to retrieve.

        Returns:
            A list of objects matching the specified type.
        """
        # This function is optimized for Word COM object access patterns.
        # It's more efficient to call the COM object once and then work with the
        # resulting collection in Python, rather than making multiple COM calls.
        # This is a best practice for performance with pywin32 and COM objects.

        candidates = []

        # Handle document-specific candidates
        if object_type == "document":
            return [self.document.Content]
        elif object_type == "document_start":
            # Return the start of the document
            start_range = self.document.Content
            start_range.Collapse(True)  # wdCollapseStart
            return [start_range]
        elif object_type == "document_end":
            # Return the end of the document
            end_range = self.document.Content
            end_range.Collapse(False)  # wdCollapseEnd (0 is correct value)
            return [end_range]

        # Handle different object types
        if object_type == "paragraph":
            candidates = self.get_all_paragraphs()
        elif object_type == "table":
            candidates = self.get_all_tables()
        elif object_type == "cell":
            tables = self.get_all_tables()
            for table in tables:
                for cell in table.Range.Cells:
                    candidates.append(cell)
        elif object_type == "inline_shape" or object_type == "image":
            # Get all inline shapes
            shapes = self.document.InlineShapes
            candidates = [shapes(i) for i in range(1, shapes.Count + 1)]
        elif object_type == "comment":
            comments = self.document.Comments
            candidates = [comments(i) for i in range(1, comments.Count + 1)]
        elif object_type == "range":
            # Default to the entire document range
            candidates = [self.document.Content]
        elif object_type == "selection":
            # Get the current selection
            candidates = [self.document.Application.Selection]

        return candidates

    def _get_range_specific_candidates(
        self, object_type: str, range_obj: CDispatch
    ) -> List[Any]:
        """Retrieve objects of a specific type within a given range.

        Args:
            object_type: The type of objects to retrieve.
            range_obj: The range within which to search.

        Returns:
            A list of objects matching the specified type within the range.
        """
        candidates = []

        # Handle different object types within the range
        if object_type == "paragraph":
            # Get all paragraphs in the range
            start_paragraph = self.document.Range(
                range_obj.Start, range_obj.Start
            ).Paragraphs(1)
            end_paragraph = self.document.Range(
                range_obj.End, range_obj.End
            ).Paragraphs(1)
            start_index = start_paragraph.Range.Start
            end_index = end_paragraph.Range.End

            # Iterate through all paragraphs and filter those that overlap with the range
            for para in self.document.Paragraphs:
                if (
                    (para.Range.Start >= start_index and para.Range.Start <= end_index)
                    or (para.Range.End >= start_index and para.Range.End <= end_index)
                    or (para.Range.Start <= start_index and para.Range.End >= end_index)
                ):
                    candidates.append(para)
        elif object_type == "table":
            # Check if the range intersects with any tables
            for table in self.document.Tables:
                table_range = table.Range
                if (
                    (
                        table_range.Start >= range_obj.Start
                        and table_range.Start <= range_obj.End
                    )
                    or (
                        table_range.End >= range_obj.Start
                        and table_range.End <= range_obj.End
                    )
                    or (
                        table_range.Start <= range_obj.Start
                        and table_range.End >= range_obj.End
                    )
                ):
                    candidates.append(table)
        elif object_type == "inline_shape" or object_type == "image":
            # Get inline shapes within the range
            candidates = self._get_inline_shapes_in_range(range_obj)

        return candidates

    def get_all_paragraphs(self) -> List[CDispatch]:
        """Retrieve all paragraphs in the document.

        Returns:
            A list of all paragraphs in the document.
        """
        return iter_com_collection(self.document.Paragraphs)

    def get_all_tables(self) -> List[CDispatch]:
        """Retrieve all tables in the document.

        Returns:
            A list of all tables in the document.
        """
        return iter_com_collection(self.document.Tables)

    def _get_inline_shapes_in_range(self, range_obj: CDispatch) -> List[CDispatch]:
        """Get all inline shapes within a specific range.

        Args:
            range_obj: The range to search within.

        Returns:
            A list of inline shapes within the specified range.
        """
        shapes_in_range = []
        for shape in iter_com_collection(self.document.InlineShapes):
            if (
                shape.Range.Start >= range_obj.Start
                and shape.Range.End <= range_obj.End
            ):
                shapes_in_range.append(shape)
        return shapes_in_range

    def apply_filters(
        self, objects: List[Any], filters: List[Dict[str, Any]]
    ) -> List[Any]:
        """Apply a series of filters to a list of objects.

        Args:
            objects: The list of objects to filter.
            filters: A list of filter definitions.

        Returns:
            A filtered list of objects.
        """
        filtered_objects = objects.copy()

        for filter_def in filters:
            filter_type = filter_def.get("type", "")
            filter_value = filter_def.get("value", None)
            filter_method_name = f"_filter_by_{filter_type}"

            # Check if the filter method exists
            if not hasattr(self, filter_method_name):
                continue

            filter_method = getattr(self, filter_method_name)
            filtered_objects = filter_method(filtered_objects, filter_value)

            # If no objects remain after filtering, we can stop early
            if not filtered_objects:
                break

        return filtered_objects

    def select_relative_to_anchor(
        self, objects: List[Any], anchor: CDispatch, relation: str
    ) -> List[Any]:
        """Select objects relative to an anchor object based on the specified relation.

        Args:
            objects: The list of objects to filter.
            anchor: The anchor object to base the selection on.
            relation: The type of relation to use (e.g., "first_occurrence_after").

        Returns:
            A list of objects that match the relation to the anchor.
        """
        # Initialize the result list
        result_objects = []

        # Handle different relation types
        if relation == "all_occurrences_within":
            # Return all objects that are within the anchor range
            result_objects = [
                el
                for el in objects
                if hasattr(el, "Range")
                and el.Range.Start >= anchor.Start
                and el.Range.End <= anchor.End
            ]
        elif relation == "first_occurrence_after":
            # Return the first object that comes after the anchor
            after_anchor = [
                el
                for el in objects
                if hasattr(el, "Range") and el.Range.Start > anchor.End
            ]
            if after_anchor:
                # Sort by position and take the first
                after_anchor.sort(key=lambda x: x.Range.Start)
                result_objects = [after_anchor[0]]
        elif relation == "parent_of":
            # Return the parent object of the anchor
            if hasattr(anchor, "Parent"):
                result_objects = [anchor.Parent]
        elif relation == "immediately_following":
            # Return the object immediately following the anchor
            # We need to find the smallest start position that is greater than the anchor's end
            min_start = float("inf")
            closest_object = None
            for el in objects:
                if (
                    hasattr(el, "Range")
                    and el.Range.Start > anchor.End
                    and el.Range.Start < min_start
                ):
                    min_start = el.Range.Start
                    closest_object = el
            if closest_object:
                result_objects = [closest_object]

        return result_objects

    def get_all_comments(self) -> List[CDispatch]:
        """Retrieve all comments in the document.

        Returns:
            A list of all comments in the document.
        """
        return iter_com_collection(self.document.Comments)

    def get_all_images(self) -> List[CDispatch]:
        """Retrieve all images in the document.

        Returns:
            A list of all images in the document.
        """
        # Images in Word are represented as InlineShapes
        # Filter to include only pictures
        images = []
        for shape in iter_com_collection(self.document.InlineShapes):
            # Check if the shape is a picture
            if shape.Type == 1:  # wdInlineShapePicture
                images.append(shape)
        return images
