"""Object finder for the selector engine.

This module contains the functionality for finding and selecting
objects within a Word document.
"""

from typing import Any, Dict, List, Optional, Set, TypeVar, Union, cast

import win32com.client
from win32com.client import CDispatch

from .exceptions import AmbiguousLocatorError, LocatorError
from .filter_handlers import FilterHandlers
from .locator_parser import LocatorParser
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
        treat_as_index = locator.get("treat_as_index", False)
        
        # 根据元素类型获取初始对象集
        objects = self._get_objects_by_type(object_type)
        
        # 应用过滤器
        if filters:
            objects = self.apply_filters(objects, filters)
        
        # 处理value参数（明确区分索引和文本内容查询）
        if value is not None and value != "":
            # 如果明确指定为索引或者value是纯数字
            if treat_as_index or str(value).isdigit():
                try:
                    index = int(value)
                    # 统一使用1-based索引
                    if 0 < index <= len(objects):
                        return [objects[index - 1]]
                    else:
                        return []
                except ValueError:
                    # 如果无法转换为整数，不做特殊处理
                    pass
            
            # 如果value不是索引，则作为额外的文本过滤器
            # 但只有在元素类型支持文本内容时才这样做
            if object_type in ["paragraph", "comment", "text"]:
                text_filtered = self.apply_filters(objects, [{"contains_text": value}])
                return text_filtered if text_filtered else []
        
        return objects
        
    def _get_objects_by_type(self, object_type: str) -> List[CDispatch]:
        """根据类型获取对象集的辅助方法"""
        if object_type == "paragraph":
            return self.get_all_paragraphs()
        elif object_type == "table":
            return self.get_all_tables()
        elif object_type == "comment":
            return self.get_all_comments()
        elif object_type == "image" or object_type == "inline_shape":
            return self.get_all_images()
        else:
            return self.get_all_paragraphs()

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
        
    def suggest_best_locator(self, target_object: Any) -> Dict[str, Any]:
        """Suggest the best possible locator for a given object.

        Args:
            target_object: The object to generate a locator for.

        Returns:
            A normalized and validated locator dictionary.
        """
        parser = LocatorParser()
        
        # Determine object type
        object_type = self._determine_object_type(target_object)
        
        # Create basic locator
        locator = {
            "type": object_type,
            "value": "",
            "filters": [],
            "treat_as_index": False
        }
        
        # Try to find a unique identifier based on content
        if hasattr(target_object, 'Range') and hasattr(target_object.Range, 'Text'):
            text = target_object.Range.Text.strip()
            if text:
                # Use a snippet of text as part of the locator
                locator["value"] = text[:50]  # Limit to first 50 characters
        
        # Add style filter if applicable
        if hasattr(target_object, 'Style') and hasattr(target_object.Style, 'NameLocal'):
            style_name = target_object.Style.NameLocal
            if style_name:
                locator["filters"].append({"type": "has_style", "value": style_name})
        
        # Test if the current locator uniquely identifies the object
        test_locator = {
            "type": locator["type"],
            "value": locator["value"],
            "filters": locator["filters"],
            "treat_as_index": False
        }
        
        test_results = self.select_core(test_locator)
        
        # If not unique, use position-based locator
        if len(test_results) != 1:
            all_objects = self._get_objects_by_type(locator["type"])
            try:
                # Find the index of the target object
                for i, obj in enumerate(all_objects):
                    # Compare Range objects if available
                    if hasattr(obj, 'Range') and hasattr(target_object, 'Range'):
                        if (obj.Range.Start == target_object.Range.Start and 
                            obj.Range.End == target_object.Range.End):
                            locator["value"] = i + 1  # Use 1-based index
                            locator["treat_as_index"] = True
                            break
            except Exception:
                # Fallback to using basic locator
                pass
        
        return locator
        
    def _determine_object_type(self, obj: Any) -> str:
        """Determine the type of an object.

        Args:
            obj: The object to analyze.

        Returns:
            A string representing the object type.
        """
        try:
            if hasattr(obj, 'Rows') and hasattr(obj, 'Columns'):
                return 'table'
            elif hasattr(obj, 'Type') and obj.Type == 1:  # wdInlineShapePicture
                return 'image'
            elif hasattr(obj, 'Author'):  # Comments have Author property
                return 'comment'
            elif hasattr(obj, 'Range') and hasattr(obj.Range, 'Text'):
                return 'paragraph'
        except Exception:
            pass
        
        return 'paragraph'  # Default fallback
