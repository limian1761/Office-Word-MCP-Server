"""Object finder for the selector engine.

This module contains the functionality for finding and selecting
objects within a Word document.
"""

from typing import Any, Dict, List, Optional, Set, TypeVar, Union, cast

import win32com.client
from win32com.client import CDispatch

from .exceptions import AmbiguousLocatorError
from .filter_handlers import FilterHandlers

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
        elif object_type == "image":
            objects = self.get_all_images()
        else:
            # 默认返回所有段落
            objects = self.get_all_paragraphs()

        # 应用过滤器
        if filters:
            objects = self.apply_filters(objects, filters)

        # 如果有值（索引），返回特定元素
        if value:
            try:
                index = int(value)
                if 0 < index <= len(objects):
                    return [objects[index - 1]]
            except ValueError:
                pass

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
            return cast(CDispatch, self.document.Content)
        elif anchor_id == "document_end":
            end_range = self.document.Content
            end_range.Collapse(Direction=1)  # wdCollapseEnd
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
        if (
            object_type == "document"
            or object_type == "document_start"
            or object_type == "document_end"
        ):
            return [self.document.Content]

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
        paragraphs = self.document.Paragraphs
        return [paragraphs(i) for i in range(1, paragraphs.Count + 1)]

    def get_all_tables(self) -> List[CDispatch]:
        """Retrieve all tables in the document.

        Returns:
            A list of all tables in the document.
        """
        tables = self.document.Tables
        return [tables(i) for i in range(1, tables.Count + 1)]

    def _get_inline_shapes_in_range(self, range_obj: CDispatch) -> List[CDispatch]:
        """Get all inline shapes within a specific range.

        Args:
            range_obj: The range to search within.

        Returns:
            A list of inline shapes within the specified range.
        """
        shapes_in_range = []
        for i in range(1, self.document.InlineShapes.Count + 1):
            shape = self.document.InlineShapes(i)
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
        comments = self.document.Comments
        return [comments(i) for i in range(1, comments.Count + 1)]

    def get_all_images(self) -> List[CDispatch]:
        """Retrieve all images in the document.

        Returns:
            A list of all images in the document.
        """
        # Images in Word are represented as InlineShapes
        shapes = self.document.InlineShapes
        # Filter to include only pictures
        images = []
        for i in range(1, shapes.Count + 1):
            shape = shapes(i)
            # Check if the shape is a picture
            if shape.Type == 1:  # wdInlineShapePicture
                images.append(shape)
        return images
