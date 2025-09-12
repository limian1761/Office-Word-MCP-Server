"""
Selector Engine for Word Document MCP Server.

This module provides a powerful and flexible object selection system
for targeting specific objects in Word documents.
"""

import logging
import weakref
from typing import (TYPE_CHECKING, Any, Dict, Generic, List, Optional, TypeVar,
                    Union)

import win32com.client

from ..mcp_service.core_utils import (ErrorCode, ObjectNotFoundError,
                                      WordDocumentError)
from .exceptions import AmbiguousLocatorError, LocatorSyntaxError
from .filter_handlers import FilterHandlers
from .locator_parser import LocatorParser
from .object_finder import ObjectFinder
from .selection import Selection

logger = logging.getLogger(__name__)

# 定义类型变量
T = TypeVar("T")  # 通用类型变量
ObjectT = TypeVar("ObjectT", bound=win32com.client.CDispatch)  # 元素类型变量


class SelectorEngine:
    """
    Engine for selecting document objects based on locator queries.
    It parses a locator, finds matching COM objects, and returns them
    wrapped in a Selection object.
    """

    def __init__(self):
        """Initializes the selector engine."""
        self._filter_handlers = FilterHandlers()
        self._object_finder = ObjectFinder(self._filter_handlers)
        self._locator_parser = LocatorParser()
        # Simple cache for selections
        self._selection_cache: Dict[str, Selection] = {}

    def parse_locator(self, locator: str) -> Dict[str, Any]:
        """Parse locator string into components.

        Args:
            locator: Locator string in format "type:value[filters]".

        Returns:
            Dictionary with locator components.

        Raises:
            LocatorSyntaxError: If the locator format is invalid.

        For guidance on proper locator syntax, please refer to:
        word_docx_tools/selector/LOCATOR_GUIDE.md
        """
        return self._locator_parser.parse_locator(locator)

    def _get_cache_key(self, locator: str) -> str:
        """
        Generate a cache key for a given locator.

        Args:
            locator: The locator string.

        Returns:
            A unique cache key string.
        """
        # Delegate to the locator parser's get_cache_key method
        return self._locator_parser.get_cache_key(locator)

    def _validate_locator(self, locator: Dict[str, Any]) -> None:
        """
        Validate the locator structure.

        Raises:
            LocatorSyntaxError: If the locator structure is invalid.

        For guidance on proper locator syntax, please refer to:
        word_docx_tools/selector/LOCATOR_GUIDE.md
        """
        # 检查locator是否为字典类型
        if not isinstance(locator, dict):
            raise LocatorSyntaxError("Locator must be a dictionary.")

        # 检查是否包含type字段
        if "type" not in locator:
            raise LocatorSyntaxError("Locator must specify an object type.")

        object_type = locator["type"]

        # Check for required type field
        if not object_type:
            raise LocatorSyntaxError("Locator must specify an object type.")

        # Validate relation if anchor is specified
        if locator.get("anchor") is not None:
            valid_relations = [
                "all_occurrences_within",
                "first_occurrence_after",
                "parent_of",
                "immediately_following",
            ]
            if (
                locator.get("relation") is not None
                and locator["relation"] not in valid_relations
            ):
                raise LocatorSyntaxError(
                    f"Invalid relation '{locator['relation']}'. Valid relations are: {', '.join(valid_relations)}"
                )

    def select(
        self,
        document: win32com.client.CDispatch,
        locator: Dict[str, Any],
        expect_single: bool = False,
    ) -> Selection:
        """
        Selects objects in the document based on a locator query.
        This is the main entry point for the selector.

        Args:
            document: The Word document COM object.
            locator: The locator dictionary specifying what to select.
            expect_single: Whether to expect a single object.

        Returns:
            Selection object containing the matched objects.

        Raises:
            LocatorSyntaxError: If the locator syntax is invalid.
            ObjectNotFoundError: If no objects match the locator.
            AmbiguousLocatorError: If multiple objects match but only one was expected.

        For guidance on proper locator syntax, please refer to:
        word_docx_tools/selector/LOCATOR_GUIDE.md
        """
        # Try to get from cache first
        cache_key = self._get_cache_key(str(locator))
        if cache_key in self._selection_cache:
            return self._selection_cache[cache_key]

        # Parse the locator and validate syntax
        self._validate_locator(locator)

        # 如果locator是简单格式（不包含target键），则转换为完整格式
        if "target" not in locator:
            target_spec = locator
        else:
            target_spec = locator["target"]

        objects: List[Any]

        # Create a copy of target_spec to avoid modifying the original
        modified_target = target_spec.copy()

        # Handle 'text' object type with 'value' property
        is_text_type = target_spec.get("type") == "text"
        text_value = target_spec.get("value") if is_text_type else None

        # If 'text' type with 'value' is specified, convert to paragraph search with filter
        if is_text_type:
            # Change type to paragraph since we'll search in paragraphs
            modified_target["type"] = "paragraph"
            # Ensure filters list exists
            if "filters" not in modified_target:
                modified_target["filters"] = []
            # Add contains_text filter if value is provided
            if text_value:
                modified_target["filters"].append({"contains_text": text_value})

        # Create ObjectFinder instance
        object_finder = ObjectFinder(document)

        # Handle special locator types
        if "type" in locator:
            if locator["type"] == "document_start":
                # Get the start of the document
                start_range = document.Content
                start_range.Collapse(True)  # wdCollapseStart
                objects = [start_range]
            elif locator["type"] == "document_end":
                # Get the end of the document
                end_range = document.Content
                end_range.Collapse(False)  # wdCollapseEnd
                objects = [end_range]
            else:
                # If no anchor, perform a global search from the start of the document
                if "anchor" not in locator:
                    candidates = object_finder.get_initial_candidates(modified_target["type"])
                    
                    # Create a copy of the filters to avoid modifying the original
                    filters = modified_target.get("filters", []).copy()
                    
                    # Convert start and end parameters to range_start and range_end filters
                    if "start" in locator:
                        filters.append({"range_start": locator["start"]})
                    if "end" in locator:
                        filters.append({"range_end": locator["end"]})
                    
                    # Add index filter if value, index or id is specified for paragraph type
                    if modified_target.get("type") == "paragraph":
                        # Check for common position parameter names
                        for param_name in ["value", "index", "id"]:
                            if param_name in modified_target:
                                param_value = modified_target[param_name]
                                try:
                                    # Try to convert to integer (1-based index)
                                    index = int(param_value)
                                    if 0 < index <= len(candidates):
                                        # Return only the specified index
                                        objects = [candidates[index - 1]]
                                    else:
                                        # Index out of range
                                        objects = []
                                except (ValueError, TypeError):
                                    # Not an integer, proceed with filters
                                    pass
                                break
                        else:
                            # No valid index parameter found, apply filters normally
                            objects = object_finder.apply_filters(candidates, filters)
                    else:
                        # For non-paragraph types, apply filters normally
                        objects = object_finder.apply_filters(candidates, filters)
                else:
                    # If anchor and relation are present, perform a relational search
                    if "relation" not in locator:
                        raise LocatorSyntaxError(
                            "Locator with 'anchor' must also have a 'relation'."
                        )

                    # 1. Find the anchor object(s) first
                    anchor_spec = locator["anchor"]
                    anchor_object = object_finder.find_anchor(anchor_spec)

                    if not anchor_object:
                        raise ObjectNotFoundError(
                            {"anchor": anchor_spec},
                            f"Anchor object not found for: {anchor_spec}"
                        )

                    # 2. Perform the relational selection
                    relation = locator["relation"]
                    candidates = object_finder.get_initial_candidates(
                        modified_target["type"], within_range=anchor_object
                    )
                    objects = object_finder.select_relative_to_anchor(
                        candidates, anchor_object, relation
                    )
                    objects = object_finder.apply_filters(
                        objects, modified_target.get("filters", [])
                    )

        if not objects:
            raise ObjectNotFoundError(
                locator, f"No objects found for locator: {locator}."
            )

        if expect_single and len(objects) > 1:
            raise AmbiguousLocatorError(
                f"Expected 1 object but found {len(objects)} for locator: {locator}."
            )

        # Apply filters if they exist
        if "filters" in locator:
            objects = self._filter_handlers.apply_filters(objects, locator["filters"])

        # 转换所有对象为Range对象
        range_objects = []
        for obj in objects:
            range_obj = None

            # 1. 检查对象是否已经是Range对象（检查是否有Text、Start和End属性）
            if hasattr(obj, "Text") and hasattr(obj, "Start") and hasattr(obj, "End"):
                # 验证这是一个有效的Range对象
                try:
                    # 简单验证Range对象的有效性
                    _ = obj.Text
                    _ = obj.Start
                    _ = obj.End
                    range_obj = obj
                except Exception:
                    logger.warning(
                        f"Object of type {type(obj).__name__} has Range-like attributes but is not a valid Range"
                    )

            # 2. 如果不是Range对象，尝试获取其Range属性
            if range_obj is None and hasattr(obj, "Range"):
                try:
                    range_property = obj.Range
                    # 验证Range属性是否为有效的Range对象
                    if (
                        hasattr(range_property, "Text")
                        and hasattr(range_property, "Start")
                        and hasattr(range_property, "End")
                    ):
                        _ = range_property.Text
                        _ = range_property.Start
                        _ = range_property.End
                        range_obj = range_property
                except Exception:
                    logger.warning(
                        f"Failed to access valid Range property from object of type {type(obj).__name__}"
                    )

            # 3. 如果以上都不行，尝试基于对象位置创建一个Range
            if range_obj is None:
                try:
                    # 尝试创建一个基于对象位置的Range
                    if hasattr(obj, "Start") and hasattr(obj, "End"):
                        start = obj.Start
                        end = obj.End
                        if start != end:  # 确保Range不为空
                            range_obj = document.Range(Start=start, End=end)
                            # 验证创建的Range
                            _ = range_obj.Text
                    else:
                        # 如果无法获取位置信息，记录警告
                        logger.warning(
                            f"Object of type {type(obj).__name__} has no position information for Range conversion"
                        )
                except Exception as e:
                    logger.warning(f"Failed to create Range from object: {e}")

            # 4. 如果成功转换为Range对象，添加到结果列表
            if range_obj is not None:
                range_objects.append(range_obj)
            else:
                logger.error(
                    f"Failed to convert object of type {type(obj).__name__} to a valid Range object"
                )

        if not range_objects:
            raise ObjectNotFoundError(
                locator, f"No valid Range objects found for locator: {locator}."
            )

        # Cache the result
        selection = Selection(range_objects, document)
        self._selection_cache[cache_key] = selection

        return selection

    def get_all_paragraphs(self, document: Any = None) -> List[Any]:
        """Get all paragraphs in the document

        Args:
            document: The Word document object.

        Returns:
            List of all paragraphs.

        For guidance on proper locator syntax, please refer to:
        word_docx_tools/selector/LOCATOR_GUIDE.md
        """
        # 修复：正确处理document参数
        if document is None:
            raise ValueError("Document parameter is required")
        return list(document.Paragraphs)

    def get_all_tables(self, document: Any = None) -> List[Any]:
        """Get all tables in the document

        Args:
            document: The Word document object.

        Returns:
            List of all tables.

        For guidance on proper locator syntax, please refer to:
        word_docx_tools/selector/LOCATOR_GUIDE.md
        """
        # 修复：正确处理document参数
        if document is None:
            raise ValueError("Document parameter is required")
        return list(document.Tables)
        
    def suggest_locator(self, document: win32com.client.CDispatch, target_object: Any) -> Dict[str, Any]:
        """Suggest the best possible locator for a given document object.

        This method analyzes the target object and generates a locator that can be used
        to consistently identify it in the document, reducing randomness in locator generation.

        Args:
            document: The Word document COM object containing the target.
            target_object: The object to generate a locator for.

        Returns:
            A normalized and validated locator dictionary.
        
        Example:
            >>> selector = SelectorEngine()
            >>> paragraph = document.Paragraphs(1)
            >>> locator = selector.suggest_locator(document, paragraph)
            >>> # locator might look like: {"type": "paragraph", "value": 1, "treat_as_index": True}
        """
        # Create ObjectFinder instance
        object_finder = ObjectFinder(document)
        
        # Use ObjectFinder's suggest_best_locator method
        suggested_locator = object_finder.suggest_best_locator(target_object)
        
        # Validate the suggested locator to ensure it's correct
        try:
            self._validate_locator(suggested_locator)
        except LocatorSyntaxError as e:
            logger.warning(f"Suggested locator validation warning: {e}")
            # Still return the locator, but log the warning
        
        return suggested_locator
