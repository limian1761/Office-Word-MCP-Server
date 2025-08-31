"""
Selector Engine for Word Document MCP Server.

This module provides a powerful and flexible element selection system
for targeting specific elements in Word documents.
"""

import logging
import weakref
from typing import TYPE_CHECKING, Any, Dict, Generic, List, Optional, TypeVar, Union

import win32com.client

from word_document_server.selector.element_finder import ElementFinder
from word_document_server.selector.exceptions import (AmbiguousLocatorError,
                                                      LocatorSyntaxError)
from word_document_server.selector.filter_handlers import FilterHandlers
from word_document_server.selector.locator_parser import LocatorParser
from word_document_server.selector.selection import Selection
from word_document_server.utils.core_utils import (ElementNotFoundError,
                                                   ErrorCode,
                                                   WordDocumentError)

# 定义类型变量
T = TypeVar("T")  # 通用类型变量
ElementT = TypeVar("ElementT", bound=win32com.client.CDispatch)  # 元素类型变量


class SelectorEngine:
    """
    Engine for selecting document elements based on locator queries.
    It parses a locator, finds matching COM objects, and returns them
    wrapped in a Selection object.
    """

    def __init__(self):
        """Initializes the selector engine."""
        self._filter_handlers = FilterHandlers()
        self._element_finder = ElementFinder(self._filter_handlers)
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
        word_document_server/selector/LOCATOR_GUIDE.md
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
        word_document_server/selector/LOCATOR_GUIDE.md
        """
        # 检查locator是否为字典类型
        if not isinstance(locator, dict):
            raise LocatorSyntaxError("Locator must be a dictionary.")
            
        # 检查是否包含type字段
        if "type" not in locator:
            raise LocatorSyntaxError("Locator must specify an element type.")
            
        element_type = locator["type"]

        # Check for required type field
        if not element_type:
            raise LocatorSyntaxError("Locator must specify an element type.")

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
        Selects elements in the document based on a locator query.
        This is the main entry point for the selector.

        Args:
            document: The Word document COM object.
            locator: The locator dictionary specifying what to select.
            expect_single: Whether to expect a single element.

        Returns:
            Selection object containing the matched elements.

        Raises:
            LocatorSyntaxError: If the locator syntax is invalid.
            ElementNotFoundError: If no elements match the locator.
            AmbiguousLocatorError: If multiple elements match but only one was expected.

        For guidance on proper locator syntax, please refer to:
        word_document_server/selector/LOCATOR_GUIDE.md
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
            
        elements: List[Any]

        # Create a copy of target_spec to avoid modifying the original
        modified_target = target_spec.copy()

        # Handle 'text' element type with 'value' property
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

        # Create ElementFinder instance
        element_finder = ElementFinder(document)

        # If no anchor, perform a global search from the start of the document
        if "anchor" not in locator:
            candidates = element_finder.get_initial_candidates(modified_target["type"])
            elements = element_finder.apply_filters(candidates, modified_target.get("filters", []))
        else:
            # If anchor and relation are present, perform a relational search
            if "relation" not in locator:
                raise LocatorSyntaxError(
                    "Locator with 'anchor' must also have a 'relation'."
                )

            # 1. Find the anchor element(s) first
            anchor_spec = locator["anchor"]
            anchor_element = element_finder.find_anchor(anchor_spec)

            if not anchor_element:
                raise ElementNotFoundError(
                    {"anchor": anchor_spec}, 
                    f"Anchor element not found for: {anchor_spec}"
                )

            # 2. Perform the relational selection
            relation = locator["relation"]
            candidates = element_finder.get_initial_candidates(modified_target["type"], within_range=anchor_element)
            elements = element_finder.select_relative_to_anchor(
                candidates, anchor_element, relation
            )
            elements = element_finder.apply_filters(elements, modified_target.get("filters", []))

        if not elements:
            raise ElementNotFoundError(
                locator, 
                f"No elements found for locator: {locator}."
            )

        if expect_single and len(elements) > 1:
            raise AmbiguousLocatorError(
                f"Expected 1 element but found {len(elements)} for locator: {locator}."
            )

        # Apply filters if they exist
        if "filters" in locator:
            elements = self._filter_handlers.apply_filters(elements, locator["filters"])

        # Cache the result
        selection = Selection(elements, document)
        self._selection_cache[cache_key] = selection

        return selection

    def get_all_paragraphs(self, document: Any = None) -> List[Any]:
        """Get all paragraphs in the document

        Args:
            document: The Word document object.

        Returns:
            List of all paragraphs.

        For guidance on proper locator syntax, please refer to:
        word_document_server/selector/LOCATOR_GUIDE.md
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
        word_document_server/selector/LOCATOR_GUIDE.md
        """
        # 修复：正确处理document参数
        if document is None:
            raise ValueError("Document parameter is required")
        return list(document.Tables)
