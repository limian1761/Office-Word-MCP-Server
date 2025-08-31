"""Filter handlers for the selector engine.

This module contains all the filter implementations that can be applied to
document elements during selection.
"""

import re
from typing import Any, Dict, List

from word_document_server.selector.exceptions import LocatorSyntaxError
from word_document_server.utils.core_utils import get_shape_types


class FilterHandlers:
    """Collection of filter methods for document element selection."""

    def __init__(self):
        """Initialize the filter handlers with a map of filter names to functions."""
        # Map filter names to their corresponding functions
        self._filter_map = {
            "index": self._filter_by_index,
            "contains_text": self._filter_by_contains_text,
            "text_matches_regex": self._filter_by_text_matches_regex,
            "shape_type": self._filter_by_shape_type,
            "style": self._filter_by_style,
            "is_bold": self._filter_by_is_bold,
            "row_index": self._filter_by_row_index,
            "column_index": self._filter_by_column_index,
            "table_index": self._filter_by_table_index,
            "is_list_item": self._filter_by_is_list_item,
            "range_start": self._filter_by_range_start,
            "range_end": self._filter_by_range_end,
        }

    def apply_filters(self, elements: List[Any], filters: List[Dict[str, Any]]) -> List[Any]:
        """Applies a series of filters to a list of elements.
        
        Args:
            elements: List of elements to filter.
            filters: List of filter dictionaries.
            
        Returns:
            Filtered list of elements.
            
        Raises:
            LocatorSyntaxError: If filter format is invalid.
            
        For guidance on proper locator syntax, please refer to:
        word_document_server/selector/LOCATOR_GUIDE.md
        """
        filtered_list = list(elements)
        for f in filters:
            if not isinstance(f, dict) or len(f) != 1:
                raise LocatorSyntaxError(f"Invalid filter format: {f}.")

            filter_name, value = next(iter(f.items()))

            if filter_name not in self._filter_map:
                raise LocatorSyntaxError(f"Unsupported filter: {filter_name}.")

            filter_func = self._filter_map[filter_name]
            filtered_list = filter_func(filtered_list, value)

        return filtered_list

    def _filter_by_index(self, elements: List[Any], index: int) -> List[Any]:
        """Filters for a single element at a specific index, supporting negative indices.

        Args:
            elements: List of elements to filter.
            index: Index of the element to select.

        Returns:
            List containing the element at the specified index, or empty list.

        For guidance on proper locator syntax, please refer to:
        word_document_server/selector/LOCATOR_GUIDE.md
        """
        # For elements like images, there is no text, so we cannot filter by "meaningful" text content.
        # We will just use the raw element list for indexing.
        if not elements:
            return []

        effective_index = index
        if index < 0:
            effective_index = len(elements) + index

        if 0 <= effective_index < len(elements):
            return [elements[effective_index]]
        return []

    def _filter_by_contains_text(self, elements: List[Any], text: str) -> List[Any]:
        """Filters elements that contain the given text (case-insensitive).

        Args:
            elements: List of elements to filter.
            text: Text to search for.

        Returns:
            List of elements containing the specified text.

        For guidance on proper locator syntax, please refer to:
        word_document_server/selector/LOCATOR_GUIDE.md
        """
        return [
            el
            for el in elements
            if hasattr(el, "Range") and text.lower() in el.Range.Text.lower()
        ]

    def _filter_by_text_matches_regex(
        self, elements: List[Any], pattern: str
    ) -> List[Any]:
        """
        Filters elements whose text matches the given regex pattern.

        Args:
            elements: List of elements to filter.
            pattern: Regex pattern to match.

        Returns:
            List of elements matching the regex pattern.

        Raises:
            LocatorSyntaxError: If the regex pattern is invalid.

        For guidance on proper locator syntax, please refer to:
        word_document_server/selector/LOCATOR_GUIDE.md
        """
        try:
            # We use re.search to find a match anywhere in the text.
            # The text from a COM object's Range can have strange carriage returns,
            # so we strip whitespace to make matching more reliable.
            return [
                el
                for el in elements
                if hasattr(el, "Range") and re.search(pattern, el.Range.Text.strip())
            ]
        except re.error as e:
            raise LocatorSyntaxError(f"Invalid regex pattern '{pattern}': {e}.")

    def _filter_by_shape_type(self, elements: List[Any], shape_type: str) -> List[Any]:
        """
        Filters elements based on their shape type.

        Args:
            elements: List of inline shape elements to filter.
            shape_type: The shape type to filter by (e.g., "Picture", "Chart").

        Returns:
            List of elements matching the specified shape type.

        For guidance on proper locator syntax, please refer to:
        word_document_server/selector/LOCATOR_GUIDE.md
        """
        # Map of Word inline shape type constants to human-readable names

        shape_types = get_shape_types()

        # Find the type code that matches the shape_type string
        type_codes = [
            code
            for code, name in shape_types.items()
            if name.lower() == shape_type.lower()
        ]

        if not type_codes:
            # No matching type code found
            return []

        type_code = type_codes[0]  # Take the first matching type code

        return [el for el in elements if hasattr(el, "Type") and el.Type == type_code]

    def _filter_by_style(self, elements: List[Any], style_name: str) -> List[Any]:
        """Filters elements that have a specific style.

        Args:
            elements: List of elements to filter.
            style_name: Name of the style to filter by.

        Returns:
            List of elements with the specified style.

        For guidance on proper locator syntax, please refer to:
        word_document_server/selector/LOCATOR_GUIDE.md
        """
        return [
            el
            for el in elements
            if hasattr(el, "Style") and el.Style.NameLocal == style_name
        ]

    def _filter_by_is_bold(self, elements: List[Any], is_bold: bool) -> List[Any]:
        """Filters elements based on whether their font is bold.

        Args:
            elements: List of elements to filter.
            is_bold: Whether to filter for bold elements.

        Returns:
            List of elements with the specified bold formatting.

        For guidance on proper locator syntax, please refer to:
        word_document_server/selector/LOCATOR_GUIDE.md
        """
        bold_value = -1 if is_bold else 0
        return [
            el
            for el in elements
            if hasattr(el, "Range")
            and el.Range.Font.Bold == bold_value
            and "heading" not in el.Style.NameLocal.lower()
            and "标题" not in el.Style.NameLocal.lower()
        ]

    def _filter_by_row_index(self, elements: List[Any], index: int) -> List[Any]:
        """Filters for cells in a specific row.

        Args:
            elements: List of elements to filter.
            index: Row index to filter by.

        Returns:
            List of cells in the specified row.

        For guidance on proper locator syntax, please refer to:
        word_document_server/selector/LOCATOR_GUIDE.md
        """
        return [
            el for el in elements if hasattr(el, "RowIndex") and el.RowIndex == index
        ]

    def _filter_by_column_index(self, elements: List[Any], index: int) -> List[Any]:
        """Filters for cells in a specific column.

        Args:
            elements: List of elements to filter.
            index: Column index to filter by.

        Returns:
            List of cells in the specified column.

        For guidance on proper locator syntax, please refer to:
        word_document_server/selector/LOCATOR_GUIDE.md
        """
        return [
            el
            for el in elements
            if hasattr(el, "ColumnIndex") and el.ColumnIndex == index
        ]

    def _filter_by_table_index(self, elements: List[Any], index: int) -> List[Any]:
        """
        Filters for cells that belong to a specific table index.

        Args:
            elements: List of elements to filter.
            index: Table index to filter by.

        Returns:
            List of cells in the specified table.

        For guidance on proper locator syntax, please refer to:
        word_document_server/selector/LOCATOR_GUIDE.md
        """
        # This filter assumes elements are cells and checks if their parent table
        # is at the specified index in the document's tables collection
        return [
            el
            for el in elements
            if hasattr(el, "Parent")
            and hasattr(el.Parent, "Index")
            and el.Parent.Index == index + 1
        ]

    def _filter_by_is_list_item(self, elements: List[Any], is_list: bool) -> List[Any]:
        """
        Filters for paragraphs that are part of a list.

        Args:
            elements: List of elements to filter.
            is_list: Whether to filter for list items.

        Returns:
            List of elements that are or are not list items.

        For guidance on proper locator syntax, please refer to:
        word_document_server/selector/LOCATOR_GUIDE.md
        """
        if not is_list:
            # Return elements that are NOT list items
            return [
                el
                for el in elements
                if hasattr(el, "Range")
                and hasattr(el.Range, "ListFormat")
                and el.Range.ListFormat.ListString == ""
            ]

        # Return elements that ARE list items
        return [
            el
            for el in elements
            if hasattr(el, "Range")
            and hasattr(el.Range, "ListFormat")
            and el.Range.ListFormat.ListString != ""
        ]

    def _filter_by_range_start(self, elements: List[Any], start_pos: int) -> List[Any]:
        """Filters range elements by start position.

        Args:
            elements: List of elements to filter.
            start_pos: Start position to filter by.

        Returns:
            List of range elements starting at or after the specified position.

        For guidance on proper locator syntax, please refer to:
        word_document_server/selector/LOCATOR_GUIDE.md
        """
        # For range elements, adjust the start position
        filtered = []
        for el in elements:
            if hasattr(el, "Start") and hasattr(el, "End"):
                # Create a new range with adjusted start position
                try:
                    doc = el.Document
                    new_range = doc.Range(max(start_pos, el.Start), el.End)
                    filtered.append(new_range)
                except Exception:
                    # If we can't create a new range, keep the original
                    if el.Start <= start_pos <= el.End:
                        filtered.append(el)
        return filtered

    def _filter_by_range_end(self, elements: List[Any], end_pos: int) -> List[Any]:
        """Filters range elements by end position.

        Args:
            elements: List of elements to filter.
            end_pos: End position to filter by.

        Returns:
            List of range elements ending at or before the specified position.

        For guidance on proper locator syntax, please refer to:
        word_document_server/selector/LOCATOR_GUIDE.md
        """
        # For range elements, adjust the end position
        filtered = []
        for el in elements:
            if hasattr(el, "Start") and hasattr(el, "End"):
                # Create a new range with adjusted end position
                try:
                    doc = el.Document
                    new_range = doc.Range(el.Start, min(end_pos, el.End))
                    filtered.append(new_range)
                except Exception:
                    # If we can't create a new range, keep the original
                    if el.Start <= end_pos <= el.End:
                        filtered.append(el)
        return filtered