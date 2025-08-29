"""
Selector Engine for Word Document MCP Server.

This module implements the logic for locating document elements based on
the Locator query language defined in the architecture.
"""

import hashlib
import json
import re
from typing import Any, Callable, Dict, List, Optional

from word_document_server.utils.core_utils import ElementNotFoundError, WordDocumentError, ErrorCode
from word_document_server.selector.selection import Selection
from word_document_server.utils.core_utils import get_shape_types
from word_document_server.tools.image import get_color_type
from pywintypes import com_error


# Custom Exception types for clarity
class LocatorSyntaxError(ValueError):
    pass


class AmbiguousLocatorError(LookupError):
    """Raised when a locator expecting a single element finds multiple."""

    pass


class SelectorEngine:
    """
    Engine for selecting document elements based on locator queries.
    It parses a locator, finds matching COM objects, and returns them
    wrapped in a Selection object.
    """

    def __init__(self):
        """Initializes the selector engine."""
        self._filter_map: Dict[str, Callable] = {
            "contains_text": self._filter_by_contains_text,
            "text_matches_regex": self._filter_by_text_matches_regex,
            "style": self._filter_by_style,
            "is_bold": self._filter_by_is_bold,
            "index_in_parent": self._filter_by_index,
            "row_index": self._filter_by_row_index,
            "column_index": self._filter_by_column_index,
            "is_list_item": self._filter_by_is_list_item,
            "table_index": self._filter_by_table_index,
            "shape_type": self._filter_by_shape_type,
            "range_start": self._filter_by_range_start,
            "range_end": self._filter_by_range_end,
        }
        # Simple cache for selections
        self._selection_cache: Dict[str, Selection] = {}

    def parse_locator(self, locator: str) -> Dict[str, Any]:
        """Parse locator string into components.

        Args:
            locator: Locator string in format "type:value[filters]".

        Returns:
            Dictionary with locator components.
        """
        if ':' not in locator:
            raise LocatorSyntaxError(f"Invalid locator format: {locator}")
        locator_type, value = locator.split(':', 1)
        filters = {}
        if '[' in value and ']' in value:
            value, filter_str = value.split('[', 1)
            filter_str = filter_str.rstrip(']')
            for filter_part in filter_str.split(','):
                if '=' in filter_part:
                    key, val = filter_part.split('=', 1)
                    filters[key.strip()] = val.strip()
        return {
            'type': locator_type.strip(),
            'value': value.strip(),
            'filters': filters
        }

    def _get_cache_key(self, document: Any, locator: Dict[str, Any]) -> str:
        """
        Generate a cache key for a given document and locator.
        """
        # Create a unique key based on document path and locator
        try:
            doc_path = document.FullName if hasattr(document, 'FullName') else "new_document"
        except Exception:
            doc_path = "new_document"
        locator_str = json.dumps(locator, sort_keys=True)
        key_str = f"{doc_path}:{locator_str}"
        return hashlib.md5(key_str.encode()).hexdigest()

    def _validate_locator(self, locator: Dict[str, Any]) -> None:
        """
        Validate the locator structure.
        """
        if "target" not in locator:
            from word_document_server.utils.file_utils import get_doc_path
            raise LocatorSyntaxError("Locator must have a 'target'.")

    def select(
        self, document: Any, locator: Dict[str, Any], expect_single: bool = False
    ) -> Selection:
        """
        Selects elements in the document based on a locator query.
        This is the main entry point for the selector.
        """
        # Try to get from cache first
        cache_key = self._get_cache_key(document, locator)
        if cache_key in self._selection_cache:
            return self._selection_cache[cache_key]

        # Parse the locator and validate syntax
        self._validate_locator(locator)

        target_spec = locator["target"]
        elements: List[Any]

        # Create a copy of target_spec to avoid modifying the original
        modified_target = target_spec.copy()

        # Handle 'text' element type with 'value' property
        is_text_type = target_spec.get("type") == 'text'
        text_value = target_spec.get('value') if is_text_type else None

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

        # If no anchor, perform a global search from the start of the document
        if "anchor" not in locator:
            elements = self._select_core(document, modified_target)
        else:
            # If anchor and relation are present, perform a relational search
            if "relation" not in locator:
                raise LocatorSyntaxError(
                    "Locator with 'anchor' must also have a 'relation'."
                )

            # 1. Find the anchor element(s) first
            anchor_spec = locator["anchor"]
            anchor_element = self._find_anchor(document, anchor_spec)

            if not anchor_element:
                raise ElementNotFoundError(
                    f"Anchor element not found for: {anchor_spec}"
                )

            # 2. Perform the relational selection
            relation = locator["relation"]
            elements = self._select_relative_to_anchor(
                document, anchor_element, modified_target, relation
            )

        if not elements:
            raise ElementNotFoundError(f"No elements found for locator: {locator}.")

        if expect_single and len(elements) > 1:
            raise AmbiguousLocatorError(
                f"Expected 1 element but found {len(elements)} for locator: {locator}."
            )

        # Apply filters if they exist
        if "filters" in locator:
            elements = self._apply_filters(elements, locator["filters"])

        # Cache the result
        selection = Selection(elements, document)
        self._selection_cache[cache_key] = selection

        return selection

    def _find_anchor(
        self, document: Any, anchor_spec: Dict[str, Any]
    ) -> Optional[Any]:
        """Finds a single anchor element based on its specification."""
        if "type" not in anchor_spec:
            raise LocatorSyntaxError("Anchor spec must have a 'type'.")

        anchor_type = anchor_spec["type"]

        # Handle special anchor types
        if anchor_type == "start_of_document":
            return document.Range(0, 0)
        if anchor_type == "end_of_document":
            end_pos = document.Content.End
            return document.Range(end_pos, end_pos)

        # For object-based anchors, find all candidates and then filter
        identifier = anchor_spec.get("identifier", {})
        if not identifier:
            raise LocatorSyntaxError("Anchor spec must have an 'identifier'.")

        # Use _select_core to get candidates based on the anchor's type
        # We create a temporary "target" spec for this purpose
        anchor_target_spec = {"type": anchor_type, "filters": []}

        # Map identifier keys to filter names
        for key, value in identifier.items():
            if key == "text":
                anchor_target_spec["filters"].append({"contains_text": value})
            elif key == "index":
                anchor_target_spec["filters"].append({"index_in_parent": value})
            # Add other identifier mappings here (e.g., level for headings)

        anchor_candidates = self._select_core(document, anchor_target_spec)

        # For now, we always use the first matching anchor candidate
        return anchor_candidates[0] if anchor_candidates else None

    def _select_core(
        self,
        document: Any,
        target_spec: Dict[str, Any],
        search_range: Optional[Any] = None,
    ) -> List[Any]:
        """
        Core selection logic that finds elements of a given type, applies filters,
        and can operate within a specific search_range.
        """
        if "type" not in target_spec:
            raise LocatorSyntaxError("Target spec must have a 'type'.")

        element_type = target_spec["type"]
        filters = target_spec.get("filters", [])

        candidate_elements = self._get_initial_candidates(
            document, element_type, search_range
        )
        filtered_elements = self._apply_filters(candidate_elements, filters)

        return filtered_elements

    def _select_relative_to_anchor(
        self,
        document: Any,
        anchor_element: Any,
        target_spec: Dict[str, Any],
        relation: Dict[str, Any],
    ) -> List[Any]:
        """Handles the logic for finding elements based on a relation to an anchor."""
        relation_type = relation.get("type")

        if relation_type == "all_occurrences_within":
            # The search scope is the anchor element's own range
            anchor_range = anchor_element.Range
            return self._select_core(document, target_spec, search_range=anchor_range)

        elif relation_type == "first_occurrence_after":
            # The search scope is from the end of the anchor to the end of the document
            doc_end = document.Content.End
            search_range = document.Range(
                Start=anchor_element.Range.End, End=doc_end
            )

            # Find all occurrences after, then return the first one
            all_after = self._select_core(
                document, target_spec, search_range=search_range
            )
            return all_after[:1] if all_after else []

        elif relation_type == "parent_of":
            # The search scope is the parent of the anchor element
            parent_element = anchor_element.Parent
            parent_range = parent_element.Range
            return self._select_core(document, target_spec, search_range=parent_range)

        elif relation_type == "immediately_following":
            # Find the element that immediately follows the anchor element
            # We'll search for elements that start right after the anchor element ends
            anchor_end = anchor_element.Range.End
            # Create a minimal search range just after the anchor
            search_range = document.Range(Start=anchor_end, End=anchor_end + 1)

            # Find all occurrences after, then return the first one
            all_after = self._select_core(
                document, target_spec, search_range=search_range
            )
            return all_after[:1] if all_after else []

        else:
            raise LocatorSyntaxError(f"Unsupported relation type: {relation_type}.")

    def _get_initial_candidates(
        self,
        document: Any,
        element_type: str,
        search_range: Optional[Any] = None,
    ) -> List[Any]:
        """
        Gets the initial list of elements, either globally or from a specific range.
        """
        if search_range:
            return self._get_range_specific_candidates(document, element_type, search_range)
        return self._get_global_candidates(document, element_type)

    def _get_range_specific_candidates(self, document: Any, element_type: str, search_range: Any) -> List[Any]:
        """Get candidates from specific range"""
        active_doc = ctx.request_context.lifespan_context.get_active_document()
        element_handlers = {
            "paragraph": active_doc.get_paragraphs_in_range,
            "table": active_doc.get_tables_in_range,
            "cell": active_doc.get_cells_in_range,
            "run": active_doc.get_runs_in_range
        }

        if element_type in element_handlers:
            try:
                return element_handlers[element_type](search_range)
            except Exception as e:
                raise WordDocumentError(
                    ErrorCode.SERVER_ERROR, 
                    f"Failed to get {element_type}s in range: {e}"
                ) from e

        if element_type in ("inline_shape", "image"):
            return self._get_inline_shapes_in_range(active_doc, search_range)

        return []

    def _get_global_candidates(self, document: Any, element_type: str) -> List[Any]:
        """Get candidates from global document scope"""
        # Handle special element types
        if element_type == "document_start":
            return [document.Range(0, 0)]
        elif element_type == "document_end":
            end_pos = document.Content.End
            return [document.Range(end_pos, end_pos)]
        elif element_type == "range":
            # For range type, we return the entire document range
            # Actual start and end positions should be handled via filters
            return [document.Range(0, document.Content.End)]
        
        handlers = {
            "text": self.get_all_paragraphs,
            "paragraph": self.get_all_paragraphs,
            "table": self.get_all_tables
        }

        if element_type not in handlers:
            return []

        try:
            return handlers[element_type]()
        except Exception as e:
            error_codes = {
                "paragraph": ErrorCode.PARAGRAPH_SELECTION_FAILED,
                "document_start": ErrorCode.SERVER_ERROR,
                "document_end": ErrorCode.SERVER_ERROR
            }
            error_code = error_codes.get(element_type, ErrorCode.SERVER_ERROR)
            raise WordDocumentError(error_code, f"Failed to get {element_type} elements: {e}") from e

    def get_all_paragraphs(self) -> List[Any]:
        """Get all paragraphs in the document"""
        active_doc = ctx.request_context.lifespan_context.get_active_document()
        return list(active_doc.Paragraphs)

    def get_all_tables(self) -> List[Any]:
        """Get all tables in the document"""
        active_doc = ctx.request_context.lifespan_context.get_active_document()
        return list(active_doc.Tables)

    def _get_inline_shapes_in_range(self, active_doc, search_range: Any) -> List[Any]:
        """Get inline shapes within specified range"""
        try:
            inline_shapes = getattr(active_doc, 'InlineShapes', None)
            if not inline_shapes:
                return []

            all_shapes = []
            for i in range(1, inline_shapes.Count + 1):
                try:
                    shape = active_doc.InlineShapes(i)
                    all_shapes.append(shape)
                except com_error as e:
                    print(f"Warning: Failed to access shape at index {i}: {e}") 
            return [shape for shape in all_shapes if hasattr(shape, "Range") and 
                   shape.Range.Start >= search_range.Start and 
                   shape.Range.End <= search_range.End]
        except Exception as e:
            raise WordDocumentError(
                ErrorCode.SERVER_ERROR, 
                f"Failed to get inline shapes in range: {e}"
            ) 

    def _apply_filters(
        self, elements: List[Any], filters: List[Dict[str, Any]]
    ) -> List[Any]:
        """Applies a series of filters to a list of elements."""
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

    # --- Filter Implementations ---

    def _filter_by_index(self, elements: List[Any], index: int) -> List[Any]:
        """Filters for a single element at a specific index, supporting negative indices."""
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
        """Filters elements that contain the given text (case-insensitive)."""
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
        """Filters elements that have a specific style."""
        return [
            el
            for el in elements
            if hasattr(el, "Style") and el.Style.NameLocal == style_name
        ]

    def _filter_by_is_bold(self, elements: List[Any], is_bold: bool) -> List[Any]:
        """Filters elements based on whether their font is bold."""
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
        """Filters for cells in a specific row."""
        return [
            el for el in elements if hasattr(el, "RowIndex") and el.RowIndex == index
        ]

    def _filter_by_column_index(self, elements: List[Any], index: int) -> List[Any]:
        """Filters for cells in a specific column."""
        return [
            el
            for el in elements
            if hasattr(el, "ColumnIndex") and el.ColumnIndex == index
        ]

    def _filter_by_table_index(self, elements: List[Any], index: int) -> List[Any]:
        """
        Filters for cells that belong to a specific table index.
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
        """Filters range elements by start position."""
        # For range elements, adjust the start position
        filtered = []
        for el in elements:
            if hasattr(el, 'Start') and hasattr(el, 'End'):
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
        """Filters range elements by end position."""
        # For range elements, adjust the end position
        filtered = []
        for el in elements:
            if hasattr(el, 'Start') and hasattr(el, 'End'):
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
