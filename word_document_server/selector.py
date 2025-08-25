"""
Selector Engine for Word Document MCP Server.

This module implements the logic for locating document elements based on
the Locator query language defined in the architecture.
"""
import re
from typing import Any, Callable, Dict, List, Optional

import win32com.client

from word_document_server.com_backend import WordBackend
from word_document_server.errors import WordDocumentError, ElementNotFoundError, ErrorCode
from word_document_server.selection import Selection


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
        }

    def select(self, backend: WordBackend, locator: Dict[str, Any], expect_single: bool = False) -> Selection:
        """
        Selects elements in the document based on a locator query.
        This is the main entry point for the selector.
        """
        if "target" not in locator:
            raise LocatorSyntaxError("Locator must have a 'target'. Please refer to docs/locator_guide.md for proper usage.")

        target_spec = locator["target"]
        elements: List[Any]

        # If no anchor, perform a global search from the start of the document
        if "anchor" not in locator:
            elements = self._select_core(backend, target_spec)
        else:
            # If anchor and relation are present, perform a relational search
            if "relation" not in locator:
                raise LocatorSyntaxError("Locator with 'anchor' must also have a 'relation'. Please refer to locator_guide.md for proper usage.")

            # 1. Find the anchor element(s) first
            anchor_spec = locator["anchor"]
            anchor_element = self._find_anchor(backend, anchor_spec)
            
            if not anchor_element:
                raise ElementNotFoundError(f"Anchor element not found for: {anchor_spec}")
            
            # 2. Perform the relational selection
            relation = locator["relation"]
            elements = self._select_relative_to_anchor(backend, anchor_element, target_spec, relation)

        if not elements:
            raise ElementNotFoundError(f"No elements found for locator: {locator}. Please refer to locator_guide.md for proper usage.")
        
        if expect_single and len(elements) > 1:
            raise AmbiguousLocatorError(f"Expected 1 element but found {len(elements)} for locator: {locator}. Please refer to locator_guide.md for proper usage.")

        return Selection(elements, backend)

    def _find_anchor(self, backend: WordBackend, anchor_spec: Dict[str, Any]) -> Optional[Any]:
        """Finds a single anchor element based on its specification."""
        if "type" not in anchor_spec:
            raise LocatorSyntaxError("Anchor spec must have a 'type'. Please refer to locator_guide.md for proper usage.")
        
        anchor_type = anchor_spec["type"]
        
        # Handle special anchor types
        if anchor_type == "start_of_document":
            return backend.document.Range(0, 0)
        if anchor_type == "end_of_document":
            end_pos = backend.document.Content.End
            return backend.document.Range(end_pos, end_pos)

        # For object-based anchors, find all candidates and then filter
        identifier = anchor_spec.get("identifier", {})
        if not identifier:
            raise LocatorSyntaxError("Anchor spec must have an 'identifier'. Please refer to locator_guide.md for proper usage.")

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

        anchor_candidates = self._select_core(backend, anchor_target_spec)
        
        # For now, we always use the first matching anchor candidate
        return anchor_candidates[0] if anchor_candidates else None

    def _select_core(self, backend: WordBackend, target_spec: Dict[str, Any], search_range: Optional[Any] = None) -> List[Any]:
        """
        Core selection logic that finds elements of a given type, applies filters,
        and can operate within a specific search_range.
        """
        if "type" not in target_spec:
            raise LocatorSyntaxError("Target spec must have a 'type'. Please refer to locator_guide.md for proper usage.")
            
        element_type = target_spec["type"]
        filters = target_spec.get("filters", [])

        candidate_elements = self._get_initial_candidates(backend, element_type, search_range)
        filtered_elements = self._apply_filters(candidate_elements, filters)
        
        return filtered_elements

    def _select_relative_to_anchor(self, backend: WordBackend, anchor_element: Any, target_spec: Dict[str, Any], relation: Dict[str, Any]) -> List[Any]:
        """Handles the logic for finding elements based on a relation to an anchor."""
        relation_type = relation.get("type")
        
        if relation_type == "all_occurrences_within":
            # The search scope is the anchor element's own range
            anchor_range = anchor_element.Range
            return self._select_core(backend, target_spec, search_range=anchor_range)
            
        elif relation_type == "first_occurrence_after":
            # The search scope is from the end of the anchor to the end of the document
            doc_end = backend.document.Content.End
            search_range = backend.document.Range(Start=anchor_element.Range.End, End=doc_end)
            
            # Find all occurrences after, then return the first one
            all_after = self._select_core(backend, target_spec, search_range=search_range)
            return all_after[:1] if all_after else []
            
        elif relation_type == "parent_of":
            # The search scope is the parent of the anchor element
            parent_element = anchor_element.Parent
            parent_range = parent_element.Range
            return self._select_core(backend, target_spec, search_range=parent_range)
            
        elif relation_type == "immediately_following":
            # Find the element that immediately follows the anchor element
            # We'll search for elements that start right after the anchor element ends
            anchor_end = anchor_element.Range.End
            # Create a minimal search range just after the anchor
            search_range = backend.document.Range(Start=anchor_end, End=anchor_end + 1)
            
            # Find all occurrences after, then return the first one
            all_after = self._select_core(backend, target_spec, search_range=search_range)
            return all_after[:1] if all_after else []

        else:
            raise LocatorSyntaxError(f"Unsupported relation type: {relation_type}. Please refer to locator_guide.md for proper usage.")

    def _get_initial_candidates(self, backend: WordBackend, element_type: str, search_range: Optional[Any] = None) -> List[Any]:
        """
        Gets the initial list of elements, either globally or from a specific range.
        """
        candidates = []
        
        if search_range:
            # Range-specific search
            if element_type == "paragraph":
                try:
                    candidates = backend.get_paragraphs_in_range(search_range)
                except Exception as e:
                    raise WordDocumentError(f"Failed to get paragraphs in range: {e}")
            elif element_type == "table":
                try:
                    candidates = backend.get_tables_in_range(search_range)
                except Exception as e:
                    raise WordDocumentError(f"Failed to get tables in range: {e}")
            elif element_type == "cell":
                try:
                    candidates = backend.get_cells_in_range(search_range)
                except Exception as e:
                    raise WordDocumentError(f"Failed to get cells in range: {e}")
            elif element_type == "run":
                try:
                    candidates = backend.get_runs_in_range(search_range)
                except Exception as e:
                    raise WordDocumentError(f"Failed to get runs in range: {e}")
            elif element_type == "inline_shape" or element_type == "image":
                # For images in a specific range, iterate through all inline shapes and check if they are within the range
                try:
                    doc_range = backend.document.Range(0, backend.document.Content.End)
                    all_shapes = []
                    # Safely get all inline shapes
                    if hasattr(backend.document, 'InlineShapes') and backend.document.InlineShapes is not None:
                        for i in range(1, backend.document.InlineShapes.Count + 1):
                            try:
                                shape = backend.document.InlineShapes(i)
                                all_shapes.append(shape)
                            except Exception as e:
                                print(f"Warning: Failed to access shape at index {i}: {e}")
                                continue
                    # Filter shapes that are within the search range
                    candidates = [shape for shape in all_shapes if hasattr(shape, 'Range') and shape.Range.Start >= search_range.Start and shape.Range.End <= search_range.End]
                except Exception as e:
                    raise WordDocumentError(f"Failed to get inline shapes in range: {e}")
        else:
            # Global search
            if element_type == "paragraph":
                try:
                    candidates = backend.get_all_paragraphs()
                except Exception as e:
                     from word_document_server.errors import ErrorCode
                     raise WordDocumentError(ErrorCode.PARAGRAPH_SELECTION_FAILED, f"Failed to get all paragraphs: {e}")
            elif element_type == "table":
                try:
                    candidates = backend.get_all_tables()
                except Exception as e:
                    raise WordDocumentError(f"Failed to get all tables: {e}")
            elif element_type == "heading":
                try:
                    all_paragraphs = backend.get_all_paragraphs()
                    candidates = [p for p in all_paragraphs if hasattr(p, 'Style') and hasattr(p.Style, 'NameLocal') and (p.Style.NameLocal.startswith("Heading") or p.Style.NameLocal.startswith("标题"))]
                except Exception as e:
                    raise WordDocumentError(f"Failed to get heading paragraphs: {e}")
            elif element_type == "cell":
                try:
                    doc_range = backend.document.Range(0, backend.document.Content.End)
                    candidates = backend.get_cells_in_range(doc_range)
                except Exception as e:
                    raise WordDocumentError(f"Failed to get all cells: {e}")
            elif element_type == "run":
                try:
                    doc_range = backend.document.Range(0, backend.document.Content.End)
                    candidates = backend.get_runs_in_range(doc_range)
                except Exception as e:
                    raise WordDocumentError(f"Failed to get all runs: {e}")
            elif element_type == "inline_shape" or element_type == "image":
                # Get all inline shapes (including images) in the document
                try:
                    candidates = []
                    if hasattr(backend.document, 'InlineShapes') and backend.document.InlineShapes is not None:
                        for i in range(1, backend.document.InlineShapes.Count + 1):
                            try:
                                shape = backend.document.InlineShapes(i)
                                candidates.append(shape)
                            except Exception as e:
                                print(f"Warning: Failed to access shape at index {i}: {e}")
                                continue
                except Exception as e:
                    raise WordDocumentError(f"Failed to get all inline shapes: {e}")

        if not candidates and element_type not in ["paragraph", "table", "heading", "cell", "run", "inline_shape", "image"]:
             raise LocatorSyntaxError(f"Unsupported element type: {element_type}. Please refer to locator_guide.md for proper usage.")
        
        return candidates

    def _apply_filters(self, elements: List[Any], filters: List[Dict[str, Any]]) -> List[Any]:
        """Applies a series of filters to a list of elements."""
        filtered_list = list(elements)
        for f in filters:
            if not isinstance(f, dict) or len(f) != 1:
                raise LocatorSyntaxError(f"Invalid filter format: {f}. Please refer to locator_guide.md for proper usage.")
            
            filter_name, value = next(iter(f.items()))
            
            if filter_name not in self._filter_map:
                raise LocatorSyntaxError(f"Unsupported filter: {filter_name}. Please refer to locator_guide.md for proper usage.")
            
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
        return [el for el in elements if hasattr(el, 'Range') and text.lower() in el.Range.Text.lower()]

    def _filter_by_text_matches_regex(self, elements: List[Any], pattern: str) -> List[Any]:
        """
        Filters elements whose text matches the given regex pattern.
        """
        try:
            # We use re.search to find a match anywhere in the text.
            # The text from a COM object's Range can have strange carriage returns,
            # so we strip whitespace to make matching more reliable.
            return [
                el for el in elements 
                if hasattr(el, 'Range') and re.search(pattern, el.Range.Text.strip())
            ]
        except re.error as e:
            raise LocatorSyntaxError(f"Invalid regex pattern '{pattern}': {e}. Please refer to locator_guide.md for proper usage.")
            
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
        shape_types = {
            1: "Picture",
            2: "LinkedPicture",
            3: "Chart",
            4: "Diagram",
            5: "OLEControlObject",
            6: "OLEObject",
            7: "ActiveXControl",
            8: "SmartArt",
            9: "3DModel"
        }
        
        # Find the type code that matches the shape_type string
        type_codes = [code for code, name in shape_types.items() if name.lower() == shape_type.lower()]
        
        if not type_codes:
            # No matching type code found
            return []
            
        type_code = type_codes[0]  # Take the first matching type code
        
        return [el for el in elements if hasattr(el, 'Type') and el.Type == type_code]

    def _filter_by_style(self, elements: List[Any], style_name: str) -> List[Any]:
        """Filters elements that have a specific style."""
        return [el for el in elements if hasattr(el, 'Style') and el.Style.NameLocal == style_name]

    def _filter_by_is_bold(self, elements: List[Any], is_bold: bool) -> List[Any]:
        """Filters elements based on whether their font is bold."""
        bold_value = -1 if is_bold else 0
        return [
            el for el in elements 
            if hasattr(el, 'Range') and el.Range.Font.Bold == bold_value and 
               "heading" not in el.Style.NameLocal.lower() and "标题" not in el.Style.NameLocal.lower()
        ]

    def _filter_by_row_index(self, elements: List[Any], index: int) -> List[Any]:
        """Filters for cells in a specific row."""
        return [el for el in elements if hasattr(el, 'RowIndex') and el.RowIndex == index]

    def _filter_by_column_index(self, elements: List[Any], index: int) -> List[Any]:
        """Filters for cells in a specific column."""
        return [el for el in elements if hasattr(el, 'ColumnIndex') and el.ColumnIndex == index]

    def _filter_by_table_index(self, elements: List[Any], index: int) -> List[Any]:
        """
        Filters for cells that belong to a specific table index.
        """
        # This filter assumes elements are cells and checks if their parent table
        # is at the specified index in the document's tables collection
        return [el for el in elements if hasattr(el, 'Parent') and hasattr(el.Parent, 'Index') and el.Parent.Index == index + 1]

    def _filter_by_is_list_item(self, elements: List[Any], is_list: bool) -> List[Any]:
        """
        Filters for paragraphs that are part of a list.
        """
        if not is_list:
            # Return elements that are NOT list items
            return [
                el for el in elements 
                if hasattr(el, 'Range') and hasattr(el.Range, 'ListFormat') and el.Range.ListFormat.ListString == ''
            ]
        
        # Return elements that ARE list items
        return [
            el for el in elements 
            if hasattr(el, 'Range') and hasattr(el.Range, 'ListFormat') and el.Range.ListFormat.ListString != ''
        ]
