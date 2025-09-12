"""Locator parser for the selector engine.

This module contains the functionality for parsing and validating
locator strings used to identify objects in a Word document.
"""

import re
from typing import Any, Dict, List, Optional, Tuple, Union

from .exceptions import LocatorSyntaxError


class LocatorParser:
    """Component responsible for parsing and validating locator strings."""

    def parse_locator(self, locator: str) -> Dict[str, Any]:
        """Parse a locator string into its components.

        Args:
            locator: The locator string to parse.

        Returns:
            A dictionary containing the parsed locator components.

        Raises:
            LocatorSyntaxError: If the locator syntax is invalid.
        """
        # Strip whitespace and validate input
        locator = locator.strip()
        if not locator:
            raise LocatorSyntaxError("Locator cannot be empty.")

        # Default parsed locator structure
        parsed_locator = {
            "type": "",
            "value": "",
            "filters": [],
            "anchor": None,
            "relation": None,
        }

        # First, check if the locator has an anchor relation
        anchor_match = re.match(r"^([^@]+)@([^\[]+)(\[.*\])?$", locator)
        if anchor_match:
            # Handle anchor relation format: type:value@anchor_id[relation]
            main_part = anchor_match.group(1).strip()
            anchor_part = anchor_match.group(2).strip()
            relation_part = anchor_match.group(3)

            # Parse the relation if it exists
            if relation_part:
                relation = relation_part.strip("[]").strip()
                if relation:
                    parsed_locator["relation"] = relation

            # Parse the anchor
            parsed_locator["anchor"] = anchor_part

            # Parse the main part (type:value)
            main_parts = self._parse_main_locator(main_part)
            parsed_locator.update(main_parts)
        else:
            # Parse the standard locator format: type:value[filter1][filter2]...
            # First, separate the main part from filters
            main_part = locator
            filters = []

            # Extract all filter parts using regex
            filter_matches = re.finditer(r"\[(.*?)\]", locator)
            for match in filter_matches:
                filter_content = match.group(1).strip()
                if filter_content:
                    filters.append(filter_content)
                # Remove the filter from the main part
                main_part = main_part.replace(match.group(0), "", 1)

            main_part = main_part.strip()
            if not main_part:
                raise LocatorSyntaxError("Locator must specify an object type.")

            # Parse the main part (type:value)
            main_parts = self._parse_main_locator(main_part)
            parsed_locator.update(main_parts)

            # Parse filters
            parsed_filters = self._parse_filters(filters)
            parsed_locator["filters"] = parsed_filters

        # Validate the parsed locator
        self._validate_locator(parsed_locator)

        return parsed_locator

    def validate_locator(self, parsed_locator: Dict[str, Any]) -> None:
        """Public method for validating a parsed locator dictionary.

        This method provides a public interface for validating locators,
        delegating to the internal _validate_locator method.

        Args:
            parsed_locator: The parsed locator dictionary to validate.

        Raises:
            LocatorSyntaxError: If the locator is invalid.
        """
        self._validate_locator(parsed_locator)

    def _parse_main_locator(self, main_part: str) -> Dict[str, str]:
        """Parse the main part of a locator (type:value).

        Args:
            main_part: The main part of the locator string.

        Returns:
            A dictionary containing the parsed type and value.

        Raises:
            LocatorSyntaxError: If the main part syntax is invalid.
        """
        # Split on the first occurrence of colon
        if ":" in main_part:
            type_part, value_part = main_part.split(":", 1)
            return {"type": type_part.strip(), "value": value_part.strip()}
        else:
            # If no colon, assume it's just the type with an empty value
            return {"type": main_part.strip(), "value": ""}

    def _parse_filters(self, filter_strings: List[str]) -> List[Dict[str, Any]]:
        """Parse a list of filter strings into filter definitions.

        Args:
            filter_strings: A list of filter strings to parse.

        Returns:
            A list of parsed filter definitions.
        """
        parsed_filters = []

        for filter_str in filter_strings:
            # Handle key=value format
            if "=" in filter_str:
                key, value = filter_str.split("=", 1)
                key = key.strip()
                value = value.strip().strip('"').strip("'")

                # Try to convert to appropriate type
                parsed_value: Union[str, bool, int]
                try:
                    # Convert to boolean if applicable
                    if value.lower() == "true":
                        parsed_value = True
                    elif value.lower() == "false":
                        parsed_value = False
                    # Convert to integer if applicable
                    elif value.isdigit() or (
                        value.startswith("-") and value[1:].isdigit()
                    ):
                        parsed_value = int(value)
                    # Otherwise keep as string
                    else:
                        parsed_value = value
                except (ValueError, TypeError):
                    parsed_value = value

                parsed_filters.append({"type": key, "value": parsed_value})
            else:
                # Handle single keyword filters
                parsed_filters.append({"type": filter_str.strip(), "value": True})

        return parsed_filters

    def _validate_locator(self, parsed_locator: Dict[str, Any]) -> None:
        """Validate a parsed locator dictionary with enhanced constraints.

        Args:
            parsed_locator: The parsed locator dictionary to validate.

        Raises:
            LocatorSyntaxError: If the locator is invalid.
        """
        # Check if parsed_locator is a dictionary
        if not isinstance(parsed_locator, dict):
            raise LocatorSyntaxError(f"Locator must be a dictionary, got {type(parsed_locator).__name__}.")

        # Validate basic structure and required fields
        required_fields = ["type"]
        for field in required_fields:
            if field not in parsed_locator:
                raise LocatorSyntaxError(f"Required field '{field}' missing in locator.")

        object_type = parsed_locator["type"]

        # Check for required type field
        if not object_type:
            raise LocatorSyntaxError("Locator must specify an object type.")

        # Validate object type
        valid_object_types = [
            "paragraph", "table", "cell", "inline_shape", 
            "image", "comment", "range", "selection", 
            "document", "document_start", "document_end"
        ]
        if object_type not in valid_object_types:
            raise LocatorSyntaxError(
                f"Invalid object type '{object_type}'. Valid types are: {', '.join(valid_object_types)}"
            )

        # Validate relation if anchor is specified
        if parsed_locator.get("anchor") is not None:
            valid_relations: List[str] = [
                "all_occurrences_within",
                "first_occurrence_after",
                "parent_of",
                "immediately_following",
            ]
            if (
                parsed_locator.get("relation") is not None
                and parsed_locator["relation"] not in valid_relations
            ):
                raise LocatorSyntaxError(
                    f"Invalid relation '{parsed_locator['relation']}'. Valid relations are: {', '.join(valid_relations)}"
                )
            
            # Ensure relation is provided with anchor
            if "relation" not in parsed_locator:
                raise LocatorSyntaxError("Locator with 'anchor' must also specify a 'relation'.")

        # Validate filters format if present
        if "filters" in parsed_locator:
            if not isinstance(parsed_locator["filters"], list):
                raise LocatorSyntaxError("'filters' must be a list.")
            
            # Validate each filter in the filters list
            valid_filter_types = [
                "index", "contains_text", "text_matches_regex", 
                "shape_type", "style", "is_bold", "row_index", 
                "column_index", "table_index", "is_list_item", 
                "range_start", "range_end", "has_style"
            ]
            
            for i, filter_item in enumerate(parsed_locator["filters"]):
                if not isinstance(filter_item, dict) or len(filter_item) != 1:
                    raise LocatorSyntaxError(f"Filter at index {i} must be a single key-value pair dictionary.")
                
                filter_name = next(iter(filter_item.keys()))
                if filter_name not in valid_filter_types:
                    raise LocatorSyntaxError(
                        f"Invalid filter type '{filter_name}' at index {i}. "
                        f"Valid filter types are: {', '.join(valid_filter_types)}"
                    )

        # Type-specific validations
        if object_type == "paragraph":
            # For paragraphs, if value is numeric and treat_as_index is True,
            # ensure it's a positive integer
            value = parsed_locator.get("value", "")
            treat_as_index = parsed_locator.get("treat_as_index", False)
            
            if treat_as_index and value:
                try:
                    index_value = int(str(value))
                    if index_value <= 0:
                        raise LocatorSyntaxError(f"Paragraph index must be a positive integer, got {index_value}.")
                except ValueError:
                    raise LocatorSyntaxError(
                        f"Cannot treat paragraph value '{value}' as index - must be a number."
                    )

        elif object_type == "table":
            # For tables, if value is numeric, ensure it's a positive integer
            value = parsed_locator.get("value", "")
            if value and str(value).isdigit():
                table_index = int(str(value))
                if table_index <= 0:
                    raise LocatorSyntaxError(f"Table index must be a positive integer, got {table_index}.")

        elif object_type in ["document_start", "document_end"]:
            # For document_start and document_end, ensure no conflicting parameters
            if parsed_locator.get("value") or parsed_locator.get("filters"):
                raise LocatorSyntaxError(
                    f"'{object_type}' cannot have 'value' or 'filters' parameters."
                )
        
    def _normalize_locator(self, parsed_locator: Dict[str, Any]) -> Dict[str, Any]:
        """Normalize a parsed locator to ensure consistent format.

        Args:
            parsed_locator: The parsed locator dictionary to normalize.

        Returns:
            A normalized locator dictionary.
        """
        normalized = parsed_locator.copy()
        
        # Ensure type field exists and is lowercase
        normalized["type"] = normalized.get("type", "").lower()
        
        # Ensure filters field is a list
        if "filters" not in normalized:
            normalized["filters"] = []
        elif not isinstance(normalized["filters"], list):
            normalized["filters"] = []
        
        # Add treat_as_index flag for numeric values
        if "value" in normalized and normalized["value"] and str(normalized["value"]).isdigit():
            # If value is a number, add treat_as_index flag if not present
            if "treat_as_index" not in normalized:
                normalized["treat_as_index"] = True
        
        # Ensure anchor and relation fields are properly set
        if "anchor" not in normalized:
            normalized["anchor"] = None
        if "relation" not in normalized:
            normalized["relation"] = None
        
        return normalized
        
    def suggest_locator(self, target_object: Any) -> Dict[str, Any]:
        """Generate a suggested locator for a target object.

        Args:
            target_object: The object for which to generate a locator.

        Returns:
            A suggested locator dictionary.
        """
        locator = {
            "type": self._determine_object_type(target_object),
            "value": self._generate_content_identifier(target_object),
            "filters": self._generate_additional_filters(target_object),
            "anchor": None,
            "relation": None
        }
        
        return self._normalize_locator(locator)
        
    def _determine_object_type(self, obj: Any) -> str:
        """Determine the type of an object.

        Args:
            obj: The object to analyze.

        Returns:
            A string representing the object type.
        """
        try:
            # This implementation depends on the actual object types in your system
            # Adjust according to your Word document object model
            if hasattr(obj, 'Range') and hasattr(obj.Range, 'Text'):
                return 'paragraph'
            elif hasattr(obj, 'Rows') and hasattr(obj, 'Columns'):
                return 'table'
            elif hasattr(obj, 'Width') and hasattr(obj, 'Height'):
                return 'image'
            elif hasattr(obj, 'Range') and hasattr(obj, 'Author'):
                return 'comment'
        except Exception:
            pass
        
        return 'paragraph'  # Default fallback
        
    def _generate_content_identifier(self, obj: Any) -> str:
        """Generate a content-based identifier for an object.

        Args:
            obj: The object to analyze.

        Returns:
            A string identifier based on content.
        """
        try:
            if hasattr(obj, 'Range') and hasattr(obj.Range, 'Text'):
                # For text objects, return a snippet of text
                text = obj.Range.Text.strip()
                return text[:50] if text else ''
        except Exception:
            pass
        
        return ''
        
    def _generate_additional_filters(self, obj: Any) -> List[Dict[str, Any]]:
        """Generate additional filters to uniquely identify an object.

        Args:
            obj: The object to analyze.

        Returns:
            A list of additional filter dictionaries.
        """
        filters = []
        
        try:
            # Add style filter if applicable
            if hasattr(obj, 'Style') and hasattr(obj.Style, 'NameLocal'):
                style_name = obj.Style.NameLocal
                if style_name:
                    filters.append({'type': 'has_style', 'value': style_name})
        except Exception:
            pass
        
        return filters
        
    def enhance_parse_locator(self, locator: Union[str, Dict[str, Any]]) -> Dict[str, Any]:
        """Enhanced locator parsing with strict constraints for AI-generated locator parameters.

        Args:
            locator: Either a locator string or a pre-parsed locator dictionary.

        Returns:
            A strictly validated and normalized locator dictionary.
        """
        # Basic type validation
        if not isinstance(locator, (str, dict)):
            raise LocatorSyntaxError(
                f"Locator must be a string or dictionary, got {type(locator).__name__}."
            )

        if isinstance(locator, dict):
            # Strict validation for dictionary type locators
            parsed = locator.copy()
            
            # Enforce required fields exist with proper types
            required_fields = ["type"]
            for field in required_fields:
                if field not in parsed:
                    raise LocatorSyntaxError(f"Required field '{field}' missing in locator.")
            
            # Validate type field is non-empty string
            if not isinstance(parsed["type"], str) or not parsed["type"].strip():
                raise LocatorSyntaxError("'type' field must be a non-empty string.")
            
            # Ensure optional fields are properly formatted if present
            optional_fields = {
                "value": (str, int, type(None)),
                "filters": (list, type(None)),
                "anchor": (dict, type(None)),
                "relation": (str, type(None)),
                "treat_as_index": (bool, type(None))
            }
            
            for field, allowed_types in optional_fields.items():
                if field in parsed:
                    if not isinstance(parsed[field], allowed_types):
                        type_names = [t.__name__ for t in allowed_types if t is not type(None)]
                        if type(None) in allowed_types:
                            type_names.append("None")
                        raise LocatorSyntaxError(
                            f"Field '{field}' must be one of: {', '.join(type_names)}, got {type(parsed[field]).__name__}."
                        )
            
            # Ensure filters is initialized as empty list if None
            if parsed.get("filters") is None:
                parsed["filters"] = []
            
            # Strict validation of anchor-relation pair
            if parsed.get("anchor") is not None:
                if parsed.get("relation") is None:
                    raise LocatorSyntaxError(
                        "Locator with 'anchor' must also specify a 'relation'."
                    )
                    
                # Validate anchor structure
                if not isinstance(parsed["anchor"], dict):
                    raise LocatorSyntaxError("'anchor' must be a dictionary.")
                    
                # Recursively validate anchor locator
                self._validate_locator(parsed["anchor"])
            
            # Apply strict validation
            self._validate_locator(parsed)
            parsed = self._normalize_locator(parsed)
        else:
            # Parse string and then normalize
            parsed = self.parse_locator(str(locator))
            parsed = self._normalize_locator(parsed)

        # Final validation with strict type checking
        valid_object_types = [
            "paragraph",
            "table",
            "cell",
            "inline_shape",
            "image",
            "comment",
            "range",
            "selection",
            "document",
            "document_start",
            "document_end",
        ]
        object_type = parsed["type"]
        if object_type not in valid_object_types:
            raise LocatorSyntaxError(
                f"Invalid object type '{object_type}'. Valid types are: {', '.join(valid_object_types)}"
            )

        # Ensure locator is deterministic and not ambiguous
        # For certain object types, enforce specific constraints
        if object_type in ["paragraph", "table"] and parsed.get("value") is None and not parsed.get("filters"):
            raise LocatorSyntaxError(
                f"Locator for '{object_type}' must specify either 'value' or 'filters' to ensure deterministic selection."
            )

        return parsed
    def get_cache_key(self, locator: str, **kwargs) -> str:
        """Generate a cache key for the given locator and additional parameters.

        Args:
            locator: The locator string.
            **kwargs: Additional parameters that affect the selection.

        Returns:
            A unique cache key string.
        """
        # Start with the locator string
        key_parts = [locator]

        # Add any additional parameters
        for k, v in sorted(kwargs.items()):
            key_parts.append(f"{k}={v}")

        # Join all parts to form the cache key
        return "|".join(key_parts)
