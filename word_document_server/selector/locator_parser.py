"""Locator parser for the selector engine.

This module contains the functionality for parsing and validating
locator strings used to identify elements in a Word document.
"""

import re
from typing import Any, Dict, List, Optional, Tuple, Union

from word_document_server.selector.exceptions import LocatorSyntaxError


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
                raise LocatorSyntaxError("Locator must specify an element type.")

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
        """Validate a parsed locator dictionary.

        Args:
            parsed_locator: The parsed locator dictionary to validate.

        Raises:
            LocatorSyntaxError: If the locator is invalid.
        """
        element_type = parsed_locator["type"]

        # Check for required type field
        if not element_type:
            raise LocatorSyntaxError("Locator must specify an element type.")

        # Validate relation if anchor is specified
        if parsed_locator["anchor"] is not None:
            valid_relations = [
                "all_occurrences_within",
                "first_occurrence_after",
                "parent_of",
                "immediately_following",
            ]
            if (
                parsed_locator["relation"] is not None
                and parsed_locator["relation"] not in valid_relations
            ):
                raise LocatorSyntaxError(
                    f"Invalid relation '{parsed_locator['relation']}'. Valid relations are: {', '.join(valid_relations)}"
                )

        # Validate certain element types
        valid_element_types = [
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
        if element_type not in valid_element_types:
            # For guidance on proper locator syntax, please refer to:
            # word_document_server/selector/LOCATOR_GUIDE.md
            raise LocatorSyntaxError(
                f"Invalid element type '{element_type}'. Valid types are: {', '.join(valid_element_types)}"
            )

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
