"""
Utility functions for Word Document MCP Server tools.
This module provides common functionality used across different tool implementations.
"""
from typing import Optional, Dict, Any

from ..selector.locator_parser import LocatorParser
from ..selector.exceptions import LocatorSyntaxError


def check_locator_param(locator_value: Optional[Dict[str, Any]]) -> Optional[Dict[str, Any]]:
    """Check and validate a locator parameter.

    This function verifies that the locator is a dictionary and validates its structure
    using the LocatorParser. It returns the validated locator dictionary that can be
    directly used by SelectorEngine without the need for re-parsing.

    Args:
        locator_value: The locator parameter to check

    Returns:
        The validated locator dictionary

    Raises:
        TypeError: If locator is not a dictionary
        ValueError: If locator format is invalid
    """
    if locator_value is not None:
        # Check if it's a dictionary type
        if not isinstance(locator_value, dict):
            raise TypeError("locator parameter must be a dictionary")
        
        # Use LocatorParser to validate locator structure
        parser = LocatorParser()
        try:
            parser.validate_locator(locator_value)
        except LocatorSyntaxError:
            # Prompt user to refer to locator guide
            raise ValueError("Invalid locator format. Please refer to the locator guide for proper syntax.")
    
    return locator_value