"""Exceptions for the selector engine.

This module defines custom exceptions used by the selector engine
for handling errors during element selection.
"""


class LocatorError(Exception):
    """Base exception class for all locator-related errors."""

    pass


class LocatorSyntaxError(LocatorError):
    """Raised when a locator string has invalid syntax."""

    pass


class AmbiguousLocatorError(LocatorError):
    """Raised when a locator matches multiple elements when a single element is expected."""

    pass
