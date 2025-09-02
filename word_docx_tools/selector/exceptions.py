"""Exceptions for the selector engine.

This module defines custom exceptions used by the selector engine
for handling errors during object selection.
"""


class LocatorError(Exception):
    """Base exception class for all locator-related errors."""

    pass


class LocatorSyntaxError(LocatorError):
    """Raised when a locator string has invalid syntax."""

    pass


class AmbiguousLocatorError(LocatorError):
    """Raised when a locator matches multiple objects when a single object is expected."""

    pass
