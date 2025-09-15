"""Backend module for Word Document MCP Server.

This module encapsulates all interactions with the Word COM interface,
providing a clean, Pythonic API for higher-level components.
"""

from .com_adapter import WordBackend
from .com_utils import *

__all__ = [
    "WordBackend",
]