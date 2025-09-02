"""
Word Document Server Package

This package provides a complete set of tools and operations for working with
Microsoft Word documents through the MCP (Model Context Protocol) server.

The main entry point for the application is word_document_server.main.run_server().
"""

# Import all operations
from .operations import *

# Import all tools
from .tools import *

# Package version
__version__ = "1.1.9"

# Define main entry point
from .main import run_server
__all__ = ['run_server']