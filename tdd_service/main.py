"""
Main entry point for the Test-Driven Development (TDD) Service.

This module provides the main entry point for running the TDD service.
"""

import mcp

from .core import tdd_server
from .test_runner import tdd_test_runner
from .auto_fixer import tdd_auto_fixer


def run_tdd_server():
    """
    Runs the TDD MCP server. This function is the entry point for the TDD service.
    """
    tdd_server.run(transport="stdio")


if __name__ == "__main__":
    run_tdd_server()