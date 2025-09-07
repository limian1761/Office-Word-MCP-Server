"""
Main entry point for the Word Document MCP Server.
"""

import mcp

from word_docx_tools.mcp_service.core import mcp_server


def run_server():
    """
    Runs the MCP server. This function is the entry point for the script
    defined in pyproject.toml.
    """
    mcp_server.run(transport="stdio")


if __name__ == "__main__":
    run_server()
