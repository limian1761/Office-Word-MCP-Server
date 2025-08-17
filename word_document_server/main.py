"""
Main entry point for the Word Document MCP Server.
"""
from mcp import cli
from word_document_server.app import mcp_server

def run_server():
    """
    Runs the MCP server. This function is the entry point for the script
    defined in pyproject.toml.
    """
    cli.run_server(mcp_server)

if __name__ == "__main__":
    run_server()
