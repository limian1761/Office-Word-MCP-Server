"""
Main entry point for the Word Document MCP Server.

This script discovers and registers all available tools, then starts the server.
Tool functions are automatically registered via the @app.tool decorator.
"""
import pkgutil
import importlib
from word_document_server.app import app
from word_document_server import tools

# Dynamically import all modules in the 'tools' package to ensure
# that the @app.tool decorators are executed and the tools are registered.
for module_info in pkgutil.iter_modules(tools.__path__, tools.__name__ + "."):
    importlib.import_module(module_info.name)

def run_server():
    """Runs the FastMCP server."""
    app.run()

if __name__ == "__main__":
    run_server()