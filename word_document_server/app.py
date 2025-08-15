"""
Centralized application instance for the Word Document MCP Server.
This module defines the FastMCP 'app' instance and its 'lifespan' manager,
making it importable by other parts of the application (e.g., tool modules).
"""
from contextlib import asynccontextmanager
import win32com.client
import pythoncom
from mcp.server.fastmcp.server import FastMCP
from word_document_server.utils.app_context import AppContext

@asynccontextmanager
async def lifespan(app: FastMCP):
    """
    Manages the lifecycle of the Word application instance and yields the AppContext.
    """
    word_app = None
    pythoncom.CoInitialize()
    try:
        print("Starting Word application...")
        try:
            word_app = win32com.client.GetActiveObject("Word.Application")
            print("Attached to an existing Word application instance.")
        except Exception:
            word_app = win32com.client.Dispatch("Word.Application")
            print("Started a new Word application instance.")
        word_app.Visible = True

        app_context = AppContext(word_app)
        yield app_context
        
    finally:
        if word_app:
            try:
                word_app.Quit(SaveChanges=0)
                print("Word application has been shut down.")
            except Exception as e:
                print(f"Error while shutting down Word: {e}")
        pythoncom.CoUninitialize()
        print("Server shutdown complete.")

# Create the main FastMCP application instance
app = FastMCP(
    lifespan=lifespan,
)
