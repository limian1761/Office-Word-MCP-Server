"""
COM Backend Adapter Layer for Word Document MCP Server.

This module encapsulates all interactions with the Word COM interface,
providing a clean, Pythonic API for higher-level components. It is designed
to be used as a context manager to ensure proper resource management.
"""

import re
from typing import Any, Dict, List, Optional

import pythoncom
import win32com.client

from word_document_server.errors import WordDocumentError


class WordBackend:
    """
    Backend adapter for interacting with Word COM interface.

    This class is designed to be used as a context manager (`with` statement)
    to ensure that the Word application is properly initialized and cleaned up.
    """

    def __init__(self, file_path: Optional[str] = None, visible: bool = True):
        """
        Initialize the Word backend adapter.

        Args:
            file_path (Optional[str]): Path to the document file to open.
                                       If None, a new document is created.
            visible (bool): Whether to make the Word application visible.
        """
        self.file_path = file_path
        self.visible = visible
        self.word_app: Optional[win32com.client.CDispatch] = None
        self.document: Optional[win32com.client.CDispatch] = None

    def __enter__(self):
        """
        Starts a new Word application instance.
        Opens or creates a document.
        """
        try:
            # First, try to get an active instance of Word
            self.word_app = win32com.client.GetActiveObject("Word.Application")
            print("Attached to an existing Word application instance.")
        except pythoncom.com_error:
            # If that fails, start a new instance
            try:
                self.word_app = win32com.client.Dispatch("Word.Application")
                print("Started a new Word application instance.")
            except Exception as e:
                raise RuntimeError(f"Failed to start Word Application: {e}")

        self.word_app.Visible = self.visible

        if self.file_path:
            try:
                # Convert to absolute path for COM
                import os

                abs_path = os.path.abspath(self.file_path)
                self.document = self.word_app.Documents.Open(abs_path)
            except pythoncom.com_error as e:
                # This can happen if the file is corrupt, password-protected, or doesn't exist.
                self.cleanup()
                raise WordDocumentError(
                    f"Word COM error while opening document: {self.file_path}. Details: {e}"
                )
            except Exception as e:
                self.cleanup()
                raise IOError(f"Failed to open document: {self.file_path}. Error: {e}")
        else:
            self.document = self.word_app.Documents.Add()

        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """
        Ensures the cleanup method is called to close Word and uninitialize COM.
        """
        self.cleanup()

    def cleanup(self):
        """
        Closes the document and quits the Word application, then uninitializes COM.
        """
        if self.document:
            try:
                self.document.Close(SaveChanges=False)
            except pythoncom.com_error as e:
                print(f"Warning: Could not close document: {e}")
            self.document = None

        # We no longer quit the app here to allow for multiple tool calls.
        # The app must be explicitly closed by a 'shutdown' tool.
        print("Word backend cleaned up (document closed).")

    def shutdown(self):
        """Closes the document and shuts down the Word application.

        This method should be called explicitly when you want to completely
        terminate the Word application instance.
        """
        # Close the document if it's open
        self.cleanup()

        # Quit the Word application
        if self.word_app:
            try:
                self.word_app.Quit()
                print("Word application has been shut down.")
            except pythoncom.com_error as e:
                print(f"Warning: Could not quit Word application: {e}")
            self.word_app = None

    def get_all_paragraphs(self) -> List[Any]:
        """Retrieves all paragraphs in the document.

        Returns:
            List of paragraph objects.
        """
        if not self.document:
            raise WordDocumentError("No active document.")
        
        try:
            return list(self.document.Paragraphs)
        except Exception as e:
            raise WordDocumentError(f"Failed to get all paragraphs: {e}")
    
    def get_paragraphs_in_range(self, search_range: Any) -> List[Any]:
        """Retrieves paragraphs within a specific range.

        Args:
            search_range: The range to search within.
        
        Returns:
            List of paragraph objects within the range.
        """
        if not self.document:
            raise WordDocumentError("No active document.")
        
        try:
            # Get all paragraphs in the document
            all_paragraphs = self.get_all_paragraphs()
            
            # Filter paragraphs that are within the search range
            filtered_paragraphs = []
            for para in all_paragraphs:
                # Check if paragraph is within the search range
                if para.Range.Start >= search_range.Start and para.Range.End <= search_range.End:
                    filtered_paragraphs.append(para)
            
            return filtered_paragraphs
        except Exception as e:
            raise WordDocumentError(f"Failed to get paragraphs in range: {e}")
    
    def get_all_tables(self) -> List[Any]:
        """Retrieves all tables in the document.

        Returns:
            List of table objects.
        """
        if not self.document:
            raise WordDocumentError("No active document.")
        
        try:
            return list(self.document.Tables)
        except Exception as e:
            raise WordDocumentError(f"Failed to get all tables: {e}")
    
    def get_tables_in_range(self, search_range: Any) -> List[Any]:
        """Retrieves tables within a specific range.

        Args:
            search_range: The range to search within.
        
        Returns:
            List of table objects within the range.
        """
        if not self.document:
            raise WordDocumentError("No active document.")
        
        try:
            # Get all tables in the document
            all_tables = self.get_all_tables()
            
            # Filter tables that are within the search range
            filtered_tables = []
            for table in all_tables:
                # Check if table is within the search range
                if table.Range.Start >= search_range.Start and table.Range.End <= search_range.End:
                    filtered_tables.append(table)
            
            return filtered_tables
        except Exception as e:
            raise WordDocumentError(f"Failed to get tables in range: {e}")
    
    def get_cells_in_range(self, search_range: Any) -> List[Any]:
        """Retrieves table cells within a specific range.

        Args:
            search_range: The range to search within.
        
        Returns:
            List of cell objects within the range.
        """
        if not self.document:
            raise WordDocumentError("No active document.")
        
        try:
            cells = []
            tables = self.get_tables_in_range(search_range)
            for table in tables:
                for row in table.Rows:
                    for cell in row.Cells:
                        # Ensure the cell is within the search range
                        if cell.Range.Start >= search_range.Start and cell.Range.End <= search_range.End:
                            cells.append(cell)
            
            return cells
        except Exception as e:
            raise WordDocumentError(f"Failed to get cells in range: {e}")
    
    def get_runs_in_range(self, search_range: Any) -> List[Any]:
        """Retrieves text runs within a specific range.

        Args:
            search_range: The range to search within.
        
        Returns:
            List of run objects within the range.
        """
        if not self.document:
            raise WordDocumentError("No active document.")
        
        try:
            runs = []
            paragraphs = self.get_paragraphs_in_range(search_range)
            for para in paragraphs:
                for run in para.Range.Runs:
                    # Ensure the run is within the search range
                    if run.Start >= search_range.Start and run.End <= search_range.End:
                        runs.append(run)
            
            return runs
        except Exception as e:
            raise WordDocumentError(f"Failed to get runs in range: {e}")
