"""
COM Backend Adapter Layer for Word Document MCP Server.

This module encapsulates all interactions with the Word COM interface,
providing a clean, Pythonic API for higher-level components. It is designed
to be used as a context manager to ensure proper resource management.
"""
import win32com.client
import pythoncom
from typing import Optional, List


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

    def get_all_paragraphs(self) -> List[win32com.client.CDispatch]:
        """
        Get all paragraphs in the document.

        Returns:
            List of paragraph COM objects.
        """
        if not self.document:
            raise RuntimeError("No document open.")
        return list(self.document.Paragraphs)

    def get_paragraphs_in_range(self, range_obj: win32com.client.CDispatch) -> List[win32com.client.CDispatch]:
        """
        Get all paragraphs within a specific COM Range.

        Args:
            range_obj: The COM Range object to search within.

        Returns:
            List of paragraph COM objects found within the range.
        """
        if not self.document:
            raise RuntimeError("No document open.")
        
        paragraphs = []
        for para in range_obj.Paragraphs:
            paragraphs.append(para)
        return paragraphs

    def get_all_tables(self) -> List[win32com.client.CDispatch]:
        """
        Get all tables in the document.

        Returns:
            List of table COM objects.
        """
        if not self.document:
            raise RuntimeError("No document open.")
        return list(self.document.Tables)

    def get_runs_in_range(self, range_obj: win32com.client.CDispatch) -> List[win32com.client.CDispatch]:
        """
        Get all runs within a specific COM Range.

        Args:
            range_obj: The COM Range object to search within.

        Returns:
            List of Run COM objects found within the range.
        """
        if not self.document:
            raise RuntimeError("No document open.")
        
        runs = []
        for run in range_obj.Runs:
            runs.append(run)
        return runs

    def get_tables_in_range(self, range_obj: win32com.client.CDispatch) -> List[win32com.client.CDispatch]:
        """
        Get all tables within a specific COM Range.

        Args:
            range_obj: The COM Range object to search within.

        Returns:
            List of table COM objects found within the range.
        """
        if not self.document:
            raise RuntimeError("No document open.")
        
        tables = []
        for table in range_obj.Tables:
            tables.append(table)
        return tables

    def get_cells_in_range(self, range_obj: win32com.client.CDispatch) -> List[win32com.client.CDispatch]:
        """
        Get all cells within a specific COM Range.

        This iterates through all tables in the range and then all cells in each table.

        Args:
            range_obj: The COM Range object to search within.

        Returns:
            List of cell COM objects found within the range.
        """
        if not self.document:
            raise RuntimeError("No document open.")
        
        cells = []
        # A more robust way to iterate COM collections.
        tables_in_range = range_obj.Tables
        for i in range(1, tables_in_range.Count + 1):
            table = tables_in_range(i)
            for row in range(1, table.Rows.Count + 1):
                for col in range(1, table.Columns.Count + 1):
                    cells.append(table.Cell(row, col))
        return cells

    def add_table(self, com_range_obj: win32com.client.CDispatch, rows: int, cols: int) -> win32com.client.CDispatch:
        """
        Adds a table after a given range.

        Args:
            com_range_obj: The range to insert the table after.
            rows: Number of rows for the table.
            cols: Number of columns for the table.

        Returns:
            The newly created table COM object.
        """
        if not self.document:
            raise RuntimeError("No document open.")
        
        # Collapse the range to its end point to insert after
        new_range = com_range_obj.Duplicate
        new_range.Collapse(0) # WdCollapseDirection.wdCollapseEnd
        new_range.InsertParagraphAfter() # Add a paragraph break to ensure table is on a new line
        return self.document.Tables.Add(new_range, rows, cols)

    def set_bold_for_range(self, com_range_obj: win32com.client.CDispatch, is_bold: bool):
        """
        Set bold formatting for a range.

        Args:
            com_range_obj: COM Range object to format.
            is_bold: Whether to set bold formatting.
        """
        com_range_obj.Font.Bold = is_bold

    def set_italic_for_range(self, com_range_obj: win32com.client.CDispatch, is_italic: bool):
        """
        Set italic formatting for a range.

        Args:
            com_range_obj: COM Range object to format.
            is_italic: Whether to set italic formatting.
        """
        com_range_obj.Font.Italic = is_italic

    def set_font_size_for_range(self, com_range_obj: win32com.client.CDispatch, size: int):
        """
        Set font size for a range.

        Args:
            com_range_obj: COM Range object to format.
            size: The font size in points.
        """
        com_range_obj.Font.Size = size

    def set_font_name_for_range(self, com_range_obj: win32com.client.CDispatch, name: str):
        """
        Set font name for a range.

        Args:
            com_range_obj: COM Range object to format.
            name: The name of the font.
        """
        com_range_obj.Font.Name = name

    def insert_paragraph_after(self, com_range_obj: win32com.client.CDispatch, text: str):
        """
        Insert a paragraph after a given range using the document's Paragraphs collection.

        Args:
            com_range_obj: COM Range object after which to insert.
            text: Text to insert.
        """
        # Define the range for the new paragraph, which is at the end of the anchor range.
        insert_range = self.document.Range(com_range_obj.End, com_range_obj.End)
        
        # Add a new paragraph at this range.
        new_para = self.document.Paragraphs.Add(insert_range)
        
        # Set the text for the new paragraph.
        new_para.Range.Text = text
