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

from word_document_server.errors import WordDocumentError, ErrorCode

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
                raise WordDocumentError(ErrorCode.DOCUMENT_OPEN_ERROR, f"Word COM error while opening document: {self.file_path}. Details: {e}")
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

    def get_text_from_range(self, start_pos: int, end_pos: int) -> str:
        """
        Get text from a specific range in the document.

        Args:
            start_pos: The start position of the range.
            end_pos: The end position of the range.

        Returns:
            The text content of the specified range.
        """
        if not self.document:
            raise RuntimeError("No document open.")

        # Validate range parameters
        if not isinstance(start_pos, int) or start_pos < 0:
            raise ValueError("start_pos must be a non-negative integer")
        if not isinstance(end_pos, int) or end_pos <= start_pos:
            raise ValueError("end_pos must be an integer greater than start_pos")

        # Get the document range
        doc_range = self.document.Range(start_pos, end_pos)
        return doc_range.Text

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

    def add_table(self, com_range_obj: win32com.client.CDispatch, rows: int, cols: int):
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
        
        try:
            # Validate row and column parameters
            if not isinstance(rows, int) or rows <= 0:
                raise ValueError("Row count must be a positive integer")
            if not isinstance(cols, int) or cols <= 0:
                raise ValueError("Column count must be a positive integer")
            
            # Validate range object
            if not com_range_obj or not hasattr(com_range_obj, 'Duplicate'):
                raise ValueError("Invalid range object")
            
            # Collapse the range to its end point to insert after
            new_range = com_range_obj.Duplicate
            new_range.Collapse(0) # WdCollapseDirection.wdCollapseEnd
            new_range.InsertParagraphAfter() # Add a paragraph break to ensure table is on a new line
            return self.document.Tables.Add(new_range, rows, cols)
        except Exception as e:
            # Check if it's a COM related error
            if "COM" in str(type(e)) or "Dispatch" in str(type(e)):
                from word_document_server.errors import WordDocumentError
                raise WordDocumentError(f"Failed to create table in Word: {str(e)}")
            raise

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

    def set_font_color_for_range(self, com_range_obj: win32com.client.CDispatch, color: str):
        """
        Set font color for a range.

        Args:
            com_range_obj: COM Range object to format.
            color: Named color (e.g., 'blue') or hex code (e.g., '#0000FF').
        """
        # Convert color name to Word's RGB color value or use hex code
        color_map = {
            'black': 0,
            'white': 16777215,
            'red': 255,
            'green': 65280,
            'blue': 16711680,
            'yellow': 65535
        }
        if color.lower() in color_map:
            com_range_obj.Font.Color = color_map[color.lower()]
        else:
            # Try to parse hex color (e.g., '#RRGGBB' or 'RRGGBB')
            color = color.lstrip('#')
            if len(color) == 6:
                try:
                    rgb = int(color, 16)
                    com_range_obj.Font.Color = rgb
                except ValueError:
                    raise ValueError(f"Invalid hex color format: {color}")
            else:
                raise ValueError(f"Unsupported color: {color}. Use named color or 6-digit hex code.")

    def set_font_name_for_range(self, com_range_obj: win32com.client.CDispatch, name: str):
        """
        Set font name for a range.

        Args:
            com_range_obj: COM Range object to format.
            name: The name of the font.
        """
        com_range_obj.Font.Name = name

    def insert_paragraph_after(self, com_range_obj: win32com.client.CDispatch, text: str, style: str = None):
        """
        Insert a paragraph after a given range using the document's Paragraphs collection.

        Args:
            com_range_obj: COM Range object after which to insert.
            text: Text to insert.
            style: Optional, paragraph style name to apply.

        Returns:
            The newly created paragraph COM object.
        """
        # Define the range for the new paragraph, which is at the end of the anchor range.
        insert_range = self.document.Range(com_range_obj.End, com_range_obj.End)
        
        # Add a new paragraph at this range.
        new_para = self.document.Paragraphs.Add(insert_range)
        
        # Set the text for the new paragraph.
        new_para.Range.Text = text
        
        # Apply style if specified
        if style:
            try:
                new_para.Style = style
            except Exception as e:
                print(f"Warning: Failed to apply paragraph style '{style}': {e}")
                
        # Return the newly created paragraph object
        return new_para

    def set_header_text(self, text: str, header_index: int = 1):
        """
        Sets the text for a specific header in all sections of the document.

        Args:
            text: The text to set in the header.
            header_index: The index of the header to modify (e.g., 1 for primary header).
        """
        if not self.document:
            raise RuntimeError("No document open.")

        # Iterate through all sections in the document
        for i in range(1, self.document.Sections.Count + 1):
            section = self.document.Sections(i)
            # Access the specified header
            header = section.Headers(header_index)
            # Set the text of the header's range
            header.Range.Text = text

    def set_footer_text(self, text: str, footer_index: int = 1):
        """
        Sets the text for a specific footer in all sections of the document.

        Args:
            text: The text to set in the footer.
            footer_index: The index of the footer to modify (e.g., 1 for primary footer).
        """
        if not self.document:
            raise RuntimeError("No document open.")

        # Iterate through all sections in the document
        for i in range(1, self.document.Sections.Count + 1):
            section = self.document.Sections(i)
            # Access the specified footer
            footer = section.Footers(footer_index)
            # Set the text of the footer's range
            footer.Range.Text = text

    def create_bulleted_list_relative_to(self, com_range_obj: win32com.client.CDispatch, items: List[str], position: str):
        """
        Creates a new bulleted list relative to a given range.

        Args:
            com_range_obj: The range to insert the list before or after.
            items: A list of strings, where each string is a list item.
            position: "before" or "after".
        """
        if position == "before":
            insertion_point = com_range_obj.Start
        elif position == "after":
            insertion_point = com_range_obj.End
        else:
            raise ValueError("Position must be 'before' or 'after'.")

        # Collapse the range to the desired insertion point
        target_range = self.document.Range(insertion_point, insertion_point)
        
        # Join items and insert the text block
        full_text = "\n".join(items) + "\n"
        target_range.InsertAfter(full_text)

        # Select the newly inserted text
        new_text_range = self.document.Range(insertion_point, insertion_point + len(full_text))
        
        # Apply list format to each paragraph in the new range
        for para in new_text_range.Paragraphs:
            para.Range.ListFormat.ApplyBulletDefault()

    def get_headings(self) -> List[Dict[str, Any]]:
        """
        Extracts all heading paragraphs from the document.

        Returns:
            A list of dictionaries, where each dictionary represents a heading
            with "text" and "level" keys.
        """
        if not self.document:
            raise RuntimeError("No document open.")

        headings = []
        for para in self.document.Paragraphs:
            style_name = para.Style.NameLocal
            # Check for both English and Chinese heading styles
            if style_name.startswith("Heading") or style_name.startswith("标题"):
                try:
                    # Extract the level number from the end of the style name
                    level = int(re.findall(r'\d+$', style_name)[0])
                    text = para.Range.Text.strip()
                    if text: # Only include headings with text
                        headings.append({"text": text, "level": level})
                except (IndexError, ValueError):
                    # Not a standard heading style like "Heading 1", so we ignore it
                    continue
        return headings

    def set_alignment_for_range(self, com_range_obj: win32com.client.CDispatch, alignment: str):
        """
        Set paragraph alignment for a range.

        Args:
            com_range_obj: COM Range object to format.
            alignment: "left", "center", or "right".
        """
        alignment_map = {
            "left": 0,    # wdAlignParagraphLeft
            "center": 1,  # wdAlignParagraphCenter
            "right": 2    # wdAlignParagraphRight
        }
        if alignment.lower() in alignment_map:
            com_range_obj.ParagraphFormat.Alignment = alignment_map[alignment.lower()]
        else:
            raise ValueError(f"Invalid alignment value: {alignment}. Must be 'left', 'center', or 'right'.")

    def accept_all_changes(self):
        """Accepts all tracked changes in the document."""
        if not self.document:
            raise RuntimeError("No document open.")
        self.document.AcceptAllRevisions()

    def enable_track_revisions(self):
        """Enables track changes (revision mode) in the document."""
        if not self.document:
            raise RuntimeError("No document open.")
        self.document.TrackRevisions = True

    def disable_track_revisions(self):
        """Disables track changes (revision mode) in the document."""
        if not self.document:
            raise RuntimeError("No document open.")
        self.document.TrackRevisions = False

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
        
    def get_all_styles(self) -> List[Dict[str, Any]]:
        """
        Retrieves all available styles in the document.
        
        Returns:
            A list of dictionaries containing style information, each with "name" and "type" keys.
        """
        if not self.document:
            raise RuntimeError("No document open.")
            
        styles = []
        # Get all styles from the document
        for i in range(1, self.document.Styles.Count + 1):
            style = self.document.Styles(i)
            try:
                style_info = {
                    "name": style.NameLocal,  # Local name of the style
                    "type": self._get_style_type(style.Type)
                }
                styles.append(style_info)
            except Exception as e:
                print(f"Warning: Failed to retrieve style information: {e}")
        return styles

    def get_protection_status(self) -> Dict[str, Any]:
        """
        Checks the protection status of the document.
        
        Returns:
            A dictionary containing protection status information:
            - is_protected: Boolean indicating if the document is protected
            - protection_type: String describing the type of protection
        """
        if not self.document:
            raise RuntimeError("No document open.")

        # Mapping of Word protection type constants to human-readable descriptions
          # Based on Microsoft's WdProtectionType enumeration: https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.wdprotectiontype
        protection_types = {
              -1: "No protection",          # wdNoProtection: Document is not protected
              0: "Allow only revisions",    # wdAllowOnlyRevisions: Allow only revisions to existing content
              1: "Allow only comments",     # wdAllowOnlyComments: Allow only comments to be added
              2: "Allow only form fields",  # wdAllowOnlyFormFields: Allow content only through form fields
              3: "Allow only reading"       # wdAllowOnlyReading: Allow read-only access
        }

        protection_type = self.document.ProtectionType
        is_protected = protection_type != -1

        return {
            "is_protected": is_protected,
            "protection_type": protection_types.get(protection_type, f"Unknown ({protection_type})")
        }

    def unprotect_document(self, password: Optional[str] = None) -> bool:
        """
        Attempts to unprotect the document.
        
        Args:
            password: Optional password to use for unprotecting the document.

        Returns:
            True if the document was successfully unprotected, False otherwise.
        """
        if not self.document:
            raise RuntimeError("No document open.")

        protection_status = self.get_protection_status()
        if not protection_status["is_protected"]:
            return True  # Already unprotected

        try:
            # Word's Unprotect method returns True if successful
            if password:
                result = self.document.Unprotect(Password=password)
            else:
                result = self.document.Unprotect()
            return result
        except Exception as e:
            print(f"Warning: Failed to unprotect document: {e}")
            return False

    def _get_style_type(self, style_type_code: int) -> str:
        """
        Converts a style type code to a human-readable string.
        
        Args:
            style_type_code: The style type code from Word's COM API.
        
        Returns:
            A human-readable string representing the style type.
        """
        style_types = {
            1: "Paragraph",
            2: "Character",
            3: "Table",
            4: "List"
        }
        return style_types.get(style_type_code, f"Unknown ({style_type_code})")
        
    def _get_style_type(self, type_code: int) -> str:
        """
        Converts Word style type code to human-readable string.
        
        Args:
            type_code: Style type code from Word COM interface.
        
        Returns:
            Human-readable style type.
        """
        # Word style type constants
        style_types = {
            1: "Paragraph",  # wdStyleTypeParagraph
            2: "Character",  # wdStyleTypeCharacter
            3: "Table",      # wdStyleTypeTable
            4: "List"         # wdStyleTypeList
        }
        return style_types.get(type_code, "Unknown")
    
    def get_all_inline_shapes(self) -> List[Dict[str, Any]]:
        """
        Retrieves all inline shapes (including pictures) in the document.
        
        Returns:
            A list of dictionaries containing shape information, each with "index", "type", and "width" keys.
        """
        if not self.document:
            raise RuntimeError("No document open.")
            
        shapes = []
        try:
            # Check if InlineShapes property exists and is accessible
            if not hasattr(self.document, 'InlineShapes'):
                return shapes
            
            # Get all inline shapes from the document safely
            shapes_count = 0
            try:
                shapes_count = self.document.InlineShapes.Count
            except Exception as e:
                print(f"Warning: Failed to access InlineShapes collection: {e}")
                return shapes
            
            for i in range(1, shapes_count + 1):
                try:
                    shape = self.document.InlineShapes(i)
                    try:
                        shape_info = {
                            "index": i - 1,  # 0-based index
                            "type": self._get_shape_type(shape.Type) if hasattr(shape, 'Type') else "Unknown",
                            "width": shape.Width if hasattr(shape, 'Width') else 0,
                            "height": shape.Height if hasattr(shape, 'Height') else 0
                        }
                        # Add additional properties based on shape type
                        if shape_info["type"] == "Picture":
                            # Try to get picture format information if available
                            if hasattr(shape, 'PictureFormat'):
                                if hasattr(shape.PictureFormat, 'ColorType'):
                                    shape_info["color_type"] = self._get_color_type(shape.PictureFormat.ColorType)
                        shapes.append(shape_info)
                    except Exception as e:
                        print(f"Warning: Failed to retrieve shape information for index {i}: {e}")
                        continue
                except Exception as e:
                    print(f"Warning: Failed to access shape at index {i}: {e}")
                    continue
        except Exception as e:
            print(f"Error: Failed to retrieve inline shapes: {e}")
            
        return shapes
        
    def insert_inline_picture(self, com_range_obj: win32com.client.CDispatch, image_path: str, position: str = "after") -> win32com.client.CDispatch:
        """
        Inserts an inline picture at the specified range.
        
        Args:
            com_range_obj: The COM Range object where the picture will be inserted.
            image_path: The absolute path to the image file.
            position: "before", "after", or "replace" to specify where to insert the picture relative to the range.
        
        Returns:
            The newly inserted InlineShape COM object.
        """
        if not self.document:
            raise RuntimeError("No document open.")
        
        if not image_path or not isinstance(image_path, str):
            raise ValueError("Invalid image path provided.")
        
        if not com_range_obj:
            raise ValueError("Invalid range object provided.")
        
        if position not in ["before", "after", "replace"]:
            raise ValueError("Invalid position. Must be 'before', 'after', or 'replace'.")
        
        try:
            # Create a duplicate of the range to avoid modifying the original
            insert_range = com_range_obj.Duplicate
            
            if position == "replace":
                # Delete the content of the range
                insert_range.Text = ""
                # Insert the picture
                return self.document.InlineShapes.AddPicture(FileName=image_path, LinkToFile=False, SaveWithDocument=True, Range=insert_range)
            elif position == "before":
                # Collapse the range to its start point
                insert_range.Collapse(1)  # wdCollapseStart
                # Insert the picture
                return self.document.InlineShapes.AddPicture(FileName=image_path, LinkToFile=False, SaveWithDocument=True, Range=insert_range)
            else:  # position == "after"
                # Collapse the range to its end point
                insert_range.Collapse(0)  # wdCollapseEnd
                # Insert the picture
                return self.document.InlineShapes.AddPicture(FileName=image_path, LinkToFile=False, SaveWithDocument=True, Range=insert_range)
        except Exception as e:
            raise WordDocumentError(f"Failed to insert picture '{image_path}': {e}")
            
    def _get_shape_type(self, type_code: int) -> str:
        """
        Converts Word shape type code to human-readable string.
        
        Args:
            type_code: Shape type code from Word COM interface.
        
        Returns:
            Human-readable shape type.
        """
        # Word inline shape type constants
        shape_types = {
            1: "Picture",       # wdInlineShapePicture
            2: "LinkedPicture", # wdInlineShapeLinkedPicture
            3: "Chart",         # wdInlineShapeChart
            4: "Diagram",       # wdInlineShapeDiagram
            5: "OLEControlObject", # wdInlineShapeOLEControlObject
            6: "OLEObject",     # wdInlineShapeOLEObject
            7: "ActiveXControl", # wdInlineShapeActiveXControl
            8: "SmartArt",      # wdInlineShapeSmartArt
            9: "3DModel"         # wdInlineShape3DModel
        }
        return shape_types.get(type_code, "Unknown")
        
    def _get_color_type(self, color_code: int) -> str:
        """
        Converts Word picture color type code to human-readable string.
        
        Args:
            color_code: Color type code from Word COM interface.
        
        Returns:
            Human-readable color type.
        """
        # Word picture color type constants
        color_types = {
            0: "Color",         # msoPictureColorTypeColor
            1: "Grayscale",     # msoPictureColorTypeGrayscale
            2: "BlackAndWhite", # msoPictureColorTypeBlackAndWhite
            3: "Watermark"       # msoPictureColorTypeWatermark
        }
        return color_types.get(color_code, "Unknown")

    def add_comment(self, com_range_obj: win32com.client.CDispatch, text: str, author: str = "User") -> win32com.client.CDispatch:
        """
        Adds a comment to the specified range.

        Args:
            com_range_obj: The COM Range object where the comment will be inserted.
            text: The text of the comment.
            author: The author of the comment (default: "User").

        Returns:
            The newly created Comment COM object.
        """
        if not self.document:
            raise RuntimeError("No document open.")
        
        if not com_range_obj:
            raise ValueError("Invalid range object provided.")
        
        try:
            # Add a comment at the specified range
            return self.document.Comments.Add(Range=com_range_obj, Text=text)
        except Exception as e:
            raise WordDocumentError(f"Failed to add comment: {e}")

    def get_comments(self) -> List[Dict[str, Any]]:
        """
        Retrieves all comments in the document.

        Returns:
            A list of dictionaries containing comment information, each with "index", "text", "author", "start_pos", "end_pos", and "scope_text" keys.
        """
        if not self.document:
            raise RuntimeError("No document open.")
            
        comments = []
        try:
            # Check if Comments property exists and is accessible
            if not hasattr(self.document, 'Comments'):
                return comments
            
            # Get all comments from the document
            comments_count = 0
            try:
                comments_count = self.document.Comments.Count
            except Exception as e:
                print(f"Warning: Failed to access Comments collection: {e}")
                return comments
            
            for i in range(1, comments_count + 1):
                try:
                    comment = self.document.Comments(i)
                    try:
                        comment_info = {
                            "index": i - 1,  # 0-based index
                            "text": comment.Range.Text if hasattr(comment, 'Range') else "",
                            "author": comment.Author if hasattr(comment, 'Author') else "Unknown",
                            "start_pos": comment.Scope.Start if hasattr(comment, 'Scope') and hasattr(comment.Scope, 'Start') else 0,
                            "end_pos": comment.Scope.End if hasattr(comment, 'Scope') and hasattr(comment.Scope, 'End') else 0,
                            "scope_text": comment.Scope.Text.strip() if hasattr(comment, 'Scope') and hasattr(comment.Scope, 'Text') else ""
                        }
                        comments.append(comment_info)
                    except Exception as e:
                        print(f"Warning: Failed to retrieve comment information for index {i}: {e}")
                        continue
                except Exception as e:
                    print(f"Warning: Failed to access comment at index {i}: {e}")
                    continue
        except Exception as e:
            print(f"Error: Failed to retrieve comments: {e}")
            
        return comments

    def get_comments_by_range(self, com_range_obj: win32com.client.CDispatch) -> List[Dict[str, Any]]:
        """
        Retrieves comments within a specific COM Range.

        Args:
            com_range_obj: The COM Range object to search within.

        Returns:
            A list of dictionaries containing comment information.
        """
        if not self.document:
            raise RuntimeError("No document open.")
            
        if not com_range_obj:
            raise ValueError("Invalid range object provided.")
            
        comments = []
        try:
            # Check if Comments property exists and is accessible
            if not hasattr(self.document, 'Comments'):
                return comments
            
            # Get all comments from the document
            comments_count = 0
            try:
                comments_count = self.document.Comments.Count
            except Exception as e:
                print(f"Warning: Failed to access Comments collection: {e}")
                return comments
            
            # Check if the range object has Start and End properties
            if not hasattr(com_range_obj, 'Start') or not hasattr(com_range_obj, 'End'):
                print("Warning: Invalid range object - missing Start or End properties")
                return comments
            
            for i in range(1, comments_count + 1):
                try:
                    comment = self.document.Comments(i)
                    try:
                        # Check if comment is within the specified range
                        if (hasattr(comment, 'Scope') and 
                            hasattr(comment.Scope, 'Start') and 
                            hasattr(comment.Scope, 'End') and 
                            comment.Scope.Start >= com_range_obj.Start and 
                            comment.Scope.End <= com_range_obj.End):
                            comment_info = {
                                "index": i - 1,  # 0-based index
                                "text": comment.Range.Text if hasattr(comment, 'Range') else "",
                                "author": comment.Author if hasattr(comment, 'Author') else "Unknown",
                                "start_pos": comment.Scope.Start if hasattr(comment, 'Scope') and hasattr(comment.Scope, 'Start') else 0,
                                "end_pos": comment.Scope.End if hasattr(comment, 'Scope') and hasattr(comment.Scope, 'End') else 0,
                                "scope_text": comment.Scope.Text.strip() if hasattr(comment, 'Scope') and hasattr(comment.Scope, 'Text') else ""
                            }
                            comments.append(comment_info)
                    except Exception as e:
                        print(f"Warning: Failed to retrieve comment information for index {i}: {e}")
                        continue
                except Exception as e:
                    print(f"Warning: Failed to access comment at index {i}: {e}")
                    continue
        except Exception as e:
            print(f"Error: Failed to retrieve comments by range: {e}")
            
        return comments
        
    def get_document_styles(self) -> List[Dict[str, Any]]:
        """
        Retrieves all available styles in the active document.
        
        Returns:
            A list of styles with their names and types.
        """
        if not self.document:
            raise RuntimeError("No document open.")
            
        styles = []
        try:
            # Iterate through all styles in the document
            for style in self.document.Styles:
                try:
                    # Skip built-in hidden styles
                    if not style.BuiltIn or style.InUse:
                        style_info = {
                            'name': style.NameLocal,
                            'type': style.Type  # wdStyleTypeParagraph (1), wdStyleTypeCharacter (2), etc.
                        }
                        styles.append(style_info)
                except Exception as e:
                    print(f"Warning: Failed to retrieve style info: {e}")
                    continue
            
            # Sort styles by name
            styles.sort(key=lambda x: x['name'])
            
        except Exception as e:
            raise WordDocumentError(f"Error retrieving document styles: {e}")
            
        return styles
        
    def get_document_structure(self) -> List[Dict[str, Any]]:
        """
        Provides a structured overview of the document by listing all headings.
        
        Returns:
            A list of dictionaries, each representing a heading with its text and level.
        """
        if not self.document:
            raise RuntimeError("No document open.")
            
        structure = []
        try:
            # Iterate through all paragraphs
            for paragraph in self.document.Paragraphs:
                try:
                    # Get paragraph style name
                    style_name = paragraph.Style.NameLocal
                    
                    # 添加调试信息
                    print(f"段落样式: {style_name}")
                    
                    # Check if it's a heading style (supports both English and Chinese styles)
                    # 优化匹配逻辑，不区分大小写并移除多余空格
                    style_name_clean = style_name.strip().lower()
                    heading_match = None
                    level = None
                    
                    # 检查英文标题样式 (Heading 1-9)
                    if style_name_clean.startswith('heading '):
                        heading_match = style_name_clean.split(' ')
                        if len(heading_match) > 1:
                            try:
                                level = int(heading_match[1])
                            except ValueError:
                                pass
                    
                    # 检查中文标题样式 (标题 1-9 或 标题1-9)
                    if '标题' in style_name_clean:
                        # 尝试匹配"标题 X"或"标题X"格式
                        import re
                        cn_heading_match = re.search(r'标题\s*(\d+)', style_name_clean)
                        if cn_heading_match:
                            level = int(cn_heading_match.group(1))
                    
                    if level and 1 <= level <= 9:
                        # Get heading text
                        text = paragraph.Range.Text.strip()
                        
                        if text:
                            structure.append({
                                'text': text,
                                'level': level
                            })
                            print(f"添加标题: {text} (级别: {level})")
                        else:
                            print(f"跳过空标题段落 (样式: {style_name})")
                except Exception as e:
                    print(f"Warning: Failed to process paragraph: {e}")
                    continue
        except Exception as e:
            raise WordDocumentError(f"Error retrieving document structure: {e}")
            
        return structure

    def delete_comment(self, comment_index: int) -> None:
        """
        Deletes a comment by its 0-based index.

        Args:
            comment_index: The 0-based index of the comment to delete.
        """
        if not self.document:
            raise RuntimeError("No document open.")
            
        try:
            # Check if Comments property exists and is accessible
            if not hasattr(self.document, 'Comments'):
                raise WordDocumentError("Comments collection is not available in this document.")
            
            # Get comments count safely
            comments_count = 0
            try:
                comments_count = self.document.Comments.Count
            except Exception as e:
                raise WordDocumentError(f"Failed to access Comments collection: {e}")
            
            # Validate comment index
            if comment_index < 0 or comment_index >= comments_count:
                raise ValueError(f"Invalid comment index: {comment_index}. Valid range is 0 to {comments_count - 1}.")
            
            # Comments are 1-based in the COM API
            try:
                self.document.Comments(comment_index + 1).Delete()
            except Exception as e:
                raise WordDocumentError(f"Failed to delete comment: {e}")
        except WordDocumentError:
            # Re-raise WordDocumentError to maintain consistency
            raise
        except Exception as e:
            raise WordDocumentError(f"Error during comment deletion: {e}")

    def delete_all_comments(self) -> int:
        """
        Deletes all comments in the document.
        
        Returns:
            The number of comments deleted.
        """
        if not self.document:
            raise RuntimeError("No document open.")
            
        try:
            # Check if Comments property exists and is accessible
            if not hasattr(self.document, 'Comments'):
                # No comments to delete
                return 0
            
            # Get initial comments count safely
            comments_count = 0
            try:
                comments_count = self.document.Comments.Count
            except Exception as e:
                raise WordDocumentError(f"Failed to access Comments collection: {e}")
            
            if comments_count == 0:
                # No comments to delete
                return 0
            
            # Store initial count for return value
            deleted_count = comments_count
            
            # Delete comments in reverse order to avoid index shifting issues
            try:
                for i in range(comments_count, 0, -1):
                    try:
                        self.document.Comments(i).Delete()
                    except Exception as e:
                        print(f"Warning: Failed to delete comment at index {i}: {e}")
                        # Continue with next comment
                        continue
            except Exception as e:
                raise WordDocumentError(f"Failed to delete all comments: {e}")
            
            return deleted_count
        except WordDocumentError:
            # Re-raise WordDocumentError to maintain consistency
            raise
        except Exception as e:
            raise WordDocumentError(f"Error during deletion of all comments: {e}")

    def add_picture_caption(self, filename: str, caption_text: str, picture_index: Optional[int] = None, paragraph_index: Optional[int] = None) -> None:
        """
        Adds a caption to a picture in the document.
        
        Args:
            filename: The filename of the document.
            caption_text: The caption text to add.
            picture_index: Optional index of the picture (0-based). If not specified, adds to first picture.
            paragraph_index: Optional index of the paragraph to add caption after.
        """
        if not self.document:
            raise RuntimeError("No document open.")
        
        try:
            # Get all inline shapes (pictures)
            inline_shapes = self.document.InlineShapes
            shape_count = inline_shapes.Count
            
            if shape_count == 0:
                raise WordDocumentError("No pictures found in the document")
            
            # Determine which picture to add caption to
            target_index = picture_index if picture_index is not None else 0
            if target_index < 0 or target_index >= shape_count:
                raise ValueError(f"Invalid picture index: {target_index}. Valid range is 0 to {shape_count - 1}.")
            
            # Get the target picture
            picture = inline_shapes(target_index + 1)  # COM is 1-based
            
            # Create a range after the picture for the caption
            caption_range = picture.Range
            caption_range.Collapse(0)  # wdCollapseEnd
            caption_range.InsertAfter("\n" + caption_text)
            
            # Apply caption style if available
            try:
                caption_range.Style = self.document.Styles("Caption")
            except:
                # If Caption style doesn't exist, continue without it
                pass
                
        except Exception as e:
            raise WordDocumentError(f"Failed to add picture caption: {e}")

    def edit_comment(self, comment_index: int, new_text: str) -> None:
        """
        Edits an existing comment by its 0-based index.

        Args:
            comment_index: The 0-based index of the comment to edit.
            new_text: The new text for the comment.

        Raises:
            IndexError: If the comment index is out of range.
            WordDocumentError: If editing the comment fails.
        """
        if not self.document:
            raise RuntimeError("No document open.")
        
        try:
            # Check if comment index is valid
            if comment_index < 0 or comment_index >= self.document.Comments.Count:
                raise IndexError(f"Comment index {comment_index} out of range.")
            
            # Get the comment (COM is 1-based)
            comment = self.document.Comments(comment_index + 1)
            
            # Update the comment text
            comment.Range.Text = new_text
        except IndexError:
            raise
        except Exception as e:
            raise WordDocumentError(f"Failed to edit comment: {e}")
    
    def reply_to_comment(self, comment_index: int, reply_text: str, author: str = "User") -> None:
        """
        Replies to an existing comment.

        Args:
            comment_index: The 0-based index of the comment to reply to.
            reply_text: The text of the reply.
            author: The author of the reply (default: "User").

        Raises:
            IndexError: If the comment index is out of range.
            WordDocumentError: If replying to the comment fails.
        """
        if not self.document:
            raise RuntimeError("No document open.")
        
        try:
            # Check if comment index is valid
            if comment_index < 0 or comment_index >= self.document.Comments.Count:
                raise IndexError(f"Comment index {comment_index} out of range.")
            
            # Get the comment (COM is 1-based)
            comment = self.document.Comments(comment_index + 1)
            
            # Add a reply to the comment
            # Note: Word COM doesn't have a direct Reply method, so we need to
            # create a new comment at the same range as the original comment
            # and set the author accordingly
            reply = self.document.Comments.Add(
                Range=comment.Scope, 
                Text=reply_text
            )
            
            # Set the author of the reply
            reply.Author = author
        except IndexError:
            raise
        except Exception as e:
            raise WordDocumentError(f"Failed to reply to comment: {e}")
    
    def get_all_text(self) -> str:
        """
        Retrieves all text from the active document.

        Returns:
            A string containing all text content from the document.

        Raises:
            RuntimeError: If no document is open.
        """
        if not self.document:
            raise RuntimeError("No document open.")

        text = []
        try:
            # Iterate through all paragraphs
            for paragraph in self.document.Paragraphs:
                try:
                    text.append(paragraph.Range.Text)
                except Exception as e:
                    print(f"Warning: Failed to retrieve text from paragraph: {e}")
                    continue
        except Exception as e:
            raise WordDocumentError(f"Error retrieving document text: {e}")

        return '\n'.join(text)

    def get_comment_thread(self, comment_index: int) -> Dict[str, Any]:
        """
        Retrieves a comment thread including the original comment and all replies.

        Args:
            comment_index: The 0-based index of the original comment.

        Returns:
            A dictionary containing the original comment and all replies.

        Raises:
            IndexError: If the comment index is out of range.
            WordDocumentError: If retrieving the comment thread fails.
        """
        if not self.document:
            raise RuntimeError("No document open.")
        
        try:
            # Check if comment index is valid
            if comment_index < 0 or comment_index >= self.document.Comments.Count:
                raise IndexError(f"Comment index {comment_index} out of range.")
            
            # Get the original comment (COM is 1-based)
            original_comment = self.document.Comments(comment_index + 1)
            
            # Get the range of the original comment's scope
            original_scope = original_comment.Scope
            
            # Create the result dictionary with the original comment
            result = {
                "original_comment": {
                    "text": original_comment.Range.Text,
                    "author": original_comment.Author,
                    "date": original_comment.Date
                },
                "replies": []
            }
            
            # Search for replies to this comment
            # We consider a reply as any comment that shares the same scope as the original
            for i in range(1, self.document.Comments.Count + 1):
                comment = self.document.Comments(i)
                
                # Skip the original comment
                if comment.Index == original_comment.Index:
                    continue
                
                # Check if this comment shares the same scope as the original
                # We compare the start and end positions of the scopes
                if (comment.Scope.Start == original_scope.Start and 
                    comment.Scope.End == original_scope.End):
                    result["replies"].append({
                        "text": comment.Range.Text,
                        "author": comment.Author,
                        "date": comment.Date
                    })
            
            return result
        except IndexError:
            raise
        except Exception as e:
            raise WordDocumentError(f"Failed to get comment thread: {e}")
