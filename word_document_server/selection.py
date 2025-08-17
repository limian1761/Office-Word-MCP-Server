"""
Selection Abstraction Layer for Word Document MCP Server.

This module provides a unified interface for operating on selected document elements.
"""
from typing import List, Dict, Any
import win32com.client
from word_document_server.com_backend import WordBackend


class Selection:
    """Represents a selection of document elements."""

    def __init__(self, raw_com_elements: List[win32com.client.CDispatch], backend: WordBackend):
        """Initialize a Selection with COM elements and a backend reference.

        Args:
            raw_com_elements: List of raw COM objects representing selected elements.
            backend: WordBackend instance for executing operations.
        """
        if not raw_com_elements:
            raise ValueError("Selection cannot be empty.")
        self._elements = raw_com_elements
        self._backend = backend

    def get_text(self) -> str:
        """
        Gets the text content of the selected elements.
        If multiple elements are selected, their text is concatenated into a single string.
        """
        if not self._elements:
            return ""
        
        # Ensure that the text from each element's range is retrieved and joined.
        texts = []
        for el in self._elements:
            if hasattr(el, 'Range'):
                if hasattr(el, 'Paragraphs'):
                    # For paragraphs, exclude the paragraph mark (last character)
                    para_text = el.Range.Text
                    if para_text.endswith('\r'):
                        texts.append(para_text[:-1])
                    else:
                        texts.append(para_text)
                else:
                    texts.append(el.Range.Text)
        return "".join(texts)

    def apply_format(self, options: Dict[str, Any]):
        """
        Applies formatting to the selected elements.

        Args:
            options: A dictionary of formatting options.
                     Supported options: "bold", "italic", "font_size", "font_name", "alignment".
        """
        for element in self._elements:
            if not hasattr(element, 'Range'):
                continue
            
            com_range = element.Range
            
            if "bold" in options:
                self._backend.set_bold_for_range(com_range, options["bold"])
            if "italic" in options:
                self._backend.set_italic_for_range(com_range, options["italic"])
            if "font_size" in options:
                self._backend.set_font_size_for_range(com_range, options["font_size"])
            if "font_name" in options:
                self._backend.set_font_name_for_range(com_range, options["font_name"])
            if "alignment" in options:
                self._backend.set_alignment_for_range(com_range, options["alignment"])

    def delete(self) -> None:
        """Delete all elements in the selection."""
        for element in self._elements:
            if hasattr(element, 'Delete'):
                element.Delete()

    def insert_text(self, text: str, position: str = "after") -> 'Selection':
        """
        Inserts text relative to the selection.

        Args:
            text: The text to insert.
            position: "before", "after", or "replace".

        Returns:
            A new Selection object representing the inserted text.
        """
        if not self._elements:
            # Cannot insert relative to an empty selection
            raise ValueError("Cannot insert text relative to an empty selection.")

        if position == "after":
            # Insert after the last element in the selection
            anchor_range = self._elements[-1].Range
            self._backend.insert_paragraph_after(anchor_range, text)
        
        elif position == "before":
            # Insert before the first element in the selection
            anchor_range = self._elements[0].Range
            # Create a new range at the beginning of the anchor range
            new_range = anchor_range.Duplicate
            new_range.Collapse(1)  # wdCollapseStart = 1
            new_range.InsertAfter(text + '\r') # Use carriage return to create a new paragraph

        elif position == "replace":
            # Replace the text of all elements in the selection
            # Delete all but the first element
            for i in range(len(self._elements) - 1, 0, -1):
                if hasattr(self._elements[i], 'Delete'):
                    self._elements[i].Delete()
            # Replace the text of the first element
            if hasattr(self._elements[0], 'Range'):
                self._elements[0].Range.Text = text
        
        else:
            raise ValueError(f"Invalid position '{position}'. Must be 'before', 'after', or 'replace'.")

        # For now, returning self. A more advanced implementation might return a
        # new Selection object representing the newly inserted text.
        return self
        
    def replace_text(self, new_text: str) -> None:
        """
        Replaces the text content of the selected elements with new text.
        
        Args:
            new_text: The new text to replace the existing content.
        """
        for element in self._elements:
            if hasattr(element, 'Range'):
                element_range = element.Range
                # Use the Find and Replace functionality to properly replace text
                # This is more robust and handles paragraph boundaries correctly.
                find = element_range.Find
                find.ClearFormatting()
                find.Replacement.ClearFormatting()
                # Execute the find and replace operation
                # wdReplaceOne = 1 (replace only the first occurrence in the range)
                find.Execute(FindText=element_range.Text.strip(), ReplaceWith=new_text, Replace=1)