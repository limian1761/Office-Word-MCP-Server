"""
Selection Abstraction Layer for Word Document MCP Server.
"""

from typing import Any, Dict, List, Optional
import win32com.client
import json

from word_document_server.utils.core_utils import ErrorCode, WordDocumentError
from word_document_server.operations import (
    set_bold_for_range, set_italic_for_range, set_font_size_for_range, 
    set_font_name_for_range, set_font_color_for_range, set_alignment_for_range,
    insert_paragraph_after, add_element_caption, insert_text_before_element, 
    insert_text_after_element, replace_element_text, set_picture_element_color_type,
    delete_element, get_element_image_info, insert_object_relative_to_element,
    get_all_inline_shapes, get_element_text, set_paragraph_style
)
from word_document_server.tools.image import get_color_type
from word_document_server.operations.comment_operations import (
    add_comment as add_comment_op,
    get_comments as get_comments_op,
    delete_comment as delete_comment_op,
    delete_all_comments as delete_all_comments_op,
    edit_comment as edit_comment_op,
    reply_to_comment as reply_to_comment_op,
    get_comment_thread as get_comment_thread_op
)
from word_document_server.operations.document_operations import get_all_text


class Selection:
    """Represents a selection of document elements."""

    def __init__(
        self, raw_com_elements: List[win32com.client.CDispatch], document: win32com.client.CDispatch
    ):
        """Initialize a Selection with COM elements and document reference.

        Args:
            raw_com_elements: List of raw COM objects representing selected elements.
            document: Word document COM object for executing operations.
        """
        if not raw_com_elements:
            raise ValueError("Selection cannot be empty.")
        self._elements = raw_com_elements
        self._document = document

    def get_text(self, join_paragraphs: bool = True) -> str:
        """
        Gets the text content of all elements in the selection.

        Args:
            join_paragraphs: Whether to join paragraphs with newlines.

        Returns:
            The text content of all selected elements.
        """
        texts = []
        for element in self._elements:
            element_text = get_element_text(element)

            # Handle paragraph joining
            if element_text:
                if join_paragraphs:
                    # Normalize paragraph breaks
                    element_text = element_text.replace("\r", "\n").replace("\n\n", "\n")
                    texts.append(element_text)
                else:
                    texts.append(element_text)
        return "".join(texts)

    def apply_format(self, options: Dict[str, Any]):
        """
        Applies formatting to the selected elements.

        Args:
            options: A dictionary of formatting options.
                     Supported options: "bold", "italic", "font_size", "font_name", "alignment".
        """
        for element in self._elements:
            # Paragraph style formatting
            if "paragraph_style" in options:
                set_paragraph_style(element, options["paragraph_style"])

            if not hasattr(element, "Range"):
                continue

            com_range = element.Range

            if "bold" in options:
                set_bold_for_range(com_range, options["bold"])
            if "italic" in options:
                set_italic_for_range(com_range, options["italic"])
            if "font_size" in options:
                set_font_size_for_range(com_range, options["font_size"])
            if "font_name" in options:
                set_font_name_for_range(com_range, options["font_name"])
            if "font_color" in options:
                set_font_color_for_range(self._document, com_range, options["font_color"])
            if "alignment" in options:
                set_alignment_for_range(self._document, com_range, options["alignment"])

    def delete(self) -> None:
        """
        Delete all elements in the selection.
        """
        if not self._elements:
            raise ValueError("No elements to delete.")

        # For paragraphs and certain element types, we need to handle them specially
        # Create a copy of the elements list to avoid issues during iteration
        elements_to_delete = list(self._elements)
        deleted_count = 0

        for element in elements_to_delete:
            try:
                delete_element(element)
                deleted_count += 1
            except Exception as e:
                raise WordDocumentError(
                    ErrorCode.ELEMENT_NOT_FOUND,
                    f"Failed to delete element: {str(e)}"
                )

    def insert_text_before(self, text: str) -> None:
        """
        Insert text before each element in the selection.

        Args:
            text: The text to insert.
        """
        for element in self._elements:
            insert_text_before_element(self._document, element, text)

    def insert_text_after(self, text: str) -> None:
        """
        Insert text after each element in the selection.

        Args:
            text: The text to insert.
        """
        for element in self._elements:
            insert_text_after_element(self._document, element, text)

    def replace_text(self, new_text: str) -> None:
        """
        Replace the text of each element in the selection.

        Args:
            new_text: The new text to replace with.
        """
        for element in self._elements:
            replace_element_text(element, new_text)

    def add_caption(self, caption_text: str, caption_style: str = "Caption", 
                   position: str = "below") -> None:
        """
        Add a caption to each element in the selection.

        Args:
            caption_text: The caption text to add.
            caption_style: The style to apply to the caption.
            position: Where to place the caption ("above" or "below").
        """
        for element in self._elements:
            add_element_caption(self._document, element, caption_text, caption_style, position)

    def set_picture_color_type(self, color_type: int) -> None:
        """
        Set the color type for picture elements in the selection.

        Args:
            color_type: The color type to set (uses Word constants).
        """
        for element in self._elements:
            if hasattr(element, 'PictureFormat'):
                set_picture_element_color_type(element, color_type)

    def get_image_info(self) -> List[Dict[str, Any]]:
        """
        Get information about image elements in the selection.

        Returns:
            A list of dictionaries containing image information.
        """
        images_info = []
        for i, element in enumerate(self._elements):
            try:
                info = get_element_image_info(element, i)
                if info:
                    images_info.append(info)
            except Exception:
                # Skip elements that are not images
                continue
        return images_info

    def insert_paragraph(self, text: str, position: str = "after", style: Optional[str] = None) -> None:
        """
        Insert a paragraph relative to elements in the selection.

        Args:
            text: The text to insert.
            position: Where to insert the paragraph ("before", "after", or "replace").
            style: Optional style to apply to the new paragraph.
        """
        for element in self._elements:
            if position == "replace":
                # Delete the element first
                element.Range.Delete()
                # Use the element's range as the insertion point
                insertion_range = element.Range
            elif position == "before":
                # Collapse the range to the start
                insertion_range = element.Range.Duplicate
                insertion_range.Collapse(1)  # wdCollapseStart = 1
            else:  # position == "after"
                # Collapse the range to the end
                insertion_range = element.Range.Duplicate
                insertion_range.Collapse(0)  # wdCollapseEnd = 0

            # Insert the text followed by a paragraph mark
            insertion_range.InsertAfter(text + "\r")  # \r is Word's paragraph mark

            # Apply style if specified
            if style:
                # Get the newly inserted paragraph to apply style
                try:
                    # Try to apply the style to the new paragraph
                    new_paragraph = self._document.Paragraphs(self._document.Paragraphs.Count)
                    new_paragraph.Style = style
                except Exception:
                    # If applying style fails, try to find it in the document styles
                    style_found = False
                    for i in range(1, self._document.Styles.Count + 1):
                        if self._document.Styles(i).NameLocal.lower() == style.lower():
                            new_paragraph.Style = self._document.Styles(i)
                            style_found = True
                            break

    def add_comment(self, text: str, author: str = "User") -> int:
        """
        Add a comment to the first element in the selection.

        Args:
            text: The comment text.
            author: The comment author.

        Returns:
            The comment ID.
        """
        if not self._elements:
            raise ValueError("No elements in selection to add comment to.")
        
        com_range_obj = self._elements[0].Range
        return add_comment_op(self._document, com_range_obj, text, author)

    def get_all_text(self) -> str:
        """
        Get all text from the document.

        Returns:
            All text in the document.
        """
        return get_all_text(self._document)

    def find_text(
        self, 
        text: str, 
        match_case: bool = False, 
        match_whole_word: bool = False,
        match_wildcards: bool = False,
        ignore_punct: bool = False,
        ignore_space: bool = False
    ) -> List[Dict[str, Any]]:
        """
        Find all occurrences of text in the document and return locator data for each match.

        Args:
            text: The text to search for.
            match_case: Whether to match case.
            match_whole_word: Whether to match whole words only.
            match_wildcards: Whether to allow wildcard characters.
            ignore_punct: Whether to ignore punctuation differences.
            ignore_space: Whether to ignore space differences.

        Returns:
            A list of dictionaries containing locator information for each found text.
        """
        if not self._document:
            raise RuntimeError("No document open.")

        if not text:
            raise ValueError("Text to find cannot be empty.")

        try:
            # Use Word's Find functionality
            finder = self._document.Content.Find
            finder.ClearFormatting()
            finder.Text = text
            finder.MatchCase = match_case
            finder.MatchWholeWord = match_whole_word
            finder.MatchWildcards = match_wildcards
            finder.IgnorePunct = ignore_punct
            finder.IgnoreSpace = ignore_space
            finder.MatchSoundsLike = False
            finder.MatchAllWordForms = False

            found_items = []
            match_index = 0
            
            # Find all occurrences
            while finder.Execute():
                range_obj = finder.Parent
                
                # Create locator data for this match
                locator_data = {
                    "target": {
                        "type": "range",
                        "filters": [
                            {"range_start": range_obj.Start},
                            {"range_end": range_obj.End}
                        ]
                    },
                    "text": range_obj.Text,
                    "start": range_obj.Start,
                    "end": range_obj.End,
                    "match_index": match_index
                }
                
                found_items.append(locator_data)
                match_index += 1

            return found_items
        except Exception as e:
            raise WordDocumentError(ErrorCode.SERVER_ERROR, f"Failed to find text '{text}': {e}")

    def create_bulleted_list(self, items: List[str], position: str = "after") -> None:
        """
        Creates a new bulleted list relative to elements in the selection.

        Args:
            items: A list of strings to become the list items.
            position: "before", "after", or "replace" the anchor element(s).
        """
        from word_document_server.operations.text_formatting import create_bulleted_list_relative_to
        
        for element in self._elements:
            if position == "replace":
                # Delete the element first
                element.Range.Delete()
                # Use the element's range as the insertion point
                insertion_range = element.Range
            elif position == "before":
                # Collapse the range to the start
                insertion_range = element.Range.Duplicate
                insertion_range.Collapse(1)  # wdCollapseStart = 1
            else:  # position == "after"
                # Collapse the range to the end
                insertion_range = element.Range.Duplicate
                insertion_range.Collapse(0)  # wdCollapseEnd = 0

            # Create the bulleted list at the insertion point
            create_bulleted_list_relative_to(self._document, insertion_range, items, "after")
