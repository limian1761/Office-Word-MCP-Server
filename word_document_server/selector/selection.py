"""
Selection Abstraction Layer for Word Document MCP Server.
"""

from typing import Any, Dict, List, Optional
import win32com.client

from word_document_server.errors import ErrorCode, WordDocumentError
from word_document_server.operations import (
    set_bold_for_range, set_italic_for_range, set_font_size_for_range, 
    set_font_name_for_range, set_font_color_for_range, set_alignment_for_range,
    insert_paragraph_after, add_element_caption, insert_text_before_element, 
    insert_text_after_element, replace_element_text, set_picture_element_color_type,
    delete_element, get_element_image_info, insert_object_relative_to_element,
    get_all_inline_shapes, get_element_text, set_paragraph_style
)


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
        errors = []

        try:
            for element in elements_to_delete:
                try:
                    # Use element_operations to handle deletion
                    success = delete_element(element)
                    if success:
                        deleted_count += 1
                    else:
                        errors.append("Element deletion returned False")

                except Exception as inner_e:
                    error_msg = f"Failed to delete element: {str(inner_e)}"
                    errors.append(error_msg)
                    import logging

                    logging.error(error_msg)

            # If we had errors but still deleted some elements, it's a partial success
            if errors and deleted_count > 0:
                raise RuntimeError(
                    f"Partial deletion success. Deleted {deleted_count} elements, but encountered errors: {', '.join(errors)}"
                )
            elif errors and deleted_count == 0:
                raise RuntimeError(f"Failed to delete any elements. Errors: {', '.join(errors)}")

        except Exception as e:
            raise RuntimeError(f"Error during deletion: {str(e)}")

    def get_image_info(self) -> List[Dict[str, Any]]:
        """
        Gets information about all inline shapes (including images) in the selection.

        Returns:
            A list of dictionaries containing information about each inline shape.
        """
        # If the selection contains the entire document, get all inline shapes
        if len(self._elements) == 1 and hasattr(self._elements[0], 'InlineShapes'):
            return get_all_inline_shapes(self._elements[0])
            
        # Otherwise, filter for inline shapes in the selection
        image_info_list = []
        for element in self._elements:
            if hasattr(element, "Type"):  # This is likely an inline shape
                element_info = get_element_image_info(element)
                image_info_list.append(element_info)

        return image_info_list

    def insert_object(
        self, object_path: str, object_type: str = "image", position: str = "after"
    ) -> None:
        """
        Inserts an object (image, file, or OLE object) relative to the selection.

        Args:
            object_path: Path to the object file to insert.
            object_type: Type of object to insert ("image", "file", or "ole").
            position: Where to insert relative to the anchor element.
                     Supported values: "before", "after", "replace".
        """
        if not self._elements:
            raise ValueError("Cannot insert object: No anchor element selected.")

        if len(self._elements) > 1:
            raise ValueError(
                "Cannot insert object: Multiple elements selected. Please select a single element as anchor."
            )

        anchor_element = self._elements[0]

        # Validate position parameter
        if position not in ["before", "after", "replace"]:
            raise ValueError("Position must be 'before', 'after', or 'replace'.")

        # Use the operation function to insert object
        success = insert_object_relative_to_element(
            element=anchor_element,
            object_path=object_path,
            object_type=object_type,
            position=position
        )
        
        if not success:
            raise RuntimeError("Failed to insert object")

    def insert_image(self, image_path: str, position: str = "after") -> None:
        """
        Inserts an inline picture at the location of the selection.

        Args:
            image_path: Path to the image file to insert.
            position: Where to insert relative to the anchor element.
                     Supported values: "before", "after", "replace".
        """
        self.insert_object(image_path, "image", position)

    from typing import Literal

    def add_caption(
        self,
        caption_text: str,
        label: str = "Figure",
        position: Literal["above", "below"] = "below",
    ) -> None:
        """
        Adds a caption to the selected object (picture, table, etc.).

        Args:
            caption_text: The caption text to add.
            label: The label for the caption (e.g., "Figure", "Table", "Equation").
            position: Where to place the caption relative to the object.
                     Supported values: "above", "below".
        """
        if not self._elements:
            raise ValueError("Cannot add caption: No element selected.")
        if len(self._elements) > 1:
            raise ValueError(
                "Cannot add caption: Multiple elements selected. Please select a single element."
            )

        # Validate position parameter
        if position not in ["above", "below"]:
            raise ValueError("Position must be 'above' or 'below'.")
        results = []
        try:
            for element in self._elements:
                # 调用元操作处理单个元素
                success = add_element_caption(
                    element=element,
                    label=label,
                    caption_text=caption_text,
                    position=position,
                    style=self._document.Styles("Caption")
                )
                results.append(success)
            return results
        except Exception as e:
            results.append(False)
            import logging
            logging.warning(f"Batch caption operation failed: {str(e)}")
            return results

    def insert_text(
        self, text: str, position: str = "after", style: str = None
    ) -> "Selection":
        """
        Inserts text relative to the selection.

        Args:
            text: The text to insert.
            position: "before", "after", or "replace".
            style: Optional, the paragraph style name to apply.

        Returns:
            A new Selection object representing the inserted text.
        """
        if not self._elements:
            # Cannot insert relative to an empty selection
            raise ValueError("Cannot insert text relative to an empty selection.")

        if position == "before":
            # 使用元操作在首个元素前插入文本
            success = insert_text_before_element(
                element=self._elements[0],
                text=text,
                style=style,
                document=self._document
            )
            if not success:
                import logging
                logging.warning("Failed to insert text before element")

        elif position == "after":
            # 使用元操作在最后元素后插入文本
            success = insert_text_after_element(
                element=self._elements[-1],
                text=text,
                style=style,
                document=self._document
            )
            if not success:
                import logging
                logging.warning("Failed to insert text after element")

        elif position == "replace":
            # Replace the text of all elements in the selection
            for element in self._elements:
                replace_element_text(element, new_text=text, style=style)

        else:
            raise ValueError(
                f"Invalid position '{position}'. Must be 'before', 'after', or 'replace'."
            )

        # For now, returning self. A more advanced implementation might return a
        # new Selection object representing the newly inserted text.
        return self

    def replace_text(self, new_text: str) -> None:
        """
        Replaces the text content of the selected elements with new text.
        Preserves paragraph structure by handling paragraph breaks intelligently.

        Args:
            new_text: The new text to replace the existing content.
        """
        for element in self._elements:
            replace_element_text(element, new_text)

    def set_picture_color_type(self, color_type: str) -> None:
        """
        Sets the color type for all picture elements in the selection.

        Args:
            color_type: The color type to set ("color", "grayscale", "black_and_white", "watermark").
        """
        color_types = {
            "color": 0,  # msoPictureColorTypeColor
            "grayscale": 1,  # msoPictureColorTypeGrayscale
            "black_and_white": 2,  # msoPictureColorTypeBlackAndWhite
            "watermark": 3,  # msoPictureColorTypeWatermark
        }

        if color_type not in color_types:
            raise ValueError(
                f"Invalid color type '{color_type}'. Must be one of: {', '.join(color_types.keys())}"
            )

        color_code = color_types[color_type]

        for element in self._elements:
            set_picture_element_color_type(element, color_code)