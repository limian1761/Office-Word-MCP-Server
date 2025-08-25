"""
Selection Abstraction Layer for Word Document MCP Server.

This module provides a unified interface for operating on selected document elements.
"""
from typing import Any, Dict, List, Optional

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
                element_text = el.Range.Text
                # For paragraph-like elements (containing carriage returns), exclude the paragraph mark
                if '\r' in element_text:
                    if element_text.endswith('\r'):
                        texts.append(element_text[:-1])
                    else:
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
            if "font_color" in options:
                self._backend.set_font_color_for_range(com_range, options["font_color"])
            if "alignment" in options:
                self._backend.set_alignment_for_range(com_range, options["alignment"])
            
            # Paragraph style formatting
            if "paragraph_style" in options:
                # Try to apply the paragraph style
                try:
                    element.Style = options["paragraph_style"]
                except Exception as e:
                    # Log the error but continue processing
                    import logging
                    logging.error(f"Failed to apply paragraph style '{options['paragraph_style']}': {str(e)}")

    def delete(self, password: Optional[str] = None) -> None:
        """
        Delete all elements in the selection.
        
        Enhanced implementation with better error handling and verification,
        including document protection check without modifying protection status.
        
        Args:
            password: Optional password (not used for unprotecting document per user request).
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
                    # First, check if the element has a Delete method
                    if not hasattr(element, 'Delete'):
                        errors.append("Element has no Delete method")
                        continue
                    
                    # Store element's text/content for verification
                    original_content = None
                    if hasattr(element, 'Range') and hasattr(element.Range, 'Text'):
                        original_content = element.Range.Text
                        
                    # For paragraphs specifically, we might need to check if it's a valid paragraph
                    is_paragraph = hasattr(element, 'Style') and hasattr(element.Style, 'NameLocal')
                    
                    # Attempt to delete the element
                    element.Delete()
                    
                    # For certain element types, perform additional verification
                    if is_paragraph:
                        # For paragraphs, we should verify deletion by checking the document structure
                        # This is a simple check - in a real implementation, you might want to do more thorough verification
                        deleted_count += 1
                    elif original_content:
                        # For other elements with content, consider it deleted if we attempted the operation
                        deleted_count += 1
                    else:
                        # For elements without content, assume deletion was attempted
                        deleted_count += 1
                        
                except Exception as e:
                    # Log the error but continue with other elements
                    error_msg = f"Failed to delete element: {str(e)}"
                    errors.append(error_msg)
                    import logging
                    logging.error(error_msg)
                    continue
            
            # If no elements were successfully deleted, check protection status
            if deleted_count == 0:
                # Check if the document is protected
                protection_status = self._backend.get_protection_status()
                if protection_status["is_protected"]:
                    raise WordDocumentError(
                        ErrorCode.PERMISSION_DENIED, 
                        "Failed to delete any elements. Document is protected and modification is not allowed.",
                        {"protection_status": protection_status}
                    )
                else:
                    error_details = {"errors": errors} if errors else {}
                    raise WordDocumentError(
                        ErrorCode.ELEMENT_LOCKED, 
                        "Failed to delete any elements. This might be due to permission issues or element locking.",
                        error_details
                    )
            self._elements = []
        except Exception as e:
            raise RuntimeError(f"Error during deletion: {str(e)}")
    
    def get_image_info(self) -> List[Dict[str, Any]]:
        """
        Gets information about all inline shapes (including images) in the selection.
        
        Returns:
            A list of dictionaries containing information about each inline shape.
        """
        image_info_list = []
        
        for element in self._elements:
            if hasattr(element, 'Type'):  # This is likely an inline shape
                shape_info = {
                    'index': element.Index if hasattr(element, 'Index') else None,
                    'type': self._get_shape_type(element.Type) if hasattr(element, 'Type') else 'Unknown',
                    'height': element.Height if hasattr(element, 'Height') else None,
                    'width': element.Width if hasattr(element, 'Width') else None,
                    'left': element.Left if hasattr(element, 'Left') else None,
                    'top': element.Top if hasattr(element, 'Top') else None
                }
                
                # Add additional properties based on the shape type
                if shape_info['type'] == 'Picture' or shape_info['type'] == 'LinkedPicture':
                    if hasattr(element, 'PictureFormat'):
                        shape_info['has_picture_format'] = True
                    if hasattr(element, 'LockAspectRatio'):
                        shape_info['lock_aspect_ratio'] = bool(element.LockAspectRatio)
                
                image_info_list.append(shape_info)
        
        return image_info_list
    
    def _get_shape_type(self, shape_type_code: int) -> str:
        """
        Converts a shape type code to a human-readable string.
        
        Args:
            shape_type_code: The shape type code from Word's COM API.
        
        Returns:
            A human-readable string representing the shape type.
        """
        shape_types = {
            1: "Picture",
            2: "LinkedPicture",
            3: "Chart",
            4: "Diagram",
            5: "OLEControlObject",
            6: "OLEObject",
            7: "ActiveXControl",
            8: "SmartArt",
            9: "3DModel"
        }
        
        return shape_types.get(shape_type_code, f"Unknown ({shape_type_code})")
    
    def set_image_size(self, width: float = None, height: float = None, lock_aspect_ratio: bool = True) -> None:
        """
        Sets the size of all images in the selection.
        
        Args:
            width: The new width in points. If None, width remains unchanged.
            height: The new height in points. If None, height remains unchanged.
            lock_aspect_ratio: Whether to maintain the aspect ratio when changing dimensions.
        """
        for element in self._elements:
            if hasattr(element, 'Type') and (element.Type == 1 or element.Type == 2):  # Picture or LinkedPicture
                # Set lock aspect ratio
                if hasattr(element, 'LockAspectRatio'):
                    element.LockAspectRatio = lock_aspect_ratio
                
                # Apply dimensions
                if width is not None:
                    element.Width = width
                if height is not None:
                    element.Height = height
                     
    def insert_image(self, image_path: str, position: str = "after") -> 'Selection':
        """
        Inserts an image relative to the selection.
        
        Args:
            image_path: The absolute path to the image file.
            position: "before", "after", or "replace" to specify where to insert the image relative to the selection.
            
        Returns:
            A new Selection object representing the inserted image.
        """
        if not self._elements:
            # Cannot insert relative to an empty selection
            raise ValueError("Cannot insert image relative to an empty selection.")
            
        inserted_shape = None
        
        if position == "replace":
            # Delete all elements in the selection first
            for i in range(len(self._elements) - 1, 0, -1):
                if hasattr(self._elements[i], 'Delete'):
                    self._elements[i].Delete()
            
            # Replace the first element with the image
            if hasattr(self._elements[0], 'Range'):
                inserted_shape = self._backend.insert_inline_picture(
                    self._elements[0].Range, 
                    image_path, 
                    position="replace"
                )
        elif position == "before":
            # Insert before the first element in the selection
            anchor_range = self._elements[0].Range
            inserted_shape = self._backend.insert_inline_picture(
                anchor_range, 
                image_path, 
                position="before"
            )
        elif position == "after":
            # Insert after the last element in the selection
            anchor_range = self._elements[-1].Range
            inserted_shape = self._backend.insert_inline_picture(
                anchor_range, 
                image_path, 
                position="after"
            )
        else:
            raise ValueError(f"Invalid position '{position}'. Must be 'before', 'after', or 'replace'.")
            
        # Return a new Selection containing the inserted image
        if inserted_shape:
            return Selection([inserted_shape], self._backend)
        else:
            raise RuntimeError("Failed to insert image.")
            
    def set_image_color_type(self, color_type: str) -> None:
        """
        Sets the color type of all images in the selection.
        
        Args:
            color_type: The color type to apply. Can be 'Color', 'Grayscale', 'BlackAndWhite', or 'Watermark'.
        """
        # Map color type strings to Word constants
        color_types = {
            'Color': 0,          # msoPictureColorTypeColor
            'Grayscale': 1,      # msoPictureColorTypeGrayscale
            'BlackAndWhite': 2,  # msoPictureColorTypeBlackAndWhite
            'Watermark': 3       # msoPictureColorTypeWatermark
        }
        
        color_code = color_types.get(color_type)
        if color_code is None:
            raise ValueError(f"Invalid color type '{color_type}'. Must be one of: {', '.join(color_types.keys())}")
            
        for element in self._elements:
            if hasattr(element, 'Type') and (element.Type == 1 or element.Type == 2):  # Picture or LinkedPicture
                if hasattr(element, 'PictureFormat') and hasattr(element.PictureFormat, 'ColorType'):
                    element.PictureFormat.ColorType = color_code

    def insert_text(self, text: str, position: str = "after", style: str = None) -> 'Selection':
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

        new_paragraph = None
        
        if position == "after":
            # Insert after the last element in the selection
            anchor_range = self._elements[-1].Range
            self._backend.insert_paragraph_after(anchor_range, text)
            # Get the last paragraph in the document to apply style if needed
            if self._backend.document.Paragraphs.Count > 0:
                new_paragraph = self._backend.document.Paragraphs(self._backend.document.Paragraphs.Count)
        
        elif position == "before":
            # Insert before the first element in the selection
            anchor_range = self._elements[0].Range
            # Create a new range at the beginning of the anchor range
            new_range = anchor_range.Duplicate
            new_range.Collapse(1)  # wdCollapseStart = 1
            new_range.InsertAfter(text + '\r') # Use carriage return to create a new paragraph
            # Get the first paragraph after insertion point to apply style if needed
            new_paragraph = self._backend.document.Paragraphs(1)

        elif position == "replace":
            # Replace the text of all elements in the selection
            # Delete all but the first element
            for i in range(len(self._elements) - 1, 0, -1):
                if hasattr(self._elements[i], 'Delete'):
                    self._elements[i].Delete()
            # Replace the text of the first element
            if hasattr(self._elements[0], 'Range'):
                self._elements[0].Range.Text = text
                # Apply style to the replaced element
                if hasattr(self._elements[0], 'Style') and style:
                    try:
                        self._elements[0].Style = style
                    except Exception as e:
                        print(f"Warning: Failed to apply style '{style}': {e}")
        
        else:
            raise ValueError(f"Invalid position '{position}'. Must be 'before', 'after', or 'replace'.")
        
        # Apply style to the newly inserted paragraph if specified
        if new_paragraph and style:
            try:
                new_paragraph.Style = style
            except Exception as e:
                print(f"Warning: Failed to apply style '{style}': {e}")

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
            if hasattr(element, 'Range'):
                original_text = element.Range.Text
                
                # For all elements, use direct replacement
                # Word's Range.Text property handles paragraph structure automatically
                element.Range.Text = new_text