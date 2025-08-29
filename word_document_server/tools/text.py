import json
import logging
from functools import wraps
from typing import Any, Dict, List, Optional
from pydantic import Field

from mcp.server.session import ServerSession

from word_document_server.mcp_service.core import mcp_server, selector
from word_document_server.utils.app_context import AppContext
from word_document_server.utils.core_utils import validate_active_document, parse_color_hex
from word_document_server.utils.errors import (ElementNotFoundError,
                                         WordDocumentError,
                                         format_error_response,
                                         handle_tool_errors)
from word_document_server.operations.document_operations import get_all_text
from word_document_server.operations.text_formatting import set_font_color_for_range
from word_document_server.utils.core_utils import parse_color_hex


# 配置日志记录
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)
# 创建控制台处理程序
console_handler = logging.StreamHandler()
console_handler.setLevel(logging.INFO)
# 创建日志格式
formatter = logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s")
console_handler.setFormatter(formatter)
# 添加控制台处理程序到日志记录器
logger.addHandler(console_handler)


@mcp_server.tool()
@handle_tool_errors
def require_active_document(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        ctx = kwargs.get('ctx') or args[0]
        from word_document_server.core_utils import validate_active_document
        error = validate_active_document(ctx)
        if error:
            return error
        return func(*args, **kwargs)
    return wrapper

@require_active_document
def get_text(
    ctx: Context[ServerSession, AppContext] = Field(description="Context object"),
    locator: Optional[Dict[str, Any]] = Field(
        description="Optional, the Locator object to find the target element(s)",
        default=None,
    ),
) -> str:
    """
    Gets text from the active document.

    Returns:
        The requested text.
    """
    active_doc = ctx.request_context.lifespan_context.get_active_document()

    # If locator is provided, use it to find elements and get their text
    if locator:
        try:
            selection = selector.select(active_doc, locator)
            return selection.get_text()
        except ElementNotFoundError as e:
            return format_error_response(e)

    # If no parameters provided, get all text
    else:
        try:
            # Create a selection with the entire document
            from word_document_server.selector.selection import Selection
            selection = Selection([active_doc], active_doc)
            return selection.get_all_text()
        except Exception as e:
            return format_error_response(e)


@mcp_server.tool()
@require_active_document
@standardize_tool_errors
def insert_text(
    ctx: Context[ServerSession, AppContext] = Field(description="Context object"),
    locator: Dict[str, Any] = Field(
        description="The Locator object to find the anchor element"
    ),
    text: str = Field(description="The text to insert"),
    position: str = Field(
        description='"before", "after", or "replace" the anchor element', default="after"
    ),
    style: str = Field(
        description="Optional, the paragraph style name to apply", default=None
    ),
) -> str:
    """
    Inserts text relative to an anchor element.

    Returns:
        A success or error message.
    """
    try:
        active_doc = ctx.request_context.lifespan_context.get_active_document()
        from word_document_server.utils.selector import SelectorEngine, AmbiguousLocatorError
        selector = SelectorEngine()
        
        # Convert locator to Selection object
        selection = selector.select(active_doc, locator, expect_single=True)

        # Validate position parameter
        from word_document_server.utils.core_utils import validate_insert_position
        pos_error = validate_insert_position(position)
        if pos_error:
            return pos_error

        # Use the Selection method to insert text
        selection.insert_text(text, position, style)

        # Save the document
        active_doc.Save()

        return "Successfully inserted text."

    except AmbiguousLocatorError as e:
        return f"Ambiguous locator - multiple elements found: {str(e)}"
    except ElementNotFoundError as e:
        return f"No elements found matching the locator: {e}. Please try simplifying your locator or use get_document_outline to check the actual document structure."
    except ValueError as e:
        return f"Invalid parameter: {e}"
    except Exception as e:
        return format_error_response(e)


@mcp_server.tool()
@require_active_document
@standardize_tool_errors
def insert_paragraph(
    ctx: Context[ServerSession, AppContext] = Field(description="Context object"),
    locator: Dict[str, Any] = Field(
        description="The Locator object to find the anchor element"
    ),
    text: str = Field(description="The text to insert"),
    position: str = Field(
        description='"before", "after", or "replace" the anchor element', default="after"
    ),
    style: str = Field(
        description="Optional, the paragraph style name to apply", default=None
    ),
) -> str:
    """
    Inserts a new paragraph relative to an anchor element.

    Returns:
        A success or error message.
    """
    try:
        active_doc = ctx.request_context.lifespan_context.get_active_document()
        from word_document_server.utils.selector import SelectorEngine, AmbiguousLocatorError
        selector = SelectorEngine()
        
        # Convert locator to Selection object
        selection = selector.select(active_doc, locator, expect_single=True)

        # Get the first element's range
        anchor_range = selection._elements[0].Range

        # Validate position parameter
        from word_document_server.utils.core_utils import validate_insert_position
        pos_error = validate_insert_position(position)
        if pos_error:
            return pos_error

        # Handle position logic
        if position == "replace":
            # Delete the anchor element first
            anchor_range.Delete()
            # Use the anchor element's range as the insertion point
            insertion_range = anchor_range
        elif position == "before":
            # Collapse the range to the start
            insertion_range = anchor_range.Duplicate
            insertion_range.Collapse(1)  # wdCollapseStart = 1
        else:  # position == "after"
            # Collapse the range to the end
            insertion_range = anchor_range.Duplicate
            insertion_range.Collapse(0)  # wdCollapseEnd = 0

        # Insert the text followed by a paragraph mark
        insertion_range.InsertAfter(text + "\r")  # \r is Word's paragraph mark

        # Apply style if specified
        if style:
            # Get the newly inserted paragraph to apply style
            # Since we just inserted text + \r, the new paragraph should be the one after the insertion point
            try:
                # Try to apply the style to the new paragraph
                # The new paragraph will be at the end of the document or just after our insertion point
                new_paragraph = active_doc.Paragraphs(active_doc.Paragraphs.Count)
                new_paragraph.Style = style
            except Exception as style_error:
                # If applying style fails, try to find it in the document styles
                style_found = False
                for i in range(1, active_doc.Styles.Count + 1):
                    if active_doc.Styles(i).NameLocal.lower() == style.lower():
                        new_paragraph.Style = active_doc.Styles(i)
                        style_found = True
                        break
                
                # If style still not found, return a warning but don't fail the operation
                if not style_found:
                    active_doc.Save()
                    return f"Successfully inserted paragraph, but style '{style}' was not found."

        # Save the document
        active_doc.Save()
        return "Successfully inserted paragraph."
    except AmbiguousLocatorError as e:
        return f"Ambiguous locator - multiple elements found: {e}. Please refine your locator to match a single element."
    except ElementNotFoundError as e:
        return f"No elements found matching the locator: {e}. Please try simplifying your locator or use get_document_structure to check the actual document structure."
    except ValueError as e:
        return f"Invalid parameter: {e}"
    except Exception as e:
        return format_error_response(e)


@mcp_server.tool()
@require_active_document
@standardize_tool_errors
def apply_formatting(
    ctx: Context[ServerSession, AppContext] = Field(description="Context object"),
    locator: Dict[str, Any] = Field(
        description="The Locator object to find the target element(s)"
    ),
    formatting: Dict[str, Any] = Field(
        description="A dictionary of formatting options to apply"
    ),
) -> str:
    """
    Applies formatting to elements found by the locator.

    Returns:
        A success or error message.
    """

    # Validate formatting parameter
    if not isinstance(formatting, dict):
        return "Formatting must be a dictionary."

    try:
        active_doc = ctx.request_context.lifespan_context.get_active_document()
        from word_document_server.utils.selector import SelectorEngine
        selector = SelectorEngine()
        selection = selector.select(active_doc, locator)

        # Validate that we have elements to format
        if not selection._elements:
            return "No elements found matching the locator. Please try simplifying your locator or use get_document_structure to check the actual document structure."

        # Use the new Selection method to apply formatting
        selection.apply_format(formatting)

        # Save the document
        active_doc.Save()

        return "Formatting applied successfully."

    except Exception as e:
        return format_error_response(e)


@mcp_server.tool()
@require_active_document
def batch_apply_format(
    ctx: Context[ServerSession, AppContext] = Field(description="Context object"),
    operations: List[Dict[str, Any]] = Field(
        description="A list of operations, each containing 'locator' and 'formatting' keys"
    ),
    save_document: bool = Field(
        description="Whether to save the document after applying all operations",
        default=True,
    ),
) -> str:
    """
    Applies formatting to multiple elements in batch.

    Returns:
        A summary of the batch operation results.
    """
    # Validate operations parameter
    from word_document_server.utils.core_utils import validate_operations
    op_error = validate_operations(operations)
    if op_error:
        return op_error

    try:
        active_doc = ctx.request_context.lifespan_context.get_active_document()
        from word_document_server.utils.selector import SelectorEngine, AmbiguousLocatorError
        selector = SelectorEngine()

        total_elements = 0
        successful_ops = 0

        for i, op in enumerate(operations):
            try:
                locator = op["locator"]
                formatting = op["formatting"]

                # Validate formatting parameter
                from word_document_server.utils.core_utils import validate_formatting
                format_error = validate_formatting(formatting)
                if format_error:
                    return f"Operation {i}: {format_error}"

                # Convert locator to Selection object
                selection = selector.select(active_doc, locator)

                # Apply formatting to each element
                for element in selection._elements:
                    # Apply bold formatting
                    if "bold" in formatting:
                        element.Range.Font.Bold = formatting["bold"]

                    # Apply italic formatting
                    if "italic" in formatting:
                        element.Range.Font.Italic = formatting["italic"]

                    # Apply font size
                    if "font_size" in formatting:
                        element.Range.Font.Size = formatting["font_size"]

                    # Apply font color
                    if "font_color" in formatting:
                        color = formatting["font_color"]
                        try:
                            set_font_color_for_range(active_doc, element.Range, color)
                        except Exception as e:
                            return f"Operation {i}: Failed to set font color: {str(e)}"

                    # Apply font name
                    if "font_name" in formatting:
                        element.Range.Font.Name = formatting["font_name"]

                    # Apply alignment
                    if "alignment" in formatting:
                        from word_document_server.operations.text_formatting import set_alignment_for_range
                        try:
                            set_alignment_for_range(active_doc, element.Range, formatting["alignment"])
                        except Exception as e:
                            return f"Operation {i}: Failed to set alignment: {str(e)}"

                    # Apply paragraph style
                    if "paragraph_style" in formatting:
                        try:
                            element.Style = formatting["paragraph_style"]
                        except:
                            # If applying style fails, try to find it in the document styles
                            style_found = False
                            for j in range(1, active_doc.Styles.Count + 1):
                                if (
                                    active_doc.Styles(j).NameLocal.lower()
                                    == formatting["paragraph_style"].lower()
                                ):
                                    element.Style = active_doc.Styles(j)
                                    style_found = True
                                    break

                total_elements += len(selection._elements)
                successful_ops += 1

            except AmbiguousLocatorError as e:
                return f"Operation {i}: Ambiguous locator - multiple elements found: {str(e)}"
            except ElementNotFoundError as e:
                return f"Operation {i}: No elements found matching the locator: {e}. Please try simplifying your locator or use get_document_outline to check the actual document structure."
            except Exception as e:
                return f"Operation {i}: Error: {str(e)}"

        # Save the document if requested
        if save_document:
            active_doc.Save()

        return f"Successfully applied formatting to {total_elements} element(s) in {successful_ops} operations."

    except Exception as e:
        return format_error_response(e)


@mcp_server.tool()
@require_active_document
def find_text(
    ctx: Context[ServerSession, AppContext] = Field(description="Context object"),
    find_text: str = Field(description="The text to search for"),
    match_case: bool = Field(description="Whether to match case", default=False),
    match_whole_word: bool = Field(
        description="Whether to match whole words only", default=False
    ),
    match_wildcards: bool = Field(
        description="Whether to allow wildcard characters", default=False
    ),
    match_synonyms: bool = Field(
        description="Whether to match synonyms (currently unsupported)", default=False
    ),
    ignore_punct: bool = Field(
        description="Whether to ignore punctuation differences", default=False
    ),
    ignore_space: bool = Field(
        description="Whether to ignore space differences", default=False
    ),
) -> str:
    """
    Finds all occurrences of text in the document.

    Returns:
        A JSON string containing information about each found text.
    """

    if not find_text:
        return "Invalid input: Text to find cannot be empty."

    try:
        active_doc = ctx.request_context.lifespan_context.get_active_document()

        # Find text using the backend method with all parameters
        # Assuming op_find_text is imported and can accept document object
        from word_document_server.operations.text_formatting import find_text as op_find_text
        found_items = op_find_text(
            active_doc, 
            find_text, 
            match_case=bool(match_case), 
            match_whole_word=bool(match_whole_word),
            match_wildcards=bool(match_wildcards),
            ignore_punct=bool(ignore_punct),
            ignore_space=bool(ignore_space)
        )

        # Format the response according to the documentation
        result = {"matches_found": len(found_items), "matches": found_items}

        # Convert to JSON string
        return json.dumps(result, ensure_ascii=False)
    except Exception as e:
        return format_error_response(e)

@mcp_server.tool()
@require_active_document
def replace_text(
    ctx: Context[ServerSession, AppContext] = Field(description="Context object"),
    locator: Dict[str, Any] = Field(
        description="The Locator object to find the target element(s)"
    ),
    new_text: str = Field(description="The text to replace with"),
) -> str:
    """
    Replaces text in elements found by the locator with new text.

    Returns:
        A success message with the number of replacements made.
    """


    if not isinstance(new_text, str):
        return "new_text must be a string."

    try:
        active_doc = ctx.request_context.lifespan_context.get_active_document()
        from word_document_server.utils.selector import SelectorEngine, ElementNotFoundError, AmbiguousLocatorError
        selector = SelectorEngine()
        selection = selector.select(active_doc, locator)

        # Validate that we have elements to replace text in
        if not selection._elements:
            return "No elements found matching the locator. Please try simplifying your locator or use get_document_structure to check the actual document structure."

        # Replace text in each element
        for element in selection._elements:
            element.Range.Text = new_text

        # Save the document
        active_doc.Save()

        count = len(selection._elements)
        if count == 1:
            return "Successfully replaced text in 1 element."
        else:
            return f"Successfully replaced text in {count} elements."
    except ElementNotFoundError as e:
        return f"No elements found matching the locator: {e}. Please try simplifying your locator or use get_document_outline to check the actual document structure."
    except ValueError as e:
        return f"Invalid parameter: {e}"
    except Exception as e:
        return format_error_response(e)


@mcp_server.tool()
@require_active_document
def create_bulleted_list(
    ctx: Context[ServerSession, AppContext] = Field(description="Context object"),
    locator: Dict[str, Any] = Field(
        description="The Locator object to find the anchor element"
    ),
    items: List[str] = Field(
        description="A list of strings, where each string is a list item"
    ),
    position: str = Field(
        description='"before", "after", or "replace" the anchor element', default="after"
    ),
) -> str:
    """
    Creates a bulleted list relative to the element found by the locator.

    Returns:
        A success or error message.
    """


    # Validate parameters
    if not isinstance(items, list) or not items:
        return "Items must be a non-empty list of strings."
    if position not in ["before", "after", "replace"]:
        return "Position must be 'before', 'after', or 'replace'."

    try:
        active_doc = ctx.request_context.lifespan_context.get_active_document()
        from word_document_server.utils.selector import SelectorEngine, ElementNotFoundError, AmbiguousLocatorError
        selector = SelectorEngine()
        selection = selector.select(active_doc, locator, expect_single=True)

        # Create bulleted list
        selection.create_bulleted_list(items, position)

        # Save the document
        active_doc.Save()
        return "Successfully created bulleted list."
    except ElementNotFoundError as e:
        return f"No elements found matching the locator: {e}. Please try simplifying your locator or use get_document_structure to check the actual document structure."
    except ValueError as e:
        return f"Invalid parameter: {e}"
    except Exception as e:
        return format_error_response(e)
