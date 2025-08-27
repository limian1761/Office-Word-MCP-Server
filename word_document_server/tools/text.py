import json
import logging
import os
from typing import Any, Dict, List, Optional

import pywintypes
from mcp.server.fastmcp.server import Context
from pydantic import Field

# Use the shared selector engine from core
from word_document_server.core import selector
from word_document_server.core_utils import get_backend_for_tool, mcp_server
from word_document_server.errors import (ElementNotFoundError,
                                         WordDocumentError,
                                         format_error_response,
                                         handle_tool_errors)
from word_document_server.operations import get_all_text
from word_document_server.selector import AmbiguousLocatorError, SelectorEngine

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
def get_text(
    ctx: Context = Field(description="Context object"),
    locator: Optional[Dict[str, Any]] = Field(
        description="Optional, the Locator object to find the target element(s)",
        default=None,
    ),
    start_pos: Optional[int] = Field(
        description="Optional, the starting position in the document", default=None
    ),
    end_pos: Optional[int] = Field(
        description="Optional, the ending position in the document", default=None
    ),
) -> str:
    """
    Retrieves text content from elements found by the locator or from a specific range in the document.

    Returns:
        The text content or an error message.
    """
    # Validate active document
    from word_document_server.core_utils import validate_active_document

    error = validate_active_document(ctx)
    if error:
        raise Exception(error)

    backend = get_backend_for_tool(
        ctx, ctx.session.document_state["active_document_path"]
    )

    # If locator is provided, use it to find elements
    if locator:
        try:
            selection = selector.select(backend, locator)
            return selection.get_text()
        except AmbiguousLocatorError as e:
            raise Exception(f"Ambiguous locator - multiple elements found: {str(e)}")
        except Exception as e:
            raise Exception(f"Error retrieving text with locator: {str(e)}")

    # If start_pos and end_pos are provided, get text from range
    elif start_pos is not None and end_pos is not None:
        try:
            # Ensure positions are within document bounds
            doc_range = backend.document.Range()
            doc_length = doc_range.End

            if start_pos < 0 or end_pos < 0:
                raise Exception("Start and end positions must be non-negative.")
            if start_pos >= doc_length or end_pos > doc_length:
                raise Exception(
                    f"Position out of bounds. Document length: {doc_length}"
                )
            if start_pos >= end_pos:
                raise Exception("Start position must be less than end position.")

            range_obj = backend.document.Range(Start=start_pos, End=end_pos)
            return range_obj.Text
        except Exception as e:
            raise Exception(f"Error retrieving text from range: {str(e)}")

    # If no parameters provided, get all text
    else:
        try:
            return get_all_text(backend)
        except Exception as e:
            raise Exception(f"Error retrieving all text: {str(e)}")


@mcp_server.tool()
def insert_paragraph(
    ctx: Context = Field(description="Context object"),
    locator: Dict[str, Any] = Field(
        description="The Locator object to find the anchor element"
    ),
    text: str = Field(description="The text to insert"),
    position: str = Field(
        description='"before" or "after" the anchor element', default="after"
    ),
    style: str = Field(
        description="Optional, the paragraph style name to apply", default=None
    ),
) -> str:
    """
    Inserts a new paragraph with the given text relative to the element found by the locator.

    Returns:
        A success or error message.
    """
    # Validate active document
    from word_document_server.core_utils import validate_active_document

    error = validate_active_document(ctx)
    if error:
        return error

    try:
        backend = get_backend_for_tool(
            ctx, ctx.session.document_state["active_document_path"]
        )
        selector = SelectorEngine()
        selection = selector.select(backend, locator, expect_single=True)
        selection.insert_paragraph(text, position, style)
        backend.document.Save()
        return "Successfully inserted paragraph."
    except ElementNotFoundError as e:
        return f"No elements found matching the locator: {e}. Please try simplifying your locator or use get_document_outline to check the actual document structure."
    except AmbiguousLocatorError as e:
        return f"The locator matched multiple elements: {e}. Please refine your locator to match a single element."
    except ValueError as e:
        return f"Invalid parameter: {e}"
    except Exception as e:
        return format_error_response(e)


@mcp_server.tool()
def delete_element(
    ctx: Context = Field(description="Context object"),
    locator: Dict[str, Any] = Field(
        description="The Locator object to find the target element(s) to delete"
    ),
) -> str:
    """
    Deletes the element(s) found by the locator.

    Returns:
        A success or error message.
    """
    # Validate active document
    from word_document_server.core_utils import validate_active_document

    error = validate_active_document(ctx)
    if error:
        return error

    try:
        backend = get_backend_for_tool(
            ctx, ctx.session.document_state["active_document_path"]
        )
        selector = SelectorEngine()
        selection = selector.select(backend, locator)

        # Validate that we have elements to delete
        if not selection._elements:
            return "No elements found matching the locator. Please try simplifying your locator or use get_document_structure to check the actual document structure."

        element_count = len(selection._elements)

        try:
            # 尝试删除元素
            selection.delete()
        except Exception as e:
            # 记录错误日志
            logger.error(f"Error occurred: {str(e)}", exc_info=True)
            # 直接汇报错误原因
            return format_error_response(e)

        # Save the document
        backend.document.Save()
        return f"Successfully deleted {element_count} element(s)."
    except ElementNotFoundError as e:
        return f"No elements found matching the locator: {e}. Please try simplifying your locator or use get_document_outline to check the actual document structure."
    except ValueError as e:
        return f"Invalid parameter: {e}"
    except Exception as e:
        return format_error_response(e)


@mcp_server.tool()
def apply_format(
    ctx: Context = Field(description="Context object"),
    locator: Dict[str, Any] = Field(
        description="The Locator object to find the target element(s)"
    ),
    formatting: Dict[str, Any] = Field(
        description="A dictionary containing formatting options: bold, italic, font_size, font_color, font_name, alignment, paragraph_style"
    ),
) -> str:
    """
    Applies formatting to elements found by the locator.

    Returns:
        A success or error message.
    """
    # Validate active document
    from word_document_server.core_utils import validate_active_document

    error = validate_active_document(ctx)
    if error:
        return error

    # Validate formatting parameter
    if not isinstance(formatting, dict):
        return "Formatting must be a dictionary."

    try:
        backend = get_backend_for_tool(
            ctx, ctx.session.document_state["active_document_path"]
        )
        selector = SelectorEngine()
        selection = selector.select(backend, locator)

        # Validate that we have elements to format
        if not selection._elements:
            return "No elements found matching the locator. Please try simplifying your locator or use get_document_structure to check the actual document structure."

        applied_formats = []

        # Apply formatting to each element
        for element in selection._elements:
            # Apply bold formatting
            if "bold" in formatting:
                if not isinstance(formatting["bold"], bool):
                    return "'bold' must be a boolean value."
                element.Range.Font.Bold = formatting["bold"]
                applied_formats.append(
                    f"bold={'enabled' if formatting['bold'] else 'disabled'}"
                )

            # Apply italic formatting
            if "italic" in formatting:
                if not isinstance(formatting["italic"], bool):
                    return "'italic' must be a boolean value."
                element.Range.Font.Italic = formatting["italic"]
                applied_formats.append(
                    f"italic={'enabled' if formatting['italic'] else 'disabled'}"
                )

            # Apply font size
            if "font_size" in formatting:
                size = formatting["font_size"]
                if not isinstance(size, int) or size <= 0:
                    return "'font_size' must be a positive integer."
                element.Range.Font.Size = size
                applied_formats.append(f"font_size={size}")

            # Apply font color
            if "font_color" in formatting:
                color = formatting["font_color"]
                if not isinstance(color, str) or not color:
                    return "'font_color' must be a non-empty string."

                # Convert color name to Word's RGB color value or use hex code
                color_map = {
                    "black": 0,
                    "white": 16777215,
                    "red": 255,
                    "green": 65280,
                    "blue": 16711680,
                    "yellow": 65535,
                }
                if color.lower() in color_map:
                    element.Range.Font.Color = color_map[color.lower()]
                else:
                    # Try to parse hex color (e.g., '#RRGGBB' or 'RRGGBB')
                    color = color.lstrip("#")
                    if len(color) == 6:
                        try:
                            rgb = int(color, 16)
                            element.Range.Font.Color = rgb
                        except ValueError:
                            return f"Invalid hex color format: {color}"
                    else:
                        return f"Unsupported color: {color}. Use named color or 6-digit hex code."
                applied_formats.append(f"font_color={color}")

            # Apply font name
            if "font_name" in formatting:
                name = formatting["font_name"]
                if not isinstance(name, str) or not name:
                    return "'font_name' must be a non-empty string."
                element.Range.Font.Name = name
                applied_formats.append(f"font_name={name}")

            # Apply alignment
            if "alignment" in formatting:
                alignment = formatting["alignment"]
                if alignment.lower() not in ["left", "center", "right"]:
                    return "'alignment' must be 'left', 'center', or 'right'."

                alignment_map = {
                    "left": 0,  # wdAlignParagraphLeft
                    "center": 1,  # wdAlignParagraphCenter
                    "right": 2,  # wdAlignParagraphRight
                }
                element.Range.ParagraphFormat.Alignment = alignment_map[
                    alignment.lower()
                ]
                applied_formats.append(f"alignment={alignment}")

            # Apply paragraph style
            if "paragraph_style" in formatting:
                style_name = formatting["paragraph_style"]
                if not isinstance(style_name, str) or not style_name:
                    return "'paragraph_style' must be a non-empty string."
                try:
                    element.Style = style_name
                    applied_formats.append(f"paragraph_style={style_name}")
                except:
                    # If applying style fails, try to find it in the document styles
                    style_found = False
                    for i in range(1, backend.document.Styles.Count + 1):
                        if (
                            backend.document.Styles(i).NameLocal.lower()
                            == style_name.lower()
                        ):
                            element.Style = backend.document.Styles(i)
                            style_found = True
                            break
                    if not style_found:
                        return f"Style '{style_name}' not found in document."

        # Save the document
        backend.document.Save()
        return f"Successfully applied formatting ({', '.join(applied_formats)}) to {len(selection._elements)} element(s)."
    except ElementNotFoundError as e:
        return f"No elements found matching the locator: {e}. Please try simplifying your locator or use get_document_outline to check the actual document structure."
    except ValueError as e:
        return f"Invalid parameter: {e}"
    except Exception as e:
        return format_error_response(e)


@mcp_server.tool()
def batch_apply_format(
    ctx: Context = Field(description="Context object"),
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
    # Validate active document
    from word_document_server.core_utils import validate_active_document

    error = validate_active_document(ctx)
    if error:
        return error

    # Validate operations parameter
    if not isinstance(operations, list):
        return "operations must be a list."

    if not operations:
        return "operations list cannot be empty."

    try:
        backend = get_backend_for_tool(
            ctx, ctx.session.document_state["active_document_path"]
        )
        selector = SelectorEngine()

        successful_ops = 0
        failed_ops = 0
        error_messages = []

        # Process each operation
        for i, operation in enumerate(operations):
            try:
                # Validate operation structure
                if not isinstance(operation, dict):
                    failed_ops += 1
                    error_messages.append(f"Operation {i}: Not a dictionary")
                    continue

                if "locator" not in operation or "formatting" not in operation:
                    failed_ops += 1
                    error_messages.append(
                        f"Operation {i}: Missing 'locator' or 'formatting' key"
                    )
                    continue

                locator = operation["locator"]
                formatting = operation["formatting"]

                # Select elements
                selection = selector.select(backend, locator)

                # Validate that we have elements to format
                if not selection._elements:
                    failed_ops += 1
                    error_messages.append(
                        f"Operation {i}: No elements found matching the locator"
                    )
                    continue

                # Apply formatting to each element
                for element in selection._elements:
                    # Apply bold formatting
                    if "bold" in formatting:
                        if isinstance(formatting["bold"], bool):
                            element.Range.Font.Bold = formatting["bold"]

                    # Apply italic formatting
                    if "italic" in formatting:
                        if isinstance(formatting["italic"], bool):
                            element.Range.Font.Italic = formatting["italic"]

                    # Apply font size
                    if "font_size" in formatting:
                        size = formatting["font_size"]
                        if isinstance(size, int) and size > 0:
                            element.Range.Font.Size = size

                    # Apply font color
                    if "font_color" in formatting:
                        color = formatting["font_color"]
                        if isinstance(color, str) and color:
                            # Convert color name to Word's RGB color value or use hex code
                            color_map = {
                                "black": 0,
                                "white": 16777215,
                                "red": 255,
                                "green": 65280,
                                "blue": 16711680,
                                "yellow": 65535,
                            }
                            if color.lower() in color_map:
                                element.Range.Font.Color = color_map[color.lower()]
                            else:
                                # Try to parse hex color (e.g., '#RRGGBB' or 'RRGGBB')
                                color = color.lstrip("#")
                                if len(color) == 6:
                                    try:
                                        rgb = int(color, 16)
                                        element.Range.Font.Color = rgb
                                    except ValueError:
                                        pass  # Ignore invalid hex color

                    # Apply font name
                    if "font_name" in formatting:
                        name = formatting["font_name"]
                        if isinstance(name, str) and name:
                            element.Range.Font.Name = name

                    # Apply alignment
                    if "alignment" in formatting:
                        alignment = formatting["alignment"]
                        if isinstance(alignment, str) and alignment.lower() in [
                            "left",
                            "center",
                            "right",
                        ]:
                            alignment_map = {
                                "left": 0,  # wdAlignParagraphLeft
                                "center": 1,  # wdAlignParagraphCenter
                                "right": 2,  # wdAlignParagraphRight
                            }
                            element.Range.ParagraphFormat.Alignment = alignment_map[
                                alignment.lower()
                            ]

                    # Apply paragraph style
                    if "paragraph_style" in formatting:
                        style_name = formatting["paragraph_style"]
                        if isinstance(style_name, str) and style_name:
                            try:
                                element.Style = style_name
                            except:
                                # If applying style fails, try to find it in the document styles
                                for j in range(1, backend.document.Styles.Count + 1):
                                    if (
                                        backend.document.Styles(j).NameLocal.lower()
                                        == style_name.lower()
                                    ):
                                        element.Style = backend.document.Styles(j)
                                        break

                successful_ops += 1

            except ElementNotFoundError as e:
                failed_ops += 1
                error_messages.append(
                    f"Operation {i}: No elements found matching the locator: {e}"
                )
            except Exception as e:
                failed_ops += 1
                error_messages.append(f"Operation {i}: {str(e)}")

        # Save the document if requested
        if save_document:
            backend.document.Save()

        # Prepare result summary
        result = f"Batch formatting completed: {successful_ops} successful, {failed_ops} failed out of {len(operations)} operations."
        if error_messages:
            result += "\nErrors:\n" + "\n".join(error_messages)

        return result
    except Exception as e:
        return format_error_response(e)


@mcp_server.tool()
def find_text(
    ctx: Context = Field(description="Context object"),
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
    # Validate active document
    from word_document_server.core_utils import validate_active_document

    error = validate_active_document(ctx)
    if error:
        return error

    if not find_text:
        return "Invalid input: Text to find cannot be empty."

    try:
        backend = get_backend_for_tool(
            ctx, ctx.session.document_state["active_document_path"]
        )

        # Import find_text function
        from word_document_server.operations import find_text as op_find_text

        # Find text using the backend method
        found_items = op_find_text(
            backend, 
            find_text, 
            bool(match_case), 
            bool(match_whole_word)
        )

        # Format the response according to the documentation
        result = {"matches_found": len(found_items), "matches": found_items}

        # Convert to JSON string
        return json.dumps(result, ensure_ascii=False)
    except Exception as e:
        return format_error_response(e)


@mcp_server.tool()
def replace_text(
    ctx: Context = Field(description="Context object"),
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
    # Validate active document
    from word_document_server.core_utils import validate_active_document

    error = validate_active_document(ctx)
    if error:
        return error

    if not isinstance(new_text, str):
        return "new_text must be a string."

    try:
        backend = get_backend_for_tool(
            ctx, ctx.session.document_state["active_document_path"]
        )
        selector = SelectorEngine()
        selection = selector.select(backend, locator)

        # Validate that we have elements to replace text in
        if not selection._elements:
            return "No elements found matching the locator. Please try simplifying your locator or use get_document_structure to check the actual document structure."

        # Replace text in each element
        for element in selection._elements:
            element.Range.Text = new_text

        # Save the document
        backend.document.Save()

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
def create_bulleted_list(
    ctx: Context = Field(description="Context object"),
    locator: Dict[str, Any] = Field(
        description="The Locator object to find the anchor element"
    ),
    items: List[str] = Field(
        description="A list of strings, where each string is a list item"
    ),
    position: str = Field(
        description='"before" or "after" the anchor element', default="after"
    ),
) -> str:
    """
    Creates a bulleted list relative to the element found by the locator.

    Returns:
        A success or error message.
    """
    # Validate active document
    from word_document_server.core_utils import validate_active_document

    error = validate_active_document(ctx)
    if error:
        return error

    # Validate parameters
    if not isinstance(items, list) or not items:
        return "Items must be a non-empty list of strings."
    if position not in ["before", "after"]:
        return "Position must be 'before' or 'after'."

    try:
        backend = get_backend_for_tool(
            ctx, ctx.session.document_state["active_document_path"]
        )
        selector = SelectorEngine()
        selection = selector.select(backend, locator, expect_single=True)

        # Create bulleted list
        selection.create_bulleted_list(items, position)

        # Save the document
        backend.document.Save()
        return "Successfully created bulleted list."
    except ElementNotFoundError as e:
        return f"No elements found matching the locator: {e}. Please try simplifying your locator or use get_document_structure to check the actual document structure."
    except ValueError as e:
        return f"Invalid parameter: {e}"
    except Exception as e:
        return format_error_response(e)
