"""
Formatting tools for Word Document Server using COM.
"""
import os
from typing import List, Optional
from word_document_server.utils import com_utils
from word_document_server.utils.file_utils import check_file_writeable, ensure_docx_extension

# Word color constants
wdColorBlack = 0
wdColorBlue = 16711680
wdColorRed = 255
# ... add other colors as needed

def hex_to_bgr(hex_color):
    """Converts a hex color string (RRGGBB) to a BGR integer."""
    hex_color = hex_color.lstrip('#')
    r = int(hex_color[0:2], 16)
    g = int(hex_color[2:4], 16)
    b = int(hex_color[4:6], 16)
    return (b << 16) | (g << 8) | r

async def format_text(filename: str, paragraph_index: int, start_pos: int, end_pos: int, 
                     bold: Optional[bool] = None, italic: Optional[bool] = None, 
                     underline: Optional[bool] = None, color: Optional[str] = None,
                     font_size: Optional[int] = None, font_name: Optional[str] = None) -> str:
    """Format a specific range of text within a paragraph using COM."""
    filename = ensure_docx_extension(filename)
    if not os.path.exists(filename):
        return f"Document {filename} does not exist"

    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return f"Cannot modify document: {error_message}."

    doc = None
    try:
        doc = com_utils.open_document(filename)
        if paragraph_index < 0 or paragraph_index >= doc.Paragraphs.Count:
            return f"Invalid paragraph index. Document has {doc.Paragraphs.Count} paragraphs."

        # COM is 1-based for paragraphs
        p = doc.Paragraphs(paragraph_index + 1)
        # Character positions are 1-based in COM Range
        text_len = len(p.Range.Text.rstrip('\r\n'))
        if start_pos < 0 or end_pos > text_len or start_pos >= end_pos:
            return f"Invalid text positions. Paragraph has {text_len} characters."

        # Create a range for the target text
        # The paragraph range includes the trailing paragraph mark, so we adjust
        start_char = p.Range.Start + start_pos
        end_char = p.Range.Start + end_pos
        text_range = doc.Range(Start=start_char, End=end_char)
        
        font = text_range.Font
        if bold is not None:
            font.Bold = bold
        if italic is not None:
            font.Italic = italic
        if underline is not None:
            font.Underline = underline
        if color:
            try:
                font.Color = hex_to_bgr(color)
            except Exception:
                return f"Invalid color format: '{color}'. Use RRGGBB hex format."
        if font_size:
            font.Size = font_size
        if font_name:
            font.Name = font_name
            
        doc.Save()
        return f"Text formatted successfully in paragraph {paragraph_index}."
    except Exception as e:
        return f"Failed to format text: {str(e)}"
    finally:
        if doc:
            doc.Close(SaveChanges=0)

async def create_custom_style(filename: str, style_name: str, 
                             bold: Optional[bool] = None, italic: Optional[bool] = None,
                             font_size: Optional[int] = None, font_name: Optional[str] = None,
                             color: Optional[str] = None, base_style: Optional[str] = None) -> str:
    """Create a custom style in the document using COM."""
    filename = ensure_docx_extension(filename)
    if not os.path.exists(filename):
        return f"Document {filename} does not exist"

    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return f"Cannot modify document: {error_message}."

    doc = None
    try:
        doc = com_utils.open_document(filename)
        # wdStyleTypeParagraph = 1
        style = doc.Styles.Add(Name=style_name, Type=1)
        font = style.Font
        if bold is not None:
            font.Bold = bold
        if italic is not None:
            font.Italic = italic
        if font_size:
            font.Size = font_size
        if font_name:
            font.Name = font_name
        if color:
            try:
                font.Color = hex_to_bgr(color)
            except Exception:
                return f"Invalid color format: '{color}'. Use RRGGBB hex format."
        if base_style:
            try:
                style.BaseStyle = base_style
            except Exception:
                return f"Base style '{base_style}' not found."

        doc.Save()
        return f"Style '{style_name}' created successfully."
    except Exception as e:
        return f"Failed to create style: {str(e)}"
    finally:
        if doc:
            doc.Close(SaveChanges=0)

async def format_table(filename: str, table_index: int, 
                      has_header_row: Optional[bool] = None,
                      border_style: Optional[str] = None,
                      shading: Optional[List[List[str]]] = None) -> str:
    """Format a table with borders, shading, and structure using COM."""
    filename = ensure_docx_extension(filename)
    if not os.path.exists(filename):
        return f"Document {filename} does not exist"

    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return f"Cannot modify document: {error_message}."

    doc = None
    try:
        doc = com_utils.open_document(filename)
        if table_index < 0 or table_index >= doc.Tables.Count:
            return f"Invalid table index. Document has {doc.Tables.Count} tables."

        table = doc.Tables(table_index + 1) # COM is 1-based
        if has_header_row:
            table.Rows(1).HeadingFormat = True
        
        if border_style:
            # This is a simplified mapping. COM offers much more control.
            # wdLineStyleSingle = 1, wdLineStyleNone = 0
            line_style = 1 if border_style != 'none' else 0
            table.Borders.InsideLineStyle = line_style
            table.Borders.OutsideLineStyle = line_style

        if shading:
            for r, row_data in enumerate(shading):
                for c, color_hex in enumerate(row_data):
                    if r < table.Rows.Count and c < table.Columns.Count:
                        cell = table.Cell(r + 1, c + 1)
                        cell.Shading.BackgroundPatternColor = hex_to_bgr(color_hex)

        doc.Save()
        return f"Table at index {table_index} formatted successfully."
    except Exception as e:
        return f"Failed to format table: {str(e)}"
    finally:
        if doc:
            doc.Close(SaveChanges=0)

# The remaining table formatting functions are complex and will be implemented
# based on user needs. For now, they return a "not implemented" message.

async def set_table_cell_shading(filename: str, table_index: int, row_index: int, 
                                col_index: int, fill_color: str, pattern: str = "clear") -> str:
    """Apply shading/filling to a specific table cell using COM."""
    filename = ensure_docx_extension(filename)
    if not os.path.exists(filename):
        return f"Document {filename} does not exist"

    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return f"Cannot modify document: {error_message}."

    doc = None
    try:
        doc = com_utils.open_document(filename)
        if table_index < 0 or table_index >= doc.Tables.Count:
            return f"Invalid table index. Document has {doc.Tables.Count} tables."

        table = doc.Tables(table_index + 1)

        if row_index < 0 or row_index >= table.Rows.Count or \
           col_index < 0 or col_index >= table.Columns.Count:
            return "Invalid row or column index."

        cell = table.Cell(row_index + 1, col_index + 1)
        try:
            cell.Shading.BackgroundPatternColor = hex_to_bgr(fill_color)
        except Exception:
            return f"Invalid color format: '{fill_color}'. Use RRGGBB hex format."
        
        doc.Save()
        return f"Cell ({row_index}, {col_index}) shading set successfully."
    except Exception as e:
        return f"Failed to set cell shading: {str(e)}"
    finally:
        if doc:
            doc.Close(SaveChanges=0)

async def apply_table_alternating_rows(filename: str, table_index: int, 
                                     color1: str = "FFFFFF", color2: str = "F2F2F2") -> str:
    """Apply alternating row colors to a table for better readability using COM."""
    filename = ensure_docx_extension(filename)
    if not os.path.exists(filename):
        return f"Document {filename} does not exist"

    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return f"Cannot modify document: {error_message}."

    doc = None
    try:
        doc = com_utils.open_document(filename)
        if table_index < 0 or table_index >= doc.Tables.Count:
            return f"Invalid table index. Document has {doc.Tables.Count} tables."

        table = doc.Tables(table_index + 1)
        
        try:
            bgr_color1 = hex_to_bgr(color1)
            bgr_color2 = hex_to_bgr(color2)
        except Exception:
            return f"Invalid color format. Use RRGGBB hex format."

        for i, row in enumerate(table.Rows):
            color = bgr_color1 if (i % 2) == 0 else bgr_color2
            row.Shading.BackgroundPatternColor = color
            
        doc.Save()
        return f"Alternating row colors applied to table {table_index}."
    except Exception as e:
        return f"Failed to apply alternating row colors: {str(e)}"
    finally:
        if doc:
            doc.Close(SaveChanges=0)

async def highlight_table_header(filename: str, table_index: int, 
                               header_color: str = "4472C4", text_color: str = "FFFFFF") -> str:
    """Apply special highlighting to table header row using COM."""
    filename = ensure_docx_extension(filename)
    if not os.path.exists(filename):
        return f"Document {filename} does not exist"

    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return f"Cannot modify document: {error_message}."

    doc = None
    try:
        doc = com_utils.open_document(filename)
        if table_index < 0 or table_index >= doc.Tables.Count:
            return f"Invalid table index. Document has {doc.Tables.Count} tables."

        table = doc.Tables(table_index + 1)
        if table.Rows.Count == 0:
            return "Cannot highlight header of an empty table."

        header_row = table.Rows(1)
        try:
            header_row.Shading.BackgroundPatternColor = hex_to_bgr(header_color)
            header_row.Range.Font.Color = hex_to_bgr(text_color)
        except Exception:
            return f"Invalid color format. Use RRGGBB hex format."
            
        doc.Save()
        return f"Table {table_index} header highlighted."
    except Exception as e:
        return f"Failed to highlight table header: {str(e)}"
    finally:
        if doc:
            doc.Close(SaveChanges=0)

async def merge_table_cells(filename: str, table_index: int, start_row: int, start_col: int, 
                          end_row: int, end_col: int) -> str:
    """Merge cells in a rectangular area of a table using COM."""
    filename = ensure_docx_extension(filename)
    if not os.path.exists(filename):
        return f"Document {filename} does not exist"

    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return f"Cannot modify document: {error_message}."

    doc = None
    try:
        doc = com_utils.open_document(filename)
        if table_index < 0 or table_index >= doc.Tables.Count:
            return f"Invalid table index. Document has {doc.Tables.Count} tables."

        table = doc.Tables(table_index + 1) # COM is 1-based

        # Validate indices
        if start_row < 0 or start_row >= table.Rows.Count or 
           end_row < 0 or end_row >= table.Rows.Count or 
           start_col < 0 or start_col >= table.Columns.Count or 
           end_col < 0 or end_col >= table.Columns.Count:
            return "Invalid row or column index."

        if start_row > end_row or start_col > end_col:
            return "Start indices must be less than or equal to end indices."

        start_cell = table.Cell(start_row + 1, start_col + 1)
        end_cell = table.Cell(end_row + 1, end_col + 1)
        
        start_cell.Merge(end_cell)
        
        doc.Save()
        return f"Cells from ({start_row}, {start_col}) to ({end_row}, {end_col}) merged successfully."
    except Exception as e:
        return f"Failed to merge cells: {str(e)}"
    finally:
        if doc:
            doc.Close(SaveChanges=0)

async def merge_table_cells_horizontal(filename: str, table_index: int, row_index: int, 
                                     start_col: int, end_col: int) -> str:
    """Merge cells horizontally in a single row using COM."""
    return await merge_table_cells(filename, table_index, row_index, start_col, row_index, end_col)

async def merge_table_cells_vertical(filename: str, table_index: int, col_index: int, 
                                   start_row: int, end_row: int) -> str:
    """Merge cells vertically in a single column using COM."""
    return await merge_table_cells(filename, table_index, start_row, col_index, end_row, col_index)

async def set_table_cell_alignment(filename: str, table_index: int, row_index: int, col_index: int,
                                 horizontal: str = "left", vertical: str = "top") -> str:
    """Set text alignment for a specific table cell using COM."""
    filename = ensure_docx_extension(filename)
    if not os.path.exists(filename):
        return f"Document {filename} does not exist"

    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return f"Cannot modify document: {error_message}."

    doc = None
    try:
        doc = com_utils.open_document(filename)
        if table_index < 0 or table_index >= doc.Tables.Count:
            return f"Invalid table index. Document has {doc.Tables.Count} tables."

        table = doc.Tables(table_index + 1)

        if row_index < 0 or row_index >= table.Rows.Count or \
           col_index < 0 or col_index >= table.Columns.Count:
            return "Invalid row or column index."

        cell = table.Cell(row_index + 1, col_index + 1)

        # Horizontal alignment constants
        h_align_map = {"left": 0, "center": 1, "right": 2, "justify": 3}
        # Vertical alignment constants
        v_align_map = {"top": 0, "center": 1, "bottom": 3}

        cell.Range.ParagraphFormat.Alignment = h_align_map.get(horizontal.lower(), 0)
        cell.VerticalAlignment = v_align_map.get(vertical.lower(), 0)
        
        doc.Save()
        return f"Cell ({row_index}, {col_index}) alignment set to {horizontal}/{vertical}."
    except Exception as e:
        return f"Failed to set cell alignment: {str(e)}"
    finally:
        if doc:
            doc.Close(SaveChanges=0)


async def set_table_alignment_all(filename: str, table_index: int, 
                                horizontal: str = "left", vertical: str = "top") -> str:
    """Set text alignment for all cells in a table using COM."""
    filename = ensure_docx_extension(filename)
    if not os.path.exists(filename):
        return f"Document {filename} does not exist"

    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return f"Cannot modify document: {error_message}."

    doc = None
    try:
        doc = com_utils.open_document(filename)
        if table_index < 0 or table_index >= doc.Tables.Count:
            return f"Invalid table index. Document has {doc.Tables.Count} tables."

        table = doc.Tables(table_index + 1)

        # Horizontal alignment constants
        h_align_map = {"left": 0, "center": 1, "right": 2, "justify": 3}
        # Vertical alignment constants
        v_align_map = {"top": 0, "center": 1, "bottom": 3}

        h_align = h_align_map.get(horizontal.lower(), 0)
        v_align = v_align_map.get(vertical.lower(), 0)

        for r in range(1, table.Rows.Count + 1):
            for c in range(1, table.Columns.Count + 1):
                cell = table.Cell(r, c)
                cell.Range.ParagraphFormat.Alignment = h_align
                cell.VerticalAlignment = v_align
        
        doc.Save()
        return f"All cells in table {table_index} alignment set to {horizontal}/{vertical}."
    except Exception as e:
        return f"Failed to set table alignment: {str(e)}"
    finally:
        if doc:
            doc.Close(SaveChanges=0)

async def set_table_column_width(filename: str, table_index: int, col_index: int, 
                                width: float, width_type: str = "points") -> str:
    """Set the width of a specific table column using COM."""
    filename = ensure_docx_extension(filename)
    if not os.path.exists(filename):
        return f"Document {filename} does not exist"

    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return f"Cannot modify document: {error_message}."

    doc = None
    try:
        doc = com_utils.open_document(filename)
        if table_index < 0 or table_index >= doc.Tables.Count:
            return f"Invalid table index. Document has {doc.Tables.Count} tables."

        table = doc.Tables(table_index + 1)

        if col_index < 0 or col_index >= table.Columns.Count:
            return "Invalid column index."

        column = table.Columns(col_index + 1)
        
        # Preferred width type constants
        # wdPreferredWidthPoints = 2, wdPreferredWidthPercent = 3
        if width_type.lower() == "points":
            column.PreferredWidthType = 2
            column.PreferredWidth = width
        elif width_type.lower() == "percent":
            column.PreferredWidthType = 3
            column.PreferredWidth = width
        elif width_type.lower() == "inches":
            column.PreferredWidthType = 2
            column.PreferredWidth = width * 72 # Convert inches to points
        else:
            return f"Unsupported width type: {width_type}. Use 'points', 'percent', or 'inches'."

        doc.Save()
        return f"Column {col_index} width set to {width} {width_type}."
    except Exception as e:
        return f"Failed to set column width: {str(e)}"
    finally:
        if doc:
            doc.Close(SaveChanges=0)

async def set_table_column_widths(filename: str, table_index: int, widths: list, 
                                 width_type: str = "points") -> str:
    """Set the widths of multiple table columns using COM."""
    filename = ensure_docx_extension(filename)
    if not os.path.exists(filename):
        return f"Document {filename} does not exist"

    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return f"Cannot modify document: {error_message}."

    doc = None
    try:
        doc = com_utils.open_document(filename)
        if table_index < 0 or table_index >= doc.Tables.Count:
            return f"Invalid table index. Document has {doc.Tables.Count} tables."

        table = doc.Tables(table_index + 1)

        if len(widths) > table.Columns.Count:
            return f"Widths list has more items ({len(widths)}) than table has columns ({table.Columns.Count})."

        for i, width in enumerate(widths):
            await set_table_column_width(filename, table_index, i, width, width_type)
        
        return f"Successfully set widths for {len(widths)} columns."
    except Exception as e:
        return f"Failed to set column widths: {str(e)}"
    finally:
        if doc:
            doc.Close(SaveChanges=0)

async def set_table_width(filename: str, table_index: int, width: float, 
                         width_type: str = "points") -> str:
    """Set the overall width of a table using COM."""
    filename = ensure_docx_extension(filename)
    if not os.path.exists(filename):
        return f"Document {filename} does not exist"

    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return f"Cannot modify document: {error_message}."

    doc = None
    try:
        doc = com_utils.open_document(filename)
        if table_index < 0 or table_index >= doc.Tables.Count:
            return f"Invalid table index. Document has {doc.Tables.Count} tables."

        table = doc.Tables(table_index + 1)
        
        if width_type.lower() == "points":
            table.PreferredWidthType = 2
            table.PreferredWidth = width
        elif width_type.lower() == "percent":
            table.PreferredWidthType = 3
            table.PreferredWidth = width
        elif width_type.lower() == "inches":
            table.PreferredWidthType = 2
            table.PreferredWidth = width * 72
        else:
            return f"Unsupported width type: {width_type}. Use 'points', 'percent', or 'inches'."

        doc.Save()
        return f"Table {table_index} width set to {width} {width_type}."
    except Exception as e:
        return f"Failed to set table width: {str(e)}"
    finally:
        if doc:
            doc.Close(SaveChanges=0)

async def auto_fit_table_columns(filename: str, table_index: int) -> str:
    """Set table columns to auto-fit based on content using COM."""
    filename = ensure_docx_extension(filename)
    if not os.path.exists(filename):
        return f"Document {filename} does not exist"

    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return f"Cannot modify document: {error_message}."

    doc = None
    try:
        doc = com_utils.open_document(filename)
        if table_index < 0 or table_index >= doc.Tables.Count:
            return f"Invalid table index. Document has {doc.Tables.Count} tables."

        table = doc.Tables(table_index + 1)
        table.AutoFitBehavior(1) # wdAutoFitContent
        
        doc.Save()
        return f"Table {table_index} columns set to auto-fit."
    except Exception as e:
        return f"Failed to auto-fit table columns: {str(e)}"
    finally:
        if doc:
            doc.Close(SaveChanges=0)

async def format_table_cell_text(filename: str, table_index: int, row_index: int, col_index: int,
                                 text_content: Optional[str] = None, bold: Optional[bool] = None, italic: Optional[bool] = None,
                                 underline: Optional[bool] = None, color: Optional[str] = None, font_size: Optional[int] = None,
                                 font_name: Optional[str] = None) -> str:
    """Format text within a specific table cell using COM."""
    filename = ensure_docx_extension(filename)
    if not os.path.exists(filename):
        return f"Document {filename} does not exist"

    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return f"Cannot modify document: {error_message}."

    doc = None
    try:
        doc = com_utils.open_document(filename)
        if table_index < 0 or table_index >= doc.Tables.Count:
            return f"Invalid table index. Document has {doc.Tables.Count} tables."

        table = doc.Tables(table_index + 1)

        if row_index < 0 or row_index >= table.Rows.Count or \
           col_index < 0 or col_index >= table.Columns.Count:
            return "Invalid row or column index."

        cell_range = table.Cell(row_index + 1, col_index + 1).Range
        
        if text_content is not None:
            cell_range.Text = text_content

        font = cell_range.Font
        if bold is not None:
            font.Bold = bold
        if italic is not None:
            font.Italic = italic
        if underline is not None:
            font.Underline = underline
        if color:
            try:
                font.Color = hex_to_bgr(color)
            except Exception:
                return f"Invalid color format: '{color}'. Use RRGGBB hex format."
        if font_size:
            font.Size = font_size
        if font_name:
            font.Name = font_name
            
        doc.Save()
        return f"Cell ({row_index}, {col_index}) text formatted successfully."
    except Exception as e:
        return f"Failed to format cell text: {str(e)}"
    finally:
        if doc:
            doc.Close(SaveChanges=0)

async def set_table_cell_padding(filename: str, table_index: int, row_index: int, col_index: int,
                                 top: Optional[float] = None, bottom: Optional[float] = None, left: Optional[float] = None, 
                                 right: Optional[float] = None, unit: str = "points") -> str:
    """Set padding/margins for a specific table cell using COM."""
    filename = ensure_docx_extension(filename)
    if not os.path.exists(filename):
        return f"Document {filename} does not exist"

    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return f"Cannot modify document: {error_message}."

    doc = None
    try:
        doc = com_utils.open_document(filename)
        if table_index < 0 or table_index >= doc.Tables.Count:
            return f"Invalid table index. Document has {doc.Tables.Count} tables."

        table = doc.Tables(table_index + 1)

        if row_index < 0 or row_index >= table.Rows.Count or \
           col_index < 0 or col_index >= table.Columns.Count:
            return "Invalid row or column index."

        cell = table.Cell(row_index + 1, col_index + 1)
        
        # Word uses points for padding.
        if unit.lower() != "points":
            return "COM implementation currently only supports 'points' for padding."

        if top is not None:
            cell.TopPadding = top
        if bottom is not None:
            cell.BottomPadding = bottom
        if left is not None:
            cell.LeftPadding = left
        if right is not None:
            cell.RightPadding = right
            
        doc.Save()
        return f"Cell ({row_index}, {col_index}) padding set successfully."
    except Exception as e:
        return f"Failed to set cell padding: {str(e)}"
    finally:
        if doc:
            doc.Close(SaveChanges=0)