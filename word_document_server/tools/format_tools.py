"""
Formatting tools for Word Document Server.
"""
import os
from typing import List, Optional
from mcp.server.fastmcp.server import Context

from word_document_server.app import app
from word_document_server.utils.app_context import AppContext
from word_document_server.utils.com_utils import handle_com_error

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

@app.tool()
async def format_text(paragraph_index: int, start_pos: int, end_pos: int,
                                 bold: Optional[bool] = None, italic: Optional[bool] = None,
                                 underline: Optional[bool] = None, color: Optional[str] = None, font_size: Optional[int] = None,
                                 font_name: Optional[str] = None, context: Context = None) -> str:
    """Format a specific range of text within a paragraph."""
    doc = None
    try:
        # 从Context获取活动文档
        app_context = context.request_context.lifespan_context.get(AppContext)
        doc = app_context.get_active_document()
        if not doc:
            return "No active document found."

        if paragraph_index < 0 or paragraph_index >= doc.Paragraphs.Count:
            return f"Invalid paragraph index. Document has {doc.Paragraphs.Count} paragraphs."

        para = doc.Paragraphs(paragraph_index + 1)
        para_range = para.Range
        
        if start_pos < 0 or end_pos > len(para_range.Text) or start_pos >= end_pos:
            return "Invalid text range specified."

        # Create a new range for the specified text
        text_range = doc.Range(Start=para_range.Start + start_pos, End=para_range.Start + end_pos)
        
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
        return f"Text from position {start_pos} to {end_pos} in paragraph {paragraph_index} formatted successfully."
    except Exception as e:
        return handle_com_error(e)
    finally:
        # No need to close the active document
        pass

@app.tool()
async def create_custom_style(style_name: str, bold: Optional[bool] = None, italic: Optional[bool] = None,
                                 font_size: Optional[int] = None, font_name: Optional[str] = None,
                                 color: Optional[str] = None, base_style: Optional[str] = None, context: Context = None) -> str:
    """Create a custom style in the document."""
    doc = None
    try:
        # 从Context获取活动文档
        app_context = context.request_context.lifespan_context.get(AppContext)
        doc = app_context.get_active_document()
        if not doc:
            return "No active document found."

        # Check if style already exists
        try:
            existing_style = doc.Styles(style_name)
            if existing_style:
                return f"Style '{style_name}' already exists."
        except Exception:
            # Style doesn't exist, which is what we want
            pass

        # Create new style based on Normal style or specified base style
        base_style_name = base_style if base_style else "Normal"
        try:
            new_style = doc.Styles.Add(Name=style_name, Type=1) # 1 = Paragraph style
            new_style.BaseStyle = base_style_name
            
            # Apply formatting
            font = new_style.Font
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
                    
            doc.Save()
            return f"Custom style '{style_name}' created successfully."
        except Exception as e:
            return handle_com_error(e)
    except Exception as e:
        return handle_com_error(e)
    finally:
        # No need to close the active document
        pass

@app.tool()
async def format_table(table_index: int, has_header_row: Optional[bool] = None,
                                 border_style: Optional[str] = None, shading: Optional[List[List[str]]] = None, context: Context = None) -> str:
    """Format a table with borders, shading, and structure."""
    doc = None
    try:
        # 从Context获取活动文档
        app_context = context.request_context.lifespan_context.get(AppContext)
        doc = app_context.get_active_document()
        if not doc:
            return "No active document found."

        if table_index < 0 or table_index >= doc.Tables.Count:
            return f"Invalid table index. Document has {doc.Tables.Count} tables."

        table = doc.Tables(table_index + 1)
        
        # Apply header row setting
        if has_header_row is not None:
            table.HeaderRowCount = 1 if has_header_row else 0
            
        # Apply border style
        if border_style:
            # Map border styles to Word constants
            border_map = {
                "single": 1,      # wdLineStyleSingle
                "double": 3,      # wdLineStyleDouble
                "dotted": 4,      # wdLineStyleDot
                "dashed": 5,      # wdLineStyleDash
                "thick": 6,       # wdLineStyleThick
            }
            line_style = border_map.get(border_style.lower(), 1)
            
            # Apply to all borders
            for border in [table.Borders(-1), table.Borders(-2), table.Borders(-3), 
                          table.Borders(-4), table.Borders(-5), table.Borders(-6)]:
                border.LineStyle = line_style
                border.Visible = True
                
        # Apply shading if provided
        if shading:
            for row_idx, row_colors in enumerate(shading):
                if row_idx < table.Rows.Count:
                    for col_idx, color in enumerate(row_colors):
                        if col_idx < table.Columns.Count:
                            cell = table.Cell(row_idx + 1, col_idx + 1)
                            try:
                                cell.Shading.BackgroundPatternColor = hex_to_bgr(color)
                            except Exception:
                                return f"Invalid color format: '{color}'. Use RRGGBB hex format."
            
        doc.Save()
        return f"Table {table_index} formatted successfully."
    except Exception as e:
        return handle_com_error(e)
    finally:
        # No need to close the active document
        pass

@app.tool()
async def set_table_cell_shading(table_index: int, row_index: int, col_index: int,
                                 fill_color: str, pattern: str = "clear", context: Context = None) -> str:
    """Apply shading/filling to a specific table cell."""
    doc = None
    try:
        # 从Context获取活动文档
        app_context = context.request_context.lifespan_context.get(AppContext)
        doc = app_context.get_active_document()
        if not doc:
            return "No active document found."

        if table_index < 0 or table_index >= doc.Tables.Count:
            return f"Invalid table index. Document has {doc.Tables.Count} tables."

        table = doc.Tables(table_index + 1)

        if row_index < 0 or row_index >= table.Rows.Count or \
           col_index < 0 or col_index >= table.Columns.Count:
            return "Invalid row or column index."

        cell = table.Cell(row_index + 1, col_index + 1)
        
        # Pattern constants
        # wdTextureNone = -1, wdTextureSolid = 1000
        pattern_map = {
            "clear": -1,
            "solid": 1000,
        }
        
        try:
            cell.Shading.BackgroundPatternColor = hex_to_bgr(fill_color)
            cell.Shading.Texture = pattern_map.get(pattern.lower(), -1)
        except Exception:
            return f"Invalid color format: '{fill_color}'. Use RRGGBB hex format."
            
        doc.Save()
        return f"Cell ({row_index}, {col_index}) shading applied successfully."
    except Exception as e:
        return handle_com_error(e)
    finally:
        # No need to close the active document
        pass

@app.tool()
async def apply_table_alternating_rows(table_index: int, color1: str = "FFFFFF", 
                                 color2: str = "F2F2F2", context: Context = None) -> str:
    """Apply alternating row colors to a table for better readability."""
    doc = None
    try:
        # 从Context获取活动文档
        app_context = context.request_context.lifespan_context.get(AppContext)
        doc = app_context.get_active_document()
        if not doc:
            return "No active document found."

        if table_index < 0 or table_index >= doc.Tables.Count:
            return f"Invalid table index. Document has {doc.Tables.Count} tables."

        table = doc.Tables(table_index + 1)
        
        for i in range(table.Rows.Count):
            row = table.Rows(i + 1)
            color = color1 if i % 2 == 0 else color2
            try:
                row.Shading.BackgroundPatternColor = hex_to_bgr(color)
            except Exception:
                return f"Invalid color format: '{color}'. Use RRGGBB hex format."
                
        doc.Save()
        return f"Alternating row colors applied to table {table_index}."
    except Exception as e:
        return handle_com_error(e)
    finally:
        # No need to close the active document
        pass

@app.tool()
async def highlight_table_header(table_index: int, header_color: str = "4472C4", 
                                 text_color: str = "FFFFFF", context: Context = None) -> str:
    """Apply special highlighting to table header row."""
    doc = None
    try:
        # 从Context获取活动文档
        app_context = context.request_context.lifespan_context.get(AppContext)
        doc = app_context.get_active_document()
        if not doc:
            return "No active document found."

        if table_index < 0 or table_index >= doc.Tables.Count:
            return f"Invalid table index. Document has {doc.Tables.Count} tables."

        table = doc.Tables(table_index + 1)
        
        if table.Rows.Count == 0:
            return "Table has no rows."
            
        header_row = table.Rows(1)
        
        try:
            header_row.Shading.BackgroundPatternColor = hex_to_bgr(header_color)
            header_row.Range.Font.Color = hex_to_bgr(text_color)
        except Exception:
            return f"Invalid color format. Use RRGGBB hex format."
            
        doc.Save()
        return f"Header row of table {table_index} highlighted successfully."
    except Exception as e:
        return handle_com_error(e)
    finally:
        # No need to close the active document
        pass

@app.tool()
async def merge_table_cells(table_index: int, start_row: int, start_col: int,
                                 end_row: int, end_col: int, context: Context = None) -> str:
    """Merge cells in a rectangular area of a table."""
    doc = None
    try:
        # 从Context获取活动文档
        app_context = context.request_context.lifespan_context.get(AppContext)
        doc = app_context.get_active_document()
        if not doc:
            return "No active document found."

        if table_index < 0 or table_index >= doc.Tables.Count:
            return f"Invalid table index. Document has {doc.Tables.Count} tables."

        table = doc.Tables(table_index + 1)

        # Validate indices
        if (start_row < 0 or start_row >= table.Rows.Count or
            end_row < 0 or end_row >= table.Rows.Count or
            start_col < 0 or start_col >= table.Columns.Count or
            end_col < 0 or end_col >= table.Columns.Count):
            return "Invalid cell range specified."
            
        if start_row > end_row or start_col > end_col:
            return "Start position must be before end position."

        # Get the range of cells to merge
        start_cell = table.Cell(start_row + 1, start_col + 1)
        end_cell = table.Cell(end_row + 1, end_col + 1)
        
        # Create a range spanning these cells
        merge_range = doc.Range(Start=start_cell.Range.Start, End=end_cell.Range.End)
        
        # Perform the merge
        merge_range.Cells.Merge()
        
        doc.Save()
        return f"Cells merged successfully in table {table_index}."
    except Exception as e:
        return handle_com_error(e)
    finally:
        # No need to close the active document
        pass

@app.tool()
async def merge_table_cells_horizontal(table_index: int, row_index: int,
                                 start_col: int, end_col: int, context: Context = None) -> str:
    """Merge cells horizontally in a single row."""
    doc = None
    try:
        # 从Context获取活动文档
        app_context = context.request_context.lifespan_context.get(AppContext)
        doc = app_context.get_active_document()
        if not doc:
            return "No active document found."

        if table_index < 0 or table_index >= doc.Tables.Count:
            return f"Invalid table index. Document has {doc.Tables.Count} tables."

        table = doc.Tables(table_index + 1)

        # Validate indices
        if (row_index < 0 or row_index >= table.Rows.Count or
            start_col < 0 or start_col >= table.Columns.Count or
            end_col < 0 or end_col >= table.Columns.Count):
            return "Invalid cell range specified."
            
        if start_col > end_col:
            return "Start column must be before end column."

        # Get the range of cells to merge
        start_cell = table.Cell(row_index + 1, start_col + 1)
        end_cell = table.Cell(row_index + 1, end_col + 1)
        
        # Create a range spanning these cells
        merge_range = doc.Range(Start=start_cell.Range.Start, End=end_cell.Range.End)
        
        # Perform the merge
        merge_range.Cells.Merge()
        
        doc.Save()
        return f"Cells merged horizontally in row {row_index} of table {table_index}."
    except Exception as e:
        return handle_com_error(e)
    finally:
        # No need to close the active document
        pass

@app.tool()
async def merge_table_cells_vertical(table_index: int, col_index: int,
                                 start_row: int, end_row: int, context: Context = None) -> str:
    """Merge cells vertically in a single column."""
    doc = None
    try:
        # 从Context获取活动文档
        app_context = context.request_context.lifespan_context.get(AppContext)
        doc = app_context.get_active_document()
        if not doc:
            return "No active document found."

        if table_index < 0 or table_index >= doc.Tables.Count:
            return f"Invalid table index. Document has {doc.Tables.Count} tables."

        table = doc.Tables(table_index + 1)

        # Validate indices
        if (col_index < 0 or col_index >= table.Columns.Count or
            start_row < 0 or start_row >= table.Rows.Count or
            end_row < 0 or end_row >= table.Rows.Count):
            return "Invalid cell range specified."
            
        if start_row > end_row:
            return "Start row must be before end row."

        # Get the range of cells to merge
        start_cell = table.Cell(start_row + 1, col_index + 1)
        end_cell = table.Cell(end_row + 1, col_index + 1)
        
        # Create a range spanning these cells
        merge_range = doc.Range(Start=start_cell.Range.Start, End=end_cell.Range.End)
        
        # Perform the merge
        merge_range.Cells.Merge()
        
        doc.Save()
        return f"Cells merged vertically in column {col_index} of table {table_index}."
    except Exception as e:
        return handle_com_error(e)
    finally:
        # No need to close the active document
        pass

@app.tool()
async def set_table_cell_alignment(table_index: int, row_index: int, col_index: int,
                                 horizontal: str = "left", vertical: str = "top", context: Context = None) -> str:
    """Set text alignment for a specific table cell."""
    doc = None
    try:
        # 从Context获取活动文档
        app_context = context.request_context.lifespan_context.get(AppContext)
        doc = app_context.get_active_document()
        if not doc:
            return "No active document found."

        if table_index < 0 or table_index >= doc.Tables.Count:
            return f"Invalid table index. Document has {doc.Tables.Count} tables."

        table = doc.Tables(table_index + 1)

        # Validate indices
        if (row_index < 0 or row_index >= table.Rows.Count or
            col_index < 0 or col_index >= table.Columns.Count):
            return "Invalid cell index specified."

        cell = table.Cell(row_index + 1, col_index + 1)
        
        # Map horizontal alignment
        horizontal_map = {
            "left": 0,      # wdCellAlignLeft
            "center": 1,    # wdCellAlignCenter
            "right": 2      # wdCellAlignRight
        }
        
        # Map vertical alignment
        vertical_map = {
            "top": 1,       # wdCellAlignVerticalTop
            "center": 3,    # wdCellAlignVerticalCenter
            "bottom": 5     # wdCellAlignVerticalBottom
        }
        
        if horizontal.lower() in horizontal_map:
            cell.VerticalAlignment = horizontal_map[horizontal.lower()]
        else:
            return f"Invalid horizontal alignment: {horizontal}"
            
        if vertical.lower() in vertical_map:
            cell.VerticalAlignment = vertical_map[vertical.lower()]
        else:
            return f"Invalid vertical alignment: {vertical}"
            
        doc.Save()
        return f"Alignment set for cell ({row_index}, {col_index}) in table {table_index}."
    except Exception as e:
        return handle_com_error(e)
    finally:
        # No need to close the active document
        pass

@app.tool()
async def set_table_alignment_all(table_index: int, horizontal: str = "left",
                                 vertical: str = "top", context: Context = None) -> str:
    """Set text alignment for all cells in a table."""
    doc = None
    try:
        # 从Context获取活动文档
        app_context = context.request_context.lifespan_context.get(AppContext)
        doc = app_context.get_active_document()
        if not doc:
            return "No active document found."

        if table_index < 0 or table_index >= doc.Tables.Count:
            return f"Invalid table index. Document has {doc.Tables.Count} tables."

        table = doc.Tables(table_index + 1)
        
        # Map horizontal alignment
        horizontal_map = {
            "left": 0,      # wdCellAlignLeft
            "center": 1,    # wdCellAlignCenter
            "right": 2      # wdCellAlignRight
        }
        
        # Map vertical alignment
        vertical_map = {
            "top": 1,       # wdCellAlignVerticalTop
            "center": 3,    # wdCellAlignVerticalCenter
            "bottom": 5     # wdCellAlignVerticalBottom
        }
        
        # Apply alignment to all cells
        for row in table.Rows:
            for cell in row.Cells:
                if horizontal.lower() in horizontal_map:
                    cell.VerticalAlignment = horizontal_map[horizontal.lower()]
                if vertical.lower() in vertical_map:
                    cell.VerticalAlignment = vertical_map[vertical.lower()]
                    
        doc.Save()
        return f"Alignment set for all cells in table {table_index}."
    except Exception as e:
        return handle_com_error(e)
    finally:
        # No need to close the active document
        pass

@app.tool()
async def set_table_column_width(table_index: int, col_index: int, width: float,
                                 width_type: str = "points", context: Context = None) -> str:
    """Set the width of a specific table column."""
    doc = None
    try:
        # 从Context获取活动文档
        app_context = context.request_context.lifespan_context.get(AppContext)
        doc = app_context.get_active_document()
        if not doc:
            return "No active document found."

        if table_index < 0 or table_index >= doc.Tables.Count:
            return f"Invalid table index. Document has {doc.Tables.Count} tables."

        table = doc.Tables(table_index + 1)

        # Validate column index
        if col_index < 0 or col_index >= table.Columns.Count:
            return "Invalid column index specified."

        # Map width type to Word constants
        width_type_map = {
            "points": 0,        # wdPreferredWidthPoints
            "pct": 1,           # wdPreferredWidthPercent
            "auto": 2           # wdPreferredWidthAuto
        }
        
        if width_type.lower() not in width_type_map:
            return f"Invalid width type: {width_type}"
            
        # Set the column width
        column = table.Columns(col_index + 1)
        column.Width = width if width_type.lower() == "points" else 0
        column.PreferredWidthType = width_type_map[width_type.lower()]
        if width_type.lower() == "pct":
            column.PreferredWidth = width
            
        doc.Save()
        return f"Width set for column {col_index} in table {table_index}."
    except Exception as e:
        return handle_com_error(e)
    finally:
        # No need to close the active document
        pass

@app.tool()
async def set_table_column_widths(table_index: int, widths: list[float],
                                 width_type: str = "points", context: Context = None) -> str:
    """Set the widths of multiple table columns."""
    doc = None
    try:
        # 从Context获取活动文档
        app_context = context.request_context.lifespan_context.get(AppContext)
        doc = app_context.get_active_document()
        if not doc:
            return "No active document found."

        if table_index < 0 or table_index >= doc.Tables.Count:
            return f"Invalid table index. Document has {doc.Tables.Count} tables."

        table = doc.Tables(table_index + 1)

        # Validate widths list
        if not isinstance(widths, list) or len(widths) == 0:
            return "Widths must be a non-empty list."
            
        if len(widths) > table.Columns.Count:
            return f"Too many widths specified. Table has {table.Columns.Count} columns."

        # Map width type to Word constants
        width_type_map = {
            "points": 0,        # wdPreferredWidthPoints
            "pct": 1,           # wdPreferredWidthPercent
            "auto": 2           # wdPreferredWidthAuto
        }
        
        if width_type.lower() not in width_type_map:
            return f"Invalid width type: {width_type}"

        # Set widths for each column
        for i, width in enumerate(widths):
            column = table.Columns(i + 1)
            column.Width = width if width_type.lower() == "points" else 0
            column.PreferredWidthType = width_type_map[width_type.lower()]
            if width_type.lower() == "pct":
                column.PreferredWidth = width
                
        doc.Save()
        return f"Widths set for {len(widths)} columns in table {table_index}."
    except Exception as e:
        return handle_com_error(e)
    finally:
        # No need to close the active document
        pass

@app.tool()
async def set_table_width(table_index: int, width: float,
                                 width_type: str = "points", context: Context = None) -> str:
    """Set the overall width of a table."""
    doc = None
    try:
        # 从Context获取活动文档
        app_context = context.request_context.lifespan_context.get(AppContext)
        doc = app_context.get_active_document()
        if not doc:
            return "No active document found."

        if table_index < 0 or table_index >= doc.Tables.Count:
            return f"Invalid table index. Document has {doc.Tables.Count} tables."

        table = doc.Tables(table_index + 1)
        
        # Map width type to Word constants
        width_type_map = {
            "points": 0,        # wdPreferredWidthPoints
            "pct": 1,           # wdPreferredWidthPercent
            "auto": 2           # wdPreferredWidthAuto
        }
        
        if width_type.lower() not in width_type_map:
            return f"Invalid width type: {width_type}"
            
        # Set the table width
        table.PreferredWidthType = width_type_map[width_type.lower()]
        if width_type.lower() == "points":
            table.PreferredWidth = width
        elif width_type.lower() == "pct":
            table.PreferredWidth = width
            
        doc.Save()
        return f"Width set for table {table_index}."
    except Exception as e:
        return handle_com_error(e)
    finally:
        # No need to close the active document
        pass

@app.tool()
async def auto_fit_table_columns(table_index: int, context: Context = None) -> str:
    """Set table columns to auto-fit based on content."""
    doc = None
    try:
        # 从Context获取活动文档
        app_context = context.request_context.lifespan_context.get(AppContext)
        doc = app_context.get_active_document()
        if not doc:
            return "No active document found."

        if table_index < 0 or table_index >= doc.Tables.Count:
            return f"Invalid table index. Document has {doc.Tables.Count} tables."

        table = doc.Tables(table_index + 1)
        
        # Auto fit columns based on content
        table.AutoFitBehavior(1)  # wdAutoFitContent
        table.Columns.AutoFit()
        
        doc.Save()
        return f"Columns auto-fitted for table {table_index}."
    except Exception as e:
        return handle_com_error(e)
    finally:
        # No need to close the active document
        pass

@app.tool()
async def format_table_cell_text(table_index: int, row_index: int, col_index: int,
                                 text_content: Optional[str] = None, bold: Optional[bool] = None, italic: Optional[bool] = None,
                                 underline: Optional[bool] = None, color: Optional[str] = None, font_size: Optional[int] = None,
                                 font_name: Optional[str] = None, context: Context = None) -> str:
    """Format text within a specific table cell."""
    doc = None
    try:
        # 从Context获取活动文档
        app_context = context.request_context.lifespan_context.get(AppContext)
        doc = app_context.get_active_document()
        if not doc:
            return "No active document found."

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
        return handle_com_error(e)
    finally:
        # No need to close the active document
        pass

@app.tool()
async def set_table_cell_padding(table_index: int, row_index: int, col_index: int,
                                 top: Optional[float] = None, bottom: Optional[float] = None, left: Optional[float] = None, 
                                 right: Optional[float] = None, unit: str = "points", context: Context = None) -> str:
    """Set padding/margins for a specific table cell."""
    doc = None
    try:
        # 从Context获取活动文档
        app_context = context.request_context.lifespan_context.get(AppContext)
        doc = app_context.get_active_document()
        if not doc:
            return "No active document found."

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
        return handle_com_error(e)
    finally:
        # No need to close the active document
        pass
        if doc:
            doc.Close(SaveChanges=0)
