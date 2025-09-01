"""
Table operations for Word Document MCP Server.
This module contains functions for table-related operations.
"""

import json
import logging
from typing import Any, Dict, List, Optional, Union

import win32com.client

from ..com_backend.com_utils import handle_com_error
from ..selector.selector import SelectorEngine
from ..utils.core_utils import ErrorCode, WordDocumentError, log_error, log_info

logger = logging.getLogger(__name__)


@handle_com_error(ErrorCode.TABLE_ERROR, "create table")
def create_table(
    document: win32com.client.CDispatch,
    rows: int,
    cols: int,
    locator: Optional[Dict[str, Any]] = None,
    position: str = "after",
) -> str:
    """创建新表格

    Args:
        document: Word文档COM对象
        rows: 表格行数
        cols: 表格列数
        locator: 定位器对象，用于指定创建位置
        position: 插入位置，可选值：'before', 'after'

    Returns:
        创建表格成功的消息

    Raises:
        ValueError: 当参数无效时抛出
        WordDocumentError: 当创建表格失败时抛出
    """
    if not document:
        raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")

    selector = SelectorEngine()

    # 验证行数和列数参数
    if rows <= 0 or cols <= 0:
        raise ValueError("Rows and columns must be positive integers")

    range_obj = None

    if locator:
        # 使用定位器获取范围
        try:
            selection = selector.select(document, locator)
            if hasattr(selection, "_com_ranges") and selection._com_ranges:
                # 所有传入的对象都是Range对象，可以直接使用
                range_obj = selection._com_ranges[0]
                
                # 根据位置参数调整范围
                if position == "before":
                    range_obj.Collapse(Direction=1)  # wdCollapseStart
                elif position == "after":
                    range_obj.Collapse(Direction=0)  # wdCollapseEnd
                # 如果是"replace"，则不折叠范围，直接替换
            else:
                raise WordDocumentError(
                    ErrorCode.OBJECT_NOT_FOUND, "No object found matching the locator"
                )
        except Exception as e:
            raise WordDocumentError(
                ErrorCode.TABLE_ERROR, f"Failed to locate position for table: {str(e)}"
            )
    else:
        # 如果没有提供定位器，在文档末尾创建表格
        range_obj = document.Range()
        range_obj.Collapse(Direction=0)  # wdCollapseEnd

    try:
        # 创建表格
        table = document.Tables.Add(Range=range_obj, NumRows=rows, NumColumns=cols)

        # 添加默认样式
        try:
            table.Style = "Table Grid"
        except Exception:
            # 如果默认样式不可用，使用第一个可用的表格样式
            try:
                if document.Styles.Count > 0:
                    table.Style = document.Styles(1).NameLocal
            except Exception:
                pass  # 如果仍然失败，就保持默认样式

        # 添加成功日志
        log_info(
            "Successfully created table with {} rows and {} columns".format(rows, cols)
        )

        return json.dumps(
            {
                "success": True,
                "message": "Table with {} rows and {} columns created successfully".format(
                    rows, cols
                ),
            },
            ensure_ascii=False,
        )

    except Exception as e:
        log_error("Failed to create table: {}".format(str(e)), exc_info=True)
        raise WordDocumentError(
            ErrorCode.TABLE_ERROR, "Failed to create table: {}".format(str(e))
        )


def add_object_caption(
    document: win32com.client.CDispatch,
    range_obj: Any,
    caption_text: str,
    caption_style: str = "Caption",
    position: str = "below",
) -> bool:
    """为元素添加标题

    Args:
        document: Word文档COM对象
        object: 要添加标题的元素
        caption_text: 标题文本
        caption_style: 标题样式
        position: 标题位置 ("above" 或 "below")

    Returns:
        操作是否成功
    """
    try:

        # 确定插入位置
        if position.lower() == "above":
            # 在元素前插入标题
            caption_range = range_obj.Duplicate
            caption_range.Collapse(1)  # wdCollapseStart
        else:
            # 在元素后插入标题
            caption_range = range_obj.Duplicate
            caption_range.Collapse(0)  # wdCollapseEnd

        # 插入标题文本
        caption_range.InsertAfter(caption_text + "\n")

        # 应用样式
        try:
            # 获取新插入的段落（标题）
            caption_paragraph = caption_range.Paragraphs(1)
            caption_paragraph.Style = caption_style
        except Exception:
            # 如果应用样式失败，记录警告但不中断操作
            log_error(f"Failed to apply caption style '{caption_style}'")

        return True
    except Exception as e:
        log_error(f"Failed to add caption to range: {str(e)}")
        return False


@handle_com_error(ErrorCode.TABLE_ERROR, "get cell text")
def get_cell_text(
    document: win32com.client.CDispatch, table_index: int, row: int, col: int
) -> str:
    """获取表格单元格文本

    Args:
        document: Word文档COM对象
        table_index: 表格索引（从1开始）
        row: 行号（从1开始）
        col: 列号（从1开始）

    Returns:
        单元格文本内容

    Raises:
        ValueError: 当参数无效时抛出
        WordDocumentError: 当获取单元格文本失败时抛出
    """
    if not document:
        raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")

    # 验证参数
    if table_index <= 0:
        raise ValueError("Table index must be a positive integer")
    if row <= 0:
        raise ValueError("Row number must be a positive integer")
    if col <= 0:
        raise ValueError("Column number must be a positive integer")

    # 检查表格数量
    table_count = document.Tables.Count
    if table_index > table_count:
        raise WordDocumentError(
            ErrorCode.TABLE_ERROR,
            f"Table index {table_index} out of range. There are {table_count} tables in the document",
        )

    # 获取表格
    table = document.Tables(table_index)

    # 检查行和列的范围
    if row > table.Rows.Count:
        raise WordDocumentError(
            ErrorCode.TABLE_ERROR,
            f"Row {row} out of range. The table has {table.Rows.Count} rows",
        )
    if col > table.Columns.Count:
        raise WordDocumentError(
            ErrorCode.TABLE_ERROR,
            f"Column {col} out of range. The table has {table.Columns.Count} columns",
        )

    # 获取单元格文本
    try:
        cell_text = table.Cell(Row=row, Column=col).Range.Text
        # 移除Word单元格末尾的特殊字符
        if cell_text.endswith("\r\x07"):
            cell_text = cell_text[:-2]
    except Exception as e:
        raise WordDocumentError(
            ErrorCode.TABLE_ERROR, f"Failed to get cell text: {str(e)}"
        )

    log_info(
        f"Successfully retrieved text from table {table_index}, cell ({row},{col})"
    )
    # 确保返回的是字符串类型
    return str(cell_text)


@handle_com_error(ErrorCode.TABLE_ERROR, "set cell text")
def set_cell_text(
    document: win32com.client.CDispatch,
    table_index: int,
    row: int,
    col: int,
    text: str,
    formatting: Optional[Dict[str, Any]] = None,
) -> str:
    """设置表格单元格文本

    Args:
        document: Word文档COM对象
        table_index: 表格索引（从1开始）
        row: 行号（从1开始）
        col: 列号（从1开始）
        text: 要设置的文本内容
        formatting: 可选的格式化参数字典

    Returns:
        设置单元格文本成功的消息

    Raises:
        ValueError: 当参数无效时抛出
        WordDocumentError: 当设置单元格文本失败时抛出
    """
    if not document:
        raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")

    # 验证参数
    if table_index <= 0:
        raise ValueError("Table index must be a positive integer")
    if row <= 0:
        raise ValueError("Row number must be a positive integer")
    if col <= 0:
        raise ValueError("Column number must be a positive integer")
    if text is None:
        raise ValueError("Text parameter cannot be None")

    # 检查表格数量
    table_count = document.Tables.Count
    if table_index > table_count:
        raise WordDocumentError(
            ErrorCode.TABLE_ERROR,
            f"Table index {table_index} out of range. There are {table_count} tables in the document",
        )

    # 获取表格
    table = document.Tables(table_index)

    # 检查行和列的范围
    if row > table.Rows.Count:
        raise WordDocumentError(
            ErrorCode.TABLE_ERROR,
            f"Row {row} out of range. The table has {table.Rows.Count} rows",
        )
    if col > table.Columns.Count:
        raise WordDocumentError(
            ErrorCode.TABLE_ERROR,
            f"Column {col} out of range. The table has {table.Columns.Count} columns",
        )

    # 设置单元格文本
    try:
        cell = table.Cell(Row=row, Column=col)
        cell.Range.Text = text
    except Exception as e:
        raise WordDocumentError(
            ErrorCode.TABLE_ERROR, f"Failed to set cell text: {str(e)}"
        )

    # 应用格式化（如果指定）
    if formatting:
        try:
            # 应用字体格式化
            if "font" in formatting:
                font_format = formatting["font"]
                font = cell.Range.Font

                if "name" in font_format:
                    font.Name = font_format["name"]
                if "size" in font_format:
                    font.Size = font_format["size"]
                if "bold" in font_format:
                    font.Bold = font_format["bold"]
                if "italic" in font_format:
                    font.Italic = font_format["italic"]
                if "color" in font_format:
                    color = font_format["color"]
                    if isinstance(color, str) and color.startswith("#"):
                        cell.Range.Font.Color = color
                    elif isinstance(color, dict) and "rgb" in color:
                        rgb = color["rgb"]
                        cell.Range.Font.Color = f"RGB({rgb[0]},{rgb[1]},{rgb[2]})"

            # 应用段落格式化
            if "paragraph" in formatting:
                para_format = formatting["paragraph"]
                paragraph = cell.Range.Paragraphs(1)

                if "alignment" in para_format:
                    alignment_map = {
                        "left": 0,  # wdAlignParagraphLeft
                        "center": 1,  # wdAlignParagraphCenter
                        "right": 2,  # wdAlignParagraphRight
                        "justify": 3,  # wdAlignParagraphJustify
                    }
                    if para_format["alignment"] in alignment_map:
                        paragraph.Alignment = alignment_map[para_format["alignment"]]
        except Exception as e:
            log_error(f"Failed to apply formatting to cell: {str(e)}")
            # 格式化应用失败不影响文本设置的成功状态

    log_info(f"Successfully set text in table {table_index}, cell ({row},{col})")
    return json.dumps(
        {
            "success": True,
            "message": "Successfully set cell text",
            "table_index": table_index,
            "cell": f"({row},{col})",
        },
        ensure_ascii=False,
    )


@handle_com_error(ErrorCode.TABLE_ERROR, "get table info")
def get_table_info(document: win32com.client.CDispatch, table_index: int) -> str:
    """获取表格信息

    Args:
        document: Word文档COM对象
        table_index: 表格索引（从1开始）

    Returns:
        包含表格信息的JSON字符串

    Raises:
        ValueError: 当参数无效时抛出
        WordDocumentError: 当获取表格信息失败时抛出
    """
    if not document:
        raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")

    # 验证参数
    if table_index <= 0:
        raise ValueError("Table index must be a positive integer")

    # 检查表格数量
    table_count = document.Tables.Count
    if table_index > table_count:
        raise WordDocumentError(
            ErrorCode.TABLE_ERROR,
            f"Table index {table_index} out of range. There are {table_count} tables in the document",
        )

    # 获取表格信息
    try:
        table = document.Tables(table_index)

        # 获取表格基本信息
        info = {
            "table_index": table_index,
            "rows": table.Rows.Count,
            "columns": table.Columns.Count,
            "has_borders": table.Borders.Enable,
   # 检查是否有嵌套表格
        "has_nested_tables": table.Cell(1, 1).Range.Tables.Count > 0
        }

        # 获取表格内容（可选择性地获取，根据需要）
        # 注意：对于大表格，获取所有单元格内容可能会影响性能
        cells_data = []
        for r in range(1, table.Rows.Count + 1):
            row_data = []
            for c in range(1, table.Columns.Count + 1):
                cell_text = table.Cell(Row=r, Column=c).Range.Text
                # 移除Word单元格末尾的特殊字符
                if cell_text.endswith("\r\x07"):
                    cell_text = cell_text[:-2]
                row_data.append(cell_text)
            cells_data.append(row_data)

        info["cells"] = cells_data
    except Exception as e:
        raise WordDocumentError(
            ErrorCode.TABLE_ERROR, f"Failed to get table info: {str(e)}"
        )

    log_info(f"Successfully retrieved info for table {table_index}")
    return json.dumps(info, ensure_ascii=False)


@handle_com_error(ErrorCode.TABLE_ERROR, "insert row")
def insert_row(
    document: win32com.client.CDispatch, table_index: int, position: Union[int, str], count: int = 1
) -> str:
    """在表格中插入行

    Args:
        document: Word文档COM对象
        table_index: 表格索引（从1开始）
        position: 插入位置（行号，从1开始）或位置描述符（"after"表示在末尾插入）
        count: 插入的行数

    Returns:
        插入行成功的消息

    Raises:
        ValueError: 当参数无效时抛出
        WordDocumentError: 当插入行失败时抛出
    """
    if not document:
        raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")

    # 验证参数
    if table_index <= 0:
        raise ValueError("Table index must be a positive integer")
    if count <= 0:
        raise ValueError("Row count must be a positive integer")

    # 检查表格数量
    table_count = document.Tables.Count
    if table_index > table_count:
        raise WordDocumentError(
            ErrorCode.TABLE_ERROR,
            f"Table index {table_index} out of range. There are {table_count} tables in the document",
        )

    # 获取表格
    table = document.Tables(table_index)

    # 处理字符串类型的position参数
    if isinstance(position, str):
        if position.lower() == "after":
            position = table.Rows.Count + 1
        else:
            raise ValueError(f"Invalid position string: {position}. Only 'after' is supported")
    elif not isinstance(position, int) or position <= 0:
        raise ValueError("Insert position must be a positive integer or 'after'")

    # 检查插入位置
    if position > table.Rows.Count + 1:
        raise WordDocumentError(
            ErrorCode.TABLE_ERROR,
            f"Insert position {position} out of range. The table has {table.Rows.Count} rows",
        )

    # 插入行
    try:
        for i in range(count):
            # 在指定位置插入行
            if position <= table.Rows.Count:
                # 插入在指定行之前
                row = table.Rows(position)
                row.Select()
                document.Application.Selection.InsertRowsAbove()
            else:
                # 插入在表格末尾
                table.Rows.Add()
    except Exception as e:
        raise WordDocumentError(
            ErrorCode.TABLE_ERROR, f"Failed to insert row(s): {str(e)}"
        )

    log_info(
        f"Successfully inserted {count} row(s) at position {position} in table {table_index}"
    )
    return json.dumps(
        {
            "success": True,
            "message": f"Successfully inserted {count} row(s)",
            "table_index": table_index,
            "inserted_rows": count,
            "position": position,
        },
        ensure_ascii=False,
    )


@handle_com_error(ErrorCode.TABLE_ERROR, "insert column")
def insert_column(
    document: win32com.client.CDispatch, table_index: int, position: int, count: int = 1
) -> str:
    """在表格中插入列

    Args:
        document: Word文档COM对象
        table_index: 表格索引（从1开始）
        position: 插入位置（列号，从1开始）
        count: 插入的列数

    Returns:
        插入列成功的消息

    Raises:
        ValueError: 当参数无效时抛出
        WordDocumentError: 当插入列失败时抛出
    """
    if not document:
        raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")

    # 验证参数
    if table_index <= 0:
        raise ValueError("Table index must be a positive integer")
    if position <= 0:
        raise ValueError("Insert position must be a positive integer")
    if count <= 0:
        raise ValueError("Column count must be a positive integer")

    # 检查表格数量
    table_count = document.Tables.Count
    if table_index > table_count:
        raise WordDocumentError(
            ErrorCode.TABLE_ERROR,
            f"Table index {table_index} out of range. There are {table_count} tables in the document",
        )

    # 获取表格
    table = document.Tables(table_index)

    # 检查插入位置
    if position > table.Columns.Count + 1:
        raise WordDocumentError(
            ErrorCode.TABLE_ERROR,
            f"Insert position {position} out of range. The table has {table.Columns.Count} columns",
        )

    # 插入列
    try:
        for i in range(count):
            # 在指定位置插入列
            if position <= table.Columns.Count:
                # 插入在指定列之前
                column = table.Columns(position)
                column.Select()
                document.Application.Selection.InsertColumnsLeft()
            else:
                # 插入在表格末尾
                table.Columns.Add()
    except Exception as e:
        raise WordDocumentError(
            ErrorCode.TABLE_ERROR, f"Failed to insert column(s): {str(e)}"
        )

    log_info(
        f"Successfully inserted {count} column(s) at position {position} in table {table_index}"
    )
    return json.dumps(
        {
            "success": True,
            "message": f"Successfully inserted {count} column(s)",
            "table_index": table_index,
            "inserted_columns": count,
            "position": position,
        },
        ensure_ascii=False,
    )
