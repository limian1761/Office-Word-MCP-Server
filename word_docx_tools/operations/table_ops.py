"""
Table operations for Word Document MCP Server.
This module contains functions for table-related operations.
"""

import json
import logging
from typing import Any, Dict, List, Optional, Union

import win32com.client

from ..com_backend.com_utils import handle_com_error, iter_com_collection
from ..com_backend.selector_utils import get_selection_range
from ..mcp_service.core_utils import (ErrorCode, WordDocumentError, log_error,
                                      log_info, AppContext)
from ..models.context import DocumentContext


logger = logging.getLogger(__name__)

def _update_document_context_for_table(table: Any, operation: str = "modify") -> None:
    """
    更新表格对应的DocumentContext
    
    Args:
        table: 表格COM对象
        operation: 操作类型（"modify", "create", "delete"等）
    """
    try:
        app_context = AppContext.get_instance()
        document = table.Document
        
        # 查找表格对应的DocumentContext
        # 基于表格的Range.Start和Range.End查找对应的上下文
        context = app_context.find_context_by_range(
            document=document,
            start=table.Range.Start,
            end=table.Range.End,
            object_type="table"
        )
        
        if context:
            # 更新上下文信息
            if operation == "delete":
                app_context.remove_context_from_tree(context)
            else:
                app_context.update_table_context(context, table)
                # 通知上下文更新处理器
                app_context.notify_context_update(context, operation)
    except Exception as e:
        log_error(f"Failed to update DocumentContext for table operation {operation}: {str(e)}")


@handle_com_error(ErrorCode.TABLE_ERROR, "create table")
def create_table(
    document: win32com.client.CDispatch,
    rows: int,
    cols: int,
    locator: Optional[Dict[str, Any]] = None,
    position: str = "replace",
    is_independent_paragraph: bool = True,
) -> str:
    """创建新表格

    Args:
        document: Word文档COM对象
        rows: 表格行数
        cols: 表格列数
        locator: 定位器，用于指定表格插入位置
        position: 插入位置相对于定位点的位置，可选值："replace"、"before"、"after"
        is_independent_paragraph: 是否作为独立段落插入

    Returns:
        包含表格信息的JSON字符串

    Raises:
        ValueError: 当参数无效时抛出
        WordDocumentError: 当创建表格失败时抛出
    """
    if not document: raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")
    if not hasattr(document, "Tables") or document.Tables is None: raise WordDocumentError(
        ErrorCode.DOCUMENT_ERROR, "Document does not support tables"
    )

    # 验证参数
    if rows <= 0: raise ValueError("Row count must be a positive integer")
    if cols <= 0: raise ValueError("Column count must be a positive integer")
    if position not in ["replace", "before", "after"]: raise ValueError(
        "Position must be one of: 'replace', 'before', 'after'"
    )

    # 处理定位器，获取插入位置的Range对象
    range_obj = get_selection_range(document, locator)

    # 处理位置参数
    if position == "before":
        # 在定位点之前插入
        temp_range = range_obj.Duplicate
        temp_range.Collapse(Direction=1)  # wdCollapseStart
        range_obj = temp_range
    elif position == "after":
        # 在定位点之后插入
        temp_range = range_obj.Duplicate
        temp_range.Collapse(Direction=0)  # wdCollapseEnd
        range_obj = temp_range
    # 对于"replace"，直接使用定位点的Range

    # 如果需要作为独立段落插入，确保在段落末尾插入
    if is_independent_paragraph:
        if position == "after":
            # 如果是在定位点之后插入，先移动到段落末尾
            range_obj.MoveEnd(Unit=12, Count=1)  # wdParagraph
        elif position == "before" or position == "replace":
            # 如果是在定位点之前或替换定位点，先移动到段落开头
            range_obj.Collapse(Direction=1)  # wdCollapseStart
            range_obj.MoveStart(Unit=12, Count=-1)  # wdParagraph
            range_obj.Collapse(Direction=0)  # wdCollapseEnd

    try:
        # 创建表格
        table = document.Tables.Add(Range=range_obj, NumRows=rows, NumColumns=cols)

        # 应用表格样式
        try:
            # 尝试应用默认的表格样式
            table.set_Style("Table Grid")
        except Exception as e:
            # 如果样式不存在，不抛出错误
            log_error(f"Failed to apply table style: {str(e)}")

        # 设置表格边框（确保所有边框都可见）
        for cell in iter_com_collection(table.Range.Cells):
            for border in iter_com_collection(cell.Borders):
                border.LineStyle = 1  # wdLineStyleSingle
                border.LineWidth = 1  # wdLineWidth025pt
                border.ColorIndex = 0  # wdColorBlack

        log_info(f"Successfully created a table with {rows} rows and {cols} columns")
        # 更新DocumentContext
        try:
            _update_document_context_for_table(table, "create")
        except Exception as e:
            log_error(f"Failed to update context after creating table: {str(e)}")
        
        return json.dumps(
            {
                "success": True,
                "message": "Successfully created table",
                "table_index": table.Index,
                "rows": rows,
                "columns": cols,
            },
            ensure_ascii=False,
        )
    except Exception as e:
        raise WordDocumentError(ErrorCode.TABLE_ERROR, f"Failed to create table: {str(e)}")





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

    if not hasattr(document, "Tables") or document.Tables is None:
        raise WordDocumentError(
            ErrorCode.DOCUMENT_ERROR, "Document does not support tables"
        )

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

    if not hasattr(document, "Tables") or document.Tables is None:
        raise WordDocumentError(
            ErrorCode.DOCUMENT_ERROR, "Document does not support tables"
        )

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

    # 更新DocumentContext
    try:
        _update_document_context_for_table(table, "modify")
    except Exception as e:
        log_error(f"Failed to update context after setting cell text: {str(e)}")
    
    log_info(f"Successfully set text in table {table_index}, cell ({row},{col})" )
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
def get_table_info(
    document: win32com.client.CDispatch, table_index: Optional[int] = None
) -> str:
    """获取表格信息

    Args:
        document: Word文档COM对象
        table_index: 表格索引（从1开始），不提供则返回所有表格信息

    Returns:
        包含表格信息的JSON字符串

    Raises:
        ValueError: 当参数无效时抛出
        WordDocumentError: 当获取表格信息失败时抛出
    """
    if not document:
        raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")

    # 检查表格数量
    table_count = document.Tables.Count
    if table_count == 0:
        return json.dumps({"tables": [], "total_tables": 0}, ensure_ascii=False)

    # 定义一个内部函数来获取单个表格的信息
    def get_single_table_info(table_idx: int) -> Dict[str, Any]:
        table = document.Tables(table_idx)

        # 获取表格基本信息
        info = {
            "table_index": table_idx,
            "rows": table.Rows.Count,
            "columns": table.Columns.Count,
            "has_borders": table.Borders.Enable,
            # 检查是否有嵌套表格
            "has_nested_tables": table.Cell(1, 1).Range.Tables.Count > 0,
        }

        # 获取表格标题（尝试获取表格前后可能的标题段落）
        try:
            # 检查表格前的段落是否可能是标题
            table_range = table.Range
            prev_range = table_range.Duplicate
            prev_range.MoveStart(Unit=12, Count=-1)  # wdParagraph
            prev_text = prev_range.Text.strip()
            if prev_text and len(prev_text) < 200:  # 简单判断，标题通常不会太长
                info["title_candidate"] = prev_text
        except Exception:
            # 获取标题失败不影响主要功能
            pass

        # 获取表格内容（可选择性地获取，根据需要）
        # 注意：对于大表格，获取所有单元格内容可能会影响性能
        cells_data = []
        for r_idx, row in enumerate(iter_com_collection(table.Rows), 1):
            row_data = []
            for c_idx, cell in enumerate(iter_com_collection(row.Cells), 1):
                cell_text = cell.Range.Text
                # 移除Word单元格末尾的特殊字符
                if cell_text.endswith("\r\x07"):
                    cell_text = cell_text[:-2]
                row_data.append(cell_text)
            cells_data.append(row_data)

        info["cells"] = cells_data
        return info

    try:
        # 如果指定了表格索引，只返回该表格的信息
        if table_index is not None:
            if table_index <= 0:
                raise ValueError("Table index must be a positive integer")

            if table_index > table_count:
                raise WordDocumentError(
                    ErrorCode.TABLE_ERROR,
                    f"Table index {table_index} out of range. There are {table_count} tables in the document",
                )

            info = get_single_table_info(table_index)
            log_info(f"Successfully retrieved info for table {table_index}")
            return json.dumps(info, ensure_ascii=False)
        else:
            # 否则返回所有表格的信息
            all_tables_info = []
            for idx, table in enumerate(iter_com_collection(document.Tables), 1):
                try:
                    table_info = get_single_table_info(idx)
                    all_tables_info.append(table_info)
                except Exception as e:
                    # 单个表格获取失败不影响其他表格
                    log_error(f"Failed to get info for table {idx}: {str(e)}")
                    all_tables_info.append({"table_index": idx, "error": str(e)})

            result = {"tables": all_tables_info, "total_tables": table_count}
            log_info(f"Successfully retrieved info for all {table_count} tables")
            return json.dumps(result, ensure_ascii=False)
    except Exception as e:
        raise WordDocumentError(
            ErrorCode.TABLE_ERROR, f"Failed to get table info: {str(e)}"
        )


@handle_com_error(ErrorCode.TABLE_ERROR, "insert row")
def insert_row(
    document: win32com.client.CDispatch,
    table_index: int,
    position: Union[int, str],
    count: int = 1,
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
            raise ValueError(
                f"Invalid position string: {position}. Only 'after' is supported"
            )
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

    # 更新DocumentContext
    try:
        _update_document_context_for_table(table, "modify")
    except Exception as e:
        log_error(f"Failed to update context after inserting rows: {str(e)}")
    
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
    # 如果position远大于表格列数，表示在末尾插入
    # 但仍需要确保position至少为1
    if position > table.Columns.Count + 1:
        # 不抛出错误，而是将position设置为表格列数+1，表示在末尾插入
        actual_position = table.Columns.Count + 1
    else:
        actual_position = position

    # 插入列
    try:
        for i in range(count):
            # 在指定位置插入列
            if actual_position <= table.Columns.Count:
                try:
                    # 方法1：尝试使用Select和InsertColumnsLeft
                    column = table.Columns(actual_position)
                    column.Select()
                    document.Application.Selection.InsertColumnsLeft()
                except Exception as e:
                    # 方法1失败，尝试方法2：使用Columns.Add并指定位置
                    try:
                        # 先保存原始列数，以便验证插入是否成功
                        original_cols = table.Columns.Count
                        # 使用Add方法添加列
                        new_column = table.Columns.Add()
                        # 如果添加成功，将新列移动到指定位置
                        if table.Columns.Count > original_cols:
                            new_column.Select()
                            # 多次执行左移，直到到达指定位置
                            for _ in range(table.Columns.Count - actual_position):
                                document.Application.CommandBars.ExecuteMso(
                                    "TableColumnsToTheLeft"
                                )
                    except Exception as inner_e:
                        # 如果两种方法都失败，尝试在末尾插入
                        table.Columns.Add()
                # 由于在指定列前插入了新列，后续插入位置需要+1
                actual_position += 1
            else:
                # 插入在表格末尾
                table.Columns.Add()
    except Exception as e:
        raise WordDocumentError(
            ErrorCode.TABLE_ERROR, f"Failed to insert column(s): {str(e)}"
        )

    # 更新DocumentContext
    try:
        _update_document_context_for_table(table, "modify")
    except Exception as e:
        log_error(f"Failed to update context after inserting columns: {str(e)}")
    
    log_info(
        f"Successfully inserted {count} column(s) at position {actual_position - count} in table {table_index}"
    )
    return json.dumps(
        {
            "success": True,
            "message": f"Successfully inserted {count} column(s)",
            "table_index": table_index,
            "inserted_columns": count,
            "position": actual_position - count,
        },
        ensure_ascii=False,
    )
