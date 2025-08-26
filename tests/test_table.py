import pytest
import json
from word_document_server.tools.table import (
    create_table, get_text_from_cell, set_cell_value
)
from word_document_server.errors import WordDocumentError

@pytest.fixture
def table_test_setup():
    from mcp.server.fastmcp.server import Context
    from word_document_server.tools.document import open_document, close_document, shutdown_word
    ctx = Context()
    
    # 清理测试环境
    try:
        close_document(ctx)
    except:
        pass
    
    # 打开测试文档
    test_file = "tests/test_docs/table_test_doc.docx"
    open_document(ctx, file_path=test_file)
    
    yield ctx
    
    # 测试后清理
    try:
        close_document(ctx)
    except:
        pass
    shutdown_word(ctx)


def test_create_table(table_test_setup):
    ctx = table_test_setup
    # 在文档末尾创建3x3表格
    result = create_table(ctx, rows=3, cols=3)
    assert "successfully" in result.lower()


def test_set_and_get_cell_value(table_test_setup):
    ctx = table_test_setup
    # 创建表格
    create_table(ctx, rows=2, cols=2)
    
    # 设置单元格值
    set_result = set_cell_value(ctx, row=0, col=0, text="Test Cell Value")
    assert "successfully" in set_result.lower()
    
    # 获取单元格值
    cell_value = get_text_from_cell(ctx, row=0, col=0)
    assert cell_value == "Test Cell Value"


def test_invalid_cell_operations(table_test_setup):
    ctx = table_test_setup
    create_table(ctx, rows=2, cols=2)
    
    # 测试无效行索引
    with pytest.raises(WordDocumentError):
        set_cell_value(ctx, row=10, col=0, text="Invalid")
    
    # 测试无效列索引
    with pytest.raises(WordDocumentError):
        get_text_from_cell(ctx, row=0, col=10)