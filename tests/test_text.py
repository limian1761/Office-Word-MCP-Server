import pytest
import json
import os
import sys
import win32com.client

# Add the project root to the Python path
current_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(current_dir)
sys.path.insert(0, project_root)

from word_document_server.tools.text import (
    insert_paragraph, replace_text, apply_format, find_text, batch_apply_format
)
from word_document_server.errors import WordDocumentError
from word_document_server.core_utils import MockContext
from word_document_server.tools.document import open_document, close_document, shutdown_word

@pytest.fixture
def text_test_setup():
    # Mock the Context object
    ctx = MockContext()
    
    # Start Word application for testing
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    
    # Get the absolute path to the test document
    test_doc_path = os.path.join(current_dir, 'test_docs', 'text_test_doc.docx')
    
    # Ensure test document exists
    if not os.path.exists(test_doc_path):
        # Create a simple test document for testing
        doc = word.Documents.Add()
        doc.SaveAs(test_doc_path)
        doc.Close()
    
    # Open the test document
    open_document(ctx, test_doc_path)
    
    yield ctx
    
    # Test cleanup
    try:
        close_document(ctx)
    except:
        pass
    shutdown_word(ctx)
    
    # Close Word application
    word.Quit()


def test_insert_paragraph(text_test_setup):
    ctx = text_test_setup
    # 在文档开头插入段落
    # 使用文档开头作为定位器
    locator = {"target": {"type": "document_start"}}
    result = insert_paragraph(ctx, locator=locator, text="Test inserted paragraph", position="before")
    assert "successfully" in result.lower()


def test_replace_text(text_test_setup):
    ctx = text_test_setup
    # 先插入可替换文本
    locator = {"target": {"type": "document_start"}}
    insert_paragraph(ctx, locator=locator, text="Original text for replacement", position="after")
    
    # 替换文本
    # 使用包含要替换文本的定位器
    locator = {"target": {"type": "text", "value": "Original text"}}
    result = replace_text(ctx, locator=locator, new_text="Replaced text")
    assert "successfully" in result.lower()


def test_batch_apply_format(text_test_setup):
    ctx = text_test_setup
    # 插入多个测试文本段落
    locator_start = {"target": {"type": "document_start"}}
    insert_paragraph(ctx, locator=locator_start, text="Batch test paragraph 1", position="after")
    insert_paragraph(ctx, locator=locator_start, text="Batch test paragraph 2", position="after")
    
    # 准备批量格式化操作
    operations = [
        {
            "locator": {"target": {"type": "text", "value": "Batch test paragraph 1"}},
            "formatting": {"bold": True, "font_size": 14}
        },
        {
            "locator": {"target": {"type": "text", "value": "Batch test paragraph 2"}},
            "formatting": {"italic": True, "font_color": "#FF5733"}
        }
    ]
    
    # 执行批量格式化
    result = batch_apply_format(ctx, operations=operations)
    assert "successfully processed 2 operations" in result.lower()


def test_apply_format(text_test_setup):
    ctx = text_test_setup
    # 插入测试文本
    locator = {"target": {"type": "document_start"}}
    insert_paragraph(ctx, locator=locator, text="Text to format", position="after")
    
    # 应用多种格式化选项
    locator = {"target": {"type": "text", "value": "Text to format"}}
    result = apply_format(ctx, locator=locator, formatting={
        "bold": True, 
        "italic": True, 
        "font_size": 16, 
        "font_color": "#2E86C1",
        "alignment": "center"
    })
    assert "successfully" in result.lower()


def test_find_text(text_test_setup):
    ctx = text_test_setup
    test_text = "Special search text 123"
    locator = {"target": {"type": "document_start"}}
    insert_paragraph(ctx, locator=locator, text=test_text, position="after")
    
    # 测试精确查找
    result = find_text(ctx, find_text=test_text, match_case=True)
    result_data = json.loads(result)
    assert len(result_data) > 0
    assert test_text in result_data[0]["context"]


def test_text_operation_errors(text_test_setup):
    ctx = text_test_setup
    # 关闭文档后尝试文本操作
    from word_document_server.tools.document import close_document
    close_document(ctx)
    
    # 使用文档开头作为定位器
    locator = {"target": {"type": "document_start"}}
    with pytest.raises(WordDocumentError):
        insert_paragraph(ctx, locator=locator, text="This should fail")