# tests/test_text.py
import pytest
import os
import sys
import json
from unittest.mock import patch, MagicMock

# Add the project root to the Python path
current_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(current_dir)
sys.path.insert(0, project_root)

from word_document_server.tools.text import insert_paragraph, replace_text, batch_apply_format, apply_format, find_text
from word_document_server.tools.document import open_document
from word_document_server.errors import WordDocumentError

# Test fixture
import pytest

@pytest.fixture
def text_test_setup():
    # Create a mock context
    from word_document_server.core_utils import MockContext, MockSession
    ctx = MockContext()
    
    # Get the absolute path to the test document
    test_doc_path = os.path.join(current_dir, 'test_docs', 'text_test_doc.docx')
    
    # Open the test document
    open_document(ctx, test_doc_path)
    
    yield ctx
    
    # Clean up would go here if needed

def test_insert_paragraph(text_test_setup):
    ctx = text_test_setup
    # Use document start as locator
    locator = {"target": {"type": "document_start"}}
    result = insert_paragraph(ctx, locator=locator, text="Test paragraph", position="after")
    assert "successfully" in result.lower()

def test_replace_text(text_test_setup):
    ctx = text_test_setup
    # Insert some text first
    locator_start = {"target": {"type": "document_start"}}
    insert_paragraph(ctx, locator=locator_start, text="Text to replace", position="after")
    
    # Replace the text
    locator = {"target": {"type": "text", "value": "Text to replace"}}
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
    assert "batch formatting completed" in result.lower()

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
    # 插入一些测试文本
    locator_start = {"target": {"type": "document_start"}}
    insert_paragraph(ctx, locator=locator_start, text="This is a sample text for finding", position="after")
    
    # 查找文本
    result = find_text(ctx, find_text="sample")
    # 验证返回的是有效的JSON
    try:
        json_result = json.loads(result)
        assert isinstance(json_result, dict)
        assert "matches" in json_result
        assert isinstance(json_result["matches"], list)
        assert len(json_result["matches"]) > 0
    except json.JSONDecodeError:
        pytest.fail("find_text should return valid JSON")

def test_text_operation_errors(text_test_setup):
    ctx = text_test_setup
    # 关闭文档后尝试文本操作
    from word_document_server.tools.document import close_document
    close_document(ctx)

    # 使用文档开头作为定位器
    locator = {"target": {"type": "document_start"}}
    result = insert_paragraph(ctx, locator=locator, text="This should fail", position="after")
    # 验证返回了错误消息
    assert "error" in result.lower()