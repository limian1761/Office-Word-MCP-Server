import pytest
import json
from word_document_server.tools.text import (
    insert_paragraph, replace_text, apply_format, find_text
)
from word_document_server.errors import WordDocumentError

@pytest.fixture
def text_test_setup():
    from mcp.server.fastmcp.server import Context
    from word_document_server.tools.document import open_document, close_document, shutdown_word
    ctx = Context()
    
    # 清理测试环境
    try:
        close_document(ctx)
    except:
        pass
    
    # 打开测试文档
    test_file = "tests/test_docs/text_test_doc.docx"
    open_document(ctx, file_path=test_file)
    
    yield ctx
    
    # 测试后清理
    try:
        close_document(ctx)
    except:
        pass
    shutdown_word(ctx)


def test_insert_paragraph(text_test_setup):
    ctx = text_test_setup
    # 在文档开头插入段落
    result = insert_paragraph(ctx, text="Test inserted paragraph", position="before")
    assert "successfully" in result.lower()


def test_replace_text(text_test_setup):
    ctx = text_test_setup
    # 先插入可替换文本
    insert_paragraph(ctx, text="Original text for replacement", position="after")
    
    # 替换文本
    result = replace_text(ctx, find_text="Original text", new_text="Replaced text")
    assert "successfully" in result.lower()
    assert "1 replacements made" in result


def test_apply_format(text_test_setup):
    ctx = text_test_setup
    # 插入测试文本
    insert_paragraph(ctx, text="Text to format", position="after")
    
    # 应用粗体和居中对齐
    result = apply_format(ctx, find_text="Text to format", formatting={"bold": True, "alignment": "center"})
    assert "successfully" in result.lower()


def test_find_text(text_test_setup):
    ctx = text_test_setup
    test_text = "Special search text 123"
    insert_paragraph(ctx, text=test_text, position="after")
    
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
    
    with pytest.raises(WordDocumentError):
        insert_paragraph(ctx, text="This should fail")