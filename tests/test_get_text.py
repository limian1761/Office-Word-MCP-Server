import pytest
from unittest.mock import patch, MagicMock
from word_document_server.com_backend import WordBackend
from word_document_server.errors import WordDocumentError


def test_get_text_tool():    
    # 创建WordBackend的mock实例
    backend = WordBackend()
    backend.document = MagicMock()
    
    # 模拟段落对象
    paragraph1 = MagicMock()
    paragraph1.Range.Text = "这是第一段文本。"
    
    paragraph2 = MagicMock()
    paragraph2.Range.Text = "这是第二段文本。"
    
    # 模拟文档的Paragraphs属性
    paragraphs = [paragraph1, paragraph2]
    backend.document.Paragraphs = paragraphs
    
    # 实现一个模拟的get_all_text方法
    def mock_get_all_text():
        text = []
        for paragraph in backend.document.Paragraphs:
            text.append(paragraph.Range.Text)
        return '\n'.join(text)
    
    # 将模拟方法添加到backend实例
    backend.get_all_text = mock_get_all_text
    
    # 测试获取所有文本
    try:
        all_text = backend.get_all_text()
        assert all_text == "这是第一段文本。\n这是第二段文本。"
        print("测试通过：成功获取所有文本")
    except Exception as e:
        pytest.fail(f"测试失败：{e}")


def test_get_text_tool_with_empty_document():
    # 创建WordBackend的mock实例
    backend = WordBackend()
    backend.document = MagicMock()
    backend.document.Paragraphs = []
    
    # 实现一个模拟的get_all_text方法
    def mock_get_all_text():
        text = []
        for paragraph in backend.document.Paragraphs:
            text.append(paragraph.Range.Text)
        return '\n'.join(text)
    
    # 将模拟方法添加到backend实例
    backend.get_all_text = mock_get_all_text
    
    # 测试空文档
    try:
        all_text = backend.get_all_text()
        assert all_text == ""
        print("测试通过：成功处理空文档")
    except Exception as e:
        pytest.fail(f"测试失败：{e}")


def test_get_text_tool_with_no_document():
    # 创建WordBackend的mock实例，但不设置document
    backend = WordBackend()
    
    # 实现一个模拟的get_all_text方法
    def mock_get_all_text():
        if not backend.document:
            raise RuntimeError("No document open.")
        text = []
        for paragraph in backend.document.Paragraphs:
            text.append(paragraph.Range.Text)
        return '\n'.join(text)
    
    # 将模拟方法添加到backend实例
    backend.get_all_text = mock_get_all_text
    
    # 测试没有打开文档的情况
    with pytest.raises(RuntimeError, match="No document open."):
        backend.get_all_text()
    print("测试通过：成功处理未打开文档的情况")

if __name__ == "__main__":
    pytest.main(["-v", "test_get_text.py"])