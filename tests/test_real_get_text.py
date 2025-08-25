import os
import pytest
from word_document_server.com_backend import WordBackend
from word_document_server.errors import WordDocumentError


@pytest.fixture
def test_document_path():    
    # 返回测试文档的路径
    return os.path.join(os.path.dirname(__file__), 'test_docs', 'valid_test_document_v2.docx')


def test_real_get_all_text(test_document_path):
    # 使用上下文管理器打开Word应用和文档
    with WordBackend(file_path=test_document_path, visible=False) as backend:
        try:
            # 调用get_all_text方法获取所有文本
            all_text = backend.get_all_text()
            
            # 验证文本不为空
            assert len(all_text) > 0, "文档文本为空"
            
            # 打印获取到的文本（可选）
            print("成功获取文档文本:")
            print(all_text)
            
        except WordDocumentError as e:
            pytest.fail(f"测试失败: {e}")


def test_real_get_all_text_with_no_document():
    # 创建WordBackend实例，但不打开文档
    backend = WordBackend(visible=False)
    
    try:
        # 尝试调用get_all_text方法
        backend.get_all_text()
        pytest.fail("应该引发RuntimeError异常")
    except RuntimeError as e:
        assert str(e) == "No document open.", f"期望的异常消息不匹配: {e}"
    finally:
        # 确保Word应用被关闭
        if backend.word_app:
            backend.word_app.Quit()


if __name__ == "__main__":
    pytest.main(["-v", "test_real_get_text.py"])