import os
import traceback
from word_document_server.com_backend import WordBackend

# 测试文档路径
TEST_DOC_PATH = os.path.join(os.path.dirname(__file__), 'tests', 'test_docs', 'test_document.docx')

# 创建一个新文档进行测试
def test_get_text_from_range():
    try:
        # 创建新文档
        with WordBackend(visible=True) as backend:
            print("成功创建测试文档")

            # 在文档中插入一些文本
            doc_range = backend.document.Range(0, 0)
            doc_range.InsertAfter("""这是测试文档的第一部分。
这是测试文档的第二部分。
这是测试文档的第三部分。""")
            print("已插入测试文本")

            # 保存文档
            backend.document.SaveAs2(TEST_DOC_PATH)
            print(f"文档已保存到: {TEST_DOC_PATH}")

            # 测试get_text_from_range方法
            # 获取前10个字符
            text1 = backend.get_text_from_range(0, 10)
            print(f"范围 [0, 10] 的文本: {text1}")

            # 获取中间部分
            text2 = backend.get_text_from_range(11, 22)
            print(f"范围 [11, 22] 的文本: {text2}")

            # 获取整个文档
            doc_end = backend.document.Content.End
            text3 = backend.get_text_from_range(0, doc_end)
            print(f"整个文档的文本: {text3}")

        print("测试完成")
    except Exception as e:
        print(f"测试失败: {e}")
        print(f"错误详细信息: {traceback.format_exc()}")

if __name__ == "__main__":
    test_get_text_from_range()