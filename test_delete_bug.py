import sys
import os
import pythoncom
from win32com.client import DispatchEx

# 添加项目根目录到Python路径
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from word_docx_tools.tools import range_tools
from word_docx_tools.com_backend.word_backend import WordBackend


def test_delete_paragraph():
    """测试删除段落功能，重点检查使用value=2定位器的行为"""
    try:
        # 初始化COM
        pythoncom.CoInitialize()
        
        # 创建Word应用程序对象
        word_app = DispatchEx('Word.Application')
        word_app.Visible = True  # 设为True以便观察
        
        # 创建新文档
        doc = word_app.Documents.Add()
        
        # 添加3个段落用于测试
        paragraphs = doc.Paragraphs
        for i in range(1, 4):
            if i > 1:
                paragraphs.Add()
            paragraphs(i).Range.Text = f"这是第{i}段文本。"
        
        # 显示初始段落数量
        print(f"初始段落数量: {doc.Paragraphs.Count}")
        for i in range(1, doc.Paragraphs.Count + 1):
            text = doc.Paragraphs(i).Range.Text.strip()
            print(f"段落{i}: '{text}'")
        
        # 创建WordBackend实例并设置文档
        backend = WordBackend()
        backend._document = doc
        
        # 测试使用value=2的定位器删除段落
        print("\n测试删除定位器为{type: 'paragraph', value: '2'}的段落...")
        result = range_tools.delete(backend, {
            "operation_type": "delete",
            "locator": {
                "type": "paragraph",
                "value": "2"
            }
        })
        
        # 显示删除结果
        print(f"删除操作结果: {result}")
        
        # 显示剩余段落数量和内容
        print(f"\n删除后的段落数量: {doc.Paragraphs.Count}")
        for i in range(1, doc.Paragraphs.Count + 1):
            text = doc.Paragraphs(i).Range.Text.strip()
            print(f"剩余段落{i}: '{text}'")
        
        # 清理：关闭文档不保存
        doc.Close(SaveChanges=0)
        word_app.Quit()
        
    except Exception as e:
        print(f"测试过程中出错: {str(e)}")
    finally:
        # 释放COM资源
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    test_delete_paragraph()