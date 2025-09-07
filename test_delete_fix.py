import sys
import os
import pythoncom
from win32com.client import DispatchEx

# 添加项目根目录到Python路径
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from word_docx_tools.operations.range_ops import delete_object_by_locator


def test_delete_paragraph():
    """测试删除段落功能，验证修复是否有效"""
    try:
        # 初始化COM
        pythoncom.CoInitialize()
        
        # 创建Word应用程序对象
        word_app = DispatchEx('Word.Application')
        word_app.Visible = True  # 设为True以便观察
        
        # 创建新文档
        doc = word_app.Documents.Add()
        
        # 添加3个段落用于测试
        for i in range(1, 4):
            if i > 1:
                doc.Paragraphs.Add()
            doc.Paragraphs(i).Range.Text = f"这是第{i}段文本。"
        
        # 显示初始段落数量
        print(f"初始段落数量: {doc.Paragraphs.Count}")
        for i in range(1, doc.Paragraphs.Count + 1):
            text = doc.Paragraphs(i).Range.Text.strip()
            print(f"段落{i}: '{text}'")
        
        # 测试使用value=2的定位器删除段落
        print("\n测试删除定位器为{type: 'paragraph', value: '2'}的段落...")
        result = delete_object_by_locator(doc, {
            "type": "paragraph",
            "value": "2"
        })
        
        # 显示删除结果
        print(f"删除操作结果: {'成功' if result else '失败'}")
        
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