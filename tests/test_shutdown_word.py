import os
import time
import pythoncom
import win32com.client
from word_document_server.com_backend import WordBackend

# 测试shutdown方法
print("开始测试shutdown方法...")

try:
    # 确保没有Word实例在运行
    try:
        word_app = win32com.client.GetActiveObject("Word.Application")
        word_app.Quit()
        print("已关闭现有Word实例")
        time.sleep(2)
    except Exception as e:
        print("没有正在运行的Word实例")
    
    # 创建一个临时文档路径
    test_docs_dir = os.path.join(os.getcwd(), "tests", "test_docs")
    # 确保目标目录存在
    os.makedirs(test_docs_dir, exist_ok=True)
    test_doc_path = os.path.join(test_docs_dir, "test_shutdown.docx")
    
    # 使用上下文管理器打开Word并创建文档
    with WordBackend(visible=True) as backend:
        print("Word应用已启动，文档已创建")
        
        # 添加一些文本到文档
        para = backend.document.Paragraphs.Add()
        para.Range.Text = "这是一个测试文档，用于验证shutdown方法。"
        print("已添加文本到文档")
        
        # 保存文档
        backend.document.SaveAs(test_doc_path)
        print(f"文档已保存到: {test_doc_path}")
        
    # 在上下文管理器退出后，文档已关闭，但Word应用仍在运行
    print("准备调用shutdown方法...")
    
    # 直接获取正在运行的Word实例
    try:
        word_app = win32com.client.GetActiveObject("Word.Application")
        print("已获取到运行中的Word实例")
        
        # 创建backend实例并关联到这个Word实例
        backend = WordBackend()
        backend.word_app = word_app
        
        # 调用shutdown方法
        backend.shutdown()
        print("shutdown方法已调用，Word应用应已关闭")
    except Exception as e:
        print(f"无法获取Word实例: {e}")
    
    # 等待一会儿让Word有时间关闭
    time.sleep(3)
    
    # 验证Word应用是否已关闭
    try:
        # 尝试获取Word实例，如果失败则说明已关闭
        win32com.client.GetActiveObject("Word.Application")
        print("警告: Word应用似乎仍在运行")
    except Exception as e:
        print("成功: Word应用已关闭")
        
    print("测试完成")

except Exception as e:
    print(f"测试失败: {e}")