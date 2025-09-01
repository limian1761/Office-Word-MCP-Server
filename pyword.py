#  测试使用com接口操作word文档
import win32com.client
import os

# 确保目标目录存在
output_dir = os.path.join(os.getcwd(), 'test_output')
os.makedirs(output_dir, exist_ok=True)

# 定义有效的文件路径
file_path = os.path.join(output_dir, 'test.docx')
print(f"将创建文件: {file_path}")

word = win32com.client.Dispatch('Word.Application')
word.Visible = 1

# 新建一个word文件
doc = word.Documents.Add()

# 插入一段文字
doc.Range().InsertAfter('hello world')

# 保存并关闭文档
doc.SaveAs(file_path)
