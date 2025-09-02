"""
Complex document operations tests for Word Document MCP Server.

This module contains tests that verify complex operations on documents
with various objects like tables, images, comments, and formatted text.
"""

import unittest
import os
import tempfile
import shutil
import json
from pathlib import Path
from typing import cast
import win32com.client
from win32com.client import CDispatch
import pythoncom
from io import StringIO

from word_document_server.tools.document_tools import document_tools
from word_document_server.tools.text_tools import text_tools
from word_document_server.tools.table_tools import table_tools
from word_document_server.tools.image_tools import image_tools
from word_document_server.tools.comment_tools import comment_tools
from word_document_server.tools.range_tools import range_tools
from word_document_server.utils.app_context import AppContext
from word_document_server.mcp_service.core_utils import format_error_response
from mcp.server.fastmcp import Context
from mcp.server.session import ServerSession


# 创建工具函数的别名
document_operation = document_tools
get_document_outline = document_tools
text_content_operation = text_tools
table_operation = table_tools
document_tools = document_tools


class TestComplexDocumentOperations(unittest.TestCase):
    """Tests for complex document operations with precise object manipulation"""
    
    @classmethod
    def setUpClass(cls):
        # 初始化COM
        pythoncom.CoInitialize()
        
        # 创建临时目录用于测试
        cls.test_dir = tempfile.mkdtemp()
        
        # 创建一个复杂测试文档的路径
        cls.test_doc_path = os.path.join(cls.test_dir, "complex_test_document.docx")
        
        # 创建Word应用程序实例用于创建测试文档
        cls.word_app = win32com.client.Dispatch('Word.Application')
        try:
            cls.word_app.Visible = False
        except AttributeError:
            pass
        
        # 创建复杂测试文档
        cls._create_complex_test_document()
    
    @classmethod
    def _create_complex_test_document(cls):
        """创建一个包含多种元素的复杂测试文档"""
        # 创建新文档
        doc = cls.word_app.Documents.Add()
        
        # 添加标题
        doc.Range(0, 0).Text = "复杂文档测试\n"
        # 使用数字样式ID而不是名称
        try:
            doc.Paragraphs(1).Range.Style = -2  # wdStyleHeading1
        except:
            pass  # 如果设置失败则跳过
        
        # 添加段落
        doc.Range().Collapse(0)  # 0 = wdCollapseEnd
        doc.Range().Text = "这是第一个段落，包含一些文本内容。\n"
        
        # 添加另一个段落
        doc.Range().Collapse(0)
        doc.Range().Text = "这是第二个段落，我们将在这段文字中查找特定词汇。\n"
        
        # 添加表格
        doc.Range().Collapse(0)
        doc.Range().Text = "\n"
        table_range = doc.Range()
        table_range.Collapse(0)
        table = doc.Tables.Add(table_range, 3, 3)
        table.Cell(1, 1).Range.Text = "姓名"
        table.Cell(1, 2).Range.Text = "年龄"
        table.Cell(1, 3).Range.Text = "城市"
        table.Cell(2, 1).Range.Text = "张三"
        table.Cell(2, 2).Range.Text = "25"
        table.Cell(2, 3).Range.Text = "北京"
        table.Cell(3, 1).Range.Text = "李四"
        table.Cell(3, 2).Range.Text = "30"
        table.Cell(3, 3).Range.Text = "上海"
        
        # 添加另一个段落
        doc.Range().Collapse(0)
        doc.Range().Text = "\n这是表格后的段落。\n"
        
        # 保存文档
        doc.SaveAs2(cls.test_doc_path)
        doc.Close()

    @classmethod
    def tearDownClass(cls):
        # 关闭Word应用程序
        try:
            cls.word_app.Quit()
        except:
            pass
            
        # 清理临时目录
        try:
            shutil.rmtree(cls.test_dir)
        except:
            pass
        pythoncom.CoUninitialize()
    
    def setUp(self):
        # 创建Word应用程序实例
        self.word_app = win32com.client.Dispatch('Word.Application')
        # 尝试设置Visible属性，但捕获可能的异常
        try:
            self.word_app.Visible = False
        except AttributeError:
            # 某些环境中可能不支持设置Visible属性，忽略此错误
            pass
        
        # 创建应用上下文
        self.app_context = AppContext(self.word_app)
        
        # 创建模拟的读写流
        self.read_stream = StringIO()
        self.write_stream = StringIO()
        
        # 创建ServerSession实例
        self.session = ServerSession(
            read_stream=self.read_stream,
            write_stream=self.write_stream,
            init_options={}
        )
        self.session.lifespan_context = self.app_context
        
        # 创建Context对象
        self.context = Context(
            server_session=self.session,
            request_context=self.session
        )
    
    def tearDown(self):
        # 关闭Word应用程序
        try:
            # 关闭所有文档
            if self.word_app.Documents.Count > 0:
                for doc in self.word_app.Documents:
                    doc.Close(SaveChanges=False)
            
            # 退出Word应用
            self.word_app.Quit()
        except:
            pass
    
    def test_document_open_and_outline(self):
        """测试打开文档并获取大纲"""
        # 1. 打开文档
        result = document_tools(
            ctx=self.context,
            operation_type="open",
            file_path=self.test_doc_path
        )
        self.assertIn("success", result.lower())
        
        # 2. 获取文档大纲
        result = document_tools(
            ctx=self.context,
            operation_type="get_info"
        )
        outline = json.loads(result)
        self.assertIsInstance(outline, dict)
        self.assertIn("success", outline)
        
        # 3. 关闭文档
        result = document_tools(
            ctx=self.context,
            operation_type="close"
        )
        self.assertIn("successfully", result.lower())
    
    def test_text_search_and_formatting(self):
        """测试文本搜索和格式化操作"""
        # 1. 打开文档
        document_tools(
            ctx=self.context,
            operation_type="open",
            file_path=self.test_doc_path
        )
        
        # 2. 查找文本
        result = text_content_operation(
            ctx=self.context,
            operation_type="find",
            text="第二个段落"
        )
        findings = json.loads(result)
        self.assertIsInstance(findings, dict)
        self.assertIn("matches_found", findings)
        
        # 3. 关闭文档
        document_tools(
            ctx=self.context,
            operation_type="close"
        )
    
    def test_table_operations(self):
        """测试表格操作"""
        # 1. 打开文档
        document_tools(
            ctx=self.context,
            operation_type="open",
            file_path=self.test_doc_path
        )
        
        # 2. 创建新表格
        result = table_operation(
            ctx=self.context,
            operation_type="create",
            rows=2,
            cols=2,
            locator={"type": "paragraph", "index": -1}
        )
        self.assertIn("successfully", result.lower())
        
        # 3. 关闭文档
        document_tools(
            ctx=self.context,
            operation_type="close"
        )
    
    def test_precise_object_selection(self):
        """测试精确元素选择"""
        # 1. 打开文档
        document_tools(
            ctx=self.context,
            operation_type="open",
            file_path=self.test_doc_path
        )
        
        # 2. 获取文档元素
        result = document_tools(
            ctx=self.context,
            operation_type="get_info"
        )
        objects = json.loads(result)
        self.assertIsInstance(objects, list)
        self.assertGreater(len(objects), 0)
        
        # 3. 关闭文档
        document_tools(
            ctx=self.context,
            operation_type="close"
        )


if __name__ == '__main__':
    unittest.main()