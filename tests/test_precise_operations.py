"""
Precise document operations tests for Word Document MCP Server.

This module contains tests that focus on precise operations using locators
to target specific objects in documents.
"""

import unittest
import os
import tempfile
import shutil
import json
from pathlib import Path
import win32com.client
import pythoncom
from io import StringIO

from word_document_server.tools.document_tools import document_tools
from word_document_server.tools.text_tools import text_tools
from word_document_server.tools.table_tools import table_tools
from word_document_server.tools.comment_tools import comment_tools
from word_document_server.tools.range_tools import range_tools
from word_document_server.utils.app_context import AppContext
from mcp.server.fastmcp import Context
from mcp.server.session import ServerSession


class TestPreciseOperations(unittest.TestCase):
    """Tests for precise document operations using locators"""
    
    @classmethod
    def setUpClass(cls):
        # 初始化COM
        pythoncom.CoInitialize()
        
        # 创建临时目录用于测试
        cls.test_dir = tempfile.mkdtemp()
        
        # 创建多个测试文件以提高测试覆盖率
        cls.test_files = {}
        
        # 创建简单文本测试文档
        cls.test_files['simple'] = os.path.join(cls.test_dir, "simple_test.docx")
        cls._create_simple_test_document(cls.test_files['simple'])
        
        # 创建带表格的测试文档
        cls.test_files['table'] = os.path.join(cls.test_dir, "table_test.docx")
        cls._create_table_test_document(cls.test_files['table'])
        
        # 创建带注释的测试文档
        cls.test_files['comment'] = os.path.join(cls.test_dir, "comment_test.docx")
        cls._create_comment_test_document(cls.test_files['comment'])
        
        # 创建复杂格式的测试文档
        cls.test_files['complex'] = os.path.join(cls.test_dir, "complex_test.docx")
        cls._create_complex_test_document(cls.test_files['complex'])
    
    @classmethod
    def _create_simple_test_document(cls, path):
        """创建一个简单的测试文档"""
        word_app = win32com.client.Dispatch('Word.Application')
        try:
            word_app.Visible = False
        except AttributeError:
            # 有些环境下无法设置Visible属性，可以忽略
            pass
            
        doc = word_app.Documents.Add()
        doc.Range(0, 0).Text = "简单测试文档\n\n这是第一段文本。\n这是第二段文本，包含一些内容。\n这是第三段文本。"
        doc.SaveAs2(path)
        doc.Close()
        word_app.Quit()
    
    @classmethod
    def _create_table_test_document(cls, path):
        """创建一个包含表格的测试文档"""
        word_app = win32com.client.Dispatch('Word.Application')
        try:
            word_app.Visible = False
        except AttributeError:
            # 有些环境下无法设置Visible属性，可以忽略
            pass
            
        doc = word_app.Documents.Add()
        doc.Range(0, 0).Text = "表格测试文档\n\n"
        
        # 添加表格
        table_range = doc.Range()
        table_range.Collapse(0)  # Collapse to end
        table = doc.Tables.Add(table_range, 3, 4)
        table.Cell(1, 1).Range.Text = "ID"
        table.Cell(1, 2).Range.Text = "姓名"
        table.Cell(1, 3).Range.Text = "年龄"
        table.Cell(1, 4).Range.Text = "城市"
        table.Cell(2, 1).Range.Text = "1"
        table.Cell(2, 2).Range.Text = "张三"
        table.Cell(2, 3).Range.Text = "25"
        table.Cell(2, 4).Range.Text = "北京"
        table.Cell(3, 1).Range.Text = "2"
        table.Cell(3, 2).Range.Text = "李四"
        table.Cell(3, 3).Range.Text = "30"
        table.Cell(3, 4).Range.Text = "上海"
        
        doc.SaveAs2(path)
        doc.Close()
        word_app.Quit()
    
    @classmethod
    def _create_comment_test_document(cls, path):
        """创建一个包含注释的测试文档"""
        word_app = win32com.client.Dispatch('Word.Application')
        try:
            word_app.Visible = False
        except AttributeError:
            # 有些环境下无法设置Visible属性，可以忽略
            pass
            
        doc = word_app.Documents.Add()
        doc.Range(0, 0).Text = "注释测试文档\n\n这是第一段文本。\n这是第二段文本，我们在这里添加注释。\n\n"
        
        # 添加注释
        second_paragraph = doc.Paragraphs(2).Range
        doc.Comments.Add(second_paragraph, "这是一个测试注释")
        
        doc.SaveAs2(path)
        doc.Close()
        word_app.Quit()
    
    @classmethod
    def _create_complex_test_document(cls, path):
        """创建一个包含多种元素的复杂测试文档"""
        word_app = win32com.client.Dispatch('Word.Application')
        try:
            word_app.Visible = False
        except AttributeError:
            # 有些环境下无法设置Visible属性，可以忽略
            pass
            
        doc = word_app.Documents.Add()
        doc.Range(0, 0).Text = "复杂测试文档\n\n"
        # 设置标题样式
        try:
            doc.Paragraphs(1).Range.Style = "Heading 1"
        except:
            pass  # 如果设置样式失败则跳过
        
        # 添加内容段落
        doc.Range().Collapse(0)
        doc.Range().Text = "这是文档正文的第一段。\n"
        
        doc.Range().Collapse(0)
        doc.Range().Text = "这是文档正文的第二段，包含一些关键词。\n"
        
        # 添加表格
        doc.Range().Collapse(0)
        doc.Range().Text = "\n"
        table_range = doc.Range()
        table_range.Collapse(0)
        table = doc.Tables.Add(table_range, 2, 2)
        table.Cell(1, 1).Range.Text = "项目"
        table.Cell(1, 2).Range.Text = "描述"
        table.Cell(2, 1).Range.Text = "测试项"
        table.Cell(2, 2).Range.Text = "这是测试项的描述"
        
        # 添加结尾段落
        doc.Range().Collapse(0)
        doc.Range().Text = "\n这是文档的最后一段。\n"
        
        doc.SaveAs2(path)
        doc.Close()
        word_app.Quit()
    
    @classmethod
    def tearDownClass(cls):
        # 清理临时目录
        try:
            shutil.rmtree(cls.test_dir)
        except:
            pass
        pythoncom.CoUninitialize()
    
    def setUp(self):
        # 创建Word应用程序实例
        self.word_app = win32com.client.Dispatch('Word.Application')
        try:
            self.word_app.Visible = False
        except AttributeError:
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
            if self.word_app.Documents.Count > 0:
                for doc in self.word_app.Documents:
                    doc.Close(SaveChanges=False)
            self.word_app.Quit()
        except:
            pass
    
    def test_open_and_close_document_simple(self):
        """测试打开和关闭简单文档"""
        # 打开文档
        result = document_tools(
            ctx=self.context,
            operation_type="open",
            file_path=self.test_files['simple']
        )
        self.assertIn("successfully", result.lower())
        
        # 关闭文档
        result = document_tools(
            ctx=self.context,
            operation_type="close"
        )
        self.assertIn("successfully", result.lower())
    
    def test_open_and_close_document_table(self):
        """测试打开和关闭表格文档"""
        # 打开文档
        result = document_tools(
            ctx=self.context,
            operation_type="open",
            file_path=self.test_files['table']
        )
        self.assertIn("successfully", result.lower())
        
        # 关闭文档
        result = document_tools(
            ctx=self.context,
            operation_type="close"
        )
        self.assertIn("successfully", result.lower())
    
    def test_open_and_close_document_comment(self):
        """测试打开和关闭注释文档"""
        # 打开文档
        result = document_tools(
            ctx=self.context,
            operation_type="open",
            file_path=self.test_files['comment']
        )
        self.assertIn("successfully", result.lower())
        
        # 关闭文档
        result = document_tools(
            ctx=self.context,
            operation_type="close"
        )
        self.assertIn("successfully", result.lower())
    
    def test_get_document_objects_simple(self):
        """测试获取简单文档元素"""
        # 打开文档
        document_tools(
            ctx=self.context,
            operation_type="open",
            file_path=self.test_files['simple']
        )
        
        # 获取段落元素
        result = document_tools(
            ctx=self.context,
            operation_type="get_objects",
            object_type="paragraphs"
        )
        objects = json.loads(result)
        self.assertIsInstance(objects, list)
        self.assertGreater(len(objects), 0)
        
        # 关闭文档
        document_tools(
            ctx=self.context,
            operation_type="close"
        )
    
    def test_get_document_objects_table(self):
        """测试获取表格文档元素"""
        # 打开文档
        document_tools(
            ctx=self.context,
            operation_type="open",
            file_path=self.test_files['table']
        )
        
        # 获取表格元素
        result = document_tools(
            ctx=self.context,
            operation_type="get_objects",
            object_type="tables"
        )
        objects = json.loads(result)
        self.assertIsInstance(objects, list)
        self.assertGreater(len(objects), 0)
        
        # 关闭文档
        document_tools(
            ctx=self.context,
            operation_type="close"
        )
    
    def test_create_table_in_simple_document(self):
        """测试在简单文档中创建表格"""
        # 打开文档
        document_tools(
            ctx=self.context,
            operation_type="open",
            file_path=self.test_files['simple']
        )

        # 在文档末尾创建表格
        result = table_tools(
            ctx=self.context,
            operation_type="create",
            rows=2,
            cols=3,
            locator={"target": {"type": "document_end"}}
        )
        # 检查结果是否包含成功信息（中英文都支持）
        self.assertTrue("成功" in result or "successfully" in result.lower())
        
        # 关闭文档
        document_tools(
            ctx=self.context,
            operation_type="close"
        )

    def test_get_comments_in_comment_document(self):
        """测试获取注释文档中的注释"""
        # 打开文档
        document_tools(
            ctx=self.context,
            operation_type="open",
            file_path=self.test_files['comment']
        )
        
        # 获取注释
        result = comment_tools(
            ctx=self.context,
            operation_type="get_all"
        )
        comments = json.loads(result)
        self.assertIsInstance(comments, dict)
        
        # 关闭文档
        document_tools(
            ctx=self.context,
            operation_type="close"
        )


if __name__ == '__main__':
    unittest.main()