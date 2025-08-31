"""
Real end-to-end tests for Word Document MCP Server.

This module contains tests that perform actual COM operations with real documents
to verify the core functionality of the system.
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
from word_document_server.tools.element_tools import element_tools
from word_document_server.utils.app_context import AppContext
from mcp.server.fastmcp import Context
from mcp.server.session import ServerSession


class TestRealE2E(unittest.TestCase):
    """Real end-to-end tests with actual COM operations"""
    
    @classmethod
    def setUpClass(cls):
        # 初始化COM
        pythoncom.CoInitialize()
        
        # 创建临时目录用于测试
        cls.test_dir = tempfile.mkdtemp()
        
        # 复制测试文档到临时目录
        source_doc = os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 
            "tests", "test_docs", "valid_test_document_v2.docx"
        )
        cls.test_doc_path = os.path.join(cls.test_dir, "valid_test_document_v2.docx")
        shutil.copy2(source_doc, cls.test_doc_path)
        
        # 创建输出文档路径
        cls.output_doc_path = os.path.join(cls.test_dir, "output_document.docx")
    
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
        # 尝试设置Visible属性，但捕获可能的异常
        try:
            self.word_app.Visible = False
        except AttributeError:
            # 某些环境中可能不支持设置Visible属性，忽略此错误
            pass
        
        # 创建应用上下文
        self.app_context = AppContext(self.word_app)
        
        # 创建session
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
    
    def test_document_operations_with_locator(self):
        """测试使用定位器的文档操作"""
        # 1. 打开文档
        result = document_tools(
            ctx=self.context,
            operation_type="open",
            file_path=self.test_doc_path
        )
        self.assertIn("successfully", result.lower())
        
        # 2. 获取文档样式
        result = document_tools(
            ctx=self.context,
            operation_type="get_styles"
        )
        # 验证返回的是有效的JSON格式样式列表
        styles = json.loads(result)
        self.assertIsInstance(styles, list)
        self.assertGreater(len(styles), 0)
        
        # 3. 获取文档元素
        result = document_tools(
            ctx=self.context,
            operation_type="get_elements",
            element_type="paragraphs"
        )
        # 验证返回的是有效的JSON格式元素列表
        elements = json.loads(result)
        self.assertIsInstance(elements, list)
        self.assertGreater(len(elements), 0)
        
        # 4. 关闭文档
        result = document_tools(
            ctx=self.context,
            operation_type="close"
        )
        self.assertIn("successfully", result.lower())
    
    def test_text_operations_with_locator(self):
        """测试使用定位器的文本操作"""
        # 1. 打开文档
        document_tools(
            ctx=self.context,
            operation_type="open",
            file_path=self.test_doc_path
        )
        
        # 2. 查找文本
        result = text_tools(
            ctx=self.context,
            operation_type="find",
            text="段落"
        )
        # 验证返回的是有效的JSON格式查找结果
        findings = json.loads(result)
        self.assertIsInstance(findings, dict)
        self.assertIn("matches_found", findings)
        
        # 3. 关闭文档
        document_tools(
            ctx=self.context,
            operation_type="close"
        )
    
    def test_table_operations_with_locator(self):
        """测试使用定位器的表格操作"""
        # 1. 打开文档
        document_tools(
            ctx=self.context,
            operation_type="open",
            file_path=self.test_doc_path
        )
        
        # 2. 创建表格
        result = table_tools(
            ctx=self.context,
            operation_type="create",
            rows=3,
            cols=3
        )
        self.assertIn("successfully", result.lower())
        
        # 3. 关闭文档
        document_tools(
            ctx=self.context,
            operation_type="close"
        )
    
    def test_comment_operations_with_locator(self):
        """测试使用定位器的注释操作"""
        # 1. 打开文档
        document_tools(
            ctx=self.context,
            operation_type="open",
            file_path=self.test_doc_path
        )
        
        # 2. 添加注释（使用简单的定位器）
        result = comment_tools(
            ctx=self.context,
            operation_type="add",
            locator={"target": "document_start"},
            text="这是一个测试注释",
            author="Test User"
        )
        self.assertIn("successfully", result.lower())
        
        # 3. 获取所有注释
        result = comment_tools(
            ctx=self.context,
            operation_type="get_all"
        )
        # 验证返回的是有效的JSON格式注释列表
        comments = json.loads(result)
        self.assertIsInstance(comments, dict)
        self.assertIn("comments_found", comments)
        
        # 4. 关闭文档
        document_tools(
            ctx=self.context,
            operation_type="close"
        )

if __name__ == '__main__':
    unittest.main()