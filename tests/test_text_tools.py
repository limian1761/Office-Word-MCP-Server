#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
测试text_tools.py中的insert_text操作
"""

import json
import os
import shutil
import tempfile
import unittest
from io import StringIO
from unittest.mock import MagicMock, patch

import pythoncom
import win32com.client
from mcp.server.fastmcp import Context
from mcp.server.session import ServerSession

from word_docx_tools.operations.text_ops import insert_text_after_object
from word_docx_tools.tools.text_tools import text_tools
from word_docx_tools.mcp_service.app_context import AppContext


class TestTextToolsInsertText(unittest.TestCase):
    """Tests for text_tools.insert_text operation"""

    @classmethod
    def setUpClass(cls):
        # 初始化COM
        pythoncom.CoInitialize()

        # 创建临时目录用于测试
        cls.test_dir = tempfile.mkdtemp()

        # 创建测试文档路径
        cls.test_doc_path = os.path.join(cls.test_dir, "test_document.docx")

        # 创建一个空的测试文档
        word_app = win32com.client.Dispatch("Word.Application")
        try:
            doc = word_app.Documents.Add()
            doc.SaveAs2(cls.test_doc_path)
            doc.Close()
        except Exception as e:
            print(f"创建测试文档失败: {str(e)}")
        finally:
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
        """测试前准备"""
        # 创建Word应用程序实例
        self.word_app = win32com.client.Dispatch("Word.Application")
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
            init_options={},
        )
        self.session.lifespan_context = self.app_context

        # 创建Context对象
        self.context = Context(
            server_session=self.session, request_context=self.session
        )

        # 创建测试文档
        self.doc = self.word_app.Documents.Add()
        # 将文档设置为活动文档
        self.app_context.active_document = self.doc

    def tearDown(self):
        # 关闭文档
        try:
            if hasattr(self, "doc"):
                self.doc.Close(SaveChanges=False)
        except:
            pass

        # 关闭Word应用程序
        try:
            self.word_app.Quit()
        except:
            pass

    @patch("word_docx_tools.operations.text_ops.insert_text")
    def test_insert_text_with_document_start_locator(self, mock_insert_text):
        """测试使用document_start定位器的insert_text操作"""
        # 设置mock返回值
        mock_result = json.dumps(
            {"success": True, "message": "Text inserted successfully"}
        )
        mock_insert_text.return_value = mock_result

        # 定义测试参数
        test_text = "这是word-docx-tools的测试文档。"
        test_locator = {"type": "document_start"}

        # 调用text_tools插入文本
        result = text_tools(
            ctx=self.context,
            operation_type="insert_text",
            text=test_text,
            locator=test_locator,
            position="after",
        )

        # 验证结果
        result_data = json.loads(result)
        self.assertTrue(result_data["success"])
        self.assertEqual(result_data["message"], "Text inserted successfully")

        # 验证insert_text被正确调用
        mock_insert_text.assert_called_once()

    def test_insert_text_after_object_function(self):
        """测试insert_text_after_object函数"""
        # 准备测试文本和元素
        test_text = "这是插入的测试文本"

        # 使用文档的Content作为测试元素
        object = self.doc.Content

        # 调用insert_text_after_object函数
        result = insert_text_after_object(object, test_text)

        # 验证结果
        result_data = json.loads(result)
        self.assertTrue(result_data["success"])
        self.assertEqual(result_data["message"], "Text inserted successfully")

        # 验证文本是否被正确插入
        # 在实际环境中，我们需要访问文档内容来验证插入是否成功
        # 由于这是测试环境，我们主要验证函数返回值和调用逻辑
        # 添加额外的验证确保插入的文本确实存在于文档中
        self.assertIn(test_text, self.doc.Content.Text)

    def test_text_tools_invalid_operation(self):
        """测试无效的操作类型"""
        # 调用text_tools并使用无效的操作类型
        result = text_tools(ctx=self.context, operation_type="invalid_operation")

        # 验证结果包含错误信息
        result_data = json.loads(result)
        self.assertFalse(result_data["success"])
        self.assertIn("error", result_data["message"].lower())


if __name__ == "__main__":
    # 创建测试套件
    test_suite = unittest.TestSuite()

    # 添加所有测试方法
    for method_name in dir(TestTextToolsInsertText):
        if method_name.startswith("test_"):
            test_case = TestTextToolsInsertText(method_name)
            test_suite.addTest(test_case)

    # 运行测试
    unittest.TextTestRunner(verbosity=2).run(test_suite)
