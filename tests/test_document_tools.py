# -*- coding: utf-8 -*-
"""
测试document_tools.py中的文档操作功能
"""

import json
import os
from unittest.mock import MagicMock, patch

from word_docx_tools.tools.document_tools import document_tools
from tests.test_utils import WordDocumentTestBase


class TestDocumentTools(WordDocumentTestBase):
    """Tests for document_tools module"""

    def setUp(self):
        """测试前准备"""
        # 调用基类的setUp方法，创建Word应用程序、文档和上下文
        super().setUp()

        # 创建模拟的文档操作
        self._setup_mock_document_operations()

    def _setup_mock_document_operations(self):
        """设置模拟的文档操作"""
        # 为文档操作创建模拟实现
        if isinstance(self.word_app, MagicMock):
            # 设置模拟的Documents集合
            self.word_app.Documents = MagicMock()
            self.word_app.Documents.Add = MagicMock(return_value=MagicMock())
            self.word_app.Documents.Open = MagicMock(return_value=MagicMock())
            self.word_app.Documents.Count = 1

    def test_create_document(self):
        """测试创建新文档的功能"""
        # 准备请求参数
        request_params = {
            "operation_type": "create_document",
            "params": {
                "template_path": ""
            }
        }

        # 调用工具函数
        result = document_tools(self.context, request_params)

        # 使用基类的验证方法验证结果
        result_data = self.verify_tool_result(result, expected_success=True)

        # 检查返回的文档信息
        self.assertIn("document_created", result_data)
        self.assertTrue(result_data["document_created"])

        # 如果是模拟对象，验证Add方法是否被调用
        if isinstance(self.word_app, MagicMock):
            self.word_app.Documents.Add.assert_called()

    def test_open_document(self):
        """测试打开现有文档的功能"""
        # 创建测试文档
        test_doc_path = self.create_test_file("test_open.docx", "Test document content")

        # 准备请求参数
        request_params = {
            "operation_type": "open_document",
            "params": {
                "file_path": test_doc_path
            }
        }

        # 调用工具函数
        result = document_tools(self.context, request_params)

        # 使用基类的验证方法验证结果
        result_data = self.verify_tool_result(result, expected_success=True)

        # 检查返回的文档信息
        self.assertIn("document_opened", result_data)
        self.assertTrue(result_data["document_opened"])

        # 如果是模拟对象，验证Open方法是否被调用
        if isinstance(self.word_app, MagicMock):
            self.word_app.Documents.Open.assert_called_with(test_doc_path)

    def test_save_document(self):
        """测试保存文档的功能"""
        # 准备请求参数
        request_params = {
            "operation_type": "save_document",
            "params": {
                "file_path": self.get_test_file_path("test_save.docx")
            }
        }

        # 调用工具函数
        result = document_tools(self.context, request_params)

        # 使用基类的验证方法验证结果
        result_data = self.verify_tool_result(result, expected_success=True)

        # 检查返回的保存信息
        self.assertIn("document_saved", result_data)
        self.assertTrue(result_data["document_saved"])

    def test_close_document(self):
        """测试关闭文档的功能"""
        # 准备请求参数
        result = document_tools(
            self.context,
            operation_type="close",
            save_changes=False
        )

        # 使用基类的验证方法验证结果
        result_data = self.verify_tool_result(result, expected_success=True)

        # 检查返回的关闭信息
        self.assertIn("document_closed", result_data)
        self.assertTrue(result_data["document_closed"])

    def test_get_document_info(self):
        """测试获取文档信息的功能"""
        # 准备请求参数
        request_params = {
            "operation_type": "get_document_info",
            "params": {}
        }

        # 调用工具函数
        result = document_tools(self.context, request_params)

        # 使用基类的验证方法验证结果
        result_data = self.verify_tool_result(result, expected_success=True)

        # 检查返回的文档信息
        self.assertIn("document_info", result_data)
        self.assertIsInstance(result_data["document_info"], dict)
        self.assertIn("paragraph_count", result_data["document_info"])

    def test_document_tools_invalid_operation(self):
        """测试无效的操作类型"""
        # 准备请求参数，使用无效的操作类型
        request_params = {
            "operation_type": "invalid_operation",
            "params": {}
        }

        # 调用工具函数
        result = document_tools(self.context, request_params)

        # 验证结果应该是失败的
        result_data = self.verify_tool_result(result, expected_success=False)

        # 检查错误信息
        self.assertIn("error", result_data)
        self.assertIn("无效的操作类型", result_data["error"])

    def test_document_tools_missing_required_params(self):
        """测试缺少必要参数的情况"""
        # 准备请求参数，缺少必要的参数
        request_params = {
            "operation_type": "save_document",
            # 故意不提供params
        }

        # 调用工具函数
        result = document_tools(self.context, request_params)

        # 验证结果应该是失败的
        result_data = self.verify_tool_result(result, expected_success=False)

        # 检查错误信息
        self.assertIn("error", result_data)
        self.assertIn("缺少必要的参数", result_data["error"])

    def test_export_to_pdf(self):
        """测试导出文档为PDF的功能"""
        # 准备请求参数
        request_params = {
            "operation_type": "export_to_pdf",
            "params": {
                "file_path": self.get_test_file_path("test_export.pdf")
            }
        }

        # 调用工具函数
        result = document_tools(self.context, request_params)

        # 使用基类的验证方法验证结果
        result_data = self.verify_tool_result(result, expected_success=True)

        # 检查返回的导出信息
        self.assertIn("pdf_exported", result_data)
        self.assertTrue(result_data["pdf_exported"])


if __name__ == "__main__":
    import unittest
    unittest.main()