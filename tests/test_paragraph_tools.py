# -*- coding: utf-8 -*-
"""
测试paragraph_tools.py中的段落操作功能
"""

import json
from unittest.mock import MagicMock, patch

from word_docx_tools.tools.paragraph_tools import paragraph_tools
from tests.test_utils import WordDocumentTestBase


class TestParagraphTools(WordDocumentTestBase):
    """Tests for paragraph_tools module"""

    def setUp(self):
        """测试前准备"""
        # 调用基类的setUp方法，创建Word应用程序、文档和上下文
        super().setUp()

        # 创建模拟的段落操作
        self._setup_mock_paragraph_operations()

    def _setup_mock_paragraph_operations(self):
        """设置模拟的段落操作"""
        # 为段落操作创建模拟实现，在测试方法中可以进一步覆盖这些实现
        pass

    def test_get_paragraphs_info(self):
        """测试获取段落信息的功能"""
        # 准备请求参数
        request_params = {
            "operation_type": "get_paragraphs_info",
            "params": {
                "start_index": 1,
                "end_index": 3
            }
        }

        # 调用工具函数
        result = paragraph_tools(self.context, request_params)

        # 使用基类的验证方法验证结果
        result_data = self.verify_tool_result(result, expected_success=True)

        # 检查返回的数据结构
        self.assertIn("paragraphs", result_data)
        self.assertEqual(len(result_data["paragraphs"]), 3)  # 应该返回3个段落的信息

        # 检查每个段落的信息
        for i, paragraph in enumerate(result_data["paragraphs"]):
            self.assertIn("index", paragraph)
            self.assertIn("text", paragraph)
            self.assertEqual(paragraph["index"], i + 1)

    def test_insert_paragraph(self):
        """测试插入段落的功能"""
        # 准备请求参数
        request_params = {
            "operation_type": "insert_paragraph",
            "params": {
                "text": "这是插入的新段落。",
                "index": 2,
                "locator": "before"
            }
        }

        # 调用工具函数
        result = paragraph_tools(self.context, request_params)

        # 使用基类的验证方法验证结果
        result_data = self.verify_tool_result(result, expected_success=True)

        # 检查返回的插入位置
        self.assertIn("inserted_at", result_data)
        self.assertEqual(result_data["inserted_at"], 2)

        # 如果是真实文档，验证段落是否真的插入了
        if not isinstance(self.doc, MagicMock):
            self.assertGreaterEqual(self.doc.Paragraphs.Count, 4)  # 至少应该有4个段落

    def test_delete_paragraph(self):
        """测试删除段落的功能"""
        # 准备请求参数
        request_params = {
            "operation_type": "delete_paragraph",
            "params": {
                "index": 2
            }
        }

        # 调用工具函数
        result = paragraph_tools(self.context, request_params)

        # 使用基类的验证方法验证结果
        result_data = self.verify_tool_result(result, expected_success=True)

        # 检查返回的删除信息
        self.assertIn("deleted", result_data)
        self.assertTrue(result_data["deleted"])

        # 如果是真实文档，验证段落是否真的删除了
        if not isinstance(self.doc, MagicMock):
            self.assertLessEqual(self.doc.Paragraphs.Count, 2)  # 最多应该有2个段落

    def test_format_paragraph(self):
        """测试格式化段落的功能"""
        # 准备请求参数
        request_params = {
            "operation_type": "format_paragraph",
            "params": {
                "index": 1,
                "format_options": {
                    "alignment": "center",
                    "indent_first_line": 18,
                    "line_spacing": 2
                }
            }
        }

        # 调用工具函数
        result = paragraph_tools(self.context, request_params)

        # 使用基类的验证方法验证结果
        result_data = self.verify_tool_result(result, expected_success=True)

        # 检查返回的格式化信息
        self.assertIn("formatted", result_data)
        self.assertTrue(result_data["formatted"])

    def test_paragraph_tools_invalid_operation(self):
        """测试无效的操作类型"""
        # 准备请求参数，使用无效的操作类型
        request_params = {
            "operation_type": "invalid_operation",
            "params": {}
        }

        # 调用工具函数
        result = paragraph_tools(self.context, request_params)

        # 验证结果应该是失败的
        result_data = self.verify_tool_result(result, expected_success=False)

        # 检查错误信息
        self.assertIn("error", result_data)
        self.assertIn("无效的操作类型", result_data["error"])

    def test_paragraph_tools_missing_required_params(self):
        """测试缺少必要参数的情况"""
        # 准备请求参数，缺少必要的参数
        request_params = {
            "operation_type": "get_paragraphs_info",
            # 故意不提供params
        }

        # 调用工具函数
        result = paragraph_tools(self.context, request_params)

        # 验证结果应该是失败的
        result_data = self.verify_tool_result(result, expected_success=False)

        # 检查错误信息
        self.assertIn("error", result_data)
        self.assertIn("缺少必要的参数", result_data["error"])


if __name__ == "__main__":
    import unittest
    unittest.main()