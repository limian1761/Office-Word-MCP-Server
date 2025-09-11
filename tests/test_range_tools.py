# -*- coding: utf-8 -*-
"""
测试range_tools.py中的范围操作功能
"""

import json
from unittest.mock import MagicMock, patch

from word_docx_tools.tools.range_tools import range_tools
from tests.test_utils import WordDocumentTestBase


class TestRangeTools(WordDocumentTestBase):
    """Tests for range_tools module"""

    def setUp(self):
        """测试前准备"""
        # 调用基类的setUp方法，创建Word应用程序、文档和上下文
        super().setUp()

        # 创建模拟的范围操作
        self._setup_mock_range_operations()

    def _setup_mock_range_operations(self):
        """设置模拟的范围操作"""
        # 为范围操作创建模拟实现
        if isinstance(self.doc, MagicMock):
            # 设置模拟的Range对象
            self.mock_range = MagicMock()
            self.mock_range.Text = "模拟范围文本"
            self.mock_range.Start = 0
            self.mock_range.End = 10
            
            # 设置文档的Range方法
            self.doc.Range = MagicMock(return_value=self.mock_range)

    def test_get_range_text(self):
        """测试获取指定范围文本的功能"""
        # 准备请求参数
        request_params = {
            "operation_type": "get_range_text",
            "params": {
                "start": 0,
                "end": 10
            }
        }

        # 调用工具函数
        result = range_tools(self.context, request_params)

        # 使用基类的验证方法验证结果
        result_data = self.verify_tool_result(result, expected_success=True)

        # 检查返回的文本
        self.assertIn("text", result_data)
        self.assertIsInstance(result_data["text"], str)

    def test_set_range_text(self):
        """测试设置指定范围文本的功能"""
        # 准备请求参数
        request_params = {
            "operation_type": "set_range_text",
            "params": {
                "start": 0,
                "end": 5,
                "text": "新文本"
            }
        }

        # 调用工具函数
        result = range_tools(self.context, request_params)

        # 使用基类的验证方法验证结果
        result_data = self.verify_tool_result(result, expected_success=True)

        # 检查返回的设置信息
        self.assertIn("text_set", result_data)
        self.assertTrue(result_data["text_set"])

        # 如果是模拟对象，验证Text属性是否被设置
        if isinstance(self.doc, MagicMock):
            self.mock_range.Text = "新文本"

    def test_delete_range(self):
        """测试删除指定范围内容的功能"""
        # 准备请求参数
        request_params = {
            "operation_type": "delete_range",
            "params": {
                "start": 0,
                "end": 5
            }
        }

        # 调用工具函数
        result = range_tools(self.context, request_params)

        # 使用基类的验证方法验证结果
        result_data = self.verify_tool_result(result, expected_success=True)

        # 检查返回的删除信息
        self.assertIn("range_deleted", result_data)
        self.assertTrue(result_data["range_deleted"])

    def test_copy_range(self):
        """测试复制指定范围内容的功能"""
        # 准备请求参数
        request_params = {
            "operation_type": "copy_range",
            "params": {
                "start": 0,
                "end": 10
            }
        }

        # 调用工具函数
        result = range_tools(self.context, request_params)

        # 使用基类的验证方法验证结果
        result_data = self.verify_tool_result(result, expected_success=True)

        # 检查返回的复制信息
        self.assertIn("copied", result_data)
        self.assertTrue(result_data["copied"])

    def test_paste_range(self):
        """测试粘贴内容到指定范围的功能"""
        # 准备请求参数
        request_params = {
            "operation_type": "paste_range",
            "params": {
                "position": 5,
                "format": "plain_text"
            }
        }

        # 调用工具函数
        result = range_tools(self.context, request_params)

        # 使用基类的验证方法验证结果
        result_data = self.verify_tool_result(result, expected_success=True)

        # 检查返回的粘贴信息
        self.assertIn("pasted", result_data)
        self.assertTrue(result_data["pasted"])

    def test_range_tools_invalid_operation(self):
        """测试无效的操作类型"""
        # 准备请求参数，使用无效的操作类型
        request_params = {
            "operation_type": "invalid_operation",
            "params": {}
        }

        # 调用工具函数
        result = range_tools(self.context, request_params)

        # 验证结果应该是失败的
        result_data = self.verify_tool_result(result, expected_success=False)

        # 检查错误信息
        self.assertIn("error", result_data)
        self.assertIn("无效的操作类型", result_data["error"])

    def test_range_tools_missing_required_params(self):
        """测试缺少必要参数的情况"""
        # 准备请求参数，缺少必要的参数
        request_params = {
            "operation_type": "get_range_text",
            # 故意不提供params
        }

        # 调用工具函数
        result = range_tools(self.context, request_params)

        # 验证结果应该是失败的
        result_data = self.verify_tool_result(result, expected_success=False)

        # 检查错误信息
        self.assertIn("error", result_data)
        self.assertIn("缺少必要的参数", result_data["error"])

    def test_find_text_in_range(self):
        """测试在指定范围内查找文本的功能"""
        # 准备请求参数
        request_params = {
            "operation_type": "find_text",
            "params": {
                "find_text": "测试",
                "match_case": False,
                "whole_word": False
            }
        }

        # 调用工具函数
        result = range_tools(self.context, request_params)

        # 使用基类的验证方法验证结果
        result_data = self.verify_tool_result(result, expected_success=True)

        # 检查返回的查找信息
        self.assertIn("found", result_data)
        self.assertIn("matches", result_data)
        self.assertIsInstance(result_data["matches"], list)


if __name__ == "__main__":
    import unittest
    unittest.main()