# -*- coding: utf-8 -*-
"""
测试styles_tools.py中的样式操作功能
"""

import json
from unittest.mock import MagicMock, patch

from word_docx_tools.tools.styles_tools import styles_tools
from tests.test_utils import WordDocumentTestBase


class TestStylesTools(WordDocumentTestBase):
    """Tests for styles_tools module"""

    def setUp(self):
        """测试前准备"""
        # 调用基类的setUp方法，创建Word应用程序、文档和上下文
        super().setUp()

        # 创建模拟的样式操作
        self._setup_mock_styles_operations()

    def _setup_mock_styles_operations(self):
        """设置模拟的样式操作"""
        # 为样式操作创建模拟实现
        if isinstance(self.doc, MagicMock):
            # 设置模拟的Style对象
            self.mock_style = MagicMock()
            self.mock_style.NameLocal = "标题 1"
            self.mock_style.Font = MagicMock()
            self.mock_style.Font.Name = "宋体"
            self.mock_style.Font.Size = 16
            self.mock_style.Font.Bold = True
            
            # 设置文档的Styles集合
            self.mock_styles = MagicMock()
            self.mock_styles.Count = 10
            self.mock_styles.Item.return_value = self.mock_style
            
            # 将Styles集合设置到文档对象上
            self.doc.Styles = self.mock_styles

    def test_get_all_styles(self):
        """测试获取所有样式的功能"""
        # 准备请求参数
        request_params = {
            "operation_type": "get_all_styles",
            "params": {}
        }

        # 调用工具函数
        result = styles_tools(self.context, request_params)

        # 使用基类的验证方法验证结果
        result_data = self.verify_tool_result(result, expected_success=True)

        # 检查返回的样式列表
        self.assertIn("styles", result_data)
        self.assertIsInstance(result_data["styles"], list)

    def test_get_style_info(self):
        """测试获取特定样式信息的功能"""
        # 准备请求参数
        request_params = {
            "operation_type": "get_style_info",
            "params": {
                "style_name": "标题 1"
            }
        }

        # 调用工具函数
        result = styles_tools(self.context, request_params)

        # 使用基类的验证方法验证结果
        result_data = self.verify_tool_result(result, expected_success=True)

        # 检查返回的样式信息
        self.assertIn("style_name", result_data)
        self.assertIn("font_name", result_data)
        self.assertIn("font_size", result_data)
        self.assertEqual(result_data["style_name"], "标题 1")

        # 验证获取样式的方法是否被调用
        if isinstance(self.doc, MagicMock):
            self.mock_styles.Item.assert_called_with("标题 1")

    def test_create_style(self):
        """测试创建新样式的功能"""
        # 准备请求参数
        request_params = {
            "operation_type": "create_style",
            "params": {
                "style_name": "自定义样式",
                "base_style": "正文",
                "font_name": "微软雅黑",
                "font_size": 14,
                "bold": True,
                "italic": False,
                "underline": False
            }
        }

        # 调用工具函数
        result = styles_tools(self.context, request_params)

        # 使用基类的验证方法验证结果
        result_data = self.verify_tool_result(result, expected_success=True)

        # 检查返回的创建信息
        self.assertIn("style_created", result_data)
        self.assertTrue(result_data["style_created"])
        self.assertIn("style_name", result_data)
        self.assertEqual(result_data["style_name"], "自定义样式")

    def test_update_style(self):
        """测试更新样式的功能"""
        # 准备请求参数
        request_params = {
            "operation_type": "update_style",
            "params": {
                "style_name": "标题 1",
                "font_name": "微软雅黑",
                "font_size": 18
            }
        }

        # 调用工具函数
        result = styles_tools(self.context, request_params)

        # 使用基类的验证方法验证结果
        result_data = self.verify_tool_result(result, expected_success=True)

        # 检查返回的更新信息
        self.assertIn("style_updated", result_data)
        self.assertTrue(result_data["style_updated"])

        # 验证获取样式的方法是否被调用
        if isinstance(self.doc, MagicMock):
            self.mock_styles.Item.assert_called_with("标题 1")

    def test_apply_style_to_paragraph(self):
        """测试应用样式到段落的功能"""
        # 准备请求参数
        request_params = {
            "operation_type": "apply_style",
            "params": {
                "element_type": "paragraph",
                "element_id": 1,
                "style_name": "标题 1"
            }
        }

        # 调用工具函数
        result = styles_tools(self.context, request_params)

        # 使用基类的验证方法验证结果
        result_data = self.verify_tool_result(result, expected_success=True)

        # 检查返回的应用信息
        self.assertIn("style_applied", result_data)
        self.assertTrue(result_data["style_applied"])

    def test_apply_style_to_text(self):
        """测试应用样式到文本的功能"""
        # 准备请求参数
        request_params = {
            "operation_type": "apply_style",
            "params": {
                "element_type": "text",
                "element_id": 1,
                "style_name": "标题 1"
            }
        }

        # 调用工具函数
        result = styles_tools(self.context, request_params)

        # 使用基类的验证方法验证结果
        result_data = self.verify_tool_result(result, expected_success=True)

        # 检查返回的应用信息
        self.assertIn("style_applied", result_data)
        self.assertTrue(result_data["style_applied"])

    def test_delete_style(self):
        """测试删除样式的功能"""
        # 准备请求参数
        request_params = {
            "operation_type": "delete_style",
            "params": {
                "style_name": "自定义样式"
            }
        }

        # 调用工具函数
        result = styles_tools(self.context, request_params)

        # 使用基类的验证方法验证结果
        result_data = self.verify_tool_result(result, expected_success=True)

        # 检查返回的删除信息
        self.assertIn("style_deleted", result_data)
        self.assertTrue(result_data["style_deleted"])

    def test_styles_tools_invalid_operation(self):
        """测试无效的操作类型"""
        # 准备请求参数，使用无效的操作类型
        request_params = {
            "operation_type": "invalid_operation",
            "params": {}
        }

        # 调用工具函数
        result = styles_tools(self.context, request_params)

        # 验证结果应该是失败的
        result_data = self.verify_tool_result(result, expected_success=False)

        # 检查错误信息
        self.assertIn("error", result_data)
        self.assertIn("无效的操作类型", result_data["error"])


if __name__ == "__main__":
    import unittest
    unittest.main()