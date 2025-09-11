# -*- coding: utf-8 -*-
"""
测试objects_tools.py中的对象操作功能
"""

import json
from unittest.mock import MagicMock, patch

from word_docx_tools.tools.objects_tools import objects_tools
from tests.test_utils import WordDocumentTestBase


class TestObjectsTools(WordDocumentTestBase):
    """Tests for objects_tools module"""

    def setUp(self):
        """测试前准备"""
        # 调用基类的setUp方法，创建Word应用程序、文档和上下文
        super().setUp()

        # 创建模拟的对象操作
        self._setup_mock_objects_operations()

    def _setup_mock_objects_operations(self):
        """设置模拟的对象操作"""
        # 为对象操作创建模拟实现
        if isinstance(self.doc, MagicMock):
            # 设置模拟的InlineShapes和Shapes集合
            self.doc.InlineShapes = MagicMock()
            self.doc.InlineShapes.Count = 0
            self.doc.InlineShapes.AddPicture = MagicMock(return_value=MagicMock())
            self.doc.InlineShapes.Item = MagicMock(return_value=MagicMock())
            
            self.doc.Shapes = MagicMock()
            self.doc.Shapes.Count = 0
            self.doc.Shapes.AddTextbox = MagicMock(return_value=MagicMock())
            self.doc.Shapes.Item = MagicMock(return_value=MagicMock())

    def test_get_objects_info(self):
        """测试获取对象信息的功能"""
        # 准备请求参数
        request_params = {
            "operation_type": "get_objects_info",
            "params": {
                "object_types": ["inline_shape", "shape"],
                "start_index": 1,
                "end_index": 10
            }
        }

        # 调用工具函数
        result = objects_tools(self.context, request_params)

        # 使用基类的验证方法验证结果
        result_data = self.verify_tool_result(result, expected_success=True)

        # 检查返回的数据结构
        self.assertIn("objects", result_data)
        self.assertIsInstance(result_data["objects"], list)

    def test_insert_inline_shape(self):
        """测试插入嵌入式形状的功能"""
        # 准备请求参数
        request_params = {
            "operation_type": "insert_inline_shape",
            "params": {
                "shape_type": "textbox",
                "width": 100,
                "height": 50,
                "text": "嵌入式文本框内容"
            }
        }

        # 调用工具函数
        result = objects_tools(self.context, request_params)

        # 使用基类的验证方法验证结果
        result_data = self.verify_tool_result(result, expected_success=True)

        # 检查返回的插入信息
        self.assertIn("inserted", result_data)
        self.assertTrue(result_data["inserted"])

    def test_delete_object(self):
        """测试删除对象的功能"""
        # 准备请求参数
        request_params = {
            "operation_type": "delete_object",
            "params": {
                "object_id": 1,
                "object_type": "inline_shape"
            }
        }

        # 调用工具函数
        result = objects_tools(self.context, request_params)

        # 使用基类的验证方法验证结果
        result_data = self.verify_tool_result(result, expected_success=True)

        # 检查返回的删除信息
        self.assertIn("deleted", result_data)
        self.assertTrue(result_data["deleted"])

    def test_update_object(self):
        """测试更新对象属性的功能"""
        # 准备请求参数
        request_params = {
            "operation_type": "update_object",
            "params": {
                "object_id": 1,
                "object_type": "inline_shape",
                "properties": {
                    "width": 150,
                    "height": 75
                }
            }
        }

        # 调用工具函数
        result = objects_tools(self.context, request_params)

        # 使用基类的验证方法验证结果
        result_data = self.verify_tool_result(result, expected_success=True)

        # 检查返回的更新信息
        self.assertIn("updated", result_data)
        self.assertTrue(result_data["updated"])

    def test_objects_tools_invalid_operation(self):
        """测试无效的操作类型"""
        # 准备请求参数，使用无效的操作类型
        request_params = {
            "operation_type": "invalid_operation",
            "params": {}
        }

        # 调用工具函数
        result = objects_tools(self.context, request_params)

        # 验证结果应该是失败的
        result_data = self.verify_tool_result(result, expected_success=False)

        # 检查错误信息
        self.assertIn("error", result_data)
        self.assertIn("无效的操作类型", result_data["error"])

    def test_objects_tools_missing_required_params(self):
        """测试缺少必要参数的情况"""
        # 准备请求参数，缺少必要的参数
        request_params = {
            "operation_type": "delete_object",
            # 故意不提供params
        }

        # 调用工具函数
        result = objects_tools(self.context, request_params)

        # 验证结果应该是失败的
        result_data = self.verify_tool_result(result, expected_success=False)

        # 检查错误信息
        self.assertIn("error", result_data)
        self.assertIn("缺少必要的参数", result_data["error"])

    def test_move_object(self):
        """测试移动对象的功能"""
        # 准备请求参数
        request_params = {
            "operation_type": "move_object",
            "params": {
                "object_id": 1,
                "object_type": "shape",
                "horizontal_position": 100,
                "vertical_position": 200,
                "relative_to": "page"
            }
        }

        # 调用工具函数
        result = objects_tools(self.context, request_params)

        # 使用基类的验证方法验证结果
        result_data = self.verify_tool_result(result, expected_success=True)

        # 检查返回的移动信息
        self.assertIn("moved", result_data)
        self.assertTrue(result_data["moved"])


if __name__ == "__main__":
    import unittest
    unittest.main()