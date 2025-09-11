# -*- coding: utf-8 -*-
"""
测试table_tools.py中的表格操作功能
"""

import json
from unittest.mock import MagicMock, patch

from word_docx_tools.tools.table_tools import table_tools
from tests.test_utils import WordDocumentTestBase


class TestTableTools(WordDocumentTestBase):
    """Tests for table_tools module"""

    def setUp(self):
        """测试前准备"""
        # 调用基类的setUp方法，创建Word应用程序、文档和上下文
        super().setUp()

        # 创建模拟的表格操作
        self._setup_mock_table_operations()

    def _setup_mock_table_operations(self):
        """设置模拟的表格操作"""
        # 为表格操作创建模拟实现
        if isinstance(self.doc, MagicMock):
            # 设置模拟的Table对象
            self.mock_table = MagicMock()
            self.mock_table.Rows.Count = 3
            self.mock_table.Columns.Count = 2
            self.mock_table.Cell = MagicMock()
            
            # 设置模拟的Cell对象
            mock_cell = MagicMock()
            mock_cell.Range.Text = "单元格文本"
            self.mock_table.Cell.return_value = mock_cell
            
            # 设置文档的Tables集合
            self.mock_tables = MagicMock()
            self.mock_tables.Count = 1
            self.mock_tables.Item.return_value = self.mock_table
            
            # 将Tables集合设置到文档对象上
            self.doc.Tables = self.mock_tables
            
            # 设置文档的Range方法
            self.mock_range = MagicMock()
            self.doc.Range = MagicMock(return_value=self.mock_range)
            
            # 设置Range的Tables.Add方法
            self.mock_range.Tables = MagicMock()
            self.mock_range.Tables.Add = MagicMock(return_value=self.mock_table)

    def test_create_table(self):
        """测试创建表格的功能"""
        # 准备请求参数
        request_params = {
            "operation_type": "create_table",
            "params": {
                "rows": 3,
                "columns": 2,
                "position": 0
            }
        }

        # 调用工具函数
        result = table_tools(self.context, request_params)

        # 使用基类的验证方法验证结果
        result_data = self.verify_tool_result(result, expected_success=True)

        # 检查返回的表格信息
        self.assertIn("table_id", result_data)
        self.assertIn("rows", result_data)
        self.assertIn("columns", result_data)
        self.assertEqual(result_data["rows"], 3)
        self.assertEqual(result_data["columns"], 2)

        # 验证创建表格的方法是否被调用
        if isinstance(self.doc, MagicMock):
            self.mock_range.Tables.Add.assert_called_once()

    def test_get_table_info(self):
        """测试获取表格信息的功能"""
        # 准备请求参数
        request_params = {
            "operation_type": "get_table_info",
            "params": {
                "table_id": 1
            }
        }

        # 调用工具函数
        result = table_tools(self.context, request_params)

        # 使用基类的验证方法验证结果
        result_data = self.verify_tool_result(result, expected_success=True)

        # 检查返回的表格信息
        self.assertIn("table_id", result_data)
        self.assertIn("rows", result_data)
        self.assertIn("columns", result_data)
        self.assertEqual(result_data["rows"], 3)
        self.assertEqual(result_data["columns"], 2)

        # 验证获取表格的方法是否被调用
        if isinstance(self.doc, MagicMock):
            self.mock_tables.Item.assert_called_with(1)

    def test_get_cell_content(self):
        """测试获取单元格内容的功能"""
        # 准备请求参数
        request_params = {
            "operation_type": "get_cell_content",
            "params": {
                "table_id": 1,
                "row": 1,
                "column": 1
            }
        }

        # 调用工具函数
        result = table_tools(self.context, request_params)

        # 使用基类的验证方法验证结果
        result_data = self.verify_tool_result(result, expected_success=True)

        # 检查返回的单元格内容
        self.assertIn("content", result_data)
        self.assertIsInstance(result_data["content"], str)

        # 验证获取单元格的方法是否被调用
        if isinstance(self.doc, MagicMock):
            self.mock_table.Cell.assert_called_with(1, 1)

    def test_set_cell_content(self):
        """测试设置单元格内容的功能"""
        # 准备请求参数
        request_params = {
            "operation_type": "set_cell_content",
            "params": {
                "table_id": 1,
                "row": 1,
                "column": 1,
                "content": "新的单元格内容"
            }
        }

        # 调用工具函数
        result = table_tools(self.context, request_params)

        # 使用基类的验证方法验证结果
        result_data = self.verify_tool_result(result, expected_success=True)

        # 检查返回的设置信息
        self.assertIn("content_set", result_data)
        self.assertTrue(result_data["content_set"])

        # 验证获取单元格的方法是否被调用
        if isinstance(self.doc, MagicMock):
            self.mock_table.Cell.assert_called_with(1, 1)

    def test_add_row(self):
        """测试添加行的功能"""
        # 准备请求参数
        request_params = {
            "operation_type": "add_row",
            "params": {
                "table_id": 1,
                "position": -1  # -1表示在表格末尾添加
            }
        }

        # 调用工具函数
        result = table_tools(self.context, request_params)

        # 使用基类的验证方法验证结果
        result_data = self.verify_tool_result(result, expected_success=True)

        # 检查返回的行数信息
        self.assertIn("rows", result_data)
        self.assertEqual(result_data["rows"], 4)  # 原先是3行，添加后应该是4行

    def test_add_column(self):
        """测试添加列的功能"""
        # 准备请求参数
        request_params = {
            "operation_type": "add_column",
            "params": {
                "table_id": 1,
                "position": -1  # -1表示在表格末尾添加
            }
        }

        # 调用工具函数
        result = table_tools(self.context, request_params)

        # 使用基类的验证方法验证结果
        result_data = self.verify_tool_result(result, expected_success=True)

        # 检查返回的列数信息
        self.assertIn("columns", result_data)
        self.assertEqual(result_data["columns"], 3)  # 原先是2列，添加后应该是3列

    def test_delete_table(self):
        """测试删除表格的功能"""
        # 准备请求参数
        request_params = {
            "operation_type": "delete_table",
            "params": {
                "table_id": 1
            }
        }

        # 调用工具函数
        result = table_tools(self.context, request_params)

        # 使用基类的验证方法验证结果
        result_data = self.verify_tool_result(result, expected_success=True)

        # 检查返回的删除信息
        self.assertIn("deleted", result_data)
        self.assertTrue(result_data["deleted"])

    def test_table_tools_invalid_operation(self):
        """测试无效的操作类型"""
        # 准备请求参数，使用无效的操作类型
        request_params = {
            "operation_type": "invalid_operation",
            "params": {}
        }

        # 调用工具函数
        result = table_tools(self.context, request_params)

        # 验证结果应该是失败的
        result_data = self.verify_tool_result(result, expected_success=False)

        # 检查错误信息
        self.assertIn("error", result_data)
        self.assertIn("无效的操作类型", result_data["error"])


if __name__ == "__main__":
    import unittest
    unittest.main()