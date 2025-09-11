# -*- coding: utf-8 -*-
"""
测试comment_tools.py中的评论操作功能
"""

import json
import asyncio
from unittest.mock import MagicMock, patch

from word_docx_tools.tools.comment_tools import comment_tools
from tests.test_utils import WordDocumentTestBase

# 辅助函数，用于在同步测试方法中运行异步代码
def run_async(coroutine):
    return asyncio.run(coroutine)


class TestCommentTools(WordDocumentTestBase):
    """Tests for comment_tools module"""

    def setUp(self):
        """测试前准备"""
        # 调用基类的setUp方法，创建Word应用程序、文档和上下文
        super().setUp()

        # 创建模拟的评论操作
        self._setup_mock_comment_operations()

    def _setup_mock_comment_operations(self):
        """设置模拟的评论操作"""
        # 为评论操作创建模拟实现，在测试方法中可以进一步覆盖这些实现
        if isinstance(self.doc, MagicMock):
            # 设置模拟的Comments集合
            self.doc.Comments = MagicMock()
            self.doc.Comments.Count = 0
            self.doc.Comments.Add = MagicMock()
            self.doc.Comments.Item = MagicMock(return_value=MagicMock())

    def test_add_comment(self):
        """测试添加评论的功能"""
        # 准备请求参数
        request_params = {
            "operation_type": "add_comment",
            "params": {
                "text": "这是一条测试评论",
                "author": "测试用户",
                "initial": "TU",
                "range": {
                    "start": 0,
                    "end": 5
                }
            }
        }

        # 调用工具函数
        result = run_async(comment_tools(self.context, request_params))

        # 使用基类的验证方法验证结果
        result_data = self.verify_tool_result(result, expected_success=True)

        # 检查返回的评论信息
        self.assertIn("comment_id", result_data)
        self.assertIn("added", result_data)
        self.assertTrue(result_data["added"])

        # 如果是模拟对象，验证Add方法是否被调用
        if isinstance(self.doc, MagicMock):
            self.doc.Comments.Add.assert_called()

    def test_delete_comment(self):
        """测试删除评论的功能"""
        # 准备请求参数
        request_params = {
            "operation_type": "delete_comment",
            "params": {
                "comment_id": 1
            }
        }

        # 调用工具函数
        result = run_async(comment_tools(self.context, request_params))

        # 使用基类的验证方法验证结果
        result_data = self.verify_tool_result(result, expected_success=True)

        # 检查返回的删除信息
        self.assertIn("deleted", result_data)
        self.assertTrue(result_data["deleted"])

    def test_update_comment(self):
        """测试更新评论的功能"""
        # 准备请求参数
        request_params = {
            "operation_type": "update_comment",
            "params": {
                "comment_id": 1,
                "text": "这是更新后的评论内容"
            }
        }

        # 调用工具函数
        result = run_async(comment_tools(self.context, request_params))

        # 使用基类的验证方法验证结果
        result_data = self.verify_tool_result(result, expected_success=True)

        # 检查返回的更新信息
        self.assertIn("updated", result_data)
        self.assertTrue(result_data["updated"])

    def test_get_comments(self):
        """测试获取评论列表的功能"""
        # 准备请求参数
        request_params = {
            "operation_type": "get_comments",
            "params": {
                "start_index": 1,
                "end_index": 10
            }
        }

        # 调用工具函数
        result = run_async(comment_tools(self.context, request_params))

        # 使用基类的验证方法验证结果
        result_data = self.verify_tool_result(result, expected_success=True)

        # 检查返回的数据结构
        self.assertIn("comments", result_data)
        self.assertIsInstance(result_data["comments"], list)

    def test_comment_tools_invalid_operation(self):
        """测试无效的操作类型"""
        # 准备请求参数，使用无效的操作类型
        request_params = {
            "operation_type": "invalid_operation",
            "params": {}
        }

        # 调用工具函数
        result = run_async(comment_tools(self.context, request_params))

        # 验证结果应该是失败的
        result_data = self.verify_tool_result(result, expected_success=False)

        # 检查错误信息
        self.assertIn("error", result_data)
        self.assertIn("无效的操作类型", result_data["error"])

    def test_comment_tools_missing_required_params(self):
        """测试缺少必要参数的情况"""
        # 准备请求参数，缺少必要的参数
        request_params = {
            "operation_type": "add_comment",
            # 故意不提供params
        }

        # 调用工具函数
        result = run_async(comment_tools(self.context, request_params))

        # 验证结果应该是失败的
        result_data = self.verify_tool_result(result, expected_success=False)

        # 检查错误信息
        self.assertIn("error", result_data)
        self.assertIn("缺少必要的参数", result_data["error"])

    def test_get_comment_by_id(self):
        """测试通过ID获取特定评论的功能"""
        # 准备请求参数
        request_params = {
            "operation_type": "get_comment",
            "params": {
                "comment_id": 1
            }
        }

        # 调用工具函数
        result = run_async(comment_tools(self.context, request_params))

        # 使用基类的验证方法验证结果
        result_data = self.verify_tool_result(result, expected_success=True)

        # 检查返回的数据结构
        self.assertIn("comment", result_data)
        self.assertIsInstance(result_data["comment"], dict)

    def test_delete_all_comments(self):
        """测试删除所有评论的功能"""
        # 准备请求参数
        request_params = {
            "operation_type": "delete_all_comments",
            "params": {}
        }

        # 调用工具函数
        result = run_async(comment_tools(self.context, request_params))

        # 使用基类的验证方法验证结果
        result_data = self.verify_tool_result(result, expected_success=True)

        # 检查返回的删除信息
        self.assertIn("deleted_count", result_data)
        self.assertEqual(result_data["deleted_count"], 0)  # 模拟文档初始没有评论


if __name__ == "__main__":
    import unittest
    unittest.main()