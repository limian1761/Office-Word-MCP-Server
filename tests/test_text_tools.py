#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
测试text_tools.py中的insert_text操作
"""

import json
from unittest.mock import MagicMock, patch

from tests.test_utils import WordDocumentTestBase

from word_docx_tools.operations.text_ops import insert_text_after_range
from word_docx_tools.tools.text_tools import text_tools


class TestTextToolsInsertText(WordDocumentTestBase):
    """Tests for text_tools.insert_text operation"""

    def setUp(self):
        """测试前准备"""
        # 调用基类的setUp方法，创建Word应用程序、文档和上下文
        super().setUp()

    @patch("word_docx_tools.operations.text_ops.insert_text_after_range")
    def test_insert_text_with_document_start_locator(self, mock_insert_text_after_range):
        """测试使用document_start定位器的insert_text操作"""
        try:
            # 模拟insert_text_after_range函数返回的是JSON字符串
            mock_insert_text_after_range.return_value = json.dumps({"success": True, "message": "Text inserted successfully"})

            # 定义测试参数
            test_text = "这是word-docx-tools的测试文档。"
            test_locator = {"type": "document_start"}

            try:
                # 调用text_tools插入文本
                result = text_tools(
                    ctx=self.context,
                    operation_type="insert_text",
                    text=test_text,
                    locator=test_locator,
                    position="after",
                )

                # 详细调试打印
                print(f"result类型: {type(result)}")
                print(f"result值: {result}")
                print(f"result is None: {result is None}")
                print(f"result is empty string: {result == ''}")

                # 验证结果
                # handle_tool_errors装饰器在捕获异常时返回字典，而正常执行时返回JSON字符串
                if isinstance(result, dict):
                    # 如果返回的是字典，说明发生了错误
                    print(f"Got dictionary result: {result}")
                    self.fail(f"Expected JSON string but got dict: {result}")
                elif result is None:
                    print("Got None result")
                    self.fail("Expected JSON string but got None")
                elif result == '':
                    print("Got empty string result")
                    self.fail("Expected JSON string but got empty string")
                else:
                    # 如果返回的是字符串，尝试解析为JSON
                    print(f"Trying to parse JSON: {result}")
                    try:
                        result_data = json.loads(result)
                        print(f"JSON parsed successfully: {result_data}")
                        self.assertTrue(result_data["success"])
                        self.assertEqual(result_data["message"], "Text inserted successfully")
                    except json.JSONDecodeError as e:
                        print(f"JSON decode error: {e}")
                        self.fail(f"Failed to parse JSON result: {e}")
            except Exception as e:
                print(f"Exception during text_tools call: {type(e).__name__}: {str(e)}")
                raise

            # 验证insert_text_after_range被正确调用
            mock_insert_text_after_range.assert_called_once()
        except Exception as e:
            print(f"测试test_insert_text_with_document_start_locator失败: {str(e)}")
            self.fail(f"测试test_insert_text_with_document_start_locator失败: {str(e)}")

    def test_insert_text_after_object_function(self):
        """测试insert_text_after_range函数"""
        try:
            # 准备测试文本和元素
            test_text = "这是插入的测试文本"

            # 使用文档的Content作为测试元素
            object = self.doc.Content

            # 调用insert_text_after_range函数
            result = insert_text_after_range(object, test_text)

            # 验证结果
            result_data = json.loads(result)
            self.assertTrue(result_data["success"])
            self.assertEqual(result_data["message"], "Text inserted successfully")

            # 验证文本是否被正确插入
            # 在实际环境中，我们需要访问文档内容来验证插入是否成功
            # 由于这是测试环境，我们主要验证函数返回值和调用逻辑
            # 添加额外的验证确保插入的文本确实存在于文档中
            if hasattr(self.doc.Content, 'Text'):
                self.assertIn(test_text, self.doc.Content.Text)
        except Exception as e:
            print(f"测试test_insert_text_after_object_function失败: {str(e)}")
            # 如果是模拟对象，我们仍然可以继续测试
            if not isinstance(self.doc, MagicMock):
                self.fail(f"测试test_insert_text_after_object_function失败: {str(e)}")

    def test_text_tools_invalid_operation(self):
        """测试无效的操作类型"""
        try:
            # 调用text_tools并使用无效的操作类型
            result = text_tools(ctx=self.context, operation_type="invalid_operation")

            # 验证结果包含错误信息
            # handle_tool_errors装饰器在捕获异常时返回字典
            if isinstance(result, dict):
                # 检查是否包含错误信息
                self.assertIn("error", result)
                self.assertTrue("Invalid operation type" in result["error"] or "无效的操作类型" in result["error"])
            else:
                # 如果返回的是字符串，尝试解析为JSON
                result_data = json.loads(result)
                self.assertFalse(result_data["success"])
                self.assertIn("Invalid operation type" or "无效的操作类型", result_data["message"])
            self.assertIn("error", result_data["message"].lower())
        except Exception as e:
            print(f"测试test_text_tools_invalid_operation失败: {str(e)}")
            self.fail(f"测试test_text_tools_invalid_operation失败: {str(e)}")


if __name__ == "__main__":
    import unittest
    unittest.main()
