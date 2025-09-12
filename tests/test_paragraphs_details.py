"""Tests for the get_paragraphs_details function."""

import json
import unittest
from unittest.mock import MagicMock, patch

from word_docx_tools.tools.paragraph_tools import paragraph_tools
from word_docx_tools.mcp_service.errors import ErrorCode

class TestParagraphsDetails(unittest.TestCase):
    """Tests for the get_paragraphs_details function."""

    def setUp(self):
        """Set up test environment."""
        # 创建一个模拟的上下文对象
        self.mock_ctx = MagicMock()
        self.mock_doc = MagicMock()
        self.mock_ctx.request_context.lifespan_context.get_active_document.return_value = self.mock_doc

    @patch('word_docx_tools.tools.paragraph_tools.get_paragraphs_details')
    def test_get_paragraphs_details_without_locator(self, mock_get_paragraphs_details):
        """Test get_paragraphs_details without locator."""
        # 设置模拟返回值
        mock_result = {
            "paragraphs": [
                {"text": "Paragraph 1", "style_name": "Normal"},
                {"text": "Paragraph 2", "style_name": "Heading 1"}
            ]
        }
        mock_get_paragraphs_details.return_value = mock_result

        # 调用函数
        result = paragraph_tools(
            ctx=self.mock_ctx,
            operation_type="get_paragraphs_details",
            locator=None,
            include_stats=False
        )

        # 验证结果
        parsed_result = json.loads(result)
        self.assertTrue(parsed_result["success"])
        self.assertEqual(parsed_result["result"], mock_result)
        mock_get_paragraphs_details.assert_called_once_with(self.mock_doc, None, False)

    @patch('word_docx_tools.tools.paragraph_tools.get_paragraphs_details')
    def test_get_paragraphs_details_with_locator(self, mock_get_paragraphs_details):
        """Test get_paragraphs_details with locator."""
        # 设置模拟返回值
        mock_result = {
            "paragraphs": [
                {"text": "Selected Paragraph", "style_name": "Quote"}
            ]
        }
        mock_get_paragraphs_details.return_value = mock_result

        # 调用函数
        locator = {"type": "paragraph", "index": 1}
        result = paragraph_tools(
            ctx=self.mock_ctx,
            operation_type="get_paragraphs_details",
            locator=locator,
            include_stats=False
        )

        # 验证结果
        parsed_result = json.loads(result)
        self.assertTrue(parsed_result["success"])
        self.assertEqual(parsed_result["result"], mock_result)
        mock_get_paragraphs_details.assert_called_once_with(self.mock_doc, locator, False)

    @patch('word_docx_tools.tools.paragraph_tools.get_paragraphs_details')
    def test_get_paragraphs_details_with_stats(self, mock_get_paragraphs_details):
        """Test get_paragraphs_details with statistics."""
        # 设置模拟返回值
        mock_result = {
            "paragraphs": [
                {"text": "Paragraph 1", "style_name": "Normal"},
                {"text": "Paragraph 2", "style_name": "Normal"},
                {"text": "Paragraph 3", "style_name": "Heading 1"}
            ],
            "stats": {
                "total_paragraphs": 3,
                "styles_used": {"Normal": 2, "Heading 1": 1}
            }
        }
        mock_get_paragraphs_details.return_value = mock_result

        # 调用函数
        result = paragraph_tools(
            ctx=self.mock_ctx,
            operation_type="get_paragraphs_details",
            locator=None,
            include_stats=True
        )

        # 验证结果
        parsed_result = json.loads(result)
        self.assertTrue(parsed_result["success"])
        self.assertEqual(parsed_result["result"], mock_result)
        mock_get_paragraphs_details.assert_called_once_with(self.mock_doc, None, True)

    @patch('word_docx_tools.tools.paragraph_tools.get_paragraphs_details')
    def test_get_paragraphs_details_with_locator_validation(self, mock_get_paragraphs_details):
        """Test get_paragraphs_details with locator validation."""
        # 设置模拟返回值
        mock_result = {"paragraphs": []}
        mock_get_paragraphs_details.return_value = mock_result

        # 调用函数
        locator = {"type": "paragraph", "index": 0}
        result = paragraph_tools(
            ctx=self.mock_ctx,
            operation_type="get_paragraphs_details",
            locator=locator,
            include_stats=False
        )

        # 验证结果
        parsed_result = json.loads(result)
        self.assertTrue(parsed_result["success"])
        mock_get_paragraphs_details.assert_called_once_with(self.mock_doc, locator, False)

    def test_get_paragraphs_details_invalid_operation_type(self):
        """Test get_paragraphs_details with invalid operation type."""
        # 调用函数，使用无效的操作类型
        result = paragraph_tools(
            ctx=self.mock_ctx,
            operation_type="invalid_operation",
        )

        # 验证结果是一个错误消息
        self.assertIsInstance(result, str)
        self.assertIn("Error [1005]", result)

if __name__ == "__main__":
    unittest.main()