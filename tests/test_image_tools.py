"""
Tests for image_tools module in Word Document MCP Server.
"""

import json
from unittest.mock import patch

from word_docx_tools.tools.image_tools import image_tools
from tests.test_utils import WordDocumentTestBase


class TestImageTools(WordDocumentTestBase):
    """Tests for image_tools module"""

    def setUp(self):
        """测试前准备，调用基类方法并进行必要的初始化"""
        # 调用基类的setUp方法，创建Word应用程序、文档和上下文
        super().setUp()

        # 创建测试图片文件
        self.test_image_path = self.create_test_file("test_image.png", "fake image content for testing")

    async def test_image_tools_get_info(self):
        """Test get_info operation"""
        # 首先确保文档中没有图片
        self.assertEqual(self.doc.InlineShapes.Count, 0)

        # 调用image_tools获取图片信息
        result = await image_tools(self.context, operation_type="get_info")

        # 验证结果
        result_data = json.loads(result)
        self.assertTrue(result_data["success"])
        self.assertEqual(len(result_data["images"]), 0)

    @patch("word_docx_tools.tools.image_tools.insert_image")
    async def test_image_tools_insert(self, mock_insert_image):
        """Test insert operation"""
        # 设置mock返回值
        mock_result = json.dumps(
            {
                "success": True,
                "message": "Image inserted successfully",
                "image_index": 1,
            }
        )
        mock_insert_image.return_value = mock_result

        # 调用image_tools插入图片
        result = await image_tools(
            self.context, operation_type="insert", image_path=self.test_image_path
        )

        # 验证结果
        result_data = json.loads(result)
        self.assertTrue(result_data["success"])
        self.assertEqual(result_data["message"], "Image inserted successfully")

        # 验证insert_image被正确调用
        mock_insert_image.assert_called_once()

    @patch("word_docx_tools.tools.image_tools.add_caption")
    async def test_image_tools_add_caption(self, mock_add_caption):
        """Test add_caption operation"""
        # 设置mock返回值
        mock_result = json.dumps(
            {
                "success": True,
                "message": "Caption added successfully",
                "caption_text": "Test Caption",
            }
        )
        mock_add_caption.return_value = mock_result

        # 调用image_tools添加题注
        result = await image_tools(
            self.context, operation_type="add_caption", caption_text="Test Caption"
        )

        # 验证结果
        result_data = json.loads(result)
        self.assertTrue(result_data["success"])
        self.assertEqual(result_data["message"], "Caption added successfully")

        # 验证add_caption被正确调用
        mock_add_caption.assert_called_once()

    @patch("word_docx_tools.tools.image_tools.resize_image")
    async def test_image_tools_resize(self, mock_resize_image):
        """Test resize operation"""
        # 设置mock返回值
        mock_result = json.dumps(
            {
                "success": True,
                "message": "Image resized successfully",
                "image_index": 1,
                "new_width": 200,
                "new_height": 150,
                "maintain_aspect_ratio": True,
            }
        )
        mock_resize_image.return_value = mock_result

        # 调用image_tools调整图片大小
        result = await image_tools(
            self.context, operation_type="resize", width=200, height=150
        )

        # 验证结果
        result_data = json.loads(result)
        self.assertTrue(result_data["success"])
        self.assertEqual(result_data["message"], "Image resized successfully")

        # 验证resize_image被正确调用
        mock_resize_image.assert_called_once()

    @patch("word_docx_tools.tools.image_tools.set_image_color_type")
    async def test_image_tools_set_color_type(self, mock_set_color_type):
        """Test set_color_type operation"""
        # 设置mock返回值
        mock_result = json.dumps(
            {
                "success": True,
                "message": "Successfully set image color type to grayscale",
                "image_index": 1,
                "color_type": "grayscale",
            }
        )
        mock_set_color_type.return_value = mock_result

        # 调用image_tools设置图片颜色类型
        result = await image_tools(
            self.context, operation_type="set_color_type", color_type="grayscale"
        )

        # 验证结果
        result_data = json.loads(result)
        self.assertTrue(result_data["success"])
        self.assertEqual(result_data["message"], "Image color type set successfully")

        # 验证set_image_color_type被正确调用
        mock_set_color_type.assert_called_once()

    async def test_image_tools_invalid_operation(self):
        """Test invalid operation type"""
        # 调用image_tools并使用无效的操作类型
        with self.assertRaises(Exception):
            await image_tools(self.context, operation_type="invalid_operation")

    async def test_image_tools_insert_missing_path(self):
        """Test insert operation with missing image path"""
        # 调用image_tools插入图片，但不提供图片路径
        with self.assertRaises(Exception):
            await image_tools(self.context, operation_type="insert")


if __name__ == "__main__":
    import unittest
    unittest.main()
