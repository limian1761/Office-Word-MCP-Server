"""
Tests for image_tools module in Word Document MCP Server.
"""

import json
import os
import shutil
import tempfile
import unittest
from io import StringIO
from pathlib import Path
from unittest.mock import MagicMock, patch

import pythoncom
import win32com.client
from mcp.server.fastmcp import Context
from mcp.server.session import ServerSession

from word_docx_tools.tools.image_tools import image_tools
from word_docx_tools.mcp_service.app_context import AppContext


class TestImageTools(unittest.TestCase):
    """Tests for image_tools module"""

    @classmethod
    def setUpClass(cls):
        # 初始化COM
        pythoncom.CoInitialize()

        # 创建临时目录用于测试
        cls.test_dir = tempfile.mkdtemp()

        # 复制测试文档到临时目录
        source_doc = os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            "tests",
            "test_docs",
            "valid_test_document_v2.docx",
        )
        cls.test_doc_path = os.path.join(cls.test_dir, "test_document.docx")
        if os.path.exists(source_doc):
            shutil.copy2(source_doc, cls.test_doc_path)
        else:
            # 如果测试文档不存在，创建一个新的
            word_app = win32com.client.Dispatch("Word.Application")
            doc = word_app.Documents.Add()
            doc.SaveAs2(cls.test_doc_path)
            doc.Close()
            word_app.Quit()

        # 创建测试图片
        cls.test_image_path = os.path.join(cls.test_dir, "test_image.png")
        # 创建一个简单的测试图片文件（实际内容不重要，只需要文件存在）
        with open(cls.test_image_path, "w") as f:
            f.write("fake image content for testing")

    @classmethod
    def tearDownClass(cls):
        # 清理临时目录
        try:
            shutil.rmtree(cls.test_dir)
        except:
            pass
        pythoncom.CoUninitialize()

    def setUp(self):
        """测试前准备"""
        # 创建Word应用程序实例
        self.word_app = win32com.client.Dispatch("Word.Application")
        # 尝试设置Visible属性，但捕获可能的异常
        try:
            self.word_app.Visible = False
        except AttributeError:
            # 某些环境中可能不支持设置Visible属性，忽略此错误
            pass

        # 创建应用上下文
        self.app_context = AppContext(self.word_app)

        # 创建模拟的读写流
        self.read_stream = StringIO()
        self.write_stream = StringIO()

        # 创建ServerSession实例
        self.session = ServerSession(
            read_stream=self.read_stream,
            write_stream=self.write_stream,
            init_options={},
        )
        self.session.lifespan_context = self.app_context

        # 创建Context对象
        self.context = Context(
            server_session=self.session, request_context=self.session
        )

        # 创建测试文档
        self.doc = self.word_app.Documents.Add()

    def tearDown(self):
        # 关闭文档
        try:
            if hasattr(self, "doc"):
                self.doc.Close(SaveChanges=False)
        except:
            pass

        # 关闭Word应用程序
        try:
            self.word_app.Quit()
        except:
            pass

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
    # 对于异步测试，需要使用异步测试运行器
    # 这里简化为直接运行测试
    import asyncio

    def run_async_test(test_case_method):
        """运行异步测试方法"""
        loop = asyncio.get_event_loop()
        return loop.run_until_complete(test_case_method)

    # 创建测试套件
    test_suite = unittest.TestSuite()

    # 添加所有测试方法
    for method_name in dir(TestImageTools):
        if method_name.startswith("test_"):
            test_case = TestImageTools(method_name)
            # 修改测试方法使其同步运行异步代码
            original_method = getattr(test_case, method_name)
            setattr(
                test_case,
                method_name,
                lambda self, method=original_method: run_async_test(method(self)),
            )
            test_suite.addTest(test_case)

    # 运行测试
    unittest.TextTestRunner(verbosity=2).run(test_suite)
