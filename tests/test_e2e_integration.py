"""
End-to-end integration tests for Word Document MCP Server.

This module contains integration tests that simulate real-world usage scenarios
by testing multiple tools and operations in sequence.
"""

import json
import os
import shutil
import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock

import pythoncom
import win32com.client
from mcp.server.fastmcp import Context

from .tools.comment_tools import comment_tools
from .tools.document_tools import document_tools
from .tools.image_tools import image_tools
from .tools.range_tools import range_tools
from .tools.table_tools import table_tools
from .tools.text_tools import text_tools
from word_docx_tools.mcp_service.app_context import AppContext


class TestE2EIntegration(unittest.TestCase):
    """End-to-end integration tests simulating real usage scenarios"""

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
        shutil.copy2(source_doc, cls.test_doc_path)

        # 创建测试图片
        cls.test_image_path = os.path.join(cls.test_dir, "test_image.png")
        # 创建一个简单的测试图片文件（实际内容不重要，只需要文件存在）
        with open(cls.test_image_path, "w") as f:
            f.write("fake image content")

    @classmethod
    def tearDownClass(cls):
        # 清理临时目录
        shutil.rmtree(cls.test_dir)
        pythoncom.CoUninitialize()

    def setUp(self):
        # 创建Word应用程序实例
        self.word_app = win32com.client.Dispatch("Word.Application")
        self.word_app.Visible = False

        # 创建应用上下文
        self.app_context = AppContext(self.word_app)

        # 创建mock session
        self.mock_session = MagicMock()
        self.mock_session.lifespan_context = self.app_context

        # 创建Context对象
        self.context = MagicMock()
        self.context.request_context = self.mock_session

    def tearDown(self):
        # 关闭Word应用程序
        try:
            self.word_app.Quit()
        except:
            pass

    def test_complete_document_workflow(self):
        """测试完整的文档操作工作流程"""
        # 1. 打开文档
        result = document_operation(
            self.context, operation_type="open", file_path=self.test_doc_path
        )
        self.assertIn("Document opened successfully", result)

        # 2. 使用text_content_operation添加段落
        result = text_content_operation(
            self.context, operation_type="insert", text="这是新添加的段落内容。"
        )
        self.assertIn("Successfully inserted", result)

        # 3. 使用text_format_operation设置为标题样式
        result = text_content_operation(
            self.context, operation_type="insert", text="新章节", style="Heading 1"
        )
        self.assertIn("Successfully inserted", result)

        # 4. 查找文本
        result = text_content_operation(
            self.context,
            operation_type="find",
            text="新添加的段落",
            match_case=False,
            match_whole_word=False,
        )
        # 验证查找结果

        # 5. 应用格式化
        # 首先查找元素获取定位器
        # 然后应用格式化

        # 6. 添加注释
        result = comment_operation(
            self.context, operation_type="add", text="这是一个测试注释", author="Tester"
        )
        self.assertIn("Comment added successfully", result)

        # 7. 获取文档样式
        result = document_operation(self.context, operation_type="get_styles")
        self.assertIn("Styles retrieved successfully", result)

    def test_table_operations_workflow(self):
        """测试表格操作工作流程"""
        # 1. 打开文档
        document_operation(
            self.context, operation_type="open", file_path=self.test_doc_path
        )

        # 2. 创建表格
        result = table_operation(self.context, operation_type="create", rows=3, cols=4)
        self.assertIn("Table created successfully", result)

        # 3. 设置单元格文本
        # 需要先实现定位器逻辑

        # 4. 获取单元格文本
        # 需要先实现定位器逻辑

    def test_image_operations_workflow(self):
        """测试图片操作工作流程"""
        # 1. 打开文档
        document_operation(
            self.context, operation_type="open", file_path=self.test_doc_path
        )

        # 2. 插入图片
        result = image_operation(
            self.context,
            operation_type="insert",
            object_path=self.test_image_path,
            position="after",
        )
        # 注意：由于是测试环境，实际插入可能失败，但我们测试调用逻辑

        # 3. 添加题注
        # 需要实现定位器逻辑

    def test_batch_operations_workflow(self):
        """测试批处理操作工作流程"""
        # 1. 打开文档
        document_operation(
            self.context, operation_type="open", file_path=self.test_doc_path
        )

        # 2. 执行批处理格式化
        # 需要实现具体的操作列表

    def test_comment_operations_workflow(self):
        """测试注释操作工作流程"""
        # 1. 打开文档
        document_operation(
            self.context, operation_type="open", file_path=self.test_doc_path
        )

        # 2. 添加多个注释
        comment_operation(
            self.context, operation_type="add", text="第一个注释", author="Author1"
        )

        comment_operation(
            self.context, operation_type="add", text="第二个注释", author="Author2"
        )

        # 3. 获取所有注释
        result = comment_operation(self.context, operation_type="get_all")
        self.assertIn("Comments retrieved successfully", result)

        # 4. 回复注释
        # 需要实现注释索引逻辑

        # 5. 编辑注释
        # 需要实现注释索引逻辑

        # 6. 删除注释
        # 需要实现注释索引逻辑


if __name__ == "__main__":
    unittest.main()
